# -*- coding: utf-8 -*-
require 'weakref'

class Hash
  def first
    to_a.first
  end
end


module RobustExcelOle

  class Book
    attr_reader :book

    class << self

      # opens a book. 
      # options: 
      #  :recycled      (boolean)  use an already open application
      #  :read_only     (boolean)  open in read-only mode
      #  :displayalerts (boolean)  allow display alerts in excel
      #  :visible       (boolean)  make visibe in Excel
      # :if_not_saved   if a book b with this name is already open:
      #                 :read_only -> let b open
      #                 :raise -> raise an exception,             if b is not saved
      #                 :accept -> let b open,                    if b is not saved
      #                 :forget -> open the new book and close b, if b is not saved
      # if the file name is nil then return
 
      def open(file, options={ :recycled => true}, &block)
        new(file, options, &block)
      end

      def close_all_excel_apps
        while running_excel_app do
          close_one_excel_app
          GC.start
        end
      end

      def close_one_excel_app
        excel = running_excel_app
        if excel then 
          excel.Workbooks.Close
          excel_hwnd = excel.HWnd
          excel.Quit
          weak_excel_ref = WeakRef.new(excel)
          excel = nil
          GC.start
          if weak_excel_ref.weakref_alive? then
            #if WIN32OLE.ole_reference_count(weak_xlapp) > 0
            begin
              weak_xlapp.ole_free
            rescue
              puts "could not do ole_free on #{weak_excel_ref}"
            end
          end
        end
        process_id = Win32API.new("user32", "GetWindowThreadProcessId", ["I","P"], "I")
        pid_puffer = " " * 32
        process_id.call(excel_hwnd, pid_puffer)
        pid = pid_puffer.unpack("L")[0]
        Process.kill("KILL", pid)
        anz_objekte = 0
        ObjectSpace.each_object(WIN32OLE) do |o|
          anz_objekte += 1
          #p [:ole_object_name, o, (o.Name rescue nil)]
          #trc_info :ole_type, o.ole_obj_help rescue nil
          #trc_info :obj_hwnd, o.HWnd rescue   nil
          #trc_info :obj_Parent, o.Parent rescue nil
          begin
            o.ole_free
          rescue
            puts "olefree_error: #{$!}"
          end
        end
      end

=begin

        
=end


      # returns nil, if no excel is running or connected to a dead Excel app
      def running_excel_app
        result = WIN32OLE.connect('Excel.Application') rescue nil 
        if result 
          begin
            result.Visible    # send any method, just to see if it responds
          rescue 
            puts "Window-handle = #{result.HWnd}"
            # dead!!!
            return nil
          end
        end
        result
      end

    end


    def initialize(file, options={ }, &block)
      #unless caller[1] =~ /book.rb:\d+:in\s+`open'$/
      #  warn "DEPRECATION WARNING: ::Book.new RobustExcelOle and RobustExcelOle::Book.open will be split. If you open existing file, please use RobustExcelOle::Book.open.(call from #{caller[1]})"
      #end
      @options = {
        :recycle => true,
        :if_not_saved => :raise,
        :read_only => true
        #:displayalerts => false,
        #:visible => false,
      }.merge(options)

      if not File.exist?(file)
        raise ExcelErrorOpen, "file #{file} not found"
      end
             
      supply_app(options)
      workbooks = @winapp.Workbooks
      @workbook_ole = workbooks.Item(File.basename(file)) rescue nil
      if @workbook_ole then
        # book open and not saved
        p "book already open"
        if (not @workbook_ole.Saved) then
          p "book not saved"
          case @options[:if_not_saved]
          when :raise
            raise ExcelErrorOpen, "book is already open but not saved (#{File.basename(file)})"
          when :accept
            #nothing
          when :forget
            @winapp.Workbooks.Close(absolute_path(file))           
          else
            raise ExcelErrorOpen, "invalid option"
          end
        end
      end
      # book not open (was not open or was closed with option :forget)
      if not @workbook_ole then
        p "open a book"                  
        @workbook_ole = @winapp.Workbooks.Open(absolute_path(file),{ 'ReadOnly' => @options[:read_only] })
      end
      if block
        begin
          yield self
        ensure
          close
        end
      end
      @workbook_ole
    end

    def supply_app(options={ })
      p "supply_app"
      if @options[:recycle] then
        p "recycle"
        @winapp = self.class.running_excel_app
        # hier noch abfragen mit Visible, ob die Application noch reagiert 
        # (siehe extzug_basis)
        p "@winapp:#{@winapp}"
        if @winapp
          p "@winapp existiert"
          @winapp.DisplayAlerts = @options[:displayalerts] unless @options[:displayalerts]==nil
          @winapp.Visible = @options[:visible] unless @options[:visible]==nil
          WIN32OLE.const_load(@winapp, RobustExcelOle) unless RobustExcelOle.const_defined?(:CONSTANTS)
          return
        end
      end
      p "kreiere neue application"
      @options = {
        :displayalerts => false,
        :visible => false,
      }.merge(options)
      @winapp = WIN32OLE.new('Excel.application')
      @winapp.DisplayAlerts = @options[:displayalerts]
      @winapp.Visible = @options[:visible]
      WIN32OLE.const_load(@winapp, RobustExcelOle) unless RobustExcelOle.const_defined?(:CONSTANTS)
    end

    def close
      @workbook_ole.close if alive?  
      #@winapp.Workbooks.Close
      #@winapp.Quit
    end

    def alive?
      @workbook_ole.Name
      true
    rescue 
      puts $!.message
      false
    end


    # saves a book
    # if a file with the same name, exists, then proceed according to :if_exists 
    #   :raise     -> raise an exception, dont't write the file
    #   :overwrite -> write the file, delete the old file
    #   :excel     -> give control to Excel 
    # if file is nil, then return
    def save(file = nil, opts = {:if_exists => :raise} )
      raise IOError, "Not opened for writing(open with :read_only option)" if @options[:read_only]
      return @workbook_ole.save unless file
      dirname, basename = File.split(file)
      file_format =
        case File.extname(basename)
        when '.xls' : RobustExcelOle::XlExcel8
        when '.xlsx': RobustExcelOle::XlOpenXMLWorkbook
        when '.xlsm': RobustExcelOle::XlOpenXMLWorkbookMacroEnabled
        end
      if File.exist?(file) then
        case opts[:if_exists]
        when :overwrite
          File.delete(file) 
          #File.delete(absolute_path(File.join(dirname, basename)))
        when :excel 
          displayalerts_value = @winapp.DisplayAlerts
          @winapp.DisplayAlerts = true 
        when :raise
          raise ExcelErrorSave, "book already exists: #{basename}"
        else
          raise ExcelErrorSave, "invalid option (#{opts[:if_exists]})"
        end
      end
      @workbook_ole.SaveAs(absolute_path(File.join(dirname, basename)), file_format)
        rescue WIN32OLERuntimeError => msg
          if not msg.message.include? "Die SaveAs-Eigenschaft des Workbook-Objektes kann nicht zugeordnet werden." then
            raise ExcelErrorSave, "unknown WIN32OELERuntimeError"
          end             
      if opts[:if_exists] == :excel then
        @winapp.DisplayAlerts = displayalerts_value
      end
    end

    def [] sheet
      sheet += 1 if sheet.is_a? Numeric
      RobustExcelOle::Sheet.new(@workbook_ole.Worksheets.Item(sheet))
    end

    def each
      @workbook_ole.Worksheets.each do |sheet|
        yield RobustExcelOle::Sheet.new(sheet)
      end
    end

    def add_sheet(sheet = nil, opts = { })
      if sheet.is_a? Hash
        opts = sheet
        sheet = nil
      end

      new_sheet_name = opts.delete(:as)

      after_or_before, base_sheet = opts.first || [:after, RobustExcelOle::Sheet.new(@workbook_ole.Worksheets.Item(@workbook_ole.Worksheets.Count))]
      base_sheet = base_sheet.sheet
      sheet ? sheet.Copy({ after_or_before.to_s => base_sheet }) : @workbook_ole.WorkSheets.Add({ after_or_before.to_s => base_sheet })

      new_sheet = RobustExcelOle::Sheet.new(@winapp.Activesheet)
      new_sheet.name = new_sheet_name if new_sheet_name
      new_sheet
    end        

    private
    def absolute_path(file)
      file = File.expand_path(file)
      file = RobustExcelOle::Cygwin.cygpath('-w', file) if RUBY_PLATFORM =~ /cygwin/
      WIN32OLE.new('Scripting.FileSystemObject').GetAbsolutePathName(file)
    end
  end

end

class ExcelErrorSave < RuntimeError
end

class ExcelErrorOpen < RuntimeError
end
