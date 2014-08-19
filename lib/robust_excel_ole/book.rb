# -*- coding: utf-8 -*-
require 'weakref'

class Hash
  def first
    to_a.first
  end
end


module RobustExcelOle

  class Book
    attr_reader :workbook
    def book
      @workbook_ole
    end


    class << self

      # opens a book. 
      # options: 
      #  :reuse      (boolean)  use an already open application
      #  :read_only     (boolean)  open in read-only mode
      #  :displayalerts (boolean)  allow display alerts in excel
      #  :visible       (boolean)  make visibe in Excel
      # :if_not_saved   if a book b with this name is already open:
      #                 :read_only -> let b open
      #                 :raise -> raise an exception,             if b is not saved
      #                 :accept -> let b open,                    if b is not saved
      #                 :forget -> open the new book and close b, if b is not saved
      # if the file name is nil then return
 
      def open(file, options={ :reuse => true}, &block)
        new(file, options, &block)
      end

    end

    def initialize(file, options={ }, &block)
      #unless caller[1] =~ /book.rb:\d+:in\s+`open'$/
      #  warn "DEPRECATION WARNING: ::Book.new RobustExcelOle and RobustExcelOle::Book.open will be split. If you open existing file, please use RobustExcelOle::Book.open.(call from #{caller[1]})"
      #end
      @options = {
        :reuse => true,
        :if_not_saved => :raise,
        :read_only => true
        #:displayalerts => false,
        #:visible => false,
      }.merge(options)

      if not File.exist?(file)
        raise ExcelErrorOpen, "file #{file} not found"
      end      
      @excel_app = ExcelApp.new
      workbooks = @excel_app.Workbooks
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
            @excel_app.Workbooks.Close(absolute_path(file))           
          else
            raise ExcelErrorOpen, "invalid option"
          end
        end
      end
      # book not open (was not open or was closed with option :forget)
      if not @workbook_ole then
        p "open a book"                  
        @workbook_ole = @excel_app.Workbooks.Open(absolute_path(file),{ 'ReadOnly' => @options[:read_only] })
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
    
    def close
      @workbook_ole.close if alive?  
      #@excel_app.Workbooks.Close
      #@excel_app.Quit
    end

    def alive?
      @workbook_ole.Name
      true
    rescue 
      puts $!.message
      false
    end

    attr_reader :excel_app


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
          displayalerts_value = @excel_app.DisplayAlerts
          @excel_app.DisplayAlerts = true 
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
        @excel_app.DisplayAlerts = displayalerts_value
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

      new_sheet = RobustExcelOle::Sheet.new(@excel_app.Activesheet)
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
