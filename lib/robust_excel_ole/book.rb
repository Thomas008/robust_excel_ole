# -*- coding: utf-8 -*-

require 'weakref'

module RobustExcelOle

  class Book
    attr_reader :workbook


    class << self
      
      # opens a book.
      # 
      # options: 
      #  :read_only     open in read-only mode          (default: false)
      #  :if_unsaved    if an unsaved book with the same name is open, then
      #                 :raise   -> raise an exception (default)
      #                 :forget  -> close the unsaved book, open the new book             
      #                 :accept  -> let the unsaved book open                  
      #                 :alert   -> give control to excel
      #                 :new_app -> open the new book in a new excel application
      #  :if_obstructed   if a book with the same name in a different path is open, then
      #                  :raise          -> raise an exception (default)             
      #                  :forget         -> close the old book, open the new book
      #                  :save           -> save the old book, close it, open the new book
      #                  :close_if_saved -> close the old book and open the new book, if the old book is saved
      #                                       raise an exception otherwise
      #                  :new_app        -> open the new book in a new excel application
      #  :reuse         use a running Excel-application (default: true)
      #  :excel         an Excel application            (default: nil) 
      #  :displayalerts allow display alerts in Excel   (default: false)
      #  :visible       make visibe in Excel            (default: false)
      def open(file, options={ :reuse => true}, &block)
        new(file, options, &block)
      end

    end

    def initialize(file, opts={ }, &block)
      @options = {
        :reuse => true,
        :excel => nil,
        :read_only => false,
        :if_unsaved => :raise,
        :if_obstructed => :raise
      }.merge(opts)
      excel_options = {:reuse => true}.merge(opts).delete_if{|k,v| 
        k== :read_only || k== :if_unsaved || k == :if_obstructed}
      if not File.exist?(file)
        raise ExcelErrorOpen, "file #{file} not found"
      end
      @excel = @options[:excel] ? excel_options[:excel] : Excel.new(excel_options)  
      workbooks = @excel.Workbooks
      @workbook = workbooks.Item(File.basename(file)) rescue nil
      @file = file

      def open_workbook
        # if book not open (was not open,was closed with option :forget or shall be opened in new application)
        #    or :if_unsaved => :alert
        if ((not alive?) || (@options[:if_unsaved] == :alert)) then
          begin
            @workbook = @excel.Workbooks.Open(RobustExcelOle::absolute_path(@file),{ 'ReadOnly' => @options[:read_only] })
          rescue WIN32OLERuntimeError
            raise ExcelUserCanceled, "open: canceled by user"
          end
        end
      end

      if @workbook then
        obstructed_by_other_book = (File.basename(file) == File.basename(@workbook.Fullname)) && 
                                   (not (RobustExcelOle::absolute_path(file) == @workbook.Fullname))
        if obstructed_by_other_book then
          # @workbook is not the desired workbook
          case @options[:if_obstructed]
          when :raise
            raise ExcelErrorOpen, "blocked by a book with the same name in a different path"
          when :forget
            @workbook.Close
          when :save
            save unless @workbook.Saved
            @workbook.Close
          when :close_if_saved
            if (not @workbook.Saved) then
              raise ExcelErrorOpen, "book with the same name in a different path is unsaved"
            else 
              @workbook.Close
            end
          when :new_app
            excel_options[:reuse] = false
            @excel = Excel.new(excel_options)
            @workbook = nil
          else
            raise ExcelErrorOpen, ":if_obstructed: invalid option"
          end
        else
          # book open, not saved, not obstructed by an other book
          if (not @workbook.Saved) then
            #p "book not saved"
            case @options[:if_unsaved]
            when :raise
              raise ExcelErrorOpen, "book is already open but not saved (#{File.basename(file)})"
            when :forget
              @workbook.Close
            when :accept
            when :alert
              @excel.with_displayalerts true do
                open_workbook
              end 
            when :new_app
              excel_options[:reuse] = false
              @excel = Excel.new(excel_options)
              @workbook = nil
            else
              raise ExcelErrorOpen, ":if_unsaved: invalid option"
            end
          end
        end
      end
      open_workbook
      if block
        begin
          yield self
        ensure
          close
        end
      end
      @workbook
    end
    
    # closes the book, if it is alive
    #
    # options:
    #  :if_unsaved    if book is unsaved
    #                      :raise   -> raise an exception       (default)             
    #                      :save    -> save the book before it is closed                  
    #                      :forget  -> close the book 
    #                      :alert   -> give control to excel
    def close(opts = {:if_unsaved => :raise})

      def close_workbook 
        @workbook.Close if alive?
        @workbook = nil unless alive?
      end

      if ((alive?) && (not @workbook.Saved) && (not @options[:read_only])) then
        case opts[:if_unsaved]
        when :raise
          raise ExcelErrorClose, "book is unsaved (#{File.basename(filename)})"
        when :save
          save
          close_workbook
        when :forget
          close_workbook
        when :alert
          @excel.with_displayalerts true do
            close_workbook
          end
        else
          raise ExcelErrorClose, ":if_unsaved: invalid option"
        end
      else
        close_workbook
      end
      raise ExcelUserCanceled, "close: canceled by user" if alive? && opts[:if_unsaved] == :alert && (not @workbook.Saved)

    end

    # returns true, if the workbook reacts to methods, false otherwise
    def alive?
      begin 
        @workbook.Name
        true
      rescue 
        @workbook = nil  # dead object won't be alive again
        #puts $!.message
        false
      end
    end

    # returns the full file name of the workbook
    def filename
      @workbook.Fullname.tr('\\','/') rescue nil
    end

    # returns true, if the full book names and excel appications are identical, false, otherwise  
    def == other_book
      other_book.is_a?(Book) &&
      @excel == other_book.excel &&
      self.filename == other_book.filename  
    end

    # make the current Excel application visible or invisible
    # option: visible_value     true -> make Excel visible, false -> make Excel invisible
    def visible visible_value
      Excel.current.Visible = visible_value
    end

    attr_reader :excel

    # saves a book.
    # returns true, if successfully saved, nil otherwise
    def save
      raise ExcelErrorSave, "Not opened for writing (opened with :read_only option)" if @options[:read_only]
      if @workbook then
        @workbook.Save 
        true
      else
        nil
      end
    end

    # saves a book.
    #
    # options:
    #  :if_exists   if a file with the same name exists, then  
    #               :raise     -> raise an exception, dont't write the file  (default)
    #               :overwrite -> write the file, delete the old file
    #               :alert     -> give control to Excel
    # returns true, if successfully saved, nil otherwise
    def save_as(file = nil, opts = {:if_exists => :raise} )
      
      raise IOError, "Not opened for writing(open with :read_only option)" if @options[:read_only]
      @dirname, @basename = File.split(file)
      @file_format =
        case File.extname(@basename)
        when '.xls' : RobustExcelOle::XlExcel8
        when '.xlsx': RobustExcelOle::XlOpenXMLWorkbook
        when '.xlsm': RobustExcelOle::XlOpenXMLWorkbookMacroEnabled
      end
      @opts = opts

      def save_as_workbook 
        begin
          @workbook.SaveAs(RobustExcelOle::absolute_path(File.join(@dirname, @basename)), @file_format)
        rescue WIN32OLERuntimeError => msg
          if msg.message =~ /SaveAs/ and msg.message =~ /Workbook/ then
            if @opts[:if_exists] == :alert then 
              raise ExcelErrorSave, "not saved or canceled by user"
            else
              return nil
            end
            # another possible semantics. raise ExcelErrorSaveFailed, "could not save Workbook"
          else
            raise ExcelErrorSaveUnknown, "unknown WIN32OELERuntimeError:\n#{msg.message}"
          end       
        end
      end

      if File.exist?(file) then
        case @opts[:if_exists]
        when :overwrite
          begin
            File.delete(file) 
          rescue Errno::EACCES
            raise ExcelErrorSave, "book is open and used in Excel"
          end
          save_as_workbook
        when :alert 
          @excel.with_displayalerts true do
            save_as_workbook
          end
        when :raise
          raise ExcelErrorSave, "book already exists: #{@basename}"
        else
          raise ExcelErrorSave, ":if_exists: invalid option"
        end
      else
        save_as_workbook
      end
      true
    end

    
    def [] sheet
      sheet += 1 if sheet.is_a? Numeric
      RobustExcelOle::Sheet.new(@workbook.Worksheets.Item(sheet))
    end

    def each
      @workbook.Worksheets.each do |sheet|
        yield RobustExcelOle::Sheet.new(sheet)
      end
    end

    def add_sheet(sheet = nil, opts = { })
      if sheet.is_a? Hash
        opts = sheet
        sheet = nil
      end

      new_sheet_name = opts.delete(:as)

      after_or_before, base_sheet = opts.to_a.first || [:after, RobustExcelOle::Sheet.new(@workbook.Worksheets.Item(@workbook.Worksheets.Count))]
      base_sheet = base_sheet.sheet
      sheet ? sheet.Copy({ after_or_before.to_s => base_sheet }) : @workbook.WorkSheets.Add({ after_or_before.to_s => base_sheet })
      new_sheet = RobustExcelOle::Sheet.new(@excel.Activesheet)
      begin
        new_sheet.name = new_sheet_name if new_sheet_name
      rescue WIN32OLERuntimeError => msg
        if msg.message =~ /OLE error code:800A03EC/ 
          raise ExcelErrorSheet, "sheet name already exists"
        end
      end
      new_sheet
    end        

  end

end

class ExcelUserCanceled < RuntimeError # :nodoc: #
end

class ExcelError < RuntimeError    # :nodoc: #
end

class ExcelErrorSave < ExcelError   # :nodoc: #
end

class ExcelErrorSaveFailed < ExcelErrorSave  # :nodoc: #
end

class ExcelErrorSaveUnknown < ExcelErrorSave  # :nodoc: #
end

class ExcelErrorOpen < ExcelError   # :nodoc: #
end

class ExcelErrorClose < ExcelError    # :nodoc: #
end

class ExcelErrorSheet < ExcelError    # :nodoc: #
end

