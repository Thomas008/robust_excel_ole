# -*- coding: utf-8 -*-
require 'weakref'

module RobustExcelOle

  class Book
    attr_reader :workbook


    class << self

      # opens a book.
      # 
      # options: 
      #  :reuse         (boolean)  use an already open Excel-application (default: true)
      #  :read_only     (boolean)  open in read-only mode                (default: false)
      #  :displayalerts (boolean)  allow display alerts in Excel         (default: false)
      #  :visible       (boolean)  make visibe in Excel                  (default: false)
      #  :if_unsaved    if an unsaved book with the same name is open, then
      #                 :raise   -> raise an exception                   (default)             
      #                 :accept  -> let the unsaved book open                  
      #                 :forget  -> close the unsaved book, open the new book
      #                 :excel   -> give control to excel
      #                 :new_app -> open the new book in a new excel application
      #  :if_unsaved_other_book   if an unsaved book with the same name in a different path is open, then
      #                  :raise   -> raise an exception                  (default)             
      #                  :save    -> save and close the unsaved book and open the new book
      #                  :forget  -> close the unsaved book, open the new book
      #                  :new_app -> open the new book in a new excel application
      def open(file, options={ :reuse => true}, &block)
        new(file, options, &block)
      end

    end

    def initialize(file, opts={ }, &block)
      @options = {
        :reuse => true,
        :read_only => false,
        :if_unsaved => :raise,
        :if_unsaved_other_book => :raise
      }.merge(opts)
      excel_app_options = {:reuse => true}.merge(opts).delete_if{|k,v| 
        k== :if_read_only || k== :unsaved || k == :if_unsaved_other_book}
      if not File.exist?(file)
        raise ExcelErrorOpen, "file #{file} not found"
      end
      @excel_app = ExcelApp.new(excel_app_options)     # :nodoc:
      workbooks = @excel_app.Workbooks
      @workbook = workbooks.Item(File.basename(file)) rescue nil
      if @workbook then
        blocked_by_other_book = (File.basename(file) == File.basename(@workbook.Fullname)) && 
                                (not (absolute_path(file) == @workbook.Fullname))
                                #(not (file == @workbook.Fullname.gsub("\\","/")))
        if blocked_by_other_book then
          case @options[:if_unsaved_other_book]
          when :raise
            raise ExcelErrorOpen, "blocked by an unsaved book with the same name in a different path"
          when :save
            #nothing
          when :forget
            @workbook.Close
          when :new_app
            @options[:reuse] = false
            @excel_app = ExcelApp.new(@options)
            @workbook = nil
          else
            raise ExcelErrorOpen, ":if_unsaved_other_book: invalid option"
          end
        else
          # book open, not saved, not blocked by other book
          if (not @workbook.Saved) then
            #p "book not saved"
            case @options[:if_unsaved]
            when :raise
              raise ExcelErrorOpen, "book is already open but not saved (#{File.basename(file)})"
            when :accept
              #nothing
            when :forget
              @workbook.Close
            when :excel
              old_displayalerts = @excel_app.DisplayAlerts  # :nodoc:
              @excel_app.DisplayAlerts = true  # :nodoc:
            when :new_app
              @options[:reuse] = false
              @excel_app = ExcelApp.new(@options)
              @workbook = nil
            else
              raise ExcelErrorOpen, ":if_unsaved: invalid option"
            end
          end
        end
      end
      begin
        # if book not open (was not open,was closed with option :forget or shall be opened in new application)
        #    or :if_unsaved => :excel
        if ((not alive?) || (@options[:if_unsaved] == :excel)) then
          begin
            @workbook = @excel_app.Workbooks.Open(absolute_path(file),{ 'ReadOnly' => @options[:read_only] })
          rescue WIN32OLERuntimeError
            raise ExcelUserCanceled, "open: canceled by user"
          end
        end
      ensure
        if @options[:if_unsaved] == :excel then
          @excel_app.DisplayAlerts = old_displayalerts  # :nodoc:
        end
      end
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
    #                      :excel   -> give control to excel
    def close(opts = {:if_unsaved => :raise})
      if ((alive?) && (not @workbook.Saved) && (not @options[:read_only])) then
        case opts[:if_unsaved]
        when :raise
          raise ExcelErrorClose, "book is unsaved (#{File.basename(filename)})"
        when :save
          save
        when :forget
          #nothing
        when :excel
          old_displayalerts = @excel_app.DisplayAlerts  # :nodoc:
          @excel_app.DisplayAlerts = true  # :nodoc:
        else
          raise ExcelErrorClose, ":if_unsaved: invalid option"
        end
      end
      begin
        @workbook.Close if alive?
        @workbook = nil unless alive?
        raise ExcelUserCanceled, "close: canceled by user" if alive? && opts[:if_unsaved] == :excel && (not @workbook.Saved)
      ensure
        if opts[:if_unsaved] == :excel then
          @excel_app.DisplayAlerts = old_displayalerts  # :nodoc:  
        end
      end
      #@excel_app.Workbooks.Close
      #@excel_app.Quit
    end

    # returns true, if the work book is alive, false otherwise
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

    # returns the full file name of the book
    def filename
      @workbook.Fullname.tr('\\','/') rescue nil
    end

    # returns true, if the full book names and excel appications are identical, false, otherwise  
    def == other_book
      other_book.is_a?(Book) &&
      @excel_app == other_book.excel_app &&
      self.filename == other_book.filename  
    end


    attr_reader :excel_app

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
    #               :excel     -> give control to Excel
    # returns true, if successfully saved, nil otherwise
    def save_as(file = nil, opts = {:if_exists => :raise} )
      raise IOError, "Not opened for writing(open with :read_only option)" if @options[:read_only]
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
          # if a book is open with the name of file, then raise error
          open_workbook = ExcelApp.reuse_if_possible.Workbooks(basename) rescue nil
          #workbook_file = @excel_app.Workbooks(basename) rescue nil
          if open_workbook == nil then
            File.delete(file) 
          else
            raise ExcelErrorSave, "book is open and used in Excel"
          end
        when :excel 
          old_displayalerts = @excel_app.DisplayAlerts  # :nodoc:
          @excel_app.DisplayAlerts = true  # :nodoc:
        when :raise
          raise ExcelErrorSave, "book already exists: #{basename}"
        else
          raise ExcelErrorSave, ":if_exists: invalid option"
        end
      end
      begin
        @workbook.SaveAs(absolute_path(File.join(dirname, basename)), file_format)
      rescue WIN32OLERuntimeError => msg
        if msg.message =~ /SaveAs/ and msg.message =~ /Workbook/ then
          if opts[:if_exists] == :excel then 
            raise ExcelErrorSave, "not saved or canceled by user"
          else
            return nil
          end
          # another possible semantics. raise ExcelErrorSaveFailed, "could not save Workbook"
        else
          raise ExcelErrorSaveUnknown, "unknown WIN32OELERuntimeError:\n#{msg.message}"
        end       
      ensure
        if opts[:if_exists] == :excel then
          @excel_app.DisplayAlerts = old_displayalerts  # :nodoc:
        end
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

      new_sheet = RobustExcelOle::Sheet.new(@excel_app.Activesheet)
      new_sheet.name = new_sheet_name if new_sheet_name
      new_sheet
    end        

    def absolute_path(file)
      file = File.expand_path(file)
      file = RobustExcelOle::Cygwin.cygpath('-w', file) if RUBY_PLATFORM =~ /cygwin/
      WIN32OLE.new('Scripting.FileSystemObject').GetAbsolutePathName(file)
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
