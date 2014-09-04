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


    class << self

      # opens a book. 
      # options: 
      #  :reuse         (boolean)  use an already open Excel-application (default: true)
      #  :read_only     (boolean)  open in read-only mode                (default: false)
      #  :displayalerts (boolean)  allow display alerts in Excel         (default: false)
      #  :visible       (boolean)  make visibe in Excel                  (default: false)
      # :if_unsaved     if an unsaved book b with this file name is already open, then
      #                 :raise -> raise an exception                     (default)             
      #                 :accept -> let b open,                  
      #                 :forget -> open the new book and close b
      # if the file name is nil then return

      def open(file, options={ :reuse => true}, &block)
        new(file, options, &block)
      end

    end

    def initialize(file, options={ }, &block)
      @options = {
        :reuse => true,
        :if_unsaved => :raise,
        :read_only => false
      }.merge(options)

      if not File.exist?(file)
        raise ExcelErrorOpen, "file #{file} not found"
      end      
      # ToDo: filter out the options specific to Book:
      @excel_app = ExcelApp.new(@options)
      workbooks = @excel_app.Workbooks
      @workbook = workbooks.Item(File.basename(file)) rescue nil
      if @workbook then
        # book open and not saved
        if (not @workbook.Saved) then
          #p "book not saved"
          case @options[:if_unsaved]
          when :raise
            raise ExcelErrorOpen, "book is already open but not saved (#{File.basename(file)})"
          when :accept
            #nothing
          when :forget
            @workbook.Close 
          else
            raise ExcelErrorOpen, "invalid option"
          end
        end
      end
      # book not open (was not open or was closed with option :forget)
      if not alive? then
        @workbook = @excel_app.Workbooks.Open(absolute_path(file),{ 'ReadOnly' => @options[:read_only] })
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
    def close
      @workbook.Close if alive?  
      @workbook = nil
      #@excel_app.Workbooks.Close
      #@excel_app.Quit
    end

    # returns true, if the work book is alive, false, otherwise
    def alive?
      begin 
        @workbook.Name
        true
      rescue 
        #puts $!.message
        false
      end
    end

    # ToConsider: different name :
    # returns the full filename of the book
    def filename
      @workbook.Fullname.tr('\\','/')
    end

    #returns true, if the full book names and excel appications are identical, false, otherwise  
    def == other_book
      other_book.is_a?(Book) &&
      @excel_app == other_book.excel_app &&
      self.filename == other_book.filename  
    end


    attr_reader :excel_app


    # saves a book.
    # options:
    #  :if_exists   if a file with the same name exists, then  
    #               :raise     -> raise an exception, dont't write the file  (default)
    #               :overwrite -> write the file, delete the old file
    #               :excel     -> give control to Excel
    # returns true, if successfully saved, nil otherwise
    def save(file = nil, opts = {:if_exists => :raise} )
      raise IOError, "Not opened for writing(open with :read_only option)" if @options[:read_only]
      unless file
        @workbook.Save 
        return true
      end

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
          old_displayalerts = @excel_app.DisplayAlerts
          @excel_app.DisplayAlerts = true 
        when :raise
          raise ExcelErrorSave, "book already exists: #{basename}"
        else
          raise ExcelErrorSave, "invalid option (#{opts[:if_exists]})"
        end
      end
      begin
        @workbook.SaveAs(absolute_path(File.join(dirname, basename)), file_format)
      rescue WIN32OLERuntimeError => msg
        if msg.message =~ /SaveAs/ and msg.message =~ /Workbook/ then
          return nil
          # another possible semantics. raise ExcelErrorSaveFailed, "could not save Workbook"
        else
          raise ExcelErrorSaveUnknown, "unknown WIN32OELERuntimeError:\n#{msg.message}"
        end       
      ensure
        if opts[:if_exists] == :excel then
          @excel_app.DisplayAlerts = old_displayalerts
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

    # adds a sheet
    def add_sheet(sheet = nil, opts = { })
      if sheet.is_a? Hash
        opts = sheet
        sheet = nil
      end

      new_sheet_name = opts.delete(:as)

      after_or_before, base_sheet = opts.first || [:after, RobustExcelOle::Sheet.new(@workbook.Worksheets.Item(@workbook.Worksheets.Count))]
      base_sheet = base_sheet.sheet
      sheet ? sheet.Copy({ after_or_before.to_s => base_sheet }) : @workbook.WorkSheets.Add({ after_or_before.to_s => base_sheet })

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

class ExcelErrorSaveFailed < ExcelErrorSave  
end

class ExcelErrorSaveUnknown < ExcelErrorSave  
end

class ExcelErrorOpen < RuntimeError
end
