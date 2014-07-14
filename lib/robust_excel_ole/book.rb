# -*- coding: utf-8 -*-

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
    # :if_book_not_saved  if a book b with this name is already open:
    #                    :read_only -> let b open
    #                    :raise -> raise an exception,             if b is not saved
    #                    :accept -> let b open,                    if b is not saved
    #                    :forget -> open the new book and close b, if b is not saved
    #  :read_only     (boolean)  open in read-only mode
    #  :displayalerts (boolean)  allow display alerts in excel
    #  :visible       (boolean)  make visibe in Excel
    # if the file name is nil then return
      def open(file, options={ }, &block)
        new(file, options, &block)
      end
    end


    def initialize(file, options={ }, &block)
      unless caller[1] =~ /book.rb:\d+:in\s+`open'$/
        warn "DEPRECATION WARNING: RobustExcelOle::Book.new and RobustExcelOle::Book.open will be split. If you open existing file, please use RobustExcelOle::Book.open.(call from #{caller[1]})"
      end

      @options = {
        :read_only => true,
        :displayalerts => false,
        :visible => false,
        :if_book_not_saved => :raise
      }.merge(options)
      @winapp = WIN32OLE.new('Excel.Application')
      @winapp.DisplayAlerts = @options[:displayalerts]
      @winapp.Visible = @options[:visible]
      WIN32OLE.const_load(@winapp, RobustExcelOle) unless RobustExcelOle.const_defined?(:CONSTANTS)
      if not File.exist?(file)
        raise ExcelErrorOpen, "file #{file} not found"
      end
      #puts "file: #{file}"
      #puts "File.basename(file): #{File.basename(file)}"
      book_already_open = begin
      #  # findet file nicht
        book_open = @winapp.Workbooks(File.basename(file))
        true
      rescue WIN32OLERuntimeError
        false
      end

      puts "book already open = #{book_already_open}"
      puts "option:#{options[:if_book_not_saved]}"


      #if book_already_open then 
      #  if @options[:if_book_not_saved] == :read_only then
      #    puts "do nothing, keep the old one open"
      #  elsif true 
      #    #not mappe.Saved then 
      #    case @options[:if_book_not_saved]
      #    when :raise : 
      #      raise ExcelErrorOpen, "A book with the same name is already open, and has unsaved changes (#{file})"
      #    when :accept : 
      #    when :forget :  
      #      book_already_open = false
      #    else
      #      raise ExcelErrorOpen, "bug: invalid option (#{opts[:if_book_not_saved]})"
      #    end
      #  end
      #end
      #
      #if not book_already_open then
      #  begin
          @book = @winapp.Workbooks.Open(absolute_path(file),{ 'ReadOnly' => @options[:read_only] })
          if block
            begin
              yield self
            ensure
              close
            end
          end
      #    if :if_book_not_saved == :read_only then
      #      # do something?
      #    end
      #  rescue WIN32OLERuntimeError
      #    raise ExcelErrorOpen, "file not found"
      #  end
      #end
    
      #if @book.nil?
      #  raise ExcelErrorOpen, "Excel has not opened the book, but also did not raised an error (file name=#{file})"
      #end

      @book
    end

    def close
      @winapp.Workbooks.Close
      @winapp.Quit
    end


    # saves a book
    # if a file with the same name, exists, then proceed according to :if_exists 
    #   :raise     -> raise an exception, dont't write the file
    #   :overwrite -> write the file, delete the old file
    #   :excel     -> give control to Excel 
    # if file is nil, then return
    def save(file = nil, opts = {:if_exists => :raise} )
      raise IOError, "Not opened for writing(open with :read_only option)" if @options[:read_only]
      return @book.save unless file
      dirname, basename = File.split(file)
      file_format =
        case File.extname(basename)
        when '.xls' : RobustExcelOle::XlExcel8
        when '.xlsx': RobustExcelOle::XlOpenXMLWorkbook
        when '.xlsm': RobustExcelOle::XlOpenXMLWorkbookMacroEnabled
        end
      if File.exist?(file) then
        displayalerts_value = @winapp.DisplayAlerts
        case opts[:if_exists]
        when :overwrite
          File.delete(file) 
          #File.delete(absolute_path(File.join(dirname, basename)))
        when :excel 
          @winapp.DisplayAlerts = true 
          #raise ExcelErrorSave, "Option nicht implementiert"
        when :raise
          raise ExcelErrorSave, "book already exists: #{basename}"
        else
          raise ExcelErrorSave, "bug: invalid option (#{opts[:if_exists]})"
        end
      end
      @book.SaveAs(absolute_path(File.join(dirname, basename)), file_format)
      @winapp.DisplayAlerts = displayalerts_value 
    end

    def [] sheet
      sheet += 1 if sheet.is_a? Numeric
      RobustExcelOle::Sheet.new(@book.Worksheets.Item(sheet))
    end

    def each
      @book.Worksheets.each do |sheet|
        yield RobustExcelOle::Sheet.new(sheet)
      end
    end

    def add_sheet(sheet = nil, opts = { })
      if sheet.is_a? Hash
        opts = sheet
        sheet = nil
      end

      new_sheet_name = opts.delete(:as)

      after_or_before, base_sheet = opts.first || [:after, RobustExcelOle::Sheet.new(@book.Worksheets.Item(@book.Worksheets.Count))]
      base_sheet = base_sheet.sheet
      sheet ? sheet.Copy({ after_or_before.to_s => base_sheet }) : @book.WorkSheets.Add({ after_or_before.to_s => base_sheet })

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
