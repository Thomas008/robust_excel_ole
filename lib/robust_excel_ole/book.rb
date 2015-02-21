
# -*- coding: utf-8 -*-

require 'weakref'

module RobustExcelOle

  class Book
    attr_reader :workbook
    attr_reader :excel

     # book management for persisten storage:
     # data structure: {filename1 => [book1,...bookn], filename2 => ...} 
     @@filename2book = {}

    class << self

      # opens a book.
      # 
      # options: 
      # :excel         determines the Excel in which to open the book 
      #                  :reuse (default) -> connect to a running Excel, if it exists, open a new Excel, otherwise
      #                  :new             -> open in a new Excel
      #                  <instance>       -> open in the given Excel instance
      # :force          if the book was already open in an Excel that is still working, then:
      #                  false     -> use this Excel (reopen), else use the Excel given in :excel
      #                  true      -> use the Excel given in :excel, even if the book was opened before
      # :if_locked      if the book is open in another Excel and writable there, then
      #                  :go_there (default) -> use the Excel in which the book is writable?
      #                  :force    -> make it writable when opening new
      #                  :raise    -> raise an exception
      # :if_locked_unsaved  if the book is open in another Excel and contains unsaved changes
      #                  :raise    -> raise an exception
      #                  :save     -> save the unsaved book 
      # :if_unsaved     if an unsaved book with the same name is open, then
      #                  :raise     -> raise an exception (default)
      #                  :forget    -> close the unsaved book, open the new book             
      #                  :accept    -> let the unsaved book open                  
      #                  :alert     -> give control to Excel
      #                  :new_excel -> open the new book in a new Excel
      # :if_obstructed  if a book with the same name in a different path is open, then
      #                  :raise          -> raise an exception (default)             
      #                  :forget         -> close the old book, open the new book
      #                  :save           -> save the old book, close it, open the new book
      #                  :close_if_saved -> close the old book and open the new book, if the old book is saved
      #                                     raise an exception otherwise
      #                  :new_excel      -> open the new book in a new Excel
      #                  :reuse_excel    -> try the next free running Excel, if it exists, open a new Excel, else
      # :read_only     open in read-only mode          (default: false)
      # :displayalerts allow display alerts in Excel   (default: false)
      # :visible       make visibe in Excel            (default: false)
      def open(file, options={ :reuse => true}, &block)
        new(file, options, &block)
      end

    end

    def initialize(file, opts={ }, &block)
      @options = {
        :excel => :reuse,
        :force => false,
        :if_locked => :go_there,
        :read_only => false,
        :if_unsaved => :raise,
        :if_obstructed => :raise,
        :displayalerts => false,
        :visible => false
      }.merge(opts)
      if not File.exist?(file)
        raise ExcelErrorOpen, "file #{file} not found"
      end  
      @file = file
      # ohne Persistenz
      # es soll auch mögich sein: ich gebe Instanz an und übergebe :displayalerts und :visible
      # mit with_displayalerts, =visible
      if @options[:excel]
        @excel = @options[:excel]
        @excel.visible = @options[:visible]
        @excel.displayalerts = @options[:dispayalerts]
      else
        excel_options = {:reuse => ((@options[:excel] == :reuse) ? true : false),
                         :displayalerts => @options[:displayalerts], :visible => @options[:visible]}
        @excel = Excel.new(excel_options)
      end
      if (not @excel.alive?)
        raise ExcelErrorOpen, "Excel instance is not alive"
      end
     
#excel_options = {:reuse => true}.merge(opts).delete_if{|k,v| k == :force || || k == :if_locked || k== :read_only || k== :if_unsaved || k == :if_obstructed}



      #book = Book.connect(@file)
      #if book
      #  p "connect"
      #  # what is with the other excel options?
      #  # nehme die Optionen von book.excel. wenn neue hier Optionen gesetzt wurden, dann nehme diese
      #  @excel = book.excel.alive? ? book.excel : (@options[:excel] ? excel_options[:excel] : Excel.new(excel_options))
      #  @workbook = book.workbook 
      #else
        @excel = @options[:excel] ? excel_options[:excel] : Excel.new(excel_options)
        @workbook = @excel.Workbooks.Item(File.basename(@file)) rescue nil
      #end
      # if book is open
      if @workbook then
        obstructed_by_other_book = (File.basename(file) == File.basename(@workbook.Fullname)) && 
                                   (not (RobustExcelOle::absolute_path(file) == @workbook.Fullname))
        # if book is obstructed by a book with same name and different path
        if obstructed_by_other_book then
          case @options[:if_obstructed]
          when :raise
            raise ExcelErrorOpen, "blocked by a book with the same name in a different path"
          when :forget
            @workbook.Close
            open_workbook
          when :save
            save unless @workbook.Saved
            @workbook.Close
            open_workbook
          when :close_if_saved
            if (not @workbook.Saved) then
              raise ExcelErrorOpen, "book with the same name in a different path is unsaved"
            else 
              @workbook.Close
              open_workbook
            end
          when :new_excel
            excel_options[:reuse] = false
            @excel = Excel.new(excel_options)
            @workbook = nil
            open_workbook
          else
            raise ExcelErrorOpen, ":if_obstructed: invalid option"
          end
        else
          # book open, not obstructed by an other book, but not saved
          if (not @workbook.Saved) then
            case @options[:if_unsaved]
            when :raise
              raise ExcelErrorOpen, "book is already open but not saved (#{File.basename(file)})"
            when :forget
              @workbook.Close
              open_workbook
            when :accept
              # do nothing
            when :alert
              @excel.with_displayalerts true do
                open_workbook
              end 
            when :new_excel
              excel_options[:reuse] = false
              @excel = Excel.new(excel_options)
              @workbook = nil
              open_workbook
            else
              raise ExcelErrorOpen, ":if_unsaved: invalid option"
            end
          end
        end
      else
        # book is not open
        open_workbook
      end
      if block
        begin
          yield self
        ensure
          close
        end
      end
    end
  
    def open_workbook
      # if book not open (was not open,was closed with option :forget or shall be opened in new application)
      #    or :if_unsaved => :alert
      if ((not alive?) || (@options[:if_unsaved] == :alert)) then
        begin
          #p "open_workbook:"
          #p "@@filename2book: #{@@filename2book.inspect}"
          filename = RobustExcelOle::absolute_path(@file)
          workbooks = @excel.Workbooks
          workbooks.Open(filename,{ 'ReadOnly' => @options[:read_only] })
          # workaround for bug in Excel 2010: workbook.Open does not always return 
          # the workbook with given file name
          @workbook = workbooks.Item(File.basename(filename))
          # book eintragen in Book-Management
          filename_key = RobustExcelOle::canonize(self.filename)
          #p "filename_key: #{filename_key}"
          if @@filename2book[filename_key]
            @@filename2book[filename_key] << self unless @@filename2book[filename_key].include?(self)
          else
            @@filename2book[filename_key] = [self]
          end
          #p "@@filename2book:"
          @@filename2book.each do |element|
            #p " filename: #{element[0]}"
            #p " books:"
            element[1].each do |book|
              #p "#{book}"
            end
          end
        rescue WIN32OLERuntimeError
          raise ExcelUserCanceled, "open: canceled by user"
        end
      end
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

    def close_workbook    
      @workbook.Close if alive?
      @workbook = nil unless alive?
    end

    # returns a book with the filename, if it was open one time
    # preference order: writable book, readonly unsaved book, readonly book (the last one), dead book
    def self.connect(filename)
      #p "connect:"
      #p "@@filename2book:"
      @@filename2book.each do |element|
        #p " filename: #{element[0]}"
        #p " books:"
        element[1].each do |book|
          #p "#{book}"
        end
      end
      filename_key = RobustExcelOle::canonize(filename)
      #p "filename_key: #{filename_key}"
      book_writable = book_readonly = book_readonly_unsaved = book_closed = nil
      books = @@filename2book[filename_key]
      #p "books: #{books}"
      return nil unless books
      books.each do |book|
        #p "book: #{book}"
        if book.alive?
          #p "book alive"
          if (not book.ReadOnly)
            #p "book writable"
            book_writable = book 
            break 
          else
            #p "book read_only"
            book.Saved ? (book_readonly = book) : (book_readonly_unsaved = book)
          end
        else
          #p "book closed"
          book_closed = book
        end
      end
      result = book_writable ? book_writable : (book_readonly_unsaved ? book_readonly_unsaved : 
                                                                        (book_readonly ? book_readonly : book_closed))
      #p "book: #{result}"
      result 
    end


    # modify a book such that its state remains unchanged.
    # options: :keep_open: let the book open after modification
    def self.unobtrusively(filename, opts = {:keep_open => false})
      book = self.connect(filename)
      was_nil = book.nil?
      was_alive = book.alive?
      was_saved = ((not was_nil) && was_alive) ? book.Saved : true
      #was_saved = book.Saved unless was_closed 
      begin
        book = open(filename, :if_unsaved => :accept, :if_obstructed => :new_excel) if (was_nil || (not was_alive))
        #book = open(filename, :if_unsaved => :accept, :if_obstructed => :new_excel) unless book 
        yield book
      ensure
        book.save if was_saved && (not book.ReadOnly)
        book.close(:if_unsaved => :save) if (was_nil && (not opts[:keep_open]))
      end
      book
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
    def visible= visible_value
      Excel.current.Visible = visible_value
    end

    # returns if the current Excel application is visible
    def visible 
      Excel.current.Visible
    end

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
      @file = file
      @opts = opts
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
          raise ExcelErrorSave, "book already exists: #{File.basename(file)}"
        else
          raise ExcelErrorSave, ":if_exists: invalid option"
        end
      else
        save_as_workbook
      end
      true
    end
  
    def save_as_workbook
      begin
        dirname, basename = File.split(@file)
        file_format =
          case File.extname(basename)
            when '.xls' : RobustExcelOle::XlExcel8
            when '.xlsx': RobustExcelOle::XlOpenXMLWorkbook
            when '.xlsm': RobustExcelOle::XlOpenXMLWorkbookMacroEnabled
          end
        filename_key = RobustExcelOle::canonize(@file)
        #p "filename_key: #{filename_key}"   
        if @@filename2book[filename_key]
          @@filename2book[filename_key] << self unless @@filename2book[filename_key].include?(self)
        else
          @@filename2book[filename_key] = [self]
        end
        #p "@@filename2book:"
        @@filename2book.each do |element|
          #p " filename: #{element[0]}"
          #p " books:"
          element[1].each do |book|
           # p "#{book}"
          end
        end                   
        @workbook.SaveAs(RobustExcelOle::absolute_path(@file), file_format)
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

    def method_missing(name, *args)
      if name.to_s[0,1] =~ /[A-Z]/ 
        begin
          @workbook.send(name, *args)
        rescue WIN32OLERuntimeError => msg
          if msg.message =~ /unknown property or method/
            raise VBAMethodMissingError, "unknown VBA property or method #{name}"
          else 
            raise msg
          end
        end
      else  
        super 
      end
    end

    private :open_workbook, :close_workbook, :save_as_workbook, :method_missing

  end
end