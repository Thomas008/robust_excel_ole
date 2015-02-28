
# -*- coding: utf-8 -*-

require 'weakref'

module RobustExcelOle

  class Book
    attr_reader :workbook
    attr_accessor :excel

     # book management for persisten storage:
     # data structure: {filename1 => [book1,...bookn], filename2 => ...} 
     @@filename2book = {}

    class << self

      # opens a book.
      # 
      # options: 
      # :default_excel   if the book was already open in a Excel, then open it there, otherwise:
      #                   :reuse (default) -> connect to a running Excel if it exists, open a new Excel otherwise
      #                   :new             -> open in a new Excel
      #                   <instance>       -> open in the given Excel instance
      # :force_excel     no matter whether the book was already open
      #                   :new (default)   -> open in a new Excel
      #                   <instance>       -> open in the given Excel instance
      # :if_locked       if the book is writable in another Excel , then
      #                   :take_writable (default) -> use the Excel in which the book is writable
      #                   :force_writability       -> make it writable in the desired Excel
      #                   :raise                   -> raise an exception
      # :if_locked_unsaved  if the book is open in another Excel and contains unsaved changes
      #                  :raise    -> raise an exception
      #                  :save     -> save the unsaved book 
      # :if_unsaved     if an unsaved book with the same name is open, then
      #                  :raise (default) -> raise an exception
      #                  :forget          -> close the unsaved book, open the new book             
      #                  :accept          -> let the unsaved book open                  
      #                  :alert           -> give control to Excel
      #                  :new_excel       -> open the new book in a new Excel
      # :if_obstructed  if a book with the same name in a different path is open, then
      #                  :raise (default) -> raise an exception 
      #                  :forget          -> close the old book, open the new book
      #                  :save            -> save the old book, close it, open the new book
      #                  :close_if_saved  -> close the old book and open the new book, if the old book is saved
      #                                      raise an exception otherwise
      #                  :new_excel       -> open the new book in a new Excel
      #                  :reuse_excel     -> try the next free running Excel, if it exists, open a new Excel, else
      # :read_only     open in read-only mode         (default: false)
      # :displayalerts enable DisplayAlerts in Excel  (default: false)
      # :visible       make visibe in Excel           (default: false)
      # If :default_excel is set, then DisplayAlerts and Visible are set only if these parameters are given,
      #                                   not set by default
      def open(file, options={ }, &block)
        p" open"
        book = nil
        if (options[:default_excel] || (not options[:force_excel]))
          p ":reuse_excel is set or not :force_excel => true"
          book = find_book(file)
          p "book: #{book}"
          if book 
            p "book exists"
            if (not book.excel.alive?)
              p "book.excel is not alive"
              book.excel = Excel.new(:excel => book.excel, 
                                     :displayalerts => book.excel.displayalerts, :visible => book.excel.visible)
            end
            p "return book"
            return book
          end
        end
        if options[:force_excel] || book.nil?
          p ":force_excel is set or book nil"
          # if :reuse_excel is set, then :excel = :reuse_excel, else :excel = :force_excel
          options[:excel] = (options[:default_excel] || options[:force_excel]) ?  
              (options[:default_excel] ? options[:default_excel] : options[:force_excel]) : :reuse
          p ":excel => #{options[:excel]}"
          new(file, options, &block)
        end
      end
    end

    def initialize(file, opts={ }, &block)
      p "initialize"
      @options = {
        :if_locked     => :take_writable,       
        :if_unsaved    => :raise,
        :if_obstructed => :raise,
        :read_only => false
      }.merge(opts)
      p ":excel => #{@options[:excel]}"
      if not File.exist?(file)
        raise ExcelErrorOpen, "file #{file} not found"
      end  
      @file = file
      if @options[:excel] == :reuse
        @excel = Excel.new(:reuse => true)
        p "@excel: #{@excel}"       
      end
      excel_options = nil
      if (not @excel)
        p "not @excel"
        if (@options[:excel] == :new || @options[:excel] == :reuse)
          p ":excel => @options[:excel]"
          excel_options = {:displayalerts => false, :visible => false}.merge(opts)
          excel_options[:reuse] = false
          @excel = Excel.new(excel_options)
          p "@excel: #{@excel}"
        else
          p "excel instance is given"
          @excel = @options[:excel]
          p "excel: #{@excel}"
        end
      end
      # if :excel => new or (:excel => :reuse but could not reuse)
      if (not excel_options)
        @excel.displayalerts = @options[:displayalerts] unless @options[:displayalerts].nil?
        @excel.visible = @options[:visible] unless @options[:visible].nil?
      end
      @workbook = @excel.Workbooks.Item(File.basename(@file)) rescue nil
      p "excel: #{@excel}  workbook: #{@workbook}"
      # book is open
      if @workbook then
        p "book is open"
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

    # returns a book with the given filename, if it was open once
    # preference order: writable book, readonly unsaved book, readonly book (the last one), closed book
    def self.find_book(filename)
      p "find_book:"
      p "@@filename2book:"
      @@filename2book.each do |element|
        p " filename: #{element[0]}"
        p " books:"
        element[1].each do |book|
          p "#{book}"
        end
      end
      filename_key = RobustExcelOle::canonize(filename)
      p "filename_key: #{filename_key}"
      readonly_book = readonly_unsaved_book = closed_book = nil
      books = @@filename2book[filename_key]
      p "books: #{books}"
      return nil  unless books
      books.each do |book|
        p "book: #{book}"
        if book.alive?
          p "book alive"
          if (not book.ReadOnly)
            p "book writable"
            return book 
          else
            p "book read_only"
            book.Saved ? (readonly_book = book) : (book_readonly_unsaved = book)
          end
        else
          p "book closed"
          closed_book = book
        end
      end
      result = readonly_unsaved_book ? readonly_unsaved_book : (readonly_book ? readonly_book : closed_book)
      p "book: #{result}"
      result
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

 
    # modify a book such that its state remains unchanged.
    # options: :keep_open: let the book open after modification
    def self.unobtrusively(filename, opts = {:keep_open => false})
      book = self.open(filename)
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

    # returns if the Excel instance is visible
    def visible 
      @excel.visible
    end

   # make the Excel instance visible or invisible
    # option: visible_value     true -> make Excel visible, false -> make Excel invisible
    def visible= visible_value
      @excel.visible = visible_value
    end

   # returns if DisplayAlerts is enabed in the Excel instance
    def displayalerts 
      @excel.displayalerts
    end

    # enable in the Excel instance Dispayalerts
    #  option: displayalerts_value     true -> enable DisplayAlerts, false -> disable DispayAlerts
    def displayalerts= displayalerts_value
      @excel.displayalerts = displayalerts_value
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


__END__


          class Object
            def update_extracted hash, key
              value = hash[param_name]
              self.send("#{key}=", value) if value
            end
          end
          @excel.visible = @options[:visible] if @options[:visible] 
          @excel.displayalerts = @options[:dispayalerts]    
          @excel.update_extracted(@options, [:visible, :dispayalerts])
          @excel.options.merge(@options.extract(:visible, :dispayalerts))
