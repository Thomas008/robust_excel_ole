
# -*- coding: utf-8 -*-

require 'weakref'

module RobustExcelOle

  class Book
    attr_reader :excel
    attr_accessor :workbook
    attr_accessor :stored_filename

    class << self
      
      # opens a book.
      # 
      # options: 
      # :default_excel   if the book was already open in an Excel instance, then open it there, otherwise:
      #                   :reuse (default) -> connect to a running Excel instance if it exists, open in a new Excel otherwise
      #                   :new             -> open in a new Excel instance
      #                   <instance>       -> open in the given Excel instance
      # :force_excel     no matter whether the book was already open
      #                   :new (default)   -> open in a new Excel
      #                   <instance>       -> open in the given Excel instance
      # :if_locked       if the book is writable in another Excel instance, then
      #                   :take_writable (default) -> use the Excel instance in which the book is writable
      #                   :force_writability       -> make it writable in the desired Excel
      #                   :raise                   -> raise an exception
      # :if_locked_unsaved  if the book is open in another Excel instance and contains unsaved changes
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

     
      def open(file, opts={ }, &block)
        p "open:"
        set_defaults(opts)
        book = nil
        if (not (@options[:force_excel] == :new))
          # reopen the book
          p "fetch book"
          book = @@bookstore.fetch(file, :readonly_excel => (@options[:read_only] ? @options[:force_excel] : nil))
          if book
            p "book found" 
            if (not @options[:force_excel] || (@options[:force_excel] == book.excel))
              if book.excel.alive?
                p "excel alive"
                if (not book.alive?)
                  p "book not alive - try to reopen"
                  @options[:reopen] = true
                  book.workbook = book.get_workbook(file, book.excel)
                  p "workbook: #{book.workbook}"
                end
                return book if book.alive?
              end
            end
          end
        end
        @options[:excel] = @options[:force_excel] ? @options[:force_excel] : @options[:default_excel]
        @options[:reopen] = false
        p "call initialize"
        new(file, @options, &block)
      end
    end

    def initialize(file, opts={ }, &block)
      p "initialize:"
      Book.set_defaults(opts)
      @excel = Book.get_excel(opts)     
      p "@excel: #{@excel}"
      # get_workbook has side effect to @excel with :if_unsaved => :new_excel, :alerted, and :if_obstructed => :new_excel
      get_workbook(file, @excel)
      p "@workbook: #{@workbook}"
      @@bookstore.store(self)
      if block
        begin
          yield self
        ensure
          close
        end
      end
    end
  
  private


    def self.set_defaults(opts)
      @@bookstore ||= BookStore.new
      @options = {
        :excel => :reuse,
        :default_excel => :reuse,
        :if_locked     => :take_writable,       
        :if_unsaved    => :raise,
        :if_obstructed => :raise,
        :read_only => false
      }.merge(opts)
    end

    def self.get_excel(opts)
      p "get_excel:"
      if @options[:excel] == :reuse
        p ":reuse"
        excel = Excel.new(:reuse => true)
        p "excel: #{excel}"
      end
      @excel_options = nil
      if (not excel)
        p "no excel"
        if @options[:excel] == :new
          p ":new"
          @excel_options = {:displayalerts => false, :visible => false}.merge(opts)
          @excel_options[:reuse] = false
          excel = Excel.new(@excel_options)
          p "excel: #{excel}"
        else
          p "else:"
          excel = @options[:excel]
          p "excel: #{excel}"
        end
      end
      # if :excel => :new or (:excel => :reuse but could not reuse)
      if (not @excel_options)
        p "no excel_options (excel => :new or (:excel => :reuse but could not reuse)"
        excel.displayalerts = @options[:displayalerts] unless @options[:displayalerts].nil?
        excel.visible = @options[:visible] unless @options[:visible].nil?
      end
      p "excel: #{excel}"
      excel
    end

    def self.get_workbook(file, excel)
      p "get_workbook:"
      workbook = excel.Workbooks.Item(File.basename(file)) rescue nil
      if workbook then
        p "workbook exists already"
        obstructed_by_other_book = (File.basename(file) == File.basename(workbook.Fullname)) && 
                                   (not (RobustExcelOle::absolute_path(file) == workbook.Fullname))
        # if book is obstructed by a book with same name and different path
        if obstructed_by_other_book then
          case @options[:if_obstructed]
          when :raise
            raise ExcelErrorOpen, "blocked by a book with the same name in a different path"
          when :forget
            workbook.Close
            open_workbook(file,excel)
          when :save
            save unless workbook.Saved
            workbook.Close
            open_workbook(file,excel)
          when :close_if_saved
            if (not workbook.Saved) then
              raise ExcelErrorOpen, "book with the same name in a different path is unsaved"
            else 
              workbook.Close
              open_workbook(file,excel)
            end
          when :new_excel    
            if (not @options[:reopen])        
              @excel_options[:reuse] = false
              excel = Excel.new(@excel_options)
              workbook = nil
              open_workbook(file,excel)
            end
          else
            raise ExcelErrorOpen, ":if_obstructed: invalid option"
          end
        else
          # book open, not obstructed by an other book, but not saved
          if (not workbook.Saved) then
            case @options[:if_unsaved]
            when :raise
              raise ExcelErrorOpen, "book is already open but not saved (#{File.basename(file)})"
            when :forget
              workbook.Close
              open_workbook(file,excel)
            when :accept
              # do nothing
            when :alert
              # ???
              # if (not @options[:reopen])  
              #   @excel.with_displayalerts true do
              #     open_workbook file
              #   end
              # end
              excel.with_displayalerts true do
                open_workbook(file,excel)
              end 
            when :new_excel
              if (not @options[:reopen])
                @excel_options[:reuse] = false
                @excel = Excel.new(@excel_options)
                workbook = nil
                open_workbook(file,@excel)
              end
            else
              raise ExcelErrorOpen, ":if_unsaved: invalid option"
            end
          end
        end
      else
        # book is not open
        p "book not open"
        open_workbook(file,excel)
      end
    end


    def self.open_workbook(file,excel)
      p "open_workbook:"
      # ... (not alive?)
      if (@options[:reopen] || (not @workbook) || (@options[:if_unsaved] == :alert)) then
        begin
          filename = RobustExcelOle::absolute_path(file)
          p "filename: #{filename}"
          workbooks = excel.Workbooks
          workbooks.Open(filename,{ 'ReadOnly' => @options[:read_only] })
          # workaround for bug in Excel 2010: workbook.Open does not always return 
          # the workbook with given file name
          workbook = workbooks.Item(File.basename(filename))
          p "workbook: #{workbook}"
          workbook
        rescue BookStoreError => e
          raise ExcelUserCanceled, "open: canceled by user: #{e}"
        end
      end
    end

  public

    # closes the book, if it is alive
    #
    # options:
    #  :if_unsaved    if book is unsaved
    #                      :raise   -> raise an exception       (default)             
    #                      :save    -> save the book before it is closed                  
    #                      :forget  -> close the book 
    #                      :alert   -> give control to excel
    def close(opts = {:if_unsaved => :raise})
      if ((alive?) && (not @workbook.Saved) && (not opts[:read_only])) then
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

  private

    def close_workbook    
      @workbook.Close if alive?
      @workbook = nil unless alive?
    end

  public

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
        #book = open(filename, :if_unsaved => 
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
      @opts = opts
      if File.exist?(file) then
        case @opts[:if_exists]
        when :overwrite
          begin
            File.delete(file) 
          rescue Errno::EACCES
            raise ExcelErrorSave, "book is open and used in Excel"
          end
          save_as_workbook(file)
        when :alert 
          @excel.with_displayalerts true do
            save_as_workbook(file)
          end
        when :raise
          raise ExcelErrorSave, "book already exists: #{File.basename(file)}"
        else
          raise ExcelErrorSave, ":if_exists: invalid option"
        end
      else
        save_as_workbook(file)
      end
      true
    end

  private

    def save_as_workbook(file)
      begin
        dirname, basename = File.split(file)
        file_format =
          case File.extname(basename)
            when '.xls' : RobustExcelOle::XlExcel8
            when '.xlsx': RobustExcelOle::XlOpenXMLWorkbook
            when '.xlsm': RobustExcelOle::XlOpenXMLWorkbookMacroEnabled
          end
        @workbook.SaveAs(RobustExcelOle::absolute_path(file), file_format)
        @@bookstore.store(self)  
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

  public

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

    def book_store
      @@bookstore
    end

  private

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

  end
  
public

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
