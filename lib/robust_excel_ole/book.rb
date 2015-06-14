
# -*- coding: utf-8 -*-

require 'weakref'

module RobustExcelOle

  class Book
    attr_accessor :excel
    attr_accessor :workbook
    attr_accessor :stored_filename
    attr_accessor :options

      DEFAULT_OPEN_OPTS = { 
        :excel => :reuse,
        :default_excel => :reuse,
        :if_locked     => :readonly,       
        :if_unsaved    => :raise,
        :if_obstructed => :raise,
        :read_only => false
      }

    class << self
      
      # opens a book.
      # 
      # when reopening a book that was opened and closed before, transparency identity is ensured:
      # same Book objects refer to the same Excel files, and vice versa
      # 
      # options: 
      # :default_excel   if the book was already open in an Excel instance, then open it there.
      #                  Otherwise, i.e. if the book was not open before or the Excel instance is not alive
      #                   :reuse (default) -> connect to a (the first opened) running Excel instance,
      #                                        excluding the hidden Excel instance, if it exists,
      #                                       otherwise open in a new Excel instance.
      #                   :new             -> open in a new Excel instance
      #                   <instance>       -> open in the given Excel instance
      # :force_excel     no matter whether the book was already open
      #                   :new (default)   -> open in a new Excel
      #                   <instance>       -> open in the given Excel instance
      # :if_unsaved     if an unsaved book with the same name is open, then
      #                  :raise (default)              -> raise an exception
      #                  :forget              -> close the unsaved book, open the new book             
      #                  :accept              -> let the unsaved book open                  
      #                  :alert               -> give control to Excel
      #                  :new_excel           -> open the new book in a new Excel
      # :if_obstructed  if a book with the same name in a different path is open, then
      #                  :raise (default)     -> raise an exception 
      #                  :forget              -> close the old book, open the new book
      #                  :save                -> save the old book, close it, open the new book
      #                  :close_if_saved      -> close the old book and open the new book, if the old book is saved,
      #                                          otherwise raise an exception.
      #                  :new_excel           -> open the new book in a new Excel    
      # :if_absent       :create (default)    -> creates a new Excel file, if it does not exists  
      #                  :raise               -> raises an exception     , if the file does not exists
      # :read_only     open in read-only mode         (default: false) 
      # :displayalerts enable DisplayAlerts in Excel  (default: false)
      # :visible       make visible in Excel           (default: false)
      # if :default_excel is set, then DisplayAlerts and Visible are set only if these parameters are given

      def open(file, opts={ }, &block)
        p "open:"
        current_options = DEFAULT_OPEN_OPTS.merge(opts)
        book = nil
        if (not (current_options[:force_excel] == :new && (not current_options[:if_locked] == :take_writable)))
          # if readonly is true, then prefer a book that is given in force_excel if this option is set
          book = book_store.fetch(file, :prefer_writable => (not current_options[:read_only]), 
                                        :prefer_excel    => (current_options[:read_only] ? current_options[:force_excel] : nil)) rescue nil
          if book
            if (((not current_options[:force_excel]) || (current_options[:force_excel] == book.excel)) &&
                 (not (book.alive? && (not book.saved) && (not current_options[:if_unsaved] == :accept))))
              book.options = DEFAULT_OPEN_OPTS.merge(opts)
              book.get_excel unless book.excel.alive?
              # if the book is opened as readonly and should be opened as writable, then close it and open the book with the new readonly mode
              book.close if (book.alive? && (not book.writable) && (not current_options[:read_only]))
              # reopen the book
              book.get_workbook unless book.alive?
              return book
            end
          end
        end
        current_options[:excel] = current_options[:force_excel] ? current_options[:force_excel] : current_options[:default_excel]
        new(file, current_options, &block)
      end
    end

    def initialize(file, opts={ }, &block)
      p "initialize"
      @options = DEFAULT_OPEN_OPTS.merge(opts)
      @file = file      
      get_excel
      get_workbook
      book_store.store(self)
      if block
        begin
          yield self
        ensure
          close
        end
      end
    end
    
    def get_excel
      p "get_excel"
      if @options[:excel] == :reuse
        @excel = Excel.new(:reuse => true)
      end
      @excel_options = nil
      if (not @excel)
        if @options[:excel] == :new
          @excel_options = {:displayalerts => false, :visible => false}.merge(@options)
          @excel_options[:reuse] = false
          @excel = Excel.new(@excel_options)
        else
          @excel = @options[:excel]
        end
      end
      # if :excel => :new or (:excel => :reuse but could not reuse)
      #   keep the old values for :visible and :displayalerts, set them only if the parameters are given
      if (not @excel_options)
        @excel.displayalerts = @options[:displayalerts] unless @options[:displayalerts].nil?
        @excel.visible = @options[:visible] unless @options[:visible].nil?
      end
    end

    def get_workbook
      p "get_workbook"
      raise ExcelErrorOpen, "file #{@file} not found" if ((not File.exist?(@file)) && @options[:if_absent] == :raise) 
      @workbook = @excel.Workbooks.Item(File.basename(@file)) rescue nil
      if @workbook then
        obstructed_by_other_book = (File.basename(@file) == File.basename(@workbook.Fullname)) && 
                                   (not (RobustExcelOle::absolute_path(@file) == @workbook.Fullname))
        # if book is obstructed by a book with same name and different path
        if obstructed_by_other_book then
          case @options[:if_obstructed]
          when :raise
            raise ExcelErrorOpen, "blocked by a book with the same name in a different path"
          when :forget
            @workbook.Close
            @workbook = nil
            open_or_create_workbook
          when :save
            save unless @workbook.Saved
            @workbook.Close
            @workbook = nil
            open_or_create_workbook
          when :close_if_saved
            if (not @workbook.Saved) then
              raise ExcelErrorOpen, "book with the same name in a different path is unsaved"
            else 
              @workbook.Close
              @workbook = nil
              open_or_create_workbook
            end
          when :new_excel 
            @excel_options = {:displayalerts => false, :visible => false}.merge(@options)   
            @excel_options[:reuse] = false
            @excel = Excel.new(@excel_options)
            open_or_create_workbook
          else
            raise ExcelErrorOpen, ":if_obstructed: invalid option"
          end
        else
          # book open, not obstructed by an other book, but not saved and writable
          if (not @workbook.Saved) then
            case @options[:if_unsaved]
            when :raise
              raise ExcelErrorOpen, "book is already open but not saved (#{File.basename(@file)})"
            when :forget
              @workbook.Close
              @workbook = nil
              open_or_create_workbook
            when :accept
              # do nothing
            when :alert
              @excel.with_displayalerts true do
                open_or_create_workbook
              end 
            when :new_excel
              @excel_options = {:displayalerts => false, :visible => false}.merge(@options)
              @excel_options[:reuse] = false
              @excel = Excel.new(@excel_options)
              open_or_create_workbook
            else
              raise ExcelErrorOpen, ":if_unsaved: invalid option"
            end
          end
        end
      else
        # book is not open
        open_or_create_workbook
      end
    end

    def open_or_create_workbook
      p "open_or_create_workbook"
      p "@file: #{@file}"
      if (not File.exist?(@file))
        p "here0"
        @workbook = Excel.current.generate_workbook(@file)
        return
      end
      p "here1"
      if ((not @workbook) || (@options[:if_unsaved] == :alert) || @options[:if_obstructed]) then
        begin
          p "here2"
          filename = RobustExcelOle::absolute_path(@file)
          workbooks = @excel.Workbooks
          p "before"
          workbooks.Open(filename,{ 'ReadOnly' => @options[:read_only] })
          p "after"
        rescue WIN32OLERuntimeError => msg 
          raise ExcelErrorOpen, "open: user canceled or open error" if msg.message =~ /OLE error code:800A03EC/
        end   
        begin
          # workaround for bug in Excel 2010: workbook.Open does not always return 
          # the workbook with given file name
          @workbook = workbooks.Item(File.basename(filename))
        rescue WIN32OLERuntimeError
          raise ExcelErrorOpen, "open: item error"
        end
      end
    end

    # closes the book, if it is alive
    #
    # options:
    #  :if_unsaved    if book is unsaved
    #                      :raise (default) -> raise an exception       
    #                      :save            -> save the book before it is closed                  
    #                      :forget          -> close the book 
    #                      :alert           -> give control to excel
    def close(opts = {:if_unsaved => :raise})
      if (alive? && (not @workbook.Saved) && writable) then
        case opts[:if_unsaved]
        when :raise
          raise ExcelErrorClose, "book is unsaved (#{File.basename(self.stored_filename)})"
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

    # modify a book such that its state (open/close, saved/unsaved, readonly/writable) remains unchanged.
    #  options: 
    #  :if_closed :  :hidden (default) : open closed books in one separate Excel instance that is not visible and has no displayaslerts
    #                :reuse            : open closed books in the Excel instance of the book, if it exists, reuse another Excel, otherwise         
    #                <excel-instance>  : open closed books in the given Excel instance
    #  :read_only: Open the book unobtrusively for reading only  (default: false)
    #  :use_readonly_excel:  if the book is opened only as ReadOnly and shall be modified, then
    #              true:  close it and open it as writable in the excel instance where it was open so far
    #              false (default)   open it as writable in another running excel instance, if it exists,
    #                                otherwise open in a new excel instance.
    #  :keep_open: let the book open after unobtrusively opening (default: false)
    def self.unobtrusively(file, opts = { })
      options = {
        :if_closed => :hidden,
        :read_only => false,
        :use_readonly_excel => false,
        :keep_open => false,
      }.merge(opts)
      book = book_store.fetch(file, :prefer_writable => (not options[:read_only]))
      was_not_alive_or_nil = book.nil? || (not book.alive?)
      was_saved = was_not_alive_or_nil ? true : book.saved
      was_writable = book.writable unless was_not_alive_or_nil
      old_visible = (book && book.excel.alive?) ? book.excel.visible : false
      begin 
        book = 
          if was_not_alive_or_nil 
            case options[:if_closed] 
            when :hidden 
              open(file, :force_excel => book_store.hidden_excel)
            when :reuse
              open(file)
            else 
              options[:if_closed].alive? ? open(file, :force_excel => options[:if_closed]) : open(file)
            end
          else
            if was_writable || options[:read_only]
              book
            else
              options[:use_readonly_excel] ? open(file, :force_excel => book.excel) : open(file, :force_excel => :new)
            end
          end
        yield book
      ensure
        book.save if (was_not_alive_or_nil || was_saved || ((not was_writable) && (not options[:read_only]))) && (not book.saved)
        # book was open, readonly and shoud be modified
        if (not was_not_alive_or_nil) && (not options[:read_only]) && (not was_writable) && options[:use_readonly_excel]
          open(file, :force_excel => book.excel, :if_obstructed => :new_excel, :read_only => true)
        end
        book.excel.visible = old_visible
        book.close if (was_not_alive_or_nil && (not opts[:keep_open]))
      end
    end

    # returns the contents of a range or cell with given name
    def nvalue(name)
      begin
        item = self.Names.Item(name)
      rescue WIN32OLERuntimeError
        raise ExcelErrorNValue, "name #{name} not in #{File.basename(self.stored_filename)}"  
      end
      begin
        referstorange = item.RefersToRange
      rescue WIN32OLERuntimeError
        raise ExcelErrorNValue, "range error in #{File.basename(self.stored_filename)}"      
      end
      begin
        value = referstorange.Value
      rescue WIN32OLERuntimeError
        raise ExcelErrorNValue, "value error in #{File.basename(self.stored_filename)}" 
      end
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

    def writable
      (not @workbook.ReadOnly) if @workbook
    end

    def saved
      @workbook.Saved if @workbook
    end

    # returns true, if the full book names and excel appications are identical, false otherwise  
    def == other_book
      other_book.is_a?(Book) &&
      @excel == other_book.excel &&
      self.filename == other_book.filename  
    end
 
    # saves a book.
    # returns true, if successfully saved, nil otherwise
    def save
      raise ExcelErrorSave, "Not opened for writing (opened with :read_only option)" if @options[:read_only]
      if @workbook then
        begin
          @workbook.Save 
        rescue WIN32OLERuntimeError => msg
          if msg.message =~ /SaveAs/ and msg.message =~ /Workbook/ then
            raise ExcelErrorSave, "workbook not saved"
          else
            raise ExcelErrorSaveUnknown, "unknown WIN32OELERuntimeError:\n#{msg.message}"
          end       
        end
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
        book_store.store(self)
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

    def self.book_store
      @@bookstore ||= BookStore.new
    end

    def book_store
      self.class.book_store
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

  class ExcelErrorNValue < WIN32OLERuntimeError # :nodoc: #
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

end