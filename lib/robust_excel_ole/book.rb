
# -*- coding: utf-8 -*-

require 'weakref'

module RobustExcelOle

  class Book
    attr_accessor :excel
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
      # :if_locked       if the book is open in another Excel instance, then
      #                   :readonly (default) -> open the book as readonly, if is should be opened in a new Excel,
      #                                          use the old book otherwise
      #                   :take_writable      -> use the Excel instance in which the book is writable,
      #                                          if such an Excel instance exists
      #                   :force_writability  -> make it writable in the desired Excel
      # :if_locked_unsaved  if the book is open in another Excel instance and contains unsaved changes
      #                  :raise    -> raise an exception
      #                  :save     -> save the unsaved book 
      #                  (not implemented yet)
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
        @options = {
          :excel => :reuse,
          :default_excel => :reuse,
          :if_locked     => :readonly,       
          :if_unsaved    => :raise,
          :if_obstructed => :raise,
          :read_only => false
        }.merge(opts)
        #self.set_defaults(opts) ???
        book = nil
        if (not (@options[:force_excel] == :new && (not @options[:if_locked] == :take_writable)))
          book = book_store.fetch(file, :readonly_excel => (@options[:read_only] ? @options[:force_excel] : nil)) rescue nil
          if book
            if (not @options[:force_excel] || (@options[:force_excel] == book.excel))
              if book.excel.alive?
                # condition: :if_unsaved is not set or :accept or workbook is not unsaved
                if_unsaved_not_set_or_accept_or_workbook_saved = (@options[:if_unsaved] == :accept || @options[:if_unsaved] == :raise || (not book.workbook) || book.workbook.Saved)
                if ((not book.alive?) || if_unsaved_not_set_or_accept_or_workbook_saved)  
                  book.set_defaults(opts)
                  # reopen the book
                  book.get_workbook          
                end
                return book if book.alive? && if_unsaved_not_set_or_accept_or_workbook_saved
              end
            end
          end
        end
        @options[:excel] = @options[:force_excel] ? @options[:force_excel] : @options[:default_excel]
        new(file, @options, &block)
      end
    end

    def initialize(file, opts={ }, &block)
      raise ExcelErrorOpen, "file #{file} not found" unless File.exist?(file)
      set_defaults(opts)
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
  
    def set_defaults(opts)
      @options = {
        :excel => :reuse,
        :default_excel => :reuse,
        :if_locked     => :readonly,       
        :if_unsaved    => :raise,
        :if_obstructed => :raise,
        :read_only => false
      }.merge(opts)
    end
    
    def get_excel
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
            open_workbook
          when :save
            save unless @workbook.Saved
            @workbook.Close
            @workbook = nil
            open_workbook
          when :close_if_saved
            if (not @workbook.Saved) then
              raise ExcelErrorOpen, "book with the same name in a different path is unsaved"
            else 
              @workbook.Close
              @workbook = nil
              open_workbook
            end
          when :new_excel 
            @excel_options = {:displayalerts => false, :visible => false}.merge(@options)   
            @excel_options[:reuse] = false
            @excel = Excel.new(@excel_options)
            open_workbook
          else
            raise ExcelErrorOpen, ":if_obstructed: invalid option"
          end
        else
          # book open, not obstructed by an other book, but not saved
          if (not @workbook.Saved) then
            case @options[:if_unsaved]
            when :raise
              raise ExcelErrorOpen, "book is already open but not saved (#{File.basename(@file)})"
            when :forget
              @workbook.Close
              @workbook = nil
              open_workbook
            when :accept
              # do nothing
            when :alert
              @excel.with_displayalerts true do
                open_workbook
              end 
            when :new_excel
              @excel_options = {:displayalerts => false, :visible => false}.merge(@options)
              @excel_options[:reuse] = false
              @excel = Excel.new(@excel_options)
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
    end

    def open_workbook
      #p "open_workbook:"
      #p "@file:#{@file}"
      if ((not @workbook) || (@options[:if_unsaved] == :alert) || @options[:if_obstructed]) then
        begin
          filename = RobustExcelOle::absolute_path(@file)
          workbooks = @excel.Workbooks
          workbooks.Open(filename,{ 'ReadOnly' => @options[:read_only] })
          # workaround for bug in Excel 2010: workbook.Open does not always return 
          # the workbook with given file name
          @workbook = workbooks.Item(File.basename(filename))
        #rescue BookStoreError => e
        #  raise ExcelUserCanceled, "open: canceled by user: #{e}"
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
      if ((alive?) && (not @workbook.Saved) && (not @workbook.ReadOnly)) then
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

    # modify a book such that its state remains unchanged.
    #  options: :keep_open: let the book open after modification
    #  if the book is read_only and modified (unsaved), then
    #    only the saved version of the book is unobtrusively modified, 
    #    not the current changed version
    # returns the block result
    def self.unobtrusively(file, opts = {:keep_open => false})
      book = book_store.fetch(file)
      was_not_alive_or_nil = book.nil? || (not book.alive?)
      was_saved = was_not_alive_or_nil ? true : book.Saved
      was_readonly = was_not_alive_or_nil ? false : book.ReadOnly
      old_book = book if was_readonly
      begin
        book = was_not_alive_or_nil ? open(file, :if_obstructed => :new_excel) : 
               (was_readonly ? open(file, :force_excel => :new) : book)
        yield book
      ensure
        book.save if (was_not_alive_or_nil || was_saved || was_readonly) && (not book.Saved)
        if was_readonly
          book.close
          book = old_book
        end
        book.close if (was_not_alive_or_nil && (not opts[:keep_open]))
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
