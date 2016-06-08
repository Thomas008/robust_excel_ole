# -*- coding: utf-8 -*-

require 'weakref'

module RobustExcelOle

  class Book < REOCommon

    attr_accessor :excel
    attr_accessor :ole_workbook
    attr_accessor :stored_filename
    attr_accessor :options
    attr_accessor :can_be_closed

    alias ole_object ole_workbook

      DEFAULT_OPEN_OPTS = { 
        :excel => :reuse,
        :default_excel => :reuse,
        :if_lockraiseed => :readonly,       
        :if_unsaved    => :raise,
        :if_obstructed => :raise,
        :if_absent     => :raise,
        :read_only => false,
        :check_compatibility => true
      }

    class << self
      
      # opens a workbook.
      # @param [String] file the file name
      # @param [Hash] opts the options
      # @option opts [Variant] :default_excel  :reuse (default), :new, or <excel-instance>     
      # @option opts [Variant] :force_excel    :new (default), or <excel-instance>
      # @option opts [Symbol]  :if_unsaved     :raise (default), :forget, :accept, :alert, :excel, or :new_excel
      # @option opts [Symbol]  :if_obstructed  :raise (default), :forget, :save, :close_if_saved, or _new_excel
      # @option opts [Symbol]  :if_absent      :raise (default), or :create
      # @option opts [Boolean] :read_only      true (default), or false
      # @option opts [Boolean] :displayalerts  true, or false (default)
      # @option opts [Boolean] :visible        true, or false (default) 
      # options: 
      # :default_excel   if the workbook was already open in an Excel instance, then open it in the Excel instance.
      #                  where it was opened most recently
      #                  Otherwise, i.e. if the workbook was not open before or the Excel instance is not alive
      #                   :reuse           -> connects to a (the first opened) running Excel instance,
      #                                        excluding the hidden Excel instance, if it exists,
      #                                       otherwise opens in a new Excel instance.
      #                   :new             -> opens in a new Excel instance
      #                   <excel-instance> -> opens in the given Excel instance
      # :force_excel     no matter whether the workbook was already open
      #                   :new             -> opens in a new Excel instance
      #                   <excel-instance> -> opens in the given Excel instance
      # :if_unsaved     if an unsaved workbook with the same name is open, then
      #                  :raise               -> raises an exception
      #                  :forget              -> close the unsaved workbook, open the new workbook             
      #                  :accept              -> lets the unsaved workbook open                  
      #                  :alert or :excel     -> gives control to Excel
      #                  :new_excel           -> opens the new workbook in a new Excel instance
      # :if_obstructed  if a workbook with the same name in a different path is open, then
      #                  :raise               -> raises an exception 
      #                  :forget              -> closes the old workbook, open the new workbook
      #                  :save                -> saves the old workbook, close it, open the new workbook
      #                  :close_if_saved      -> closes the old workbook and open the new workbook, if the old workbook is saved,
      #                                          otherwise raises an exception.
      #                  :new_excel           -> opens the new workbook in a new Excel instance   
      # :if_absent       :raise               -> raises an exception     , if the file does not exists
      #                  :create              -> creates a new Excel file, if it does not exists  
      #                  
      # :read_only     opens in read-only mode         
      # :displayalerts enables DisplayAlerts in Excel  
      # :visible       makes visible in Excel        
      # :check_compatibility  check compatibility when saving
      # if :default_excel is set, then DisplayAlerts and Visible are set only if these parameters are given
      # @return [Book] a workbook
      def open(file, opts={ }, &block)
        options = DEFAULT_OPEN_OPTS.merge(opts)
        book = nil
        if (not (options[:force_excel] == :new && (not options[:if_locked] == :take_writable)))
          # if readonly is true, then prefer a book that is given in force_excel if this option is set
          book = bookstore.fetch(file, 
                  :prefer_writable => (not options[:read_only]), 
                  :prefer_excel    => (options[:read_only] ? excel_of(options[:force_excel]) : nil)) rescue nil
          if book
            if (((not options[:force_excel]) || (excel_of(options[:force_excel]) == book.excel)) &&
                 (not (book.alive? && (not book.saved) && (not options[:if_unsaved] == :accept))))
              book.options = DEFAULT_OPEN_OPTS.merge(opts)
              book.ensure_excel(options) unless book.excel.alive?
              # if the book is opened as readonly and should be opened as writable, then close it and open the book with the new readonly mode
              book.close if (book.alive? && (not book.writable) && (not options[:read_only]))
              # reopens the book
              book.ensure_workbook(file,options) unless book.alive?
              return book
            end
          end
        end
        new(file, options, &block)
      end
    end    

    # creates a Book object for a given workbook or file name
    # @param [WIN32OLE] workbook a workbook
    # @param [Hash] opts the options
    # @option opts [Symbol] see above
    # @return [Book] a workbook
    def self.new(workbook, opts={ }, &block)      
      if workbook && (workbook.is_a? WIN32OLE)
        filename = workbook.Fullname.tr('\\','/') rescue nil
        if filename
          book = bookstore.fetch(filename)
          if book && book.alive?
            book.apply_options(opts)
            return book 
          end
        end
      end
      super
    end

    # creates a new Book object, if a file name is given
    # Promotes the workbook to a Book object, if a win32ole-workbook is given    
    # @param [Variant] file_or_workbook  file name or workbook
    # @param [Hash]    opts              the options
    # @option opts [Symbol] see above
    # @return [Book] a workbook
    def initialize(file_or_workbook, opts={ }, &block)
      options = DEFAULT_OPEN_OPTS.merge(opts)
      options[:excel] = options[:force_excel] ? options[:force_excel] : options[:default_excel]
      if file_or_workbook.is_a? WIN32OLE
        workbook = file_or_workbook
        @ole_workbook = workbook        
        # use the Excel instance where the workbook is opened
        win32ole_excel = WIN32OLE.connect(workbook.Fullname).Application rescue nil   
        @excel = excel_class.new(win32ole_excel)     
        self.apply_options(options)       
        # if the Excel could not be promoted, then create it         
        ensure_excel(options)
      else
        file = file_or_workbook
        ensure_excel(options)
        ensure_workbook(file, options)
      end
      bookstore.store(self)
      if block
        begin
          yield self
        ensure
          close
        end
      end
    end

    def apply_options(options) # :nodoc: #
      @excel.visible = options[:visible] unless options[:visible].nil?
      @excel.displayalerts = options[:displayalerts] unless options[:displayalerts].nil? 
    end

  private

    # returns an Excel object when given Excel, Book or Win32ole object representing a Workbook or an Excel
    def self.excel_of(object)  # :nodoc: #
      if object.is_a? WIN32OLE
        case object.ole_obj_help.name
        when /Workbook/i 
          new(object).excel 
        when /Application/i
          excel_class.new(object)
        else
          object.excel
        end
      else
        begin
          object.excel
        rescue
          raise ExcelErrorOpen, "given object is neither an Excel, a Workbook, nor a Win32ole"
        end
      end
      #rescue
        # trace "no Excel, Book, or WIN32OLE object representing a Workbook or an Excel instance"
    end

  public

    def ensure_excel(options)   # :nodoc: #
      return if @excel && @excel.alive?
      if options[:excel] == :reuse
        @excel = excel_class.new(:reuse => true)
      end
      excel_options = nil
      if @excel 
        dead_or_recycled = begin
          (not @excel.alive?)
        rescue WeakRef::RefError => msg
          true
        end
      end
      if (not @excel) || dead_or_recycled
        if options[:excel] == :new || dead_or_recycled
          excel_options = {:displayalerts => false, :visible => false}.merge(options)
          excel_options[:reuse] = false
          @excel = excel_class.new(excel_options)
        else 
          @excel = self.class.excel_of(options[:excel])
        end
      end
      apply_options(options) unless excel_options
    end

    def ensure_workbook(file, options)     # :nodoc: #
      file = @stored_filename ? @stored_filename : file
      unless File.exist?(file)
        if options[:if_absent] == :create
          @ole_workbook = excel_class.current.generate_workbook(file)
        else 
          raise ExcelErrorOpen, "file #{file.inspect} not found"
        end
      end
      @ole_workbook = @excel.Workbooks.Item(File.basename(file)) rescue nil
      if @ole_workbook then
        obstructed_by_other_book = (File.basename(file) == File.basename(@ole_workbook.Fullname)) && 
                                   (not (General::absolute_path(file) == @ole_workbook.Fullname))
        # if book is obstructed by a book with same name and different path
        if obstructed_by_other_book then
          case options[:if_obstructed]
          when :raise
            raise ExcelErrorOpen, "blocked by a book with the same name in a different path: #{@ole_workbook.Fullname.tr('\\','/')}"
          when :forget
            @ole_workbook.Close
            @ole_workbook = nil
            open_or_create_workbook(file, options)
          when :save
            save unless @ole_workbook.Saved
            @ole_workbook.Close
            @ole_workbook = nil
            open_or_create_workbook(file, options)
          when :close_if_saved
            if (not @ole_workbook.Saved) then
              raise ExcelErrorOpen, "workbook with the same name in a different path is unsaved: #{@ole_workbook.Fullname.tr('\\','/')}"
            else 
              @ole_workbook.Close
              @ole_workbook = nil
              open_or_create_workbook(file, options)
            end
          when :new_excel 
            excel_options = {:displayalerts => false, :visible => false}.merge(options)   
            excel_options[:reuse] = false
            @excel = excel_class.new(excel_options)
            open_or_create_workbook(file, options)
          else
            raise ExcelErrorOpen, ":if_obstructed: invalid option: #{options[:if_obstructed].inspect}"
          end
        else
          # book open, not obstructed by an other book, but not saved and writable
          if (not @ole_workbook.Saved) then
            case options[:if_unsaved]
            when :raise
              raise ExcelErrorOpen, "workbook is already open but not saved: #{File.basename(file).inspect}"
            when :forget
              @ole_workbook.Close
              @ole_workbook = nil
              open_or_create_workbook(file, options)
            when :accept
              # do nothing
            when :alert, :excel
              @excel.with_displayalerts true do
                open_or_create_workbook(file,options)
              end 
            when :new_excel
              excel_options = {:displayalerts => false, :visible => false}.merge(options)
              excel_options[:reuse] = false
              @excel = excel_class.new(excel_options)
              open_or_create_workbook(file, options)
            else
              raise ExcelErrorOpen, ":if_unsaved: invalid option: #{options[:if_unsaved].inspect}"
            end
          end
        end
      else
        # book is not open
        open_or_create_workbook(file, options)
      end
    end

  private

    def open_or_create_workbook(file, options)   # :nodoc: #
      if ((not @ole_workbook) || (options[:if_unsaved] == :alert) || options[:if_obstructed]) then
        begin
          filename = General::absolute_path(file)
          begin
            workbooks = @excel.Workbooks
          rescue RuntimeError => msg
            trace "RuntimeError: #{msg.message}" 
            if msg.message =~ /method missing: Excel not alive/
              raise ExcelErrorOpen, "Excel instance not alive or damaged" 
            else
              raise ExcelErrorOpen, "unknown RuntimeError"
            end
          rescue WeakRef::RefError => msg
            trace "WeakRefError: #{msg.message}"
            raise ExcelErrorOpen, "#{msg.message}"
          end
          # workaround for linked workbooks for Excel 2007: 
          # opening and closing a dummy workbook if Excel has no workbooks.
          # delay: with visible: 0.2 sec, without visible almost none
          count = workbooks.Count
          workbooks.Add if @excel.Version == "12.0" && count == 0
          workbooks.Open(filename,{ 'ReadOnly' => options[:read_only] })
          workbooks.Item(1).Close if @excel.Version == "12.0" && count == 0
          workbooks.Item(1).CheckCompatibility = options[:check_compatibility]
          @can_be_closed = false if @can_be_closed.nil?
        rescue WIN32OLERuntimeError => msg
          trace "WIN32OLERuntimeError: #{msg.message}" 
          if msg.message =~ /800A03EC/
            raise ExcelErrorOpen, "open: user canceled or open error"
          else 
            raise ExcelErrorOpen, "unknown WIN32OLERuntimeError"
          end
        end   
        begin
          # workaround for bug in Excel 2010: workbook.Open does not always return 
          # the workbook with given file name
          @ole_workbook = workbooks.Item(File.basename(filename))
        rescue WIN32OLERuntimeError
          raise ExcelErrorOpen, "cannot find the file #{File.basename(filename).inspect}"
        end
      end
    end

  public

    # closes the workbook, if it is alive
    # @param [Hash] opts the options
    # @option opts [Symbol] :if_unsaved :raise (default), :save, :forget, :keep_open, or :alert
    # options:
    #  :if_unsaved    if the workbook is unsaved
    #                      :raise           -> raises an exception       
    #                      :save            -> saves the workbook before it is closed                  
    #                      :forget          -> closes the workbook 
    #                      :keep_open       -> keep the workbook open
    #                      :alert or :excel -> gives control to excel
    # @raise ExcelErrorClose if the option :if_unsaved is :raise and the workbook is unsaved, or option is invalid
    # @raise ExcelErrorCanceled if the user has canceled 
    def close(opts = {:if_unsaved => :raise})
      if (alive? && (not @ole_workbook.Saved) && writable) then
        case opts[:if_unsaved]
        when :raise
          raise ExcelErrorClose, "workbook is unsaved: #{File.basename(self.stored_filename).inspect}"
        when :save
          save
          close_workbook
        when :forget
          close_workbook
        when :keep_open
          # nothing
        when :alert, :excel
          @excel.with_displayalerts true do
            close_workbook
          end
        else
          raise ExcelErrorClose, ":if_unsaved: invalid option: #{opts[:if_unsaved].inspect}"
        end
      else
        close_workbook
      end
      raise ExcelUserCanceled, "close: canceled by user" if alive? && 
      (opts[:if_unsaved] == :alert || opts[:if_unsaved] == :excel) && (not @ole_workbook.Saved)
    end

  private

    def close_workbook    
      @ole_workbook.Close if alive?
      @ole_workbook = nil unless alive?
    end

  public

    def self.for_reading(*args, &block)
      args = args.dup
      opts = args.last.is_a?(Hash) ? args.pop : {}
      opts = {:read_only => true}.merge(opts)
      args.push opts
      unobtrusively(*args, &block)
    end

    def self.for_modifying(*args, &block)
      args = args.dup
      opts = args.last.is_a?(Hash) ? args.pop : {}
      opts = {:read_only => false}.merge(opts)
      args.push opts
      unobtrusively(*args, &block)
    end

    # modifies a workbook such that its state (open/close, saved/unsaved, readonly/writable) remains unchanged
    # @param [String] file        the file name
    # @param [Hash]   if_closed   an option
    # @param [Hash]   opts        the options
    # @option opts [Variant] :if_closed  :reuse (default), :hidden or a Excel instance
    # @option opts [Boolean] :read_only whether the file is opened for read-only
    # @option opts [Boolean] :readonly_excel behaviour when workbook is opened read-only and shall be modified
    # @option opts [Boolean] :keep_open whether the workbook shall be kept open after unobtrusively opening
    # @option opts [Boolean] :displayalerts  true, or false (default)
    # @option opts [Boolean] :visible        true, or false (default) 
    #  options: 
    #   :if_closed :   if the workbook is closed, then open it in
    #                    :reuse  -> the Excel instance of the workbook, if it exists, 
    #                               reuse another Excel, otherwise          
    #                    :hidden -> a separate Excel instance that is not visible and has no displayaslerts
    #                    <excel-instance> -> the given Excel instance
    #  :read_only        : opens the workbook unobtrusively for reading only  (default: false)
    #  :readonly_excel:  if the workbook is opened only as ReadOnly and shall be modified, then
    #                    true:  closes it and open it as writable in the Excel instance where it was open so far
    #                    false (default)   opens it as writable in another running excel instance, if it exists,
    #                                      otherwise open in a new Excel instance.
    # :displayalerts      enables DisplayAlerts in Excel  
    # :visible            makes visible in Excel 
    # :check_compatibility checks compatibility when saving
    # @return [Book] a workbook
    def self.unobtrusively(file, if_closed = nil, opts = { }, &block) 
      if if_closed.is_a? Hash
        opts = if_closed
        if_closed = nil
      end
      if_closed = :reuse unless if_closed
      options = {
        :read_only => false,
        :readonly_excel => false,
        :keep_open => false,
        :check_compatibility => true
      }.merge(opts)
      book = bookstore.fetch(file, :prefer_writable => (not options[:read_only]))
      was_not_alive_or_nil = book.nil? || (not book.alive?)
      workbook = book.excel.Workbooks.Item(File.basename(file)) rescue nil
      now_alive = 
        begin 
          workbook.Name
          true
        rescue 
          false
        end
      was_saved = was_not_alive_or_nil ? true : book.saved
      was_writable = book.writable unless was_not_alive_or_nil      
      begin 
        book = 
          if was_not_alive_or_nil 
            case if_closed
            when :reuse
              open(file, :read_only => options[:read_only])
            when :hidden 
              open(file, :force_excel => bookstore.hidden_excel, :read_only => options[:read_only])
            else 
              open(file, :force_excel => if_closed, :read_only => options[:read_only])
            end
          else
            if was_writable || options[:read_only]
              book
            else
              options[:readonly_excel] ? open(file, :force_excel => book.excel, :read_only => options[:read_only]) : 
                                         open(file, :force_excel => :new, :read_only => options[:read_only])
            end
          end
        book.excel.displayalerts = options[:displayalerts] unless options[:displayalerts].nil?
        book.excel.visible = options[:visible] unless options[:visible].nil?
        old_check_compatibility = book.CheckCompatibility
        book.CheckCompatibility = options[:check_compatibility]
        yield book
      ensure
        was_saved_or_appeared = was_saved || was_not_alive_or_nil || (not was_writable)
        book.save if book && (not book.saved) && (not options[:read_only]) && was_saved_or_appeared
        # book was open, readonly and shoud be modified
        if (not was_not_alive_or_nil) && (not options[:read_only]) && (not was_writable) && options[:readonly_excel]
          open(file, :force_excel => book.excel, :if_obstructed => :new_excel, :read_only => true)
        end
        @can_be_closed = true if options[:keep_open] && book
        book.close if (was_not_alive_or_nil && (not now_alive) && (not options[:keep_open]) && book)
        book.CheckCompatibility = old_check_compatibility if book && book.alive?
      end
    end

    # reopens a closed workbook
    def reopen
      self.class.open(self.stored_filename)
    end

    # simple save of a workbook.
    # @raise ExcelErrorSave if workbook is not alive or opened for read-only, or another error occurs
    # @return [Boolean] true, if successfully saved, nil otherwise
    def save      
      raise ExcelErrorSave, "Workbook is not alive" if (not alive?)
      raise ExcelErrorSave, "Not opened for writing (opened with :read_only option)" if @ole_workbook.ReadOnly
      begin
        @ole_workbook.Save 
      rescue WIN32OLERuntimeError => msg
        if msg.message =~ /SaveAs/ and msg.message =~ /Workbook/ then
          raise ExcelErrorSave, "workbook not saved"
        else
          raise ExcelErrorSaveUnknown, "unknown WIN32OELERuntimeError:\n#{msg.message}"
        end       
      end
      true
    end

    # saves a workbook with a given file name.
    # @param [String] file   file name
    # @param [Hash]   opts   the options
    # @option opts [Symbol] :if_exists      :raise (default), :overwrite, or :alert, :excel
    # @option opts [Symbol] :if_obstructed  :raise (default), :forget, :save, or :close_if_saved
    # options: 
    # :if_exists  if a file with the same name exists, then  
    #               :raise     -> raises an exception, dont't write the file  (default)
    #               :overwrite -> writes the file, delete the old file
    #               :alert or :excel -> gives control to Excel
    #  :if_obstructed   if a workbook with the same name and different path is already open and blocks the saving, then
    #                  :raise               -> raises an exception 
    #                  :forget              -> closes the blocking workbook
    #                  :save                -> saves the blocking workbook and closes it
    #                  :close_if_saved      -> closes the blocking workbook, if it is saved, 
    #                                          otherwise raises an exception
    # @raise ExcelErrorSave if workbook is not alive, opened in read-only mode, invalid options,
    #                          the file already exists (with option :if_exists :raise),
    #                          the workbook is blocked by another one (with option :if_obstructed :raise)
    # @return [Book], the book itself, if successfully saved, raises an exception otherwise
    def save_as(file = nil, opts = { } )
      raise ExcelErrorSave, "Workbook is not alive" if (not alive?)
      raise ExcelErrorSave, "Not opened for writing (opened with :read_only option)" if @ole_workbook.ReadOnly
      options = {
        :if_exists => :raise,
        :if_obstructed => :raise,
      }.merge(opts)
      if File.exist?(file) then
        case options[:if_exists]
        when :overwrite
          if file == self.filename
            save
            return self
          else
            begin
              File.delete(file)
            rescue Errno::EACCES
              raise ExcelErrorSave, "workbook is open and used in Excel"
            end
          end
        when :alert, :excel 
          @excel.with_displayalerts true do
            save_as_workbook(file, options)
          end
          return self
        when :raise
          raise ExcelErrorSave, "file already exists: #{File.basename(file).inspect}"
        else
          raise ExcelErrorSave, ":if_exists: invalid option: #{options[:if_exists].inspect}"
        end
      end
      blocking_workbook = 
        begin
          @excel.Workbooks.Item(File.basename(file))
        rescue WIN32OLERuntimeError => msg
          nil
        end
      if blocking_workbook then
        case options[:if_obstructed]
        when :raise
          raise ExcelErrorSave, "blocked by another workbook: #{blocking_workbook.Fullname.tr('\\','/')}"
        when :forget
          # nothing
        when :save
          blocking_workbook.Save
        when :close_if_saved
          raise ExcelErrorSave, "blocking workbook is unsaved: #{File.basename(file).inspect}" unless blocking_workbook.Saved
        else
          raise ExcelErrorSave, ":if_obstructed: invalid option: #{options[:if_obstructed].inspect}"
        end
        blocking_workbook.Close
      end
      save_as_workbook(file, options)
      self
    end

  private

    def save_as_workbook(file, options)   # :nodoc: #
      begin
        dirname, basename = File.split(file)
        file_format =
          case File.extname(basename)
            when '.xls' ; RobustExcelOle::XlExcel8
            when '.xlsx'; RobustExcelOle::XlOpenXMLWorkbook
            when '.xlsm'; RobustExcelOle::XlOpenXMLWorkbookMacroEnabled
          end
        @ole_workbook.SaveAs(General::absolute_path(file), file_format)
        bookstore.store(self)
      rescue WIN32OLERuntimeError => msg
        if msg.message =~ /SaveAs/ and msg.message =~ /Workbook/ then
          if options[:if_exists] == :alert || options[:if_exists] == :excel then 
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

    # returns a sheet, if a sheet name or a number is given
    # @param [String] or [Number]
    # @returns [Sheet]
    def sheet(name)
      begin
        sheet_class.new(@ole_workbook.Worksheets.Item(name))
      rescue WIN32OLERuntimeError => msg
        raise ExcelError, "could not return a sheet with name #{name.inspect}"
        trace "#{msg.message}"
      end
    end    

    def each
      @ole_workbook.Worksheets.each do |sheet|
        yield sheet_class.new(sheet)
      end
    end
  
    # copies a sheet to another position
    # default: copied sheet is appended
    # @param [Sheet] sheet a sheet that shall be copied
    # @param [Hash]  opts  the options
    # @option opts [Symbol] :as     new name of the copied sheet
    # @option opts [Symbol] :before a sheet before which the sheet shall be inserted
    # @option opts [Symbol] :after  a sheet after which the sheet shall be inserted
    # @raise  ExcelErrorSheet if the sheet name already exists
    # @return [Sheet] the copied sheet
    def copy_sheet(sheet, opts = { })
      new_sheet_name = opts.delete(:as)
      after_or_before, base_sheet = opts.to_a.first || [:after, last_sheet]
      sheet.Copy({ after_or_before.to_s => base_sheet.worksheet })
      new_sheet = sheet_class.new(@excel.Activesheet)
      begin
        new_sheet.name = new_sheet_name if new_sheet_name
      rescue WIN32OLERuntimeError => msg
        msg.message =~ /800A03EC/ ? raise(ExcelErrorSheet, "sheet name already exists") : raise(ExcelErrorSheetUnknown)
      end
      new_sheet
    end      

    # adds an empty sheet
    # default: empty sheet is appended
    # @param [Hash]  opts  the options
    # @option opts [Symbol] :as     new name of the copied added sheet
    # @option opts [Symbol] :before a sheet before which the sheet shall be inserted
    # @option opts [Symbol] :after  a sheet after which the sheet shall be inserted
    # @raise  ExcelErrorSheet if the sheet name already exists
    # @return [Sheet] the added sheet
    def add_empty_sheet(opts = { })
      new_sheet_name = opts.delete(:as)
      after_or_before, base_sheet = opts.to_a.first || [:after, last_sheet]
      @ole_workbook.Worksheets.Add({ after_or_before.to_s => base_sheet.worksheet })
      new_sheet = sheet_class.new(@excel.Activesheet)
      begin
        new_sheet.name = new_sheet_name if new_sheet_name
      rescue WIN32OLERuntimeError => msg
        msg.message =~ /800A03EC/ ? raise(ExcelErrorSheet, "sheet name already exists") : raise(ExcelErrorSheetUnknown)
      end
      new_sheet
    end    

    # copies a sheet to another position if a sheet is given, or adds an empty sheet
    # default: copied or empty sheet is appended, i.e. added behind the last sheet
    # @param [Sheet] sheet a sheet that shall be copied (optional)
    # @param [Hash]  opts  the options
    # @option opts [Symbol] :as     new name of the copied or added sheet
    # @option opts [Symbol] :before a sheet before which the sheet shall be inserted
    # @option opts [Symbol] :after  a sheet after which the sheet shall be inserted
    # @raise  ExcelErrorSheet if the sheet name already exists
    # @return [Sheet] the copied or added sheet
    def add_or_copy_sheet(sheet = nil, opts = { })
      if sheet.is_a? Hash
        opts = sheet
        sheet = nil
      end
      sheet ? copy_sheet(sheet, opts) : add_empty_sheet(opts)
    end      

    # for compatibility to older versions
    def add_sheet(sheet = nil, opts = { })
      add_or_copy_sheet(sheet, opts)
    end 

    def last_sheet
      sheet_class.new(@ole_workbook.Worksheets.Item(@ole_workbook.Worksheets.Count))
    end

    def first_sheet
      sheet_class.new(@ole_workbook.Worksheets.Item(1))
    end

    # returns the value of a range
    # @param [String] name the name of a range
    # @returns [Variant] the value of the range
    def [] name
      nameval(name)
    end

    # sets the value of a range
    # @param [String]  name  the name of the range
    # @param [Variant] value the contents of the range
    def []= (name, value)
      set_nameval(name,value)
    end

    # returns the contents of a range with given name
    # evaluates formula contents of the range is a formula
    # if no contents could be returned, then return default value, if provided, raise error otherwise
    # @param  [String]      name      the name of the range
    # @param  [Hash]        opts      the options
    # @option opts [Symbol] :default  the default value that is provided if no contents could be returned
    # @raise  ExcelError if range name is not in the workbook or if range value could not be evaluated
    # @return [Variant] the contents of a range with given name
    def nameval(name, opts = {:default => nil})
      begin
        name_obj = self.Names.Item(name)
      rescue WIN32OLERuntimeError
        return opts[:default] if opts[:default]
        raise ExcelError, "name #{name.inspect} not in #{File.basename(self.stored_filename).inspect}"
      end
      begin
        value = name_obj.RefersToRange.Value
      rescue  WIN32OLERuntimeError
        begin
          value = self.sheet(1).Evaluate(name_obj.Name)
        rescue WIN32OLERuntimeError
          return opts[:default] if opts[:default]
          raise SheetError, "cannot evaluate name #{name.inspect} in #{File.basename(self.stored_filename).inspect}"
        end
      end
      if value == -2146826259
        return opts[:default] if opts[:default]
        raise SheetError, "cannot evaluate name #{name.inspect} in #{File.basename(self.stored_filename).inspect}"
      end 
      return opts[:default] if (value.nil? && opts[:default])
      value      
    end

    # sets the contents of a range
    # @param [String]  name  the name of a range
    # @param [Variant] value the contents of the range
    # @raise ExcelError if range name is not in the workbook or if value could not be assigned to range
    def set_nameval(name, value) 
      begin
        name_obj = self.Names.Item(name)
      rescue WIN32OLERuntimeError
        raise ExcelError, "name #{name.inspect} not in #{File.basename(self.stored_filename).inspect}"  
      end
      begin
        name_obj.RefersToRange.Value = value
      rescue WIN32OLERuntimeError
        raise ExcelError, "cannot assign value to range named #{name.inspect} in #{File.basename(self.stored_filename).inspect}"    
      end
    end    

    # renames a range
    # @param [String] name     the previous range name
    # @param [String] new_name the new range name
    # @raise ExcelError if name is not in the file, or if new_name cannot be set
    def rename_range(name, new_name)
      begin
        item = self.Names.Item(name)
      rescue WIN32OLERuntimeError
        raise ExcelError, "name #{name.inspect} not in #{File.basename(self.stored_filename).inspect}"  
      end
      begin
        item.Name = new_name
      rescue WIN32OLERuntimeError
        raise ExcelError, "name error in #{File.basename(self.stored_filename).inspect}"      
      end
    end

    # brings workbook to foreground, makes it available for heyboard inputs, makes the Excel instance visible
    # @raise ExcelError if workbook cannot be activated    
    def activate      
      @excel.visible = true
      begin
        Win32API.new("user32","SetForegroundWindow","I","I").call(@excel.hwnd)     # Excel  2010
        @ole_workbook.Activate   # Excel 2007
      rescue WIN32OLERuntimeError
        raise ExcelError, "cannot activate"
      end
    end

    # returns true, if the workbook is visible, false otherwise 
    def visible
      @excel.Windows(@ole_workbook.Name).Visible
    end

    # makes a workbook visible or invisible
    # @param [Boolean] visible_value value that determines whether the workbook shall be visible
    def visible= visible_value
      saved = @ole_workbook.Saved
      @excel.Windows(@ole_workbook.Name).Visible = visible_value
      save if saved 
    end

    # returns true, if the workbook reacts to methods, false otherwise
    def alive?
      begin 
        @ole_workbook.Name
        true
      rescue 
        @ole_workbook = nil  # dead object won't be alive again
        #t $!.message
        false
      end
    end

    # returns the full file name of the workbook
    def filename
      @ole_workbook.Fullname.tr('\\','/') rescue nil
    end

    def writable   # :nodoc: #
      (not @ole_workbook.ReadOnly) if @ole_workbook
    end

    def saved   # :nodoc: #
      @ole_workbook.Saved if @ole_workbook
    end

    # @return [Boolean] true, if the full book names and excel Instances are identical, false otherwise  
    def == other_book
      other_book.is_a?(Book) &&
      @excel == other_book.excel &&
      self.filename == other_book.filename  
    end

    def self.books
      bookstore.books
    end

    def self.bookstore   # :nodoc: #
      @@bookstore ||= Bookstore.new
    end

    def bookstore    # :nodoc: #
      self.class.bookstore
    end   

    def self.show_books   # :nodoc: #
      bookstore.books
    end

    def to_s    # :nodoc: #
      "#{self.filename}"
    end

    def inspect    # :nodoc: #
      "#<Book: " + "#{"not alive " unless alive?}" + "#{File.basename(self.filename) if alive?}" + " #{@ole_workbook} #{@excel}"  + ">"
    end

    def self.in_context(klass)  # :nodoc: #
      
    end

    def self.excel_class    # :nodoc: #
      @excel_class ||= begin
        module_name = self.parent_name
        "#{module_name}::Excel".constantize
      rescue NameError => e
        #trace "excel_class: NameError: #{e}"
        Excel
      end
    end

    def self.sheet_class    # :nodoc: #
      @sheet_class ||= begin
        module_name = self.parent_name
        "#{module_name}::Sheet".constantize
      rescue NameError => e
        Sheet
      end
    end

    def excel_class        # :nodoc: #
      self.class.excel_class
    end

    def sheet_class        # :nodoc: #
      self.class.sheet_class
    end

    include MethodHelpers

  private

    def method_missing(name, *args)   # :nodoc: #
      if name.to_s[0,1] =~ /[A-Z]/ 
        begin
          raise ExcelError, "method missing: workbook not alive" unless alive?
          @ole_workbook.send(name, *args)
        rescue WIN32OLERuntimeError => msg
          if msg.message =~ /unknown property or method/
            raise VBAMethodMissingError, "unknown VBA property or method #{name.inspect}"
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

  Workbook = Book

  class ExcelError < RuntimeError    # :nodoc: #
  end

  class ExcelErrorOpen < ExcelError   # :nodoc: #
  end

  class ExcelErrorClose < ExcelError    # :nodoc: #
  end

  class ExcelErrorSave < ExcelError   # :nodoc: #
  end

  class ExcelErrorSaveFailed < ExcelErrorSave  # :nodoc: #
  end

  class ExcelErrorSaveUnknown < ExcelErrorSave  # :nodoc: #
  end

  class ExcelUserCanceled < RuntimeError # :nodoc: #
  end
  
  class ExcelErrorSheet < ExcelError    # :nodoc: #
  end

  class ExcelErrorSheetUnknown < ExcelErrorSheet    # :nodoc: #
  end

end
