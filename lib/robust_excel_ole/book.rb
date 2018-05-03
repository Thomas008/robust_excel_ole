# -*- coding: utf-8 -*-

require 'weakref'

module RobustExcelOle

  class Book < REOCommon

    attr_accessor :excel
    attr_accessor :ole_workbook
    attr_accessor :stored_filename
    attr_accessor :options    
    attr_accessor :modified_cells

    alias ole_object ole_workbook

    DEFAULT_OPEN_OPTS = { 
      :default => {:excel => :current},
      :force => {},      
      :if_unsaved    => :raise,
      :if_obstructed => :raise,
      :if_absent     => :raise,
      :read_only => false,
      :check_compatibility => false,       
      :update_links => :never
    }

    SYNONYMS_OPTS = [[[:default],[:d]], [[:force], [:f]], 
                     [[:default,:excel],[:default_excel]],[[:force,:excel],[:force_excel]],                    
                     [[:default,:excel],[:default,:e]], [[:force,:excel],[:force,:e]],
                     [[:default,:visible],[:default,:v]], [[:force,:visible],[:force,:v]],
                     [[:force,:visible],[:visible]], [[:force,:visible],[:v]]
    ]

    class << self
      
      # opens a workbook.
      # @param [String] file the file name
      # @param [Hash] opts the options
      # @option opts [Hash] :default or :d
      # @option opts [Hash] :force or :f
      # @option opts [Symbol]  :if_unsaved     :raise (default), :forget, :accept, :alert, :excel, or :new_excel
      # @option opts [Symbol]  :if_obstructed  :raise (default), :forget, :save, :close_if_saved, or _new_excel
      # @option opts [Symbol]  :if_absent      :raise (default) or :create
      # @option opts [Boolean] :read_only      true (default) or false
      # @option opts [Boolean] :update_links   :never (default), :always, :alert
      # @option opts [Boolean] :calculation    :manual, :automatic, or nil (default) 
      # options: 
      # :default : if the workbook was already open before, then use (unchange) its properties,
      #            otherwise, i.e. if the workbook cannot be reopened, use the properties stated in :default
      # :force   : no matter whether the workbook was already open before, use the properties stated in :force 
      # :default and :force contain: :excel, :visible
      #  :excel   :current (or :active or :reuse) 
      #                    -> connects to a running (the first opened) Excel instance,
      #                       excluding the hidden Excel instance, if it exists,
      #                       otherwise opens in a new Excel instance.
      #           :new     -> opens in a new Excel instance 
      #           <excel-instance> -> opens in the given Excel instance
      #  :visible true, false, or nil (default)
      #  alternatives: :default_excel, :force_excel, :visible, :d, :f, :e, :v
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
      # :read_only            true -> opens in read-only mode               
      # :visible              true -> makes the workbook visible
      # :check_compatibility  true -> check compatibility when saving
      # :update_links         true -> user is being asked how to update links, false -> links are never updated
      # @return [Book] a workbook
      def open(file, opts={ }, &block)
        options = process_options(opts)
        book = nil
        if (not (options[:force][:excel] == :new))
          # if readonly is true, then prefer a book that is given in force_excel if this option is set
          forced_excel = if options[:force][:excel]
            options[:force][:excel] == :current ? excel_class.new(:reuse => true) : excel_of(options[:force][:excel])
          end
          book = bookstore.fetch(file, 
                  :prefer_writable => (not options[:read_only]), 
                  :prefer_excel    => (options[:read_only] ? forced_excel : nil)) rescue nil
          if book
            if (((not options[:force][:excel]) || (forced_excel == book.excel)) &&
                 (not (book.alive? && (not book.saved) && (not options[:if_unsaved] == :accept))))
              book.options = options
              book.ensure_excel(options) # unless book.excel.alive?
              # if the book is opened as readonly and should be opened as writable, then close it and open the book with the new readonly mode
              book.close if (book.alive? && (not book.writable) && (not options[:read_only]))
              # reopens the book
              book.ensure_workbook(file,options) unless book.alive?
              book.visible = options[:force][:visible] unless options[:force][:visible].nil?
              book.excel.calculation = options[:calculation] unless options[:calculation].nil?
              return book
            end
          end
        end
        new(file, options, &block)
      end
    end    

    # creates a Book object by opening an Excel file given its filename workbook or 
    # by lifting a Win32OLE object representing an Excel file
    # @param [WIN32OLE] workbook a workbook
    # @param [Hash] opts the options
    # @option opts [Symbol] see above
    # @return [Book] a workbook
    def self.new(workbook, opts={ }, &block)
      if workbook && (workbook.is_a? WIN32OLE)
        opts = process_options(opts)
        filename = workbook.Fullname.tr('\\','/') rescue nil
        if filename
          book = bookstore.fetch(filename)
          if book && book.alive?
            book.visible = opts[:force][:visible] unless opts[:force][:visible].nil?
            #book.excel.calculation = opts[:calculation].nil? ? book.excel.calculation : opts[:calculation]
            book.excel.calculation = opts[:calculation] unless opts[:calculation].nil?
            return book 
          end
        end
      else
        super
      end
    end

    # creates a new Book object, if a file name is given
    # Promotes the workbook to a Book object, if a win32ole-workbook is given    
    # @param [Variant] file_or_workbook  file name or workbook
    # @param [Hash]    opts              the options
    # @option opts [Symbol] see above
    # @return [Book] a workbook
    def initialize(file_or_workbook, opts={ }, &block)
      options = self.class.process_options(opts)
      if file_or_workbook.is_a? WIN32OLE        
        workbook = file_or_workbook
        @ole_workbook = workbook        
        # use the Excel instance where the workbook is opened
        win32ole_excel = WIN32OLE.connect(workbook.Fullname).Application rescue nil   
        @excel = excel_class.new(win32ole_excel)     
        @excel.visible = options[force][:visible] unless options[:force][:visible].nil? 
        @excel.calculation = options[:calculation] unless options[:calculation].nil?
        ensure_excel(options)
      else
        file = file_or_workbook
        ensure_excel(options)
        ensure_workbook(file, options)        
      end
      bookstore.store(self)
      @modified_cells = []
      if block
        begin
          yield self
        ensure
          close
        end
      end
    end

  private

    # merges options with defaults and translates abbreviations and synonyms
    def self.process_options(options) # :nodoc: #
      translator = proc do |opts|
        SYNONYMS_OPTS.each do |a|
          synonym = a[1][1].nil? ? opts[a[1][0]] : opts[a[1][0]][a[1][1]] unless opts[a[1][0]].nil?
          unless synonym.nil?
            if a[0][1].nil?
              opts[a[0][0]] = synonym if opts[a[0][0]].nil?
            else
              opts[a[0][0]] = {} if opts[a[0][0]].nil?
              opts[a[0][0]][a[0][1]] = synonym if opts[a[0][0]][a[0][1]].nil?
            end
          end          
        end
        opts[:default][:excel] = :current if (not opts[:default].nil?) && (opts[:default][:excel] == :reuse || opts[:default][:excel] == :active)
        opts[:force][:excel] = :current if (not opts[:force].nil?) && (opts[:force][:excel] == :reuse || opts[:force][:excel] == :active)        
        opts
      end
      default_opts = translator.call(DEFAULT_OPEN_OPTS)
      given_opts = translator.call(options)
      opts = default_opts.merge(given_opts)
      opts[:default] = default_opts[:default].merge(opts[:default]) unless opts[:default].nil?
      opts[:force] = default_opts[:force].merge(opts[:force]) unless opts[:force].nil?
      opts
    end

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
          raise TypeErrorREO, "given object is neither an Excel, a Workbook, nor a Win32ole"
        end
      end
    end

  public

=begin
    # work in progress#
    def self.open_in_current_excel(file, opts = { }) # :nodoc: # 
      options = DEFAULT_OPEN_OPTS.merge(opts)
      filename = General::absolute_path(file)
      ole_workbook = WIN32OLE.connect(filename)
      workbook = Book.new(ole_workbook)
      workbook.visible = options[:force][:visible] unless options[:force][:visible].nil?
      update_links_opt =
            case options[:update_links]
            when :alert; RobustExcelOle::XlUpdateLinksUserSetting
            when :never; RobustExcelOle::XlUpdateLinksNever
            when :always; RobustExcelOle::XlUpdateLinksAlways
            else RobustExcelOle::XlUpdateLinksNever
          end
      workbook.UpdateLinks = update_links_opt
      workbook.CheckCompatibility = options[:check_compatibility]
      workbook
    end
=end

    def ensure_excel(options)   # :nodoc: #
      if @excel && @excel.alive?
        @excel.created = false
        return
      end
      excel_option = (options[:force].nil? or options[:force][:excel].nil?) ? options[:default][:excel] : options[:force][:excel]
      @excel = self.class.excel_of(excel_option) unless (excel_option == :current || excel_option == :new)
      @excel = excel_class.new(:reuse => (excel_option == :current)) unless (@excel && @excel.alive?)

      #options[:excel] = options[:force_excel] ? options[:force_excel] : options[:default_excel]
      #options[:excel] = :current if (options[:excel] == :reuse || options[:excel] == :active)
      #@excel = self.class.excel_of(options[:excel]) unless (options[:excel] == :current || options[:excel] == :new)
      #@excel = excel_class.new(:reuse => (options[:excel] == :current)) unless (@excel && @excel.alive?)
    end    

    def ensure_workbook(file, options)     # :nodoc: #
      file = @stored_filename ? @stored_filename : file
      raise(FileNameNotGiven, "filename is nil") if file.nil?
      raise(FileNotFound, "file #{General::absolute_path(file).inspect} is a directory") if File.directory?(file)
      unless File.exist?(file)
        if options[:if_absent] == :create
          @ole_workbook = excel_class.current.generate_workbook(file)
        else 
          raise FileNotFound, "file #{General::absolute_path(file).inspect} not found"
        end
      end
      @ole_workbook = @excel.Workbooks.Item(File.basename(file)) rescue nil
      if @ole_workbook then
        obstructed_by_other_book = (File.basename(file) == File.basename(@ole_workbook.Fullname)) && 
                                   (not (General::absolute_path(file) == @ole_workbook.Fullname))
        # if workbook is obstructed by a workbook with same name and different path
        if obstructed_by_other_book then
          case options[:if_obstructed]
          when :raise
            raise WorkbookBlocked, "blocked by a workbook with the same name in a different path: #{@ole_workbook.Fullname.tr('\\','/')}"
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
              raise WorkbookBlocked, "workbook with the same name in a different path is unsaved: #{@ole_workbook.Fullname.tr('\\','/')}"
            else 
              @ole_workbook.Close
              @ole_workbook = nil
              open_or_create_workbook(file, options)
            end
          when :new_excel 
            @excel = excel_class.new(:reuse => false)
            open_or_create_workbook(file, options)
          else
            raise OptionInvalid, ":if_obstructed: invalid option: #{options[:if_obstructed].inspect}"
          end
        else
          # book open, not obstructed by an other book, but not saved and writable
          if (not @ole_workbook.Saved) then
            case options[:if_unsaved]
            when :raise
              raise WorkbookNotSaved, "workbook is already open but not saved: #{File.basename(file).inspect}"
            when :forget
              @ole_workbook.Close
              @ole_workbook = nil
              open_or_create_workbook(file, options)
            when :accept
              # do nothing
            when :alert, :excel
              @excel.with_displayalerts(true) { open_or_create_workbook(file,options) }
            when :new_excel
              @excel = excel_class.new(:reuse => false)
              open_or_create_workbook(file, options)
            else
              raise OptionInvalid, ":if_unsaved: invalid option: #{options[:if_unsaved].inspect}"
            end
          end
        end
      else
        # open a new workbook
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
          rescue WIN32OLERuntimeError => msg
            raise UnexpectedError, "WIN32OLERuntimeError: #{msg.message} #{msg.backtrace}"
          end
          begin
            with_workaround_linked_workbooks_excel2007(options) do
              workbooks.Open(filename, { 'ReadOnly' => options[:read_only] ,
                                         'UpdateLinks' => updatelinks_vba(options[:update_links]) })
            end
          rescue WIN32OLERuntimeError => msg
            # for Excel2007: for option :if_unsaved => :alert and user cancels: this error appears?
            # if yes: distinguish these events
            raise UnexpectedError, "WIN32OLERuntimeError: #{msg.message} #{msg.backtrace}"
          end
          begin
            # workaround for bug in Excel 2010: workbook.Open does not always return the workbook when given file name
            begin
              @ole_workbook = workbooks.Item(File.basename(filename))
            rescue WIN32OLERuntimeError => msg
              raise UnexpectedError, "WIN32OLERuntimeError: #{msg.message}"
            end
            if options[:force][:visible].nil? && (not options[:default][:visible].nil?)
              if @excel.created   
                self.visible = options[:default][:visible] 
              else
                self.window_visible = options[:default][:visible]
              end
            else
              self.visible = options[:force][:visible] unless options[:force][:visible].nil?
            end
            @ole_workbook.CheckCompatibility = options[:check_compatibility]
            @excel.calculation = options[:calculation] unless options[:calculation].nil?
            self.Saved = true # unless self.Saved # ToDo: this is too hard
          rescue WIN32OLERuntimeError => msg
            raise UnexpectedError, "WIN32OLERuntimeError: #{msg.message} #{msg.backtrace}"
          end       
        end
      end
    end

    # translating the option UpdateLinks from REO to VBA 
    # setting UpdateLinks works only if calculation mode is automatic,
    # parameter 'UpdateLinks' has no effect
    def updatelinks_vba(updatelinks_reo)
      case updatelinks_reo
      when :alert; RobustExcelOle::XlUpdateLinksUserSetting
      when :never; RobustExcelOle::XlUpdateLinksNever
      when :always; RobustExcelOle::XlUpdateLinksAlways
      else RobustExcelOle::XlUpdateLinksNever
      end
    end

    # workaround for linked workbooks for Excel 2007: 
    # opening and closing a dummy workbook if Excel has no workbooks.
    # delay: with visible: 0.2 sec, without visible almost none
    def with_workaround_linked_workbooks_excel2007(options)
      old_visible_value = @excel.Visible
      workbooks = @excel.Workbooks
      workaround_condition = @excel.Version.split(".").first.to_i == 12 && workbooks.Count == 0
      if workaround_condition
        workbooks.Add 
        @excel.calculation = options[:calculation].nil? ? @excel.calculation : options[:calculation] 
      end
      begin
        #@excel.with_displayalerts(update_links_opt == :alert ? true : @excel.displayalerts) do
        yield self
      ensure
        @excel.with_displayalerts(false){workbooks.Item(1).Close} if workaround_condition    
        @excel.visible = old_visible_value       
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
    # @raise WorkbookNotSaved if the option :if_unsaved is :raise and the workbook is unsaved
    # @raise OptionInvalid if the options is invalid
    def close(opts = {:if_unsaved => :raise})
      if (alive? && (not @ole_workbook.Saved) && writable) then
        case opts[:if_unsaved]
        when :raise
          raise WorkbookNotSaved, "workbook is unsaved: #{File.basename(self.stored_filename).inspect}"
        when :save
          save
          close_workbook
        when :forget
          @excel.with_displayalerts(false) { close_workbook }
        when :keep_open
          # nothing
        when :alert, :excel
          @excel.with_displayalerts(true) { close_workbook }
        else
          raise OptionInvalid, ":if_unsaved: invalid option: #{opts[:if_unsaved].inspect}"
        end
      else
        close_workbook
      end
      #trace "close: canceled by user" if alive? &&  
      #  (opts[:if_unsaved] == :alert || opts[:if_unsaved] == :excel) && (not @ole_workbook.Saved)
    end

  private

    def close_workbook    
      @ole_workbook.Close if alive?
      @ole_workbook = nil unless alive?
    end

  public

    # keeps the saved-status unchanged
    def retain_saved
      saved = self.Saved
      begin
         yield self
      ensure
        self.Saved = saved
      end
    end

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

    # allows to read or modify a workbook such that its state remains unchanged
    # state comprises: open, saved, writable, visible, calculation mode, check compatibility 
    # @param [String] file        the file name
    # @param [Hash]   opts        the options
    # @option opts [Variant] :if_closed  :current (default), :new or an Excel instance
    # @option opts [Boolean] :read_only whether the workbook is opened for read-only
    # @option opts [Boolean] :readonly_excel behaviour when workbook is opened read-only and shall be modified
    #               true:  closes it and open it as writable in the Excel instance where it was open so far
    #               false (default)  opens it as writable in another running excel instance, if it exists,
    #                                otherwise open in a new Excel instance.
    # @option opts [Boolean] :keep_open whether the workbook shall be kept open after unobtrusively opening 
    # @return [Book] a workbook

    # state = [:open, :saved, :writable, :visible, :calculation, :check_compatibility]

    def self.unobtrusively(file, opts = { }, &block) 
      options = {:if_closed => :current, 
                 :read_only => false,
                 :readonly_excel => false,
                 :keep_open => false}.merge(opts)
      book = bookstore.fetch(file, :prefer_writable => (not options[:read_only]))
      was_open = book && book.alive?
      if was_open
        was_saved = book.saved
        was_writable = book.writable
        was_visible = book.visible
        was_calculation = book.calculation
        was_check_compatibility = book.check_compatibility
      end     
      begin 
        book = 
          if was_open 
            if (not was_writable) && (not options[:read_only])
              open(file, :force => {:excel => (options[:readonly_excel] ? book.excel : :new)}, :read_only => false)
            else
              book
            end
          else
            open(file, :force => {:excel => options[:if_closed]}, :read_only => false)
          end
        yield book
      ensure
        if book && book.alive?
          unless book.saved
            book.save unless options[:read_only]
            book.Saved = true if (was_saved || (not was_open)) && options[:read_only]
            book.Saved = false if (not was_saved) && (not options[:read_only]) && was_open
          end
          if was_open
            if (not was_writable) && (not options[:read_only]) && options[:readonly_excel]
              book.close
              open(file, :force => {:excel => book.excel}, :if_obstructed => :new_excel, :read_only => true)
            end         
            book.excel.calculation = was_calculation
            book.CheckCompatibility = was_check_compatibility
            #book.visible = was_visible  # not necessary
          end          
          
          book.close unless was_open || options[:keep_open]
        end
      end
    end

    # reopens a closed workbook
    def reopen
      self.class.open(self.stored_filename)
    end

    # simple save of a workbook.
    # @option opts [Boolean]  states, whether colored ranges shall be discolored
    # @return [Boolean] true, if successfully saved, nil otherwise
    def save(opts = {:discoloring => false})      
      raise ObjectNotAlive, "workbook is not alive" if (not alive?)
      raise WorkbookReadOnly, "Not opened for writing (opened with :read_only option)" if @ole_workbook.ReadOnly
      begin
        discoloring if opts[:discoloring] 
        @modified_cells = []
        @ole_workbook.Save 
      rescue WIN32OLERuntimeError => msg
        if msg.message =~ /SaveAs/ and msg.message =~ /Workbook/ then
          raise WorkbookNotSaved, "workbook not saved"
        else
          raise UnexpectedError, "unknown WIN32OLERuntimeError:\n#{msg.message}"
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
    # :discoloring     states, wheter colored ranges shall be discolored
    # @return [Book], the book itself, if successfully saved, raises an exception otherwise
    def save_as(file, opts = { } )
      raise FileNameNotGiven, "filename is nil" if file.nil?
      raise ObjectNotAlive, "workbook is not alive" unless alive?
      raise WorkbookReadOnly, "Not opened for writing (opened with :read_only option)" if @ole_workbook.ReadOnly
      options = {
        :if_exists => :raise,
        :if_obstructed => :raise,
      }.merge(opts)
      if File.exist?(file) then
        case options[:if_exists]
        when :overwrite
          if file == self.filename
            save({:discoloring => opts[:discoloring]})
            return self
          else
            begin
              File.delete(file)
            rescue Errno::EACCES
              raise WorkbookBeingUsed, "workbook is open and used in Excel"
            end
          end
        when :alert, :excel 
          @excel.with_displayalerts true do
            save_as_workbook(file, options)
          end
          return self
        when :raise
          raise FileAlreadyExists, "file already exists: #{File.basename(file).inspect}"
        else
          raise OptionInvalid, ":if_exists: invalid option: #{options[:if_exists].inspect}"
        end
      end
      other_workbook = @excel.Workbooks.Item(File.basename(file)) rescue nil
      if other_workbook && (not(self.filename == other_workbook.Fullname.tr('\\','/'))) then
        case options[:if_obstructed]
        when :raise
          raise WorkbookBlocked, "blocked by another workbook: #{other_workbook.Fullname.tr('\\','/')}"
        when :forget
          # nothing
        when :save
          other_workbook.Save
        when :close_if_saved
          raise WorkbookBlocked, "blocking workbook is unsaved: #{File.basename(file).inspect}" unless other_workbook.Saved
        else
          raise OptionInvalid, ":if_obstructed: invalid option: #{options[:if_obstructed].inspect}"
        end
        other_workbook.Close
      end
      save_as_workbook(file, options)
      self
    end

  private

    def discoloring
      # self.each{|sheet| sheet.UsedRange.each{|cell| cell.Interior.ColorIndex = XlNone}}
      @modified_cells.each{|cell| cell.Interior.ColorIndex = XlNone}
    end

    def save_as_workbook(file, options)   # :nodoc: #
      begin
        dirname, basename = File.split(file)
        file_format =
          case File.extname(basename)
            when '.xls' ; RobustExcelOle::XlExcel8
            when '.xlsx'; RobustExcelOle::XlOpenXMLWorkbook
            when '.xlsm'; RobustExcelOle::XlOpenXMLWorkbookMacroEnabled
          end
        discoloring if options[:discoloring]  
        @modified_cells = []
        @ole_workbook.SaveAs(General::absolute_path(file), file_format)
        bookstore.store(self)
      rescue WIN32OLERuntimeError => msg
        if msg.message =~ /SaveAs/ and msg.message =~ /Workbook/ then
          # trace "save: canceled by user" if options[:if_exists] == :alert || options[:if_exists] == :excel
          # another possible semantics. raise WorkbookError, "could not save Workbook"
        else
          raise UnexpectedError, "unknown WIN32OELERuntimeError:\n#{msg.message}"
        end       
      end
    end

  public

    # closes a given file if it is open
    def self.close(file, opts = {:if_unsaved => :raise})
      book = bookstore.fetch(file) rescue nil
      book.close(opts) if book && book.alive?
    end

    # saves a given file if it is open
    def self.save(file)
      book = bookstore.fetch(file) rescue nil
      book.save if book && book.alive?
    end

    # saves a given file under a new name if it is open
    def self.save_as(file, new_file, opts = { })
      book = bookstore.fetch(file) rescue nil
      book.save_as(new_file, opts) if book && book.alive?
    end

    # returns a sheet, if a sheet name or a number is given
    # @param [String] or [Number]
    # @returns [Sheet]
    def sheet(name)
      begin
        sheet_class.new(@ole_workbook.Worksheets.Item(name))
      rescue WIN32OLERuntimeError => msg
        raise NameNotFound, "could not return a sheet with name #{name.inspect}"
      end
    end    

    def each
      @ole_workbook.Worksheets.each do |sheet|
        yield sheet_class.new(sheet)
      end
    end

    def each_with_index(offset = 0)
      i = offset
      @ole_workbook.Worksheets.each do |sheet|
        yield sheet_class.new(sheet), i
        i += 1
      end
    end
  
    # copies a sheet to another position
    # default: copied sheet is appended
    # @param [Sheet] sheet a sheet that shall be copied
    # @param [Hash]  opts  the options
    # @option opts [Symbol] :as     new name of the copied sheet
    # @option opts [Symbol] :before a sheet before which the sheet shall be inserted
    # @option opts [Symbol] :after  a sheet after which the sheet shall be inserted
    # @raise  NameAlreadyExists if the sheet name already exists
    # @return [Sheet] the copied sheet
    def copy_sheet(sheet, opts = { })
      new_sheet_name = opts.delete(:as)
      after_or_before, base_sheet = opts.to_a.first || [:after, last_sheet]
      sheet.Copy({ after_or_before.to_s => base_sheet.worksheet })
      new_sheet = sheet_class.new(@excel.Activesheet)
      new_sheet.name = new_sheet_name if new_sheet_name
      new_sheet
    end      

    # adds an empty sheet
    # default: empty sheet is appended
    # @param [Hash]  opts  the options
    # @option opts [Symbol] :as     new name of the copied added sheet
    # @option opts [Symbol] :before a sheet before which the sheet shall be inserted
    # @option opts [Symbol] :after  a sheet after which the sheet shall be inserted
    # @raise  NameAlreadyExists if the sheet name already exists
    # @return [Sheet] the added sheet
    def add_empty_sheet(opts = { })
      new_sheet_name = opts.delete(:as)
      after_or_before, base_sheet = opts.to_a.first || [:after, last_sheet]
      @ole_workbook.Worksheets.Add({ after_or_before.to_s => base_sheet.worksheet })
      new_sheet = sheet_class.new(@excel.Activesheet)
      new_sheet.name = new_sheet_name if new_sheet_name
      new_sheet
    end    

    # copies a sheet to another position if a sheet is given, or adds an empty sheet
    # default: copied or empty sheet is appended, i.e. added behind the last sheet
    # @param [Sheet] sheet a sheet that shall be copied (optional)
    # @param [Hash]  opts  the options
    # @option opts [Symbol] :as     new name of the copied or added sheet
    # @option opts [Symbol] :before a sheet before which the sheet shall be inserted
    # @option opts [Symbol] :after  a sheet after which the sheet shall be inserted
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
      set_nameval(name,value, :color => 42)   # 42 - aqua-marin, 4-green
    end

    # returns the contents of a range with given name
    # evaluates formula contents of the range is a formula
    # if no contents could be returned, then return default value, if provided, raise error otherwise
    # Excel Bug: if a local name without a qualifier is given, then by default Excel takes the first worksheet,
    #            even if a different worksheet is active
    # @param  [String]      name      the name of the range
    # @param  [Hash]        opts      the options
    # @option opts [Symbol] :default  the default value that is provided if no contents could be returned
    # @return [Variant] the contents of a range with given name
    def nameval(name, opts = {:default => nil})
      name_obj = name_object(name)
      value = begin
        name_obj.RefersToRange.Value
      rescue  WIN32OLERuntimeError
        begin
          self.sheet(1).Evaluate(name_obj.Name)
        rescue WIN32OLERuntimeError
          return opts[:default] if opts[:default]
          raise RangeNotEvaluatable, "cannot evaluate range named #{name.inspect} in #{File.basename(self.stored_filename).inspect}"
        end
      end
      if value.is_a?(Bignum)  #RobustExcelOle::XlErrName  
        return opts[:default] if opts[:default]
        raise RangeNotEvaluatable, "cannot evaluate range named #{name.inspect} in #{File.basename(self.stored_filename).inspect}"
      end 
      return opts[:default] if opts[:default] && value.nil?
      value      
    end

    # sets the contents of a range
    # @param [String]  name  the name of a range
    # @param [Variant] value the contents of the range
    # @param [FixNum]  color the color when setting a value
    # @param [Hash]    opts :color [FixNum]  the color when setting the contents
    def set_nameval(name, value, opts = {:color => 0})
      begin
        cell = name_object(name).RefersToRange
        cell.Interior.ColorIndex = opts[:color] 
        @modified_cells << cell unless cell_included?(cell)
        cell.Value = value
      rescue WIN32OLERuntimeError
        raise RangeNotEvaluatable, "cannot assign value to range named #{name.inspect} in #{File.basename(self.stored_filename).inspect}"    
      end
    end

  private

    def cell_included?(cell)
      @modified_cells.each{|c| return true if c.Name.Value == cell.Name.Value}
      false
    end        

    def name_object(name)
      begin
        self.Parent.Names.Item(name)
      rescue WIN32OLERuntimeError
        begin
          self.Names.Item(name)
        rescue WIN32OLERuntimeError
          raise NameNotFound, "name #{name.inspect} not in #{File.basename(self.stored_filename).inspect}"  
        end
      end
    end

  public    

    # renames a range
    # @param [String] name     the previous range name
    # @param [String] new_name the new range name
    def rename_range(name, new_name)
      begin
        item = self.Names.Item(name)
      rescue WIN32OLERuntimeError
        raise NameNotFound, "name #{name.inspect} not in #{File.basename(self.stored_filename).inspect}"  
      end
      begin
        item.Name = new_name
      rescue WIN32OLERuntimeError
        raise UnexpectedError, "name error in #{File.basename(self.stored_filename).inspect}"      
      end
    end

    # brings workbook to foreground, makes it available for heyboard inputs, makes the Excel instance visible
    def focus
      self.visible = true
      @excel.focus
      @ole_workbook.Activate
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

    def calculation
      @excel.calculation if @ole_workbook
    end

    def check_compatibility
      @ole_workbook.CheckCompatibility if @ole_workbook
    end

        # returns true, if the workbook is visible, false otherwise 
    def visible
      @excel.visible && @ole_workbook.Windows(@ole_workbook.Name).Visible
    end

    # makes both the Excel instance and the window of the workbook visible, or the window invisible
    # @param [Boolean] visible_value determines whether the workbook shall be visible
    def visible= visible_value
      @excel.visible = true if visible_value
      self.window_visible = visible_value
    end

    # returns true, if the window of the workbook is set to visible, false otherwise
    def window_visible
      return @ole_workbook.Windows(@ole_workbook.Name).Visible
    end

    # makes the window of the workbook visible or invisible
    # @param [Boolean] visible_value determines whether the window of the workbook shall be visible
    def window_visible= visible_value
      retain_saved do
        @ole_workbook.Windows(@ole_workbook.Name).Visible = visible_value if @ole_workbook.Windows.Count > 0
      end
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

    def to_s    # :nodoc: #
      "#{self.filename}"
    end

    def inspect    # :nodoc: #
      "#<Book: " + "#{"not alive " unless alive?}" + "#{File.basename(self.filename) if alive?}" + " #{@ole_workbook} #{@excel}"  + ">"
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
          raise ObjectNotAlive, "method missing: workbook not alive" unless alive?
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

end
