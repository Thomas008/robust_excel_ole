# -*- coding: utf-8 -*-

require 'weakref'

module RobustExcelOle

  # This class essentially wraps a Win32Ole Workbook object. 
  # You can apply all VBA methods (starting with a capital letter) 
  # that you would apply for a Workbook object. 
  # See https://docs.microsoft.com/en-us/office/vba/api/excel.workbook#methods

  class Workbook < RangeOwners

    #include General 

    attr_accessor :excel
    attr_accessor :ole_workbook
    attr_accessor :stored_filename
    attr_accessor :color_if_modified
    attr_accessor :was_open
    attr_reader :workbook

    alias ole_object ole_workbook

    DEFAULT_OPEN_OPTS = {
      :default => {:excel => :current},
      :force => {},
      :update_links => :never,
      :if_unsaved    => :raise,
      :if_obstructed => :raise,
      :if_absent     => :raise,
      :if_exists => :raise,
      :check_compatibility => false
    }.freeze

    CORE_DEFAULT_OPEN_OPTS = {
      :default => {:excel => :current}, :force => {}, :update_links => :never
    }.freeze

    ABBREVIATIONS = [[:default,:d], [:force, :f], [:excel, :e], [:visible, :v],
                     [:if_obstructed, :if_blocked]].freeze


    # opens a workbook.
    # @param [String] file_or_workbook a file name or WIN32OLE workbook
    # @param [Hash] opts the options
    # @option opts [Hash] :default or :d
    # @option opts [Hash] :force or :f
    # @option opts [Symbol]  :if_unsaved     :raise (default), :forget, :save, :accept, :alert, :excel, or :new_excel
    # @option opts [Symbol]  :if_blocked     :raise (default), :forget, :save, :close_if_saved, or _new_excel
    # @option opts [Symbol]  :if_absent      :raise (default) or :create
    # @option opts [Boolean] :read_only      true (default) or false
    # @option opts [Boolean] :update_links   :never (default), :always, :alert
    # @option opts [Boolean] :calculation    :manual, :automatic, or nil (default)
    # options:
    # :default : if the workbook was already open before, then use (unchange) its properties,
    #            otherwise, i.e. if the workbook cannot be reopened, use the properties stated in :default
    # :force   : no matter whether the workbook was already open before, use the properties stated in :force
    # :default and :force contain: :excel
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
    # or               :raise               -> raises an exception
    # :if_blocked      :forget              -> closes the old workbook, open the new workbook
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
    # @return [Workbook] a representation of a workbook
    def self.new(file_or_workbook, opts = { }, &block)
      options = process_options(opts)
      if file_or_workbook.is_a? WIN32OLE
        file = file_or_workbook.Fullname.tr('\\','/') 
      else
        file = file_or_workbook
        raise(FileNameNotGiven, 'filename is nil') if file.nil?
        raise(FileNotFound, "file #{General.absolute_path(file).inspect} is a directory") if File.directory?(file)
      end
      # try to fetch the workbook from the bookstore
      book = nil
      if options[:force][:excel] != :new
        # if readonly is true, then prefer a book that is given in force_excel if this option is set              
        forced_excel = 
          (options[:force][:excel].nil? || options[:force][:excel] == :current) ? 
            (excel_class.new(:reuse => true) if !::JRUBY_BUG_CONNECT) : excel_of(options[:force][:excel])              
        begin
          book = if File.exists?(file)
            bookstore.fetch(file, :prefer_writable => !(options[:read_only]),
                                  :prefer_excel    => (options[:read_only] ? forced_excel : nil))
          end
        rescue
          trace "#{$!.message}"
        end
        if book 
          # hack (unless condition): calling Worksheet[]= causes calling Worksheet#workbook which calls Workbook#new(ole_workbook)
          book.was_open = book.alive? unless file_or_workbook.is_a? WIN32OLE 
          # drop the fetched workbook if it shall be opened in another Excel instance
          # or the workbook is an unsaved workbook that should not be accepted
          if (options[:force][:excel].nil? || options[:force][:excel] == :current || forced_excel == book.excel) &&
            !(book.alive? && !book.saved && (options[:if_unsaved] != :accept))
            options[:force][:excel] = book.excel if book.excel && book.excel.alive?
            book.ensure_workbook(file,options)
            book.set_options(file,options)
            return book
          end
        end
      end        
      super(file_or_workbook, options, &block)
    end

    def self.open(file_or_workbook, opts = { }, &block)
      new(file_or_workbook, opts, &block)
    end

    # creates a new Workbook object, if a file name is given
    # Promotes the win32ole workbook to a Workbook object, if a win32ole-workbook is given
    # @param [Variant] file_or_workbook  file name or workbook
    # @param [Hash]    opts             
    # @option opts [Symbol] see above
    # @return [Workbook] a workbook
    def initialize(file_or_workbook, options = { }, &block)
      if file_or_workbook.is_a? WIN32OLE
        @ole_workbook = file_or_workbook
        ole_excel = begin
          WIN32OLE.connect(@ole_workbook.Fullname).Application
        rescue
          raise ExcelREOError, 'could not determine the Excel instance'
        end
        @excel = excel_class.new(ole_excel)
        filename = file_or_workbook.Fullname.tr('\\','/') 
      else
        filename = file_or_workbook            
        ensure_workbook(filename, options)        
      end      
      set_options(filename, options)
      bookstore.store(self)
      @workbook = @excel.workbook = self
      r1c1_letters = @ole_workbook.Worksheets.Item(1).Cells.Item(1,1).Address(true,true,XlR1C1).gsub(/[0-9]/,'') #('ReferenceStyle' => XlR1C1).gsub(/[0-9]/,'')
      address_class.new(r1c1_letters)
      if block
        begin
          yield self
        ensure
          close
        end
      end
    end
 
  private    

    # translates abbreviations and synonyms and merges with default options
    def self.process_options(options, proc_opts = {:use_defaults => true})
      translator = proc do |opts|
        erg = {}
        opts.each do |key,value|
          new_key = key
          ABBREVIATIONS.each { |long,short| new_key = long if key == short }
          if value.is_a?(Hash)
            erg[new_key] = {}
            value.each do |k,v|
              new_k = k
              ABBREVIATIONS.each { |l,s| new_k = l if k == s }
              erg[new_key][new_k] = v
            end
          else
            erg[new_key] = value
          end
        end
        erg[:default] ||= {}
        erg[:force] ||= {}
        force_list = [:visible, :excel]
        erg.each { |key,value| erg[:force][key] = value if force_list.include?(key) }
        erg[:default][:excel] = erg[:default_excel] unless erg[:default_excel].nil?
        erg[:force][:excel] = erg[:force_excel] unless erg[:force_excel].nil?
        erg[:default][:excel] = :current if erg[:default][:excel] == :reuse || erg[:default][:excel] == :active
        erg[:force][:excel] = :current if erg[:force][:excel] == :reuse || erg[:force][:excel] == :active
        erg
      end
      opts = translator.call(options)
      default_open_opts = proc_opts[:use_defaults] ? DEFAULT_OPEN_OPTS : CORE_DEFAULT_OPEN_OPTS        
      default_opts = translator.call(default_open_opts)
      opts = default_opts.merge(opts)
      opts[:default] = default_opts[:default].merge(opts[:default]) unless opts[:default].nil?
      opts[:force] = default_opts[:force].merge(opts[:force]) unless opts[:force].nil?
      opts
    end

    # returns an Excel object when given Excel, Workbook or Win32ole object representing a Workbook or an Excel
    # @private
    def self.excel_of(object) 
      begin
        object = object.to_reo if object.is_a? WIN32OLE
        object.excel
      rescue
        raise TypeREOError, 'given object is neither an Excel, a Workbook, nor a Win32ole'
      end
    end

  public

    # @private
    # ensures an excel but not for jruby if current Excel shall be used
    def ensure_excel(options)
      return if @excel && @excel.alive?
      excel_option = options[:force][:excel].nil? ? options[:default][:excel] : options[:force][:excel]
      @excel = if excel_option == :new
        excel_class.new(:reuse => false) 
      elsif excel_option.nil? || excel_option == :current
        excel_class.new(:reuse => true)
      else
        self.class.excel_of(excel_option)
      end
      raise ExcelREOError, "excel is not alive" unless @excel && @excel.alive?
    end

    # @private    
    def ensure_workbook(filename, options)  
      return if (@ole_workbook && alive? && (options[:read_only].nil? || @ole_workbook.ReadOnly == options[:read_only]))
      filename = @stored_filename ? @stored_filename : filename 
      manage_nonexisting_file(filename,options)
      excel_option = options[:force][:excel].nil? ? options[:default][:excel] : options[:force][:excel]        
      ensure_excel(options)
      workbooks = @excel.Workbooks
      @ole_workbook = workbooks.Item(File.basename(filename)) rescue nil if @ole_workbook.nil?
      if @ole_workbook
        @was_open = true if @was_open.nil? # necessary?
        manage_blocking_or_unsaved_workbook(filename,options)
        open_or_create_workbook(filename,options) if @ole_workbook.ReadOnly != options[:read_only]
      else
        if excel_option.nil? || excel_option == :current &&  
          (!::JRUBY_BUG_CONNECT || filename[0] != '/')
          workbooks_number_before_connect = @excel.Workbooks.Count
          connect(filename,options)
          @was_open = @excel.Workbooks.Count == workbooks_number_before_connect
        else 
          open_or_create_workbook(filename,options)
        end
      end       
    end

    # @private
    def set_options(filename, options)
      # changing read-only mode
      if (!options[:read_only].nil?) && options[:read_only] != @ole_workbook.ReadOnly
        ensure_workbook(filename, options)
      end
      retain_saved do
        self.visible = options[:force][:visible].nil? ?
          (@excel.Visible && @ole_workbook.Windows(@ole_workbook.Name).Visible) : options[:force][:visible]
        @excel.calculation = options[:calculation] unless options[:calculation].nil?
        @ole_workbook.CheckCompatibility = options[:check_compatibility] unless options[:check_compatibility].nil?
      end      
    end

  private

    # @private
    # connects to an unknown workbook
    def connect(filename,options)
      excels_number = excel_class.excels_number
      workbooks_number = if excels_number>0 
        excel_class.current.Workbooks.Count 
      else 0
      end
      abs_filename = General.absolute_path(filename)
      @ole_workbook = begin
        WIN32OLE.connect(abs_filename)
      rescue
        if $!.message =~ /moniker/
          raise WorkbookConnectingBlockingError
        else
          raise WorkbookConnectingUnknownError
        end
      end
      ole_excel = begin
        @ole_workbook.Application     
      rescue 
        if $!.message =~ /dispid/
          raise WorkbookConnectingUnsavedError
        else
          raise WorkbookConnectingUnknownError
        end
      end
      @excel = excel_class.new(ole_excel)
    end

    # @private
    def manage_nonexisting_file(filename,options)   
      return if File.exist?(filename)
      abs_filename = General.absolute_path(filename)
      if options[:if_absent] == :create
        ensure_excel(options) unless @excel && @excel.alive?
        @excel.Workbooks.Add
        empty_ole_workbook = excel.Workbooks.Item(excel.Workbooks.Count)
        begin
          empty_ole_workbook.SaveAs(abs_filename)
        rescue WIN32OLERuntimeError, Java::OrgRacobCom::ComFailException => msg
          raise FileNotFound, "could not save workbook with filename #{filename.inspect}"
        end
      else
        raise FileNotFound, "file #{abs_filename.inspect} not found" +
          "\nHint: If you want to create a new file, use option :if_absent => :create or Workbook::create"
      end
    end

    # @private
    def manage_blocking_or_unsaved_workbook(filename,options)
      obstructed_by_other_book = if (File.basename(filename) == File.basename(@ole_workbook.Fullname)) 
        General.absolute_path(filename) != @ole_workbook.Fullname
      end      
      if obstructed_by_other_book
        # workbook is being obstructed by a workbook with same name and different path
        manage_blocking_workbook(filename,options)        
      else
        unless @ole_workbook.Saved
          # workbook open and writable, not obstructed by another workbook, but not saved
          manage_unsaved_workbook(filename,options)
        end
      end        
    end

    # @private
    def manage_blocking_workbook(filename,options)
      case options[:if_obstructed]
      when :raise
        raise WorkbookBlocked, "can't open workbook #{filename},
        because it is being blocked by #{@ole_workbook.Fullname.tr('\\','/')} with the same name in a different path." +
        "\nHint: Use the option :if_blocked with values :forget or :save,
         to allow automatic closing of the old workbook (without or with saving before, respectively),
         before the new workbook is being opened."
      when :forget
        manage_forgetting_workbook(filename, options)       
      when :save
        manage_saving_workbook(filename, options)        
      when :close_if_saved
        if !@ole_workbook.Saved
          raise WorkbookBlocked, "workbook with the same name in a different path is unsaved: #{@ole_workbook.Fullname.tr('\\','/')}"
        else
          manage_forgetting_workbook(filename, options)
        end
      when :new_excel
        manage_new_excel(filename, options)        
      else
        raise OptionInvalid, ":if_blocked: invalid option: #{options[:if_obstructed].inspect}" +
        "\nHint: Valid values are :raise, :forget, :save, :close_if_saved, :new_excel"
      end
    end

    # @private
    def manage_unsaved_workbook(filename,options)
      case options[:if_unsaved]
      when :raise
        raise WorkbookNotSaved, "workbook is already open but not saved: #{File.basename(filename).inspect}" +
        "\nHint: Save the workbook or open the workbook using option :if_unsaved with values :forget and :accept to
         close the unsaved workbook and reopen it, or to let the unsaved workbook open, respectively"
      when :forget
        manage_forgetting_workbook(filename,options)
      when :accept
        # do nothing
      when :save
        manage_saving_workbook(filename, options)
      when :alert, :excel
        @excel.with_displayalerts(true) { open_or_create_workbook(filename,options) }
      when :new_excel
        manage_new_excel(filename, options)
      else
        raise OptionInvalid, ":if_unsaved: invalid option: #{options[:if_unsaved].inspect}" +
        "\nHint: Valid values are :raise, :forget, :save, :accept, :alert, :excel, :new_excel"
      end
    end

    # @private
    def manage_forgetting_workbook(filename, options)
      @excel.with_displayalerts(false) { @ole_workbook.Close }
      @ole_workbook = nil
      open_or_create_workbook(filename, options)
    end

    # @private
    def manage_saving_workbook(filename, options)
      save unless @ole_workbook.Saved
      manage_forgetting_workbook(filename, options)
    end

    # @private
    def manage_new_excel(filename, options)
      @excel = excel_class.new(:reuse => false)
      @ole_workbook = nil
      open_or_create_workbook(filename, options)
    end
    
    # @private
    def open_or_create_workbook(filename, options)
      return if @ole_workbook && options[:if_unsaved] != :alert && options[:if_unsaved] != :excel &&
        (options[:read_only].nil? || options[:read_only]==@ole_workbook.ReadOnly )
      begin
        abs_filename = General.absolute_path(filename)
        begin
          @was_open = false if @was_open.nil?
          workbooks = @excel.Workbooks
        rescue WIN32OLERuntimeError, Java::OrgRacobCom::ComFailException => msg
          raise UnexpectedREOError, "cannot access workbooks: #{msg.message} #{msg.backtrace}"
        end
        begin
          with_workaround_linked_workbooks_excel2007(options) do
            # temporary workaround until jruby-win32ole implements named parameters (Java::JavaLang::RuntimeException (createVariant() not implemented for class org.jruby.RubyHash)
            workbooks.Open(abs_filename, 
                                      updatelinks_vba(options[:update_links]), 
                                      options[:read_only] )
          end
        rescue WIN32OLERuntimeError => msg
          # for Excel2007: for option :if_unsaved => :alert and user cancels: this error appears?
          # if yes: distinguish these events
          raise UnexpectedREOError, "cannot open workbook: #{msg.message} #{msg.backtrace}"
        end
        begin
          # workaround for bug in Excel 2010: workbook.Open does not always return the workbook when given file name
          begin
            @ole_workbook = workbooks.Item(File.basename(filename))
          rescue WIN32OLERuntimeError => msg
            raise UnexpectedREOError, "WIN32OLERuntimeError: #{msg.message}"
          end
        rescue WIN32OLERuntimeError => msg
          raise UnexpectedREOError, "WIN32OLERuntimeError: #{msg.message} #{msg.backtrace}"
        end
      end
    end    
           
    # @private
    # translating the option UpdateLinks from REO to VBA
    # setting UpdateLinks works only if calculation mode is automatic,
    # parameter 'UpdateLinks' has no effect
    def updatelinks_vba(updatelinks_reo)
      case updatelinks_reo
      when :alert then RobustExcelOle::XlUpdateLinksUserSetting
      when :never then RobustExcelOle::XlUpdateLinksNever
      when :always then RobustExcelOle::XlUpdateLinksAlways
      else RobustExcelOle::XlUpdateLinksNever
      end
    end

    # @private
    # workaround for linked workbooks for Excel 2007:
    # opening and closing a dummy workbook if Excel has no workbooks.
    # delay: with visible: 0.2 sec, without visible almost none
    def with_workaround_linked_workbooks_excel2007(options)
      old_visible_value = @excel.Visible
      workbooks = @excel.Workbooks
      workaround_condition = @excel.Version.split('.').first.to_i == 12 && workbooks.Count == 0
      if workaround_condition
        workbooks.Add
        @excel.calculation = options[:calculation].nil? ? @excel.calculation : options[:calculation]
      end
      begin
        # @excel.with_displayalerts(update_links_opt == :alert ? true : @excel.displayalerts) do
        yield self
      ensure
        @excel.with_displayalerts(false) { workbooks.Item(1).Close } if workaround_condition
        @excel.visible = old_visible_value
      end
    end

  public

    # creates, i.e., opens a new, empty workbook, and saves it under a given filename
    # @param [String] filename the filename under which the new workbook should be saved
    # @param [Hash] opts the options as in Workbook::open
    def self.create(filename, opts = { })
      open(filename, :if_absent => :create)
    end

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
      if alive? && !@ole_workbook.Saved && writable
        case opts[:if_unsaved]
        when :raise
          raise WorkbookNotSaved, "workbook is unsaved: #{File.basename(self.stored_filename).inspect}" +
          "\nHint: Use option :save or :forget to close the workbook with or without saving"
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
          raise OptionInvalid, ":if_unsaved: invalid option: #{opts[:if_unsaved].inspect}" +
          "\nHint: Valid values are :raise, :save, :keep_open, :alert, :excel"
        end
      else
        close_workbook
      end
      # trace "close: canceled by user" if alive? &&
      #  (opts[:if_unsaved] == :alert || opts[:if_unsaved] == :excel) && (not @ole_workbook.Saved)
    end

  private

    # @private
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
      opts = {:writable => false}.merge(opts)
      args.push opts
      unobtrusively(*args, &block)
    end

    def self.for_modifying(*args, &block)
      args = args.dup
      opts = args.last.is_a?(Hash) ? args.pop : {}
      opts = {:writable => true}.merge(opts)
      args.push opts
      unobtrusively(*args, &block)
    end

    # allows to read or modify a workbook such that its state remains unchanged
    # the state comprises, whether the workbok was open, saved, writable, visible,
    # includes calculation mode, and check compatibility
    # @param [String] file_or_workbook     a file name or WIN32OLE workbook
    # @param [Hash]   opts        the options
    # @option opts [Variant] :if_closed  Excel where to open the workbook, if the workbook was closed
    #                           :current (default), :new or an Excel instance (this option is for known workbooks only) 
    # @option opts [Boolean] :read_only true/false (default), open the workbook in read-only/read-write modus (save changes)
    # @option opts [Boolean] :writable  true (default)/false changes of the workbook shall be saved/not saved
    # @option opts [Boolean] :keep_open whether the workbook shall be kept open after unobtrusively opening (default: false)
    # @return [Workbook] a workbook
    def self.unobtrusively(file_or_workbook, opts = { })
      file = (file_or_workbook.is_a? WIN32OLE) ? file_or_workbook.Fullname.tr('\\','/') : file_or_workbook
      opts = process_options(opts)
      raise OptionInvalid, 'contradicting options' if opts[:writable] && opts[:read_only]       
      opts = {:if_closed => :current, :keep_open => false}.merge(opts)      
      prefer_writable = ((!(opts[:read_only]) || opts[:writable] == true) &&
                         !(opts[:read_only].nil? && opts[:writable] == false))
      book = bookstore.fetch(file, :prefer_writable => prefer_writable)    
      excel_opts = (book.nil? || (book && !book.alive?)) ? {:force => {:excel => opts[:if_closed]}} :
          {:force => {:excel => opts[:force][:excel]}, :default => {:excel => opts[:default][:excel]}}
      open_opts = excel_opts.merge({:if_unsaved => :accept})
      begin  
        book = open(file, open_opts)
        was_visible = book.visible
        was_writable = book.writable
        was_saved = book.saved
        was_check_compatibility = book.check_compatibility
        was_calculation = book.excel.calculation
        book.set_options(file,opts)       
        yield book
      ensure
        if book && book.alive?
          do_not_write = opts[:read_only] || opts[:writable]==false
          book.save unless book.saved || do_not_write || book.ReadOnly
          if (opts[:read_only] && was_writable) || (!opts[:read_only] && !was_writable)
            book.set_options(file, opts.merge({:read_only => !was_writable, 
                                   :if_unsaved => (opts[:writable]==false ? :forget : :save)}))
          end
          if book.was_open
            book.visible = was_visible    
            book.CheckCompatibility = was_check_compatibility
            book.excel.calculation = was_calculation
          end
          book.Saved = (was_saved || !book.was_open)
          book.close unless book.was_open || opts[:keep_open]
        end
      end
    end

=begin
    # allows to read or modify a workbook such that its state remains unchanged
    # state comprises: open, saved, writable, visible, calculation mode, check compatibility
    # @param [String] file_or_workbook     a file name or WIN32OLE workbook
    # @param [Hash]   opts        the options
    # @option opts [Variant] :if_closed  :current (default), :new or an Excel instance
    # @option opts [Boolean] :read_only true/false (default), open the workbook in read-only/read-write modus (save changes)
    # @option opts [Boolean] :writable  true (default)/false changes of the workbook shall be saved/not saved
    # @option opts [Boolean] :keep_open whether the workbook shall be kept open after unobtrusively opening (default: false)
    # @return [Workbook] a workbook
    def self.unobtrusively(file_or_workbook, opts = { })
      file = (file_or_workbook.is_a? WIN32OLE) ? file_or_workbook.Fullname.tr('\\','/') : file_or_workbook
      opts = process_options(opts)
      raise OptionInvalid, 'contradicting options' if opts[:writable] && opts[:read_only]       
      opts = {:if_closed => :current, :keep_open => false}.merge(opts)      
      prefer_writable = ((!(opts[:read_only]) || opts[:writable] == true) &&
                         !(opts[:read_only].nil? && opts[:writable] == false))
      book = bookstore.fetch(file, :prefer_writable => prefer_writable)    
      excel_opts = (book.nil? || (book && !book.alive?)) ? {:force => {:excel => opts[:if_closed]}} :
          {:force => {:excel => opts[:force][:excel]}, :default => {:excel => opts[:default][:excel]}}
      open_opts = excel_opts.merge({:if_unsaved => :accept})
      begin  
        book = open(file, open_opts)
        book.unobtrusively
      end
    end
=end

    def unobtrusively(opts = { })
      file = stored_filename
      opts = self.class.process_options(opts)
      raise OptionInvalid, 'contradicting options' if opts[:writable] && opts[:read_only]       
      opts = {:if_closed => :current, :keep_open => false}.merge(opts)      
      excel_opts = !book.alive? ? {:force => {:excel => opts[:if_closed]}} :
        {:force => {:excel => opts[:force][:excel]}, :default => {:excel => opts[:default][:excel]}}
      open_opts = excel_opts.merge({:if_unsaved => :accept})
      begin  
        was_open = alive?
        was_visible = visible
        was_writable = writable
        was_saved = saved
        was_check_compatibility = check_compatibility
        was_calculation = calculation
        set_options(file,opts)       
        yield self
      ensure
        if was_open
          do_not_write = opts[:read_only] || opts[:writable]==false
          save unless saved || do_not_write || !writable
          if (opts[:read_only] && was_writable) || (!opts[:read_only] && !was_writable)
            set_options(file, opts.merge({:read_only => !was_writable, 
                              :if_unsaved => (opts[:writable]==false ? :forget : :save)}))
          end
          if was_open
            visible = was_visible    
            @ole_workbook.CheckCompatibility = was_check_compatibility
            @excel.calculation = was_calculation
          end
          @ole_workbook.Saved = (was_saved || !was_open)
          close unless was_open || opts[:keep_open]
        end
      end
    end

     


    # reopens a closed workbook
    # @options options
    def reopen(options = { })
      book = self.class.open(@stored_filename, options)
      raise WorkbookREOError('cannot reopen book') unless book && book.alive?
      book
    end

    # simple save of a workbook.
    # @return [Boolean] true, if successfully saved, nil otherwise
    def save(opts = { })  # option opts is deprecated #
      raise ObjectNotAlive, 'workbook is not alive' unless alive?
      raise WorkbookReadOnly, 'Not opened for writing (opened with :read_only option)' if @ole_workbook.ReadOnly   
      # if you have open the workbook with :read_only => true,
      # then you could close the workbook and open it again with option :read_only => false
      # otherwise the workbook may already be open writable in an another Excel instance
      # then you could use this workbook or close the workbook there
      begin
        @ole_workbook.Save
      rescue WIN32OLERuntimeError, Java::OrgRacobCom::ComFailException => msg
        if msg.message =~ /SaveAs/ && msg.message =~ /Workbook/
          raise WorkbookNotSaved, 'workbook not saved'
        else
          raise UnexpectedREOError, "unknown WIN32OLERuntimeError:\n#{msg.message}"
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
    #  or              :raise               -> raises an exception
    #  :if_blocked     :forget              -> closes the blocking workbook
    #                  :save                -> saves the blocking workbook and closes it
    #                  :close_if_saved      -> closes the blocking workbook, if it is saved,
    #                                          otherwise raises an exception   
    # @return [Workbook], the book itself, if successfully saved, raises an exception otherwise
    def save_as(file, opts = { })
      raise FileNameNotGiven, 'filename is nil' if file.nil?
      raise ObjectNotAlive, 'workbook is not alive' unless alive?
      raise WorkbookReadOnly, 'Not opened for writing (opened with :read_only option)' if @ole_workbook.ReadOnly
      raise(FileNotFound, "file #{General.absolute_path(file).inspect} is a directory") if File.directory?(file)
      options = self.class.process_options(opts)      
      if File.exist?(file)
        case options[:if_exists]
        when :overwrite
          if file == self.filename
            save
            return self
          else
            begin
              File.delete(file)
            rescue Errno::EACCES
              raise WorkbookBeingUsed, 'workbook is open and being used in an Excel instance'
            end
          end
        when :alert, :excel
          @excel.with_displayalerts true do
            save_as_workbook(file, options)
          end
          return self
        when :raise
          raise FileAlreadyExists, "file already exists: #{File.basename(file).inspect}" +
          "\nHint: Use option :if_exists => :overwrite, if you want to overwrite the file" 
        else
          raise OptionInvalid, ":if_exists: invalid option: #{options[:if_exists].inspect}" +
          "\nHint: Valid values are :raise, :overwrite, :alert, :excel"
        end
      end
      other_workbook = @excel.Workbooks.Item(File.basename(file)) rescue nil
      if other_workbook && self.filename != other_workbook.Fullname.tr('\\','/')
        case options[:if_obstructed]
        when :raise
          raise WorkbookBlocked, "blocked by another workbook: #{other_workbook.Fullname.tr('\\','/')}" +
          "\nHint: Use the option :if_blocked with values :forget or :save to
           close or save and close the blocking workbook"
        when :forget
          # nothing
        when :save
          other_workbook.Save
        when :close_if_saved
          unless other_workbook.Saved
            raise WorkbookBlocked, "blocking workbook is unsaved: #{File.basename(file).inspect}" +
            "\nHint: Use option :if_blocked => :save to save the blocking workbooks"
          end
        else
          raise OptionInvalid, ":if_blocked: invalid option: #{options[:if_obstructed].inspect}" +
          "\nHint: Valid values are :raise, :forget, :save, :close_if_saved"
        end
        other_workbook.Close
      end
      save_as_workbook(file, options)
      self
    end    

  private

    # @private
    def save_as_workbook(file, options)  
      dirname, basename = File.split(file)
      file_format =
        case File.extname(basename)
        when '.xls' then RobustExcelOle::XlExcel8
        when '.xlsx' then RobustExcelOle::XlOpenXMLWorkbook
        when '.xlsm' then RobustExcelOle::XlOpenXMLWorkbookMacroEnabled
        end
      @ole_workbook.SaveAs(General.absolute_path(file), file_format)
      bookstore.store(self)
    rescue WIN32OLERuntimeError, Java::OrgRacobCom::ComFailException => msg
      if msg.message =~ /SaveAs/ && msg.message =~ /Workbook/
        # trace "save: canceled by user" if options[:if_exists] == :alert || options[:if_exists] == :excel
        # another possible semantics. raise WorkbookREOError, "could not save Workbook"
      else
        raise UnexpectedREOError, "unknown WIN32OELERuntimeError:\n#{msg.message}"
      end
    end

  public

    # closes a given file if it is open
    # @options opts [Symbol] :if_unsaved
    def self.close(file, opts = {:if_unsaved => :raise})
      book = begin
        bookstore.fetch(file)
        rescue
          nil
        end
      book.close(opts) if book && book.alive?
    end

    # saves a given file if it is open
    def self.save(file)
      book = bookstore.fetch(file) rescue nil
      book.save if book && book.alive?
    end

    # saves a given file under a new name if it is open
    def self.save_as(file, new_file, opts = { })
      book = begin
        bookstore.fetch(file)
      rescue 
        nil
      end
      book.save_as(new_file, opts) if book && book.alive?
    end

    # returns a sheet, if a sheet name or a number is given
    # @param [String] or [Number]
    # @returns [Worksheet]
    def sheet(name)
      worksheet_class.new(@ole_workbook.Worksheets.Item(name))
    rescue WIN32OLERuntimeError, Java::OrgRacobCom::ComFailException => msg
      raise NameNotFound, "could not return a sheet with name #{name.inspect}"
    end

    def each
      @ole_workbook.Worksheets.each do |sheet|
        yield worksheet_class.new(sheet)
      end
    end

    def each_with_index(offset = 0)
      i = offset
      @ole_workbook.Worksheets.each do |sheet|
        yield worksheet_class.new(sheet), i
        i += 1
      end
    end

    # copies a sheet to another position if a sheet is given, or adds an empty sheet
    # default: copied or empty sheet is appended, i.e. added behind the last sheet
    # @param [Worksheet] sheet a sheet that shall be copied (optional)
    # @param [Hash]  opts  the options
    # @option opts [Symbol] :as     new name of the copied or added sheet
    # @option opts [Symbol] :before a sheet before which the sheet shall be inserted
    # @option opts [Symbol] :after  a sheet after which the sheet shall be inserted
    # @return [Worksheet] the copied or added sheet
    def add_or_copy_sheet(sheet = nil, opts = { })
      if sheet.is_a? Hash
        opts = sheet
        sheet = nil
      end
      new_sheet_name = opts.delete(:as)
      last_sheet_local = last_sheet
      after_or_before, base_sheet = opts.to_a.first || [:after, last_sheet_local]
      begin
        if !::JRUBY_BUG_COPYSHEETS
          if sheet
            sheet.Copy({ after_or_before.to_s => base_sheet.ole_worksheet })
          else
            @ole_workbook.Worksheets.Add({ after_or_before.to_s => base_sheet.ole_worksheet })
            #@ole_workbook.Worksheets.Item(ole_workbook.Worksheets.Count).Activate
          end
        else
          if after_or_before == :before 
            if sheet
              sheet.Copy(base_sheet.ole_worksheet)
            else
              ole_workbook.Worksheets.Add(base_sheet.ole_worksheet)
            end
          else
            #not_given = WIN32OLE_VARIANT.new(nil, WIN32OLE::VARIANT::VT_NULL)
            #ole_workbook.Worksheets.Add(not_given,base_sheet.ole_worksheet)          
            if base_sheet.name != last_sheet_local.name
              if sheet
                sheet.Copy(base_sheet.Next)
              else
                ole_workbook.Worksheets.Add(base_sheet.Next)
              end
            else
              if sheet
                sheet.Copy(base_sheet.ole_worksheet)  
              else
                ole_workbook.Worksheets.Add(base_sheet.ole_worksheet) 
              end
              base_sheet.Move(ole_workbook.Worksheets.Item(ole_workbook.Worksheets.Count-1))
              ole_workbook.Worksheets.Item(ole_workbook.Worksheets.Count).Activate
            end
          end
        end
      rescue WIN32OLERuntimeError, NameNotFound, Java::OrgRacobCom::ComFailException
        raise WorksheetREOError, "could not add given worksheet #{sheet.inspect}"
      end
      #ole_sheet = @excel.Activesheet
      ole_sheet = ole_workbook.Activesheet
      #ole_sheet = ole_sheet.nil? ? ole_workbook.Worksheets.Item(ole_workbook.Worksheets.Count) : ole_sheet
      new_sheet = worksheet_class.new(ole_sheet)
      new_sheet.name = new_sheet_name if new_sheet_name
      new_sheet
    end

    # for compatibility to older versions
    def add_sheet(sheet = nil, opts = { })
      add_or_copy_sheet(sheet, opts)
    end

    # for compatibility to older versions
    def copy_sheet(sheet, opts = { })
      add_or_copy_sheet(sheet, opts)
    end

    def last_sheet
      worksheet_class.new(@ole_workbook.Worksheets.Item(@ole_workbook.Worksheets.Count))
    end

    def first_sheet
      worksheet_class.new(@ole_workbook.Worksheets.Item(1))
    end

    # returns the value of a range
    # @param [String] name the name of a range
    # @returns [Variant] the value of the range
    def [] name
      namevalue_glob(name)
    end

    # sets the value of a range
    # @param [String]  name  the name of the range
    # @param [Variant] value the contents of the range
    def []= (name, value)
      old_color_if_modified = @color_if_modified
      workbook.color_if_modified = 42  # 42 - aqua-marin, 4-green
      set_namevalue_glob(name,value)   
      workbook.color_if_modified = old_color_if_modified
    end

    # sets options
    # @param [Hash] opts
    def for_this_workbook(opts)
      return unless alive?
      opts = self.class.process_options(opts, :use_defaults => false)
      visible_before = visible
      check_compatibility_before = check_compatibility
      unless opts[:read_only].nil?
        # if the ReadOnly status shall be changed, then close and reopen it
        if (!writable && !(opts[:read_only])) || (writable && opts[:read_only])
          opts[:check_compatibility] = check_compatibility if opts[:check_compatibility].nil?
          close(:if_unsaved => true)
          open_or_create_workbook(@stored_filename, opts)
        end
      end
      self.visible = opts[:force][:visible].nil? ? visible_before : opts[:force][:visible]
      self.CheckCompatibility = opts[:check_compatibility].nil? ? check_compatibility_before : opts[:check_compatibility]
      @excel.calculation = opts[:calculation] unless opts[:calculation].nil?
    end

    # brings workbook to foreground, makes it available for heyboard inputs, makes the Excel instance visible
    def focus
      self.visible = true
      @excel.focus
      @ole_workbook.Activate
    end

    # returns true, if the workbook reacts to methods, false otherwise
    def alive?
      @ole_workbook.Name
      true
    rescue
      @ole_workbook = nil  # dead object won't be alive again
      # t $!.message
      false
    end

    # returns the full file name of the workbook
    def filename
      @ole_workbook.Fullname.tr('\\','/') rescue nil
    end

    # @private
    def writable   
      !@ole_workbook.ReadOnly if @ole_workbook
    end

    # @private
    def saved  
      @ole_workbook.Saved if @ole_workbook
    end

    def calculation
      @excel.calculation if @ole_workbook
    end

    # @private
    def check_compatibility
      @ole_workbook.CheckCompatibility if @ole_workbook
    end

    # returns true, if the workbook is visible, false otherwise
    def visible
      @excel.Visible && @ole_workbook.Windows(@ole_workbook.Name).Visible
    end

    # makes both the Excel instance and the window of the workbook visible, or the window invisible
    # does not do anything if geben visible_value is nil
    # @param [Boolean] visible_value determines whether the workbook shall be visible
    def visible= visible_value
      return if visible_value.nil?
      @excel.visible = true if visible_value
      self.window_visible = @excel.Visible ? visible_value : true
    end

    # returns true, if the window of the workbook is set to visible, false otherwise
    def window_visible
      @ole_workbook.Windows(@ole_workbook.Name).Visible
    end

    # makes the window of the workbook visible or invisible
    # @param [Boolean] visible_value determines whether the window of the workbook shall be visible
    def window_visible= visible_value
      retain_saved do
        @ole_workbook.Windows(@ole_workbook.Name).Visible = visible_value if @ole_workbook.Windows.Count > 0
      end
    end

    # @return [Boolean] true, if the full workbook names and excel Instances are identical, false otherwise
    def == other_book
      other_book.is_a?(Workbook) &&
        @excel == other_book.excel &&
        self.filename == other_book.filename
    end

    def self.books
      bookstore.books
    end

    # @private
    def self.bookstore   
      @@bookstore ||= Bookstore.new
    end

    # @private
    def bookstore    
      self.class.bookstore
    end

    # @private
    def to_s    
      self.filename.to_s
    end

    # @private
    def inspect    
      '#<Workbook: ' + ('not alive ' unless alive?).to_s + (File.basename(self.filename) if alive?).to_s + " #{@ole_workbook} #{@excel}" + '>'
    end

    # @private
    def self.excel_class    
      @excel_class ||= begin
        module_name = self.parent_name
        "#{module_name}::Excel".constantize
      rescue NameError => e
        # trace "excel_class: NameError: #{e}"
        Excel
      end
    end

    # @private
    def self.worksheet_class    
      @worksheet_class ||= begin
        module_name = self.parent_name
        "#{module_name}::Worksheet".constantize
      rescue NameError => e
        Worksheet
      end
    end

    # @private
    def self.address_class    
      @address_class ||= begin
        module_name = self.parent_name
        "#{module_name}::Address".constantize
      rescue NameError => e
        Address
      end
    end

    # @private
    def excel_class        
      self.class.excel_class
    end

    # @private
    def worksheet_class        
      self.class.worksheet_class
    end

    # @private
    def address_class        
      self.class.address_class
    end

    include MethodHelpers

  private

    def method_missing(name, *args) 
      if name.to_s[0,1] =~ /[A-Z]/
        raise ObjectNotAlive, 'method missing: workbook not alive' unless alive?
        if ::JRUBY_BUG_ERRORMESSAGE 
          begin
            @ole_workbook.send(name, *args)
          rescue Java::OrgRacobCom::ComFailException 
            raise VBAMethodMissingError, "unknown VBA property or method #{name.inspect}"
          end
        else
          begin
            @ole_workbook.send(name, *args)
          rescue NoMethodError 
            raise VBAMethodMissingError, "unknown VBA property or method #{name.inspect}"
          end
        end
      else
        super
      end
    end

  end

public

  Book = Workbook

end
