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
        :if_lockraiseed     => :readonly,       
        :if_unsaved    => :raise,
        :if_obstructed => :raise,
        :if_absent     => :raise,
        :read_only => false
      }

    class << self
      
      # opens a workbook.
      # 
      # when reopening a workbook that was opened and closed before, transparency identity is ensured:
      # same Book objects refer to the same Excel files, and vice versa
      # 
      # options: 
      # :default_excel   if the workbook was already open in an Excel instance, then open it there.
      #                  Otherwise, i.e. if the workbook was not open before or the Excel instance is not alive
      #                   :reuse (default) -> connects to a (the first opened) running Excel instance,
      #                                        excluding the hidden Excel instance, if it exists,
      #                                       otherwise opens in a new Excel instance.
      #                   :new             -> opens in a new Excel instance
      #                   <excel-instance> -> opens in the given Excel instance
      # :force_excel     no matter whether the workbook was already open
      #                   :new (default)   -> opens in a new Excel instance
      #                   <excel-instance> -> opens in the given Excel instance
      # :if_unsaved     if an unsaved workbook with the same name is open, then
      #                  :raise (default)     -> raises an exception
      #                  :forget              -> close the unsaved workbook, open the new workbook             
      #                  :accept              -> lets the unsaved workbook open                  
      #                  :alert               -> gives control to Excel
      #                  :new_excel           -> opens the new workbook in a new Excel instance
      # :if_obstructed  if a workbook with the same name in a different path is open, then
      #                  :raise (default)     -> raises an exception 
      #                  :forget              -> closes the old workbook, open the new workbook
      #                  :save                -> saves the old workbook, close it, open the new workbook
      #                  :close_if_saved      -> closes the old workbook and open the new workbook, if the old workbook is saved,
      #                                          otherwise raises an exception.
      #                  :new_excel           -> opens the new workbook in a new Excel instance   
      # :if_absent       :raise (default)     -> raises an exception     , if the file does not exists
      #                  :create              -> creates a new Excel file, if it does not exists  
      #                  
      # :read_only     opens in read-only mode         (default: false) 
      # :displayalerts enables DisplayAlerts in Excel  (default: false)
      # :visible       makes visible in Excel          (default: false)
      # if :default_excel is set, then DisplayAlerts and Visible are set only if these parameters are given

      def open(file, opts={ }, &block)
        current_options = DEFAULT_OPEN_OPTS.merge(opts)
        book = nil
        if (not (current_options[:force_excel] == :new && (not current_options[:if_locked] == :take_writable)))
          # if readonly is true, then prefer a book that is given in force_excel if this option is set
          book = bookstore.fetch(file, 
                  :prefer_writable => (not current_options[:read_only]), 
                  :prefer_excel    => (current_options[:read_only] ? current_options[:force_excel].excel : nil)) rescue nil
          if book
            if (((not current_options[:force_excel]) || (current_options[:force_excel].excel == book.excel)) &&
                 (not (book.alive? && (not book.saved) && (not current_options[:if_unsaved] == :accept))))
              book.options = DEFAULT_OPEN_OPTS.merge(opts)
              book.get_excel unless book.excel.alive?
              # if the book is opened as readonly and should be opened as writable, then close it and open the book with the new readonly mode
              book.close if (book.alive? && (not book.writable) && (not current_options[:read_only]))
              # reopens the book
              book.get_workbook(file) unless book.alive?
              return book
            end
          end
        end
        current_options[:excel] = current_options[:force_excel] ? current_options[:force_excel] : current_options[:default_excel]
        new(file, current_options, &block)
      end
    end

    def initialize(file, opts={ }, &block)
      @options = DEFAULT_OPEN_OPTS.merge(opts)      
      get_excel
      get_workbook file
      bookstore.store(self)
      if block
        begin
          yield self
        ensure
          close
        end
      end
    end

    def self.excel_class
      @excel_class ||= begin
        module_name = self.parent_name
        "#{module_name}::Excel".constantize
      rescue NameError => e
        Excel
      end
    end

    def excel_class
      self.class.excel_class
    end

    
    def get_excel
      if @options[:excel] == :reuse
        @excel = excel_class.new(:reuse => true)
      end
      @excel_options = nil
      if (not @excel)
        if @options[:excel] == :new
          @excel_options = {:displayalerts => false, :visible => false}.merge(@options)
          @excel_options[:reuse] = false
          @excel = excel_class.new(@excel_options)
        else
          @excel = @options[:excel].excel
        end
      end
      # if :excel => :new or (:excel => :reuse but could not reuse)
      #   keep the old values for :visible and :displayalerts, set them only if the parameters are given
      if (not @excel_options)
        @excel.displayalerts = @options[:displayalerts] unless @options[:displayalerts].nil?
        @excel.visible = @options[:visible] unless @options[:visible].nil?
      end
    end

    def get_workbook file
      file = @stored_filename ? @stored_filename : file
      unless File.exist?(file)
        if @options[:if_absent] == :create
          @workbook = excel_class.current.generate_workbook(file)
        else 
          raise ExcelErrorOpen, "file #{file} not found"
        end
      end
      @workbook = @excel.Workbooks.Item(File.basename(file)) rescue nil
      if @workbook then
        obstructed_by_other_book = (File.basename(file) == File.basename(@workbook.Fullname)) && 
                                   (not (RobustExcelOle::absolute_path(file) == @workbook.Fullname))
        # if book is obstructed by a book with same name and different path
        if obstructed_by_other_book then
          case @options[:if_obstructed]
          when :raise
            raise ExcelErrorOpen, "blocked by a book with the same name in a different path: #{File.basename(file)}"
          when :forget
            @workbook.Close
            @workbook = nil
            open_or_create_workbook file
          when :save
            save unless @workbook.Saved
            @workbook.Close
            @workbook = nil
            open_or_create_workbook file
          when :close_if_saved
            if (not @workbook.Saved) then
              raise ExcelErrorOpen, "workbook with the same name in a different path is unsaved: #{File.basename(file)}"
            else 
              @workbook.Close
              @workbook = nil
              open_or_create_workbook file
            end
          when :new_excel 
            @excel_options = {:displayalerts => false, :visible => false}.merge(@options)   
            @excel_options[:reuse] = false
            @excel = excel_class.new(@excel_options)
            open_or_create_workbook file
          else
            raise ExcelErrorOpen, ":if_obstructed: invalid option: #{@options[:if_obstructed]}"
          end
        else
          # book open, not obstructed by an other book, but not saved and writable
          if (not @workbook.Saved) then
            case @options[:if_unsaved]
            when :raise
              raise ExcelErrorOpen, "workbook is already open but not saved (#{File.basename(file)})"
            when :forget
              @workbook.Close
              @workbook = nil
              open_or_create_workbook file
            when :accept
              # do nothing
            when :alert
              @excel.with_displayalerts true do
                open_or_create_workbook file
              end 
            when :new_excel
              @excel_options = {:displayalerts => false, :visible => false}.merge(@options)
              @excel_options[:reuse] = false
              @excel = excel_class.new(@excel_options)
              open_or_create_workbook file
            else
              raise ExcelErrorOpen, ":if_unsaved: invalid option: #{@options[:if_unsaved]}"
            end
          end
        end
      else
        # book is not open
        open_or_create_workbook file
      end
    end

    def open_or_create_workbook file
      if ((not @workbook) || (@options[:if_unsaved] == :alert) || @options[:if_obstructed]) then
        begin
          filename = RobustExcelOle::absolute_path(file)
          begin
            workbooks = @excel.Workbooks
          rescue RuntimeError => msg
            puts "RuntimeError: #{msg.message}" 
            if msg.message =~ /failed to get Dispatch Interface/
              raise ExcelErrorOpen, "Excel instance not alive or damaged" 
            else
              raise ExcelErrorOpen, "unknown RuntimeError"
            end
          end
          workbooks.Open(filename,{ 'ReadOnly' => @options[:read_only] })
        rescue WIN32OLERuntimeError => msg
          puts "WIN32OLERuntimeError: #{msg.message}" 
          if msg.message =~ /800A03EC/
            raise ExcelErrorOpen, "open: user canceled or open error"
          else 
            raise ExcelErrorOpen, "unknown WIN32OLERuntimeError"
          end
        end   
        begin
          # workaround for bug in Excel 2010: workbook.Open does not always return 
          # the workbook with given file name
          @workbook = workbooks.Item(File.basename(filename))
        rescue WIN32OLERuntimeError
          raise ExcelErrorOpen, "cannot find the file #{File.basename(filename)}"
        end
      end
    end

    # closes the workbook, if it is alive
    #
    # options:
    #  :if_unsaved    if the workbook is unsaved
    #                      :raise (default) -> raises an exception       
    #                      :save            -> saves the workbook before it is closed                  
    #                      :forget          -> closes the workbook 
    #                      :alert           -> gives control to excel
    def close(opts = {:if_unsaved => :raise})
      if (alive? && (not @workbook.Saved) && writable) then
        case opts[:if_unsaved]
        when :raise
          raise ExcelErrorClose, "workbook is unsaved (#{File.basename(self.stored_filename)})"
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
          raise ExcelErrorClose, ":if_unsaved: invalid option: #{opts[:if_unsaved]}"
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

    # modifies a workbook such that its state (open/close, saved/unsaved, readonly/writable) remains unchanged.
    #  options:
    #  :reuse (default)  : opens closed workbooks in the Excel instance of the workbook, if it exists, reuse another Excel, otherwise          
    #  :hidden           : opens closed workbooks in one separate Excel instance that is not visible and has no displayaslerts    
    #  <excel-instance>  : opens closed workbooks in the given Excel instance
    #  :read_only        : opens the workbook unobtrusively for reading only  (default: false)
    #  :readonly_excel:  if the workbook is opened only as ReadOnly and shall be modified, then
    #                    true:  closes it and open it as writable in the Excel instance where it was open so far
    #                    false (default)   opens it as writable in another running excel instance, if it exists,
    #                                      otherwise open in a new Excel instance.
    #  :keep_open: lets the workbook open after unobtrusively opening (default: false)
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
        yield book
      ensure
        book.save if (was_not_alive_or_nil || was_saved || ((not options[:read_only]) && (not was_writable))) && (not options[:read_only]) && book && (not book.saved)
        # book was open, readonly and shoud be modified
        if (not was_not_alive_or_nil) && (not options[:read_only]) && (not was_writable) && options[:readonly_excel]
          open(file, :force_excel => book.excel, :if_obstructed => :new_excel, :read_only => true)
        end
        book.close if (was_not_alive_or_nil && (not now_alive) && (not options[:keep_open]) && book)
      end
    end

    # reopens a closed workbook
    def reopen
      self.class.open(self.stored_filename)
    end

    # renames a range
    def rename_range(name,new_name)
      begin
        item = self.Names.Item(name)
      rescue WIN32OLERuntimeError
        raise ExcelError, "name #{name} not in #{File.basename(self.stored_filename)}"  
      end
      begin
        item.Name = new_name
      rescue WIN32OLERuntimeError
        raise ExcelError, "name error in #{File.basename(self.stored_filename)}"      
      end
    end

    # returns the contents of a range with given name
    # if no contents could returned, then return default value, if a default value was provided
    #                                raise an error, otherwise
    def nvalue(name, opts = {:default => nil})
      begin
        item = self.Names.Item(name)
      rescue WIN32OLERuntimeError
        return opts[:default] if opts[:default]
        raise ExcelError, "name #{name} not in #{File.basename(self.stored_filename)}"
      end
      begin
        value = item.RefersToRange.Value
      rescue  WIN32OLERuntimeError
        begin
          sheet = self[0]
          value = sheet.Evaluate(name)
        rescue WIN32OLERuntimeError
          return opts[:default] if opts[:default]
          raise SheetError, "cannot evaluate name #{name} in sheet"
        end
      end
      if value == -2146826259
        return opts[:default] if opts[:default]
        raise SheetError, "cannot evaluate name #{name} in sheet"
      end 
      return opts[:default] if (value.nil? && opts[:default])
      value      
    end

    # set the contents of a range with given name
    def set_nvalue(name,value) 
      begin
        item = self.Names.Item(name)
      rescue WIN32OLERuntimeError
        raise ExcelError, "name #{name} not in #{File.basename(self.stored_filename)}"  
      end
      begin
        item.RefersToRange.Value = value
      rescue WIN32OLERuntimeError
        raise ExcelError, "RefersToRange error of name #{name} in #{File.basename(self.stored_filename)}"    
      end
    end

    # brings the workbook to the foreground and available for heyboard inputs, and makes the Excel instance visible
    def activate      
      @excel.visible = true
      begin
        Win32API.new("user32","SetForegroundWindow","I","I").call(@excel.hwnd)     # Excel  2010
        @workbook.Activate   # Excel 2007
      rescue WIN32OLERuntimeError
        raise ExcelError, "cannot activate"
      end
    end

    # returns whether the workbook is visible or invisible
    def visible
      @excel.Windows(@workbook.Name).Visible
    end

    # makes a workbook visible or invisible
    def visible= visible_value
      saved = @workbook.Saved
      @excel.Windows(@workbook.Name).Visible = visible_value
      save if saved 
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
 
    # simple save of a workbook.
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

    # saves a workbook with a given file name.
    #
    # options:
    #  :if_exists   if a file with the same name exists, then  
    #               :raise     -> raises an exception, dont't write the file  (default)
    #               :overwrite -> writes the file, delete the old file
    #               :alert     -> gives control to Excel
    #  :if_obstructed   if a workbook with the same name and different path is already open and blocks the saving, then
    #                  :raise (default)     -> raises an exception 
    #                  :forget              -> closes the blocking workbook
    #                  :save                -> saves the blocking workbook and closes it
    #                  :close_if_saved      -> closes the blocking workbook, if it is saved, 
    #                                          otherwise raises an exception.
    # returns true, if successfully saved, nil otherwise
    def save_as(file = nil, opts = { } )
      raise ExcelErrorSave, "Not opened for writing (opened with :read_only option)" if @options[:read_only]
      options = {
        :if_exists => :raise,
        :if_obstructed => :raise,
      }.merge(opts)
      if File.exist?(file) then
        case options[:if_exists]
        when :overwrite
          if file == self.filename
            save
          else
            begin
              File.delete(file)
            rescue Errno::EACCES
              raise ExcelErrorSave, "workbook is open and used in Excel"
            end
            blocking_workbook = 
              begin
                @excel.Workbooks.Item(File.basename(file))
              rescue WIN32OLERuntimeError => msg
                #puts "#{msg.message}"
                nil
              end
            puts "blocking_workbook: #{blocking_workbook}"
            puts "name: #{blocking_workbook.Name}" if blocking_workbook
            if blocking_workbook then
              case options[:if_obstructed]
              when :raise
                raise ExcelErrorSave, "blocked by another workbook (#{File.basename(file)})"
              when :forget
                # nothing
              when :save
                blocking_workbook.Save
              when :close_if_saved
                raise ExcelErrorSave, "blocking workbook is unsaved (#{File.basename(file)})" unless blocking_workbook.Saved
              else
                raise ExcelErrorSave, ":if_obstructed: invalid option (#{options[:if_obstructed]})"
              end
              blocking_workbook.Close
            end
            save_as_workbook(file)
          end
        when :alert 
          @excel.with_displayalerts true do
            save_as_workbook(file)
          end
        when :raise
          raise ExcelErrorSave, "workbook already exists: #{File.basename(file)}"
        else
          raise ExcelErrorSave, ":if_exists: invalid option: #{options[:if_exists]}"
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
        bookstore.store(self)
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

    # returns a sheet, if a name of a sheet or a number is given
    # returns the value of the range, if a global name of a range in the book is given 
    def [] name
      name += 1 if name.is_a? Numeric
      begin
        RobustExcelOle::Sheet.new(@workbook.Worksheets.Item(name))
      rescue WIN32OLERuntimeError => msg
        if msg.message =~ /8002000B/
          nvalue(name)
        else
          raise ExcelError, "could neither return a sheet nor a value of a range when giving the name #{name}"
        end
      end
    end

    # sets the value of a range given its name
    def []= (name, value)
      set_nvalue(name,value)
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
      ws = @workbook.Worksheets
      after_or_before, base_sheet = opts.to_a.first || [:after, Sheet.new(ws.Item(ws.Count))]
      base_sheet = base_sheet.worksheet
      sheet ? sheet.Copy({ after_or_before.to_s => base_sheet }) : @workbook.WorkSheets.Add({ after_or_before.to_s => base_sheet })
      new_sheet = RobustExcelOle::Sheet.new(@excel.Activesheet)
      begin
        new_sheet.name = new_sheet_name if new_sheet_name
      rescue WIN32OLERuntimeError => msg
        if msg.message =~ /800A03EC/ 
          raise ExcelErrorSheet, "sheet name already exists"
        else
          puts "#{msg.message}"
          raise ExcelErrorSheetUnknown
        end
      end
      new_sheet
    end      

    def self.bookstore
      @@bookstore ||= Bookstore.new
    end

    def bookstore
      self.class.bookstore
    end   

    def to_s
      "#{self.filename}"
    end

    def inspect
      "<#Book: " + "#{"not alive " unless alive?}" + "#{File.basename(@stored_filename)}" + " #{@workbook} #{@excel}"  + ">"
    end

  private

    def method_missing(name, *args)
      if name.to_s[0,1] =~ /[A-Z]/ 
        begin
          raise ExcelError, "workbook not alive" unless alive?
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