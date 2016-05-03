# -*- coding: utf-8 -*-

require 'timeout'

module RobustExcelOle      

  class Excel < REOCommon    

    attr_accessor :ole_excel

    alias ole_object ole_excel

    @@hwnd2excel = {}    

    # creates a new Excel instance
    # @return [Excel] a new Excel instance
    def self.create
      new(:reuse => false)
    end

    # uses the current Excel instance (connects), if such a running Excel instance exists
    # creates a new one, otherwise 
    # @return [Excel] an Excel instance
    def self.current
      new(:reuse => true)
    end

    # returns an Excel instance  
    # given a WIN32OLE object representing an Excel instance, or a Hash representing options:
    # @param [Hash] options the options
    # @option options [Boolean] :reuse  
    # @option options [Boolean] :displayalerts 
    # @option options [Boolean] :visible 
    # options: 
    #  :reuse          connects to an already running Excel instance (true) or
    #                  creates a new Excel instance (false)   (default: true)
    #  :displayalerts  allows display alerts in Excel         (default: false)
    #  :visible        makes the Excel visible                (default: false)
    #  if :reuse => true, then DisplayAlerts and Visible are set only if they are given
    # @return [Excel] an Excel instance
    def self.new(options = {})
      if options.is_a? WIN32OLE
        excel = options
      else
        options = {:reuse => true}.merge(options)
        if options[:reuse] == true then
          excel = current_excel
        end
      end
      if not (excel)
        excel = WIN32OLE.new('Excel.Application')
        options = {
          :displayalerts => false,
          :visible => false,
        }.merge(options)
      end
      unless options.is_a? WIN32OLE
        excel.DisplayAlerts = options[:displayalerts] unless options[:displayalerts].nil?
        excel.Visible = options[:visible] unless options[:visible].nil?
      end

      hwnd = excel.HWnd
      stored = hwnd2excel(hwnd)

      if stored 
        result = stored
      else
        result = super(options)
        result.instance_variable_set(:@ole_excel, excel)
        WIN32OLE.const_load(excel, RobustExcelOle) unless RobustExcelOle.const_defined?(:CONSTANTS)
        @@hwnd2excel[hwnd] = WeakRef.new(result)
      end
      result
    end

    def initialize(options= {}) # :nodoc: #
      @excel = self
    end

    # reopens a closed Excel instance
    # @param [Hash] opts the options
    # @option opts [Boolean] :reopen_workbooks
    # @option opts [Boolean] :displayalerts
    # @option opts [Boolean] :visible
    # options: reopen_workbooks (default: false): reopen the workbooks in the Excel instances
    #          :visible (default: false), :displayalerts (default: false)
    # @return [Excel] an Excel instance
    def recreate(opts = {})      
      unless self.alive?
        opts = {
          :displayalerts => @displayalerts ? @displayalerts : false,
          :visible => @visible ? @visible : false
        }.merge(opts)
        new_excel = WIN32OLE.new('Excel.Application')
        new_excel.DisplayAlerts = opts[:displayalerts]
        new_excel.Visible = opts[:visible]
        @ole_excel = new_excel 
        if opts[:reopen_workbooks]
          books = book_class.books
          books.each do |book|
            book.reopen if ((not book.alive?) && book.excel.alive? && book.excel == self)
          end        
        end
      end
      self 
    end

  private
    
    # returns an Excel instance to which one 'connect' was possible
    def self.current_excel   # :nodoc: #
      #p "current_excel:"
      result = WIN32OLE.connect('Excel.Application') rescue nil
      #p "result: #{result}"
      if result
        begin
          result.Visible    # send any method, just to see if it responds
        rescue 
          trace "dead excel " + ("Window-handle = #{result.HWnd}" rescue "without window handle")
          return nil
        end
      end
      #p "result: #{result}"
      result
    end

  public

    # closes all Excel instances
    # @param [Hash] options the options
    # @option options [Symbol]  :if_unsaved :raise, :save, :forget, :alert, or :keep_open
    # @option options [Boolean] :hard
    # @option options [Boolean] :kill_if_timeout
    # options:
    #  :if_unsaved    if unsaved workbooks are open in an Excel instance
    #                      :raise (default) -> raises an exception       
    #                      :save            -> saves the workbooks before closing
    #                      :forget          -> closes the excel instance without saving the workbooks 
    #                      :keep_open       -> let the workbooks open
    #                      :alert           -> give control to Excel
    #  :hard          closes Excel instances soft (default: false), or, additionally kills the Excel processes hard (true)
    #  :kill_if_timeout:  kills Excel instances hard if the closing process exceeds a certain time limit (default: true)
    # @raise ExcelError if time limit has exceeded, some Excel instance cannot be closed, or
    #                   unsaved workbooks exist and option :if_unsaved is :raise
    def self.close_all(options={})
      options = {
        :if_unsaved => :raise,
        :hard => false,
        :kill_if_timeout => false
      }.merge(options)      
      excels_number = excel_processes.size
      timeout = false
      begin
        status = Timeout::timeout(5) {
          while current_excel do
            close_one_excel(options)
            GC.start
            sleep 0.3
            current_excels_number = excel_processes.size
            if current_excels_number == excels_number && excels_number > 0
              raise ExcelError, "some Excel instance cannot be closed"
            end
            excels_number = current_excels_number
          end   
        }
      rescue Timeout::Error
        raise ExcelError, "close_all: timeout" unless options[:kill_if_timeout]
        timeout = true
      end
      kill_all if options[:hard] || (timeout && options[:kill_if_timeout])
      init
    end

    def self.init
      @@hwnd2excel = {}
    end

  private

    def self.manage_unsaved_workbooks(excel, options)     # :nodoc: #
      unsaved_workbooks = []
      begin
        excel.Workbooks.each {|w| unsaved_workbooks << w unless (w.Saved || w.ReadOnly)}
      rescue RuntimeError => msg
        trace "RuntimeError: #{msg.message}" 
        raise ExcelErrorOpen, "Excel instance not alive or damaged" if msg.message =~ /failed to get Dispatch Interface/
      end
      unless unsaved_workbooks.empty? 
        case options[:if_unsaved]
        when :raise
          raise ExcelErrorClose, "Excel contains unsaved workbooks"
        when :save
          unsaved_workbooks.each do |workbook|
            workbook.Save
          end
        when :forget
          # nothing
        when :keep_open
          return
        when :alert
          begin
            excel.DisplayAlerts = true
            yield
          ensure
            begin
              excel.DisplayAlerts = false 
            rescue RuntimeError => msg
              trace "RuntimeError: #{msg.message}" if msg.message =~ /failed to get Dispatch Interface/
            end
          end
          return
        else
          raise ExcelErrorClose, ":if_unsaved: invalid option: #{options[:if_unsaved].inspect}"
        end
      end
      yield
    end

    # closes one Excel instance to which one was connected
    def self.close_one_excel(options={})
      excel = current_excel
      return unless excel
      manage_unsaved_workbooks(excel, options) do
        weak_ole_excel = WeakRef.new(excel)
        excel = nil
        close_excel_ole_instance(weak_ole_excel.__getobj__)
      end
    end

    def self.close_excel_ole_instance(ole_excel)  # :nodoc: #
      @@hwnd2excel.delete(ole_excel.Hwnd)
      excel = ole_excel
      ole_excel = nil
      begin
        excel.Workbooks.Close
        excel_hwnd = excel.HWnd
        excel.Quit
        weak_excel_ref = WeakRef.new(excel)
        excel = nil
        GC.start
        sleep 0.2
        if weak_excel_ref.weakref_alive? then
          begin
            weak_excel_ref.ole_free
            trace "successfully ole_freed #{weak_excel_ref}"
          rescue
            trace "could not do ole_free on #{weak_excel_ref}"
          end
        end
        @@hwnd2excel.delete(excel_hwnd)
      rescue => e
        trace "Error when closing Excel: #{e.message}"
        #t e.backtrace
      end
      free_all_ole_objects
    end

    # frees all OLE objects in the object space
    def self.free_all_ole_objects     # :nodoc: #
      anz_objekte = 0
      ObjectSpace.each_object(WIN32OLE) do |o|
        anz_objekte += 1
        #t [:Name, (o.Name rescue (o.Count rescue "no_name"))]
        #t [:ole_object_name, (o.ole_object_name rescue nil)]
        #t [:methods, (o.ole_methods rescue nil)] unless (o.Name rescue false)
        #t o.ole_type rescue nil
        begin
          o.ole_free
          #trace "olefree OK"
        rescue
          #trace "olefree_error: #{$!}"
          #trace $!.backtrace.first(9).join "\n"
        end
      end
      trace "went through #{anz_objekte} OLE objects"
    end   

  public

    # closes the Excel
    # @param [Hash] options the options
    # @option options [Symbol] :if_unsaved :raise, :save, :forget, or :keep_open
    # @option options [Boolean] :hard      
    #  :if_unsaved    if unsaved workbooks are open in an Excel instance
    #                      :raise (default) -> raises an exception       
    #                      :save            -> saves the workbooks before closing
    #                      :forget          -> closes the Excel instance without saving the workbooks 
    #                      :keep_open       -> keeps the Excel instance open 
    #  :hard          kill the Excel instance hard (default: false) 
    def close(options = {})
      options = {
        :if_unsaved => :raise,
        :hard => false
      }.merge(options)
      self.class.manage_unsaved_workbooks(@ole_excel, options) do 
        close_excel(options)
      end
    end

  private

    def close_excel(options)
      excel = @ole_excel
      begin
        excel.Workbooks.Close
      rescue WIN32OLERuntimeError => msg
        raise ExcelUserCanceled, "close: canceled by user" if msg.message =~ /80020009/ && 
              options[:if_unsaved] == :alert && (not self.unsaved_workbooks.empty?)
      end
      excel_hwnd = excel.HWnd
      excel.Quit
      weak_excel_ref = WeakRef.new(excel)
      excel = nil
      GC.start
      sleep 0.2
      if weak_excel_ref.weakref_alive? then
        begin
          weak_excel_ref.ole_free
          trace "successfully ole_freed #{weak_excel_ref}"
        rescue => msg
          trace "#{msg.message}"
          trace "could not do ole_free on #{weak_excel_ref}"
        end
      end
      @@hwnd2excel.delete(excel_hwnd)      
      if options[:hard] then
        Excel.free_all_ole_objects
        process_id = Win32API.new("user32", "GetWindowThreadProcessId", ["I","P"], "I")
        pid_puffer = " " * 32
        process_id.call(excel_hwnd, pid_puffer)
        pid = pid_puffer.unpack("L")[0]
        Process.kill("KILL", pid) rescue nil   
      end
    end

  public

    # kill all Excel instances
    # @return [Fixnum] number of killed Excel processes
    def self.kill_all
      procs = WIN32OLE.connect("winmgmts:\\\\.")
      processes = procs.InstancesOf("win32_process")
      number = processes.select{|p| (p.name == "EXCEL.EXE")}.size
      procs.InstancesOf("win32_process").each do |p|
        Process.kill('KILL', p.processid) if p.name == "EXCEL.EXE"        
      end
      init
      number
    end

    # provide Excel objects 
    # (so far restricted to all Excel instances opened with RobustExcelOle,
    #  not for Excel instances opened by the user)
    def self.excel_processes
      pid2excel = {}
      @@hwnd2excel.each do |hwnd,wr_excel|
        excel = wr_excel.__getobj__
        process_id = Win32API.new("user32", "GetWindowThreadProcessId", ["I","P"], "I")
        pid_puffer = " " * 32
        process_id.call(hwnd, pid_puffer)
        pid = pid_puffer.unpack("L")[0]
        pid2excel[pid] = excel
      end
      procs = WIN32OLE.connect("winmgmts:\\\\.")
      processes = procs.InstancesOf("win32_process")     
      result = []
      processes.each do |p|
        if p.name == "EXCEL.EXE"
          if pid2excel.include?(p.processid)
            excel = pid2excel[p.processid]
            result << excel
          end
          # how to connect to an (interactively opened) Excel instance and get a WIN32OLE object?
          # after that, lift it to an Excel object
        end
      end
      result
    end

    def excel   # :nodoc: #
      self
    end

    def self.hwnd2excel(hwnd)   # :nodoc: #
      excel_weakref = @@hwnd2excel[hwnd]
      if excel_weakref
        if excel_weakref.weakref_alive?
          excel_weakref.__getobj__
        else
          trace "dead reference to an Excel"
          begin 
            @@hwnd2excel.delete(hwnd)
          rescue
            trace "Warning: deleting dead reference failed! (hwnd: #{hwnd.inspect})"
          end
        end
      end
    end

    def hwnd   # :nodoc: #
      self.Hwnd rescue nil
    end

    def self.print_hwnd2excel    # :nodoc: #
      @@hwnd2excel.each do |hwnd,wr_excel|
        excel_string = (wr_excel.weakref_alive? ? wr_excel.__getobj__.to_s : "weakref not alive") 
        printf("hwnd: %8i => excel: %s\n", hwnd, excel_string)
      end
      @@hwnd2excel.size
    end

    # returns true, if the Excel instances are alive and identical, false otherwise
    def == other_excel
      self.Hwnd == other_excel.Hwnd  if other_excel.is_a?(Excel) && self.alive? && other_excel.alive?
    end

    # returns true, if the Excel instances responds to VBA methods, false otherwise
    def alive?
      @ole_excel.Name
      true
    rescue
      #t $!.message
      false
    end

    
    # returns all unsaved workbooks in Excel instances
    def self.unsaved_workbooks_all    # :nodoc: #
      result = []
      @@hwnd2excel.each do |hwnd,wr_excel| 
        excel = wr_excel.__getobj__
        result << excel.unsaved_workbooks
      end
      result
    end

    # returns unsaved workbooks
    def unsaved_workbooks
      result = []
      begin
        self.Workbooks.each {|w| result << w unless (w.Saved || w.ReadOnly)}
      rescue RuntimeError => msg
        trace "RuntimeError: #{msg.message}" 
        raise ExcelErrorOpen, "Excel instance not alive or damaged" if msg.message =~ /failed to get Dispatch Interface/
      end
      result      
    end

    def print_workbooks
      self.Workbooks.each {|w| puts "#{w.Name} #{w}"}
    end


    # generates, saves, and closes empty workbook
    def generate_workbook file_name                  
      self.Workbooks.Add                           
      empty_workbook = self.Workbooks.Item(self.Workbooks.Count)          
      filename = General::absolute_path(file_name).gsub("/","\\")
      unless File.exists?(filename)
        begin
          empty_workbook.SaveAs(filename) 
        rescue WIN32OLERuntimeError => msg
          if msg.message =~ /SaveAs/ and msg.message =~ /Workbook/ then
            raise ExcelErrorSave, "could not save workbook with filename #{file_name.inspect}"
          else
            # todo some time: find out when this occurs : 
            raise ExcelErrorSaveUnknown, "unknown WIN32OELERuntimeError with filename #{file_name.inspect}: \n#{msg.message}"
          end
        end      
      end
      empty_workbook                               
    end

    # sets DisplayAlerts in a block
    def with_displayalerts displayalerts_value
      old_displayalerts = @ole_excel.DisplayAlerts
      @ole_excel.DisplayAlerts = displayalerts_value
      begin
         yield self
      ensure
        @ole_excel.DisplayAlerts = old_displayalerts if alive?
      end
    end    

    # enables DisplayAlerts in the current Excel instance
    def displayalerts= displayalerts_value
      @displayalerts = @ole_excel.DisplayAlerts = displayalerts_value
    end

    # return whether DisplayAlerts is enabled in the current Excel instance
    def displayalerts 
      @displayalerts = @ole_excel.DisplayAlerts
    end

    # makes the current Excel instance visible or invisible
    def visible= visible_value
      @visible = @ole_excel.Visible = visible_value
    end

    # returns whether the current Excel instance is visible
    def visible 
      @visible = @ole_excel.Visible
    end    

    # sets calculation mode
    def with_calculation(calculation_mode = :automatic)
      if @ole_excel.Workbooks.Count > 0
        old_calculation_mode = @ole_excel.Calculation
        old_calculation_before_save_mode = @ole_excel.CalculateBeforeSave
        @ole_excel.Calculation = calculation_mode == :automatic ? XlCalculationAutomatic : XlCalculationManual
        @ole_excel.CalculateBeforeSave = (calculation_mode == :automatic)
        begin
          yield self
        ensure
          @ole_excel.Calculation = old_calculation_mode if alive?
          @ole_excel.CalculateBeforeSave = old_calculation_before_save_mode if alive?
        end
      end
    end

    def to_s              # :nodoc: #
      "#<Excel: " + "#{hwnd}" + ("#{"not alive" unless self.alive?}") + ">"
    end

    def inspect           # :nodoc: #
      self.to_s
    end

    def self.book_class   # :nodoc: #
      @book_class ||= begin
        module_name = self.parent_name
        "#{module_name}::Book".constantize
      rescue NameError => e
        book
      end
    end

    def book_class        # :nodoc: #
      self.class.book_class
    end

    include MethodHelpers

  private

    def method_missing(name, *args)    # :nodoc: #
      if name.to_s[0,1] =~ /[A-Z]/ 
        begin          
          raise ExcelError, "method missing: Excel not alive" unless alive?
          @ole_excel.send(name, *args)
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
end
