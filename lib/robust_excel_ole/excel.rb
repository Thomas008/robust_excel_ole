# -*- coding: utf-8 -*-

require 'weakref'
require 'Win32API'

def ka
  Excel.kill_all
end

module RobustExcelOle

  # This class essentially wraps a Win32Ole Application object. 
  # You can apply all VBA methods (starting with a capital letter) 
  # that you would apply for an Application object. 
  # See https://docs.microsoft.com/en-us/office/vba/api/excel.application(object)#methods

  class Excel < VbaObjects

    attr_reader :ole_excel
    attr_reader :properties
    attr_reader :address_tool

    alias ole_object ole_excel

    @@hwnd2excel = {}

    PROPERTIES = [:visible, :displayalerts, :calculation, :screenupdating]

    # creates a new Excel instance
    # @param [Hash] options the options
    # @option options [Variant] :displayalerts
    # @option options [Boolean] :visible
    # @option options [Symbol]  :calculation
    # @option options [Boolean] :screenupdating
    # @return [Excel] a new Excel instance
    def self.create(options = {})
      new(options.merge(reuse: false))
    end

    # connects to the current (first opened) Excel instance, if such a running Excel instance exists
    # returns a new Excel instance, otherwise
    # @option options [Variant] :displayalerts
    # @option options [Boolean] :visible
    # @option options [Symbol] :calculation
    # @option options [Boolean] :screenupdating
    # @return [Excel] an Excel instance
    def self.current(options = {})
      new(options.merge(reuse: true))
    end

    # returns an Excel instance
    # @param [Win32Ole] (optional) a WIN32OLE object representing an Excel instance
    # @param [Hash] options the options
    # @option options [Boolean] :reuse
    # @option options [Boolean] :visible
    # @option options [Variant] :displayalerts
    # @option options [Boolean] :screenupdating
    # @option options [Symbol]  :calculation
    # options:
    #  :reuse            connects to an already running Excel instance (true) or
    #                    creates a new Excel instance (false)  (default: true)
    #  :visible          makes the Excel visible               (default: false)
    #  :displayalerts    enables or disables DisplayAlerts     (true, false, :if_visible (default))
    #  :calculation      calculation mode is being forced to be manual (:manual) or automatic (:automtic)
    #                    or is not being forced (default: nil)
    #  :screenupdating  turns on or off screen updating (default: true)
    # @return [Excel] an Excel instance
    def self.new(win32ole_excel = nil, options = {})
      if win32ole_excel.is_a? Hash
        options = win32ole_excel
        win32ole_excel = nil
      end
      options = { reuse: true }.merge(options)
      ole_xl = if !win32ole_excel.nil? 
        win32ole_excel
      elsif options[:reuse] == true
        current_ole_excel
      end
      connected = (not ole_xl.nil?) && win32ole_excel.nil?
      ole_xl ||= WIN32OLE.new('Excel.Application')
      hwnd = ole_xl.Hwnd
      stored = hwnd2excel(hwnd)
      if stored && stored.alive?
        result = stored
      else
        result = super(options)
        result.instance_variable_set(:@ole_excel, ole_xl)
        WIN32OLE.const_load(ole_xl, RobustExcelOle) unless RobustExcelOle.const_defined?(:CONSTANTS)
        @@hwnd2excel[hwnd] = WeakRef.new(result)
      end
      reused = options[:reuse] && stored && stored.alive? 
      options = { displayalerts: :if_visible, visible: false, screenupdating: true }.merge(options) unless reused || connected
      result.set_options(options)        
      result
    end    

    def initialize(options = {}); end

    # reopens a closed Excel instance
    # @param [Hash] opts the options
    # @option opts [Boolean] :reopen_workbooks
    # @option opts [Boolean] :displayalerts
    # @option opts [Boolean] :visible
    # @option opts [Boolean] :calculation
    # options: reopen_workbooks (default: false): reopen the workbooks in the Excel instances
    # :visible (default: false), :displayalerts (default: :if_visible), :calculation (default: false)
    # @return [Excel] an Excel instance
    def recreate(opts = {})
      unless alive?
        opts = {visible: false, displayalerts: :if_visible}.merge(
               {visible: @properties[:visible], displayalerts: @properties[:displayalerts]}).merge(opts)        
        @ole_excel = WIN32OLE.new('Excel.Application')
        set_options(opts)
        if opts[:reopen_workbooks]
          workbook_class.books.each{ |book| book.reopen if !book.alive? && book.excel.alive? && book.excel == self }
        end
      end
      self
    end

    # @private
    def address_tool
      raise(ExcelREOError, "Excel contains no workbook") unless @ole_excel.Workbooks.Count > 0
      @address_tool ||= begin
        address_string = @ole_excel.Workbooks.Item(1).Worksheets.Item(1).Cells.Item(1,1).Address(true,true,XlR1C1)
        address_tool_class.new(address_string)
      end
    end

  private

    # retain the saved status of all workbooks
    def retain_saved_workbooks
      saved_stati = @ole_excel.Workbooks.map { |w| w.Saved }
      begin
        yield self
      ensure
        @ole_excel.Workbooks.zip(saved_stati) { |w,s| w.Saved = s }
      end
    end

    def ole_workbooks
      @ole_excel.Workbooks
    rescue WIN32OLERuntimeError, Java::OrgRacobCom::ComFailException => msg
      if msg.message =~ /failed to get Dispatch Interface/
        raise ExcelDamaged, "Excel instance not alive or damaged\n#{$!.message}"
      else
        raise ExcelREOError, "workbooks could not be determined\n#{$!.message}"
      end
    end

  public

    # @private
    # returns unsaved workbooks (win32ole objects)
    def self.contains_unsaved_workbooks?
      !Excel.current.unsaved_workbooks.empty?
    end

    # @private
    # returns unsaved workbooks (win32ole objects)   
    def unsaved_workbooks
      @ole_excel.Workbooks.reject { |w| w.Saved || w.ReadOnly }
    rescue RuntimeError => msg
      raise ExcelDamaged, "Excel instance not alive or damaged\n#{$!.message}" if msg.message =~ /failed to get Dispatch Interface/
    end

    # closes workbooks
    # @option options [Symbol] :if_unsaved :raise, :save, :forget, :alert, Proc
    #  :if_unsaved    if unsaved workbooks are open in an Excel instance
    #                      :raise (default) -> raises an exception
    #                      :forget          -> closes the Excel instance without saving the workbooks
    #                      :save            -> saves the workbooks before closing
    #                      :alert           -> let Excel do it
    # @private
    def close_workbooks(options = { if_unsaved: :raise })
      return unless alive?

      weak_wkbks = @ole_excel.Workbooks
      unless unsaved_workbooks.empty?
        case options[:if_unsaved]
        when Proc
          options[:if_unsaved].call(self, unsaved_workbooks)
        when :raise
          raise UnsavedWorkbooks, "Excel contains unsaved workbooks" +
          "\nHint: Use option :if_unsaved with values :forget and :save to close the 
           Excel instance without or with saving the unsaved workbooks before, respectively"
        when :alert
          # nothing
        when :forget
          unsaved_workbooks.each { |m| m.Saved = true }
        when :save
          unsaved_workbooks.each { |m| m.Save }
        else
          raise OptionInvalid, ":if_unsaved: invalid option: #{options[:if_unsaved].inspect}" +
          "\nHint: Valid values are :raise, :forget, :save and :alert"
        end
      end

      begin
        @ole_excel.Workbooks.Close
      rescue
        if $!.message =~ /kann nicht zugeordnet werden/ or $!.message =~ /800A03EC/
          raise ExcelREOError, "user canceled or runtime error"
        else
          raise UnexpectedREOError, "unknown WIN32OLERuntimeError: #{msg.message}"
        end
      end
      weak_wkbks = nil
      weak_wkbks = @ole_excel.Workbooks
      weak_wkbks = nil
     end

    # closes all Excel instances
    # @return [Integer,Integer] number of closed Excel instances, number of errors
    # remark: the returned number of closed Excel instances is valid only for known Excel instances
    # if there are unknown Excel instances (opened not via this class), then they are counted as 1
    # @param [Hash] options the options
    # @option options [Symbol]  :if_unsaved :raise, :save, :forget, or :alert
    # options:
    #  :if_unsaved    if unsaved workbooks are open in an Excel instance
    #                      :raise (default) -> raises an exception
    #                      :save            -> saves the workbooks before closing
    #                      :forget          -> closes the excel instance without saving the workbooks
    #                      :alert           -> give control to Excel
    # @option options [Proc] block
    def self.close_all(options = { if_unsaved: :raise }, &blk)
      options[:if_unsaved] = blk if blk
      finished_number = error_number = overall_number = 0
      first_error = nil
      finishing_action = proc do |excel|
        if excel
          begin
            overall_number += 1
            finished_number += excel.close(if_unsaved: options[:if_unsaved])
          rescue
            first_error = $!
            error_number += 1
          end
        end
      end

      # known Excel-instances
      @@hwnd2excel.each do |hwnd, wr_excel|
        if wr_excel.weakref_alive?
          excel = wr_excel.__getobj__
          if excel.alive?
            excel.displayalerts = false
            finishing_action.call(excel)
          end
        else
          @@hwnd2excel.delete(hwnd)
        end
      end

      # unknown Excel-instances
      old_error_number = error_number
      9.times do |_index|
        sleep 0.1
        excel = new(WIN32OLE.connect('Excel.Application')) rescue nil
        finishing_action.call(excel) if excel
        free_all_ole_objects unless (error_number > 0) && (options[:if_unsaved] == :raise)
        break unless excel
        break if error_number > old_error_number # + 3
      end
      raise first_error if ((options[:if_unsaved] == :raise) && first_error) || (first_error.class == OptionInvalid)
      [finished_number, error_number]
    end

    # closes the Excel
    # @param [Hash] options the options
    # @option options [Symbol] :if_unsaved :raise, :save, :forget, :alert
    # @option options [Boolean] :hard
    #  :if_unsaved    if unsaved workbooks are open in an Excel instance
    #                      :raise (default) -> raises an exception
    #                      :save            -> saves the workbooks before closing
    #                      :forget          -> closes the Excel instance without saving the workbooks
    #                      :alert           -> Excel takes over    
    def close(options = { if_unsaved: :raise })
      finishing_living_excel = alive?
      if finishing_living_excel
        hwnd = @ole_excel.Hwnd rescue nil
        close_workbooks(if_unsaved: options[:if_unsaved])
        @ole_excel.Quit
        weak_wkbks.ole_free if false && defined?(weak_wkbks) && weak_wkbks.weakref_alive?
        weak_xl = WeakRef.new(@ole_excel)
      else
        weak_xl = nil
      end
      @ole_excel = nil
      GC.start
      sleep 0.1
      if finishing_living_excel
        if hwnd
          process_id = Win32API.new('user32', 'GetWindowThreadProcessId', %w[I P], 'I')
          pid_puffer = ' ' * 32
          process_id.call(hwnd, pid_puffer)
          pid = pid_puffer.unpack('L')[0]
          Process.kill('KILL', pid) rescue nil
        end
        @@hwnd2excel.delete(hwnd)
        weak_xl.ole_free if weak_xl.weakref_alive?
      end
      weak_xl ? 1 : 0
    end

    # frees all OLE objects in the object space
    # @private
    def self.free_all_ole_objects
      anz_objekte = 0
      ObjectSpace.each_object(WIN32OLE) do |o|
        anz_objekte += 1
        # trace "#{anz_objekte} name: #{(o.Name rescue (o.Count rescue "no_name"))} ole_object_name: #{(o.ole_object_name rescue nil)} type: #{o.ole_type rescue nil}"
        # trace [:Name, (o.Name rescue (o.Count rescue "no_name"))]
        # trace [:ole_object_name, (o.ole_object_name rescue nil)]
        # trace [:methods, (o.ole_methods rescue nil)] unless (o.Name rescue false)
        # trace o.ole_type rescue nil
        begin
          o.ole_free
          # trace "olefree OK"
        rescue
          # trace "olefree_error: #{$!}"
          # trace $!.backtrace.first(9).join "\n"
        end
      end
      # trace "went through #{anz_objekte} OLE objects"
    end

    def self.init
      @@hwnd2excel = {}
    end

    # kill all Excel instances
    # @return [Integer] number of killed Excel processes
    def self.kill_all
      number = 0
      WIN32OLE.connect('winmgmts:\\\\.').InstancesOf('win32_process').each do |p|
        begin
          if p.Name == 'EXCEL.EXE'
            Process.kill('KILL', p.processid)
            number += 1
          end
        rescue
          # trace "kill error: #{$!}"
        end
      end
      init
      number
    end

    def self.excels_number
      WIN32OLE.connect('winmgmts:\\\\.').InstancesOf('win32_process').select { |p| p.Name == 'EXCEL.EXE' }.size
    end

    def self.known_excels_number
      @@hwnd2excel.size
    end

  private

    # returns a Win32OLE object that represents a Excel instance to which Excel connects
    # connects to the first opened Excel instance
    # if this Excel instance is being closed, then Excel creates a new Excel instance
    def self.current_ole_excel   
      if ::CONNECT_EXCEL_JRUBY_BUG
        result = known_excel_instance
        if result.nil?
          if excels_number > 0
            dummy_ole_workbook = WIN32OLE.connect(General.absolute_path('___dummy_workbook.xls')) rescue nil
            result = dummy_ole_workbook.Application
            visible_status = result.Visible
            dummy_ole_workbook.Close
            dummy_ole_workbook = nil
            result.Visible = visible_status
          end
        end
      else
        result = WIN32OLE.connect('Excel.Application') rescue nil
      end
      if result
        begin
          result.Visible # send any method, just to see if it responds
        rescue
          trace 'dead excel ' + (begin
                                 "Window-handle = #{result.Hwnd}"
                                 rescue
                                 'without window handle'
                                 end)
          return nil
        end
      end
      result
    end


    # returns an Excel instance opened with RobustExcelOle
    def self.known_excel_instance
      @@hwnd2excel.each do |hwnd, wr_excel|
        if wr_excel.weakref_alive?
          excel = wr_excel.__getobj__
          return excel if excel.alive?
        end
      end
      nil
    end

    def self.hwnd2excel(hwnd)
      excel_weakref = @@hwnd2excel[hwnd]
      if excel_weakref
        if excel_weakref.weakref_alive?
          excel_weakref.__getobj__
        else
          trace 'dead reference to an Excel'
          begin
            @@hwnd2excel.delete(hwnd)
            nil
          rescue
            trace "Warning: deleting dead reference failed! (hwnd: #{hwnd.inspect})"
          end
        end
      end
    end

  public

    # returns all Excel objects for all Excel instances opened with RobustExcelOle
    def self.known_excel_instances
      pid2excel = {}
      @@hwnd2excel.each do |hwnd,wr_excel|
        next unless wr_excel.weakref_alive?
        excel = wr_excel.__getobj__
        process_id = Win32API.new('user32', 'GetWindowThreadProcessId', %w[I P], 'I')
        pid_puffer = ' ' * 32
        process_id.call(hwnd, pid_puffer)
        pid = pid_puffer.unpack('L')[0]
        pid2excel[pid] = excel
      end
      processes = WIN32OLE.connect('winmgmts:\\\\.').InstancesOf('win32_process')     
      processes.map{ |p| pid2excel[p.ProcessId] if p.Name == 'EXCEL.EXE'}.compact
    end

    # @private
    def excel
      self
    end

    # @private
    def hwnd 
      self.Hwnd
    rescue
      nil
    end

    # @private
    def self.print_hwnd2excel
      @@hwnd2excel.each do |hwnd,wr_excel|
        excel_string = (wr_excel.weakref_alive? ? wr_excel.__getobj__.to_s : 'weakref not alive')
        printf("hwnd: %8i => excel: %s\n", hwnd, excel_string)
      end
      @@hwnd2excel.size
    end

    # returns true, if the Excel instances are alive and identical, false otherwise
    def == other_excel
      self.Hwnd == other_excel.Hwnd if other_excel.is_a?(Excel) && alive? && other_excel.alive?
    end

    # returns true, if the Excel instances responds to VBA methods, false otherwise
    def alive?
      @ole_excel.Name
      true
    rescue
      # trace $!.message
      false
    end

    # returns unsaved workbooks in known (not opened by user) Excel instances
    # @private
    def self.unsaved_known_workbooks       
      @@hwnd2excel.values.map{ |wk_exl| wk_exl.__getobj__.unsaved_workbooks if wk_exl.weakref_alive? }.compact.flatten
    end


    # @private
    def print_workbooks
      self.Workbooks.each { |w| trace "#{w.Name} #{w}" }
    end

    # @private
    def generate_workbook file_name  # :deprecated: #
      workbook_class.open(file_name, if_absent: :create, force: {excel: self})
    end

    # sets DisplayAlerts in a block
    def with_displayalerts displayalerts_value
      old_displayalerts = @properties[:displayalerts]
      self.displayalerts = displayalerts_value
      begin
        yield self
      ensure
        self.displayalerts = old_displayalerts if alive?
      end
    end

    # makes the current Excel instance visible or invisible
    def visible= visible_value
      return if visible_value.nil?
      @ole_excel.Visible = @properties[:visible] = visible_value
      @ole_excel.DisplayAlerts = @properties[:visible] if @properties[:displayalerts] == :if_visible
    end

    # enables DisplayAlerts in the current Excel instance
    def displayalerts= displayalerts_value
      return if displayalerts_value.nil?
      @properties[:displayalerts] = displayalerts_value
      @ole_excel.DisplayAlerts = @properties[:displayalerts] == :if_visible ? @ole_excel.Visible : displayalerts_value
    end

    # sets ScreenUpdating
    def screenupdating= screenupdating_value
      return if screenupdating_value.nil?
      @ole_excel.ScreenUpdating = @properties[:screenupdating] = screenupdating_value
    end

    # sets calculation mode
    # retains the saved-status of the workbooks when set to manual
    def calculation= calculation_mode
      return if calculation_mode.nil?
      @properties[:calculation] = calculation_mode
      calc_mode_changable = @ole_excel.Workbooks.Count > 0 && @ole_excel.Calculation.is_a?(Integer)
      return unless calc_mode_changable
      retain_saved_workbooks do
        begin
          best_wb_to_make_visible = @ole_excel.Workbooks.sort_by {|wb|
            score =
              (wb.Saved    ? 0 : 40) +  # an unsaved workbooks is most likely the main workbook
              (wb.ReadOnly ? 0 : 20) +  # the main wb is usually writable
              case wb.Name.split(".").last.downcase
                when "xlsm" then 10  # the main workbook is more likely to have macros
                when "xls"  then  8
                when "xlsx" then  4
                when "xlam" then -2  # libraries are not normally the main workbook
                else 0
              end
            score
          }.last
          best_wb_to_make_visible.Windows(1).Visible = true
        rescue => e
          trace "error setting calculation=#{calculation_mode} msg: " + e.message
          trace e.backtrace
          # continue on errors here, failing would usually disrupt too much
        end
        @ole_excel.CalculateBeforeSave = false
        @ole_excel.Calculation = calculation_mode == :automatic ? XlCalculationAutomatic : XlCalculationManual
      end
    end

    # VBA method overwritten
    def Calculation= calculation_vba_mode
      case calculation_vba_mode
      when XlCalculationManual
        @properties[:calculation] = :manual
      when XlCalculationAutomatic
        @properties[:calculation] = :automatic
      end
      @ole_excel.Calculation = calculation_vba_mode
    end

    # sets calculation mode in a block
    def with_calculation(calculation_mode)
      return unless calculation_mode
      old_calculation_mode = @ole_excel.Calculation
      begin
        self.calculation = calculation_mode
        yield self
      ensure
        @ole_excel.Calculation = old_calculation_mode if @ole_excel.Calculation.is_a?(Integer)
      end
    end

    # set options in this Excel instance
    def for_this_instance(options)
      set_options(options)
    end

    def set_options(options)      
      @properties ||= { }
      PROPERTIES.each do |property|
        method = (property.to_s + '=').to_sym
        send(method, options[property]) 
      end
    end
  
    # set options in all workbooks
    def for_all_workbooks(options)
      each_workbook(options)
    end

    def workbooks
      ole_workbooks.map {|ole_workbook| workbook_class.new(ole_workbook) }
    end

    # traverses over all workbooks and sets options if provided
    def each_workbook(opts = { })
      ole_workbooks.each do |ow|
        wb = workbook_class.new(ow, opts)
        block_given? ? (yield wb) : wb
      end
    end

    def each_workbook_with_index(opts = { }, offset = 0)
      i = offset
      ole_workbooks.each do |ow| 
        yield workbook_class.new(ow, opts), i 
        i += 1
      end
    end

    def focus
      self.visible = true
      # if not Windows10 then
      Win32API.new('user32','SetForegroundWindow','I','I').call(@ole_excel.Hwnd)
      # else
      # Win32API.new("user32","SetForegroundWindow","","I").call
      # end
    end
    
    # @private
    # returns active workbook
    def workbook
      @workbook ||= workbook_class.new(@ole_excel.ActiveWorkbook) if @ole_excel.Workbooks.Count > 0
    end    

    alias active_workbook workbook

    # @private
    def to_s            
      "#<Excel: #{hwnd}#{ ("not alive" unless alive?)}>"
    end

    # @private
    def inspect         
      to_s
    end

    using ParentRefinement
    using StringRefinement

    # @private
    def self.workbook_class  
      @workbook_class ||= begin
        module_name = parent_name
        "#{module_name}::Workbook".constantize
      rescue NameError => e
        Workbook
      end
    end

    # @private
    def workbook_class       
      self.class.workbook_class
    end

    # @private
    def self.address_tool_class  
      @address_tool_class ||= begin
        module_name = parent_name
        "#{module_name}::AddressTool".constantize
      rescue NameError => e
        AddressTool
      end
    end

    # @private
    def address_tool_class       
      self.class.address_tool_class
    end


    include MethodHelpers

  private

    def method_missing(name, *args) 
      super unless name.to_s[0,1] =~ /[A-Z]/
      raise ObjectNotAlive, 'method missing: Excel not alive' unless alive?
      if ::ERRORMESSAGE_JRUBY_BUG
        begin
          @ole_excel.send(name, *args)
        rescue Java::OrgRacobCom::ComFailException
          raise VBAMethodMissingError, "unknown VBA property or method #{name.inspect}"
        end
      else
        begin
          @ole_excel.send(name, *args)
        rescue NoMethodError
          raise VBAMethodMissingError, "unknown VBA property or method #{name.inspect}"
        end
      end
    end
  end

public

  # @private
  class ExcelDamaged < ExcelREOError               
  end

  # @private
  class UnsavedWorkbooks < ExcelREOError           
  end


  Application = Excel

end
