# -*- coding: utf-8 -*-

require 'weakref'
require 'fiddle/import'

def ka
  Excel.kill_all
end

module User32
  # Extend this module to an importer
  extend Fiddle::Importer
  # Load 'user32' dynamic library into this importer
  dlload 'user32'
  # Set C aliases to this importer for further understanding of function signatures
  typealias 'HWND', 'HANDLE'
  typealias 'HANDLE', 'void*'
  typealias 'LPCSTR', 'const char*'
  typealias 'LPCWSTR', 'const wchar_t*'
  typealias 'UINT', 'unsigned int'
  typealias 'HANDLE', 'void*'
  typealias 'ppvObject', 'void**'
  typealias 'DWORD', 'unsigned long'
  typealias 'LPDWORD', 'DWORD*'
  typealias 'REFIID', 'GUID&'
  typealias 'GUID&', 'struct _GUID {unsigned long, unsigned short, unsigned short, unsigned *char}'
  # Import C functions from loaded libraries and set them as module functions
  extern 'DWORD GetWindowThreadProcessId(HWND, LPDWORD)'
  extern 'HWND FindWindowExA(HWND, HWND, LPCSTR, LPCSTR)'
end

module Oleacc
  # Extend this module to an importer
  extend Fiddle::Importer
  # Load 'oleacc' dynamic library into this importer
  dlload 'oleacc'
  # Set C aliases to this importer for further understanding of function signatures
  typealias 'HWND', 'HANDLE'
  typealias 'HANDLE', 'void*'
  typealias 'LPCSTR', 'const char*'
  typealias 'LPCWSTR', 'const wchar_t*'
  typealias 'UINT', 'unsigned int'
  typealias 'HANDLE', 'void*'
  typealias 'ppvObject', 'void**'
  typealias 'DWORD', 'unsigned long'
  typealias 'HRESULT', 'long'
  #typealias 'REFIID', 'const GUID'
  #typealias 'REFIID', 'const GUID*'
  typealias 'REFIID', 'IID*'
  typealias 'IID', 'GUID'
  typealias 'GUID', 'struct {unsigned long, unsigned short, unsigned short, unsigned *char}'
  # Import C functions from loaded libraries and set them as module functions
  #extern 'HRESULT AccessibleObjectFromWindow(HWND, DWORD, REFIID, ppvObject)'
  extern 'HRESULT AccessibleObjectFromWindow(HWND, DWORD, REFIID, ppvObject)'
end

typedef IID* REFIID;
     typedef const char* LPCSTR;

module RobustExcelOle

  # This class essentially wraps a Win32Ole Application object. 
  # You can apply all VBA methods (starting with a capital letter) 
  # that you would apply for an Application object. 
  # See https://docs.microsoft.com/en-us/office/vba/api/excel.application(object)#methods

  class Excel < VbaObjects

    include Enumerable

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
          workbook_class.books.each{ |book| book.open if !book.alive? && book.excel.alive? && book.excel == self }
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
        raise ExcelDamaged, "Excel instance not alive or damaged"
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
      raise ExcelDamaged, "Excel instance not alive or damaged" if msg.message =~ /failed to get Dispatch Interface/
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
          pid_puffer = ' ' * 32
          User32::GetWindowThreadProcessId(hwnd, pid_puffer)         
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

    def self.instance_count
      WIN32OLE.connect('winmgmts:\\\\.').InstancesOf('win32_process').select { |p| p.Name == 'EXCEL.EXE' }.size
    end

    def self.known_instance_count
      @@hwnd2excel.size
    end

    # returns running Excel instances
    # currently restricted to visible Excel instances with at least one workbook
=begin
    def self.known_running_instances
      find_windows_method = Win32API.new('user32', 'FindWindowExA', %w[I I P P P], 'I')
      #acc_obj_fr_window = Win32API.new('oleacc', 'AccessibleObjectFromWindow', %w[I I I P P], 'I')
      acc_obj_fr_window = Win32API.new('oleacc', 'AccessibleObjectFromWindow', %w[I P P P P], 'I')
      acc_obj_addr = nil
      hwnd = 0
      win32ole_excel_instances = []
      loop do
        hwnd_puffer = ' ' * 32
        find_windows_method.call(0, hwnd, "XLMAIN", "", hwnd_puffer)
        hwnd = hwnd_puffer.unpack('L')[0]
        break if hwnd == 0
        hwnd2_puffer = ' ' * 32
        find_windows_method.call(hwnd, 0, "XLDESK", "", hwnd2_puffer)
        hwnd2 = hwnd2_puffer.unpack('L')[0]
        hwnd3_puffer = ' ' * 32
        find_windows_method.call(hwnd2, 0, "EXCEL7", "", hwnd3_puffer)
        hwnd3 = hwnd3_puffer.unpack('L')[0]
        status_puffer = ' ' * 32
        #acc_obj_fr_window.call(hwnd3, '&HFFFFFFF0', '&H20400', acc_obj_addr, status_puffer)
        acc_obj_fr_window.call(hwnd3, 0xFFFFFFF0, 0x20400, acc_obj_addr, status_puffer)
        status = status_puffer.unpack('L')
        acc_obj = acc_obj_addr.unpack('L')[0]
        win32ole_excel_instances << acc_obj.Application if status == 0   # == '&H0'
      end
      win32ole_excel_instances.map{|w| w.to_reo}
    end
=end

    def self.known_running_instances
      win32ole_excel_instances = []
      hwnd = 0
      loop do
        hwnd = User32::FindWindowExA(0, hwnd, "XLMAIN", nil).to_i
        break if hwnd == 0
        hwnd2 = User32::FindWindowExA(hwnd, 0, "XLDESK", nil).to_i
        hwnd3 = User32::FindWindowExA(hwnd2, 0, "EXCEL7", nil).to_i
        acc_obj_addr_puffer = ' ' * 32
        status = Oleacc::AccessibleObjectFromWindow(hwnd3, 0xFFFFFFF0, 0x20400, acc_obj_addr_puffer)
        if status == 0 # == '&H0'
          acc_obj = acc_obj_addr_puffer.unpack('L')[0]
          win32ole_excel_instances << acc_obj.Application 
        end
      end
      win32ole_excel_instances.map{|w| w.to_reo}
    end

    # returns a running Excel instance opened with RobustExcelOle
    def self.known_running_instance     
      self.known_running_instances.first
    end

    # @return [Enumerator] known running Excel instances
    def self.known_running_instances
      pid2excel = {}
      @@hwnd2excel.each do |hwnd,wr_excel|
        next unless wr_excel.weakref_alive?
        excel = wr_excel.__getobj__
        pid_puffer = ' ' * 32
        User32::GetWindowThreadProcessId(hwnd, pid_puffer)
        pid = pid_puffer.unpack('L')[0]
        pid2excel[pid] = excel
      end
      processes = WIN32OLE.connect('winmgmts:\\\\.').InstancesOf('win32_process')     
      processes.map{ |p| pid2excel[p.ProcessId] if p.Name == 'EXCEL.EXE'}.compact.lazy.each
    end

    class << self
      alias excels_number instance_count                  # :deprecated: #
      alias known_excels_number known_instance_count      # :deprecated: #
      alias known_excel_instance known_running_instance   # :deprecated: #
      alias known_excel_instances known_running_instances # :deprecated: #
    end

  private

    # returns a Win32OLE object that represents a Excel instance to which Excel connects
    # connects to the first opened Excel instance
    # if this Excel instance is being closed, then Excel creates a new Excel instance
    def self.current_ole_excel   
      if ::CONNECT_EXCEL_JRUBY_BUG
        result = known_running_instance
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
    def set_options(options)      
      @properties ||= { }
      PROPERTIES.each do |property|
        method = (property.to_s + '=').to_sym
        send(method, options[property]) 
      end
    end

    alias for_this_instance set_options  # :deprecated: #    

    # @return [Enumerator] traversing all workbook objects
    def each
      if block_given?
        ole_workbooks.lazy.each do |ole_workbook|
          yield workbook_class.new(ole_workbook)
        end
      else
        to_enum(:each).lazy
      end
    end

    # @return [Array] all workbook objects
    def workbooks
      to_a
    end

    # traverses all workbooks and sets options if provided
    def each_workbook(opts = { })
      ole_workbooks.lazy.each do |ow|
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

    alias for_all_workbooks each_workbook   # :deprecated: #

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
