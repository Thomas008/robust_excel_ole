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

  class Excel < RangeOwners
    attr_accessor :ole_excel
    attr_accessor :created
    attr_accessor :workbook

    # setter methods are implemented below
    attr_reader :visible
    attr_reader :displayalerts
    attr_reader :calculation
    attr_reader :screenupdating

    alias ole_object ole_excel

    @@hwnd2excel = {}

    # creates a new Excel instance
    # @param [Hash] options the options
    # @option options [Variant] :displayalerts
    # @option options [Boolean] :visible
    # @option options [Symbol]  :calculation
    # @option options [Boolean] :screenupdating
    # @return [Excel] a new Excel instance
    def self.create(options = {})
      new(options.merge(:reuse => false))
    end

    # connects to the current (first opened) Excel instance, if such a running Excel instance exists
    # returns a new Excel instance, otherwise
    # @option options [Variant] :displayalerts
    # @option options [Boolean] :visible
    # @option options [Symbol] :calculation
    # @option options [Boolean] :screenupdating
    # @return [Excel] an Excel instance
    def self.current(options = {})
      new(options.merge(:reuse => true))
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
      ole_xl = win32ole_excel unless win32ole_excel.nil?
      options = { :reuse => true }.merge(options)
      ole_xl = current_excel if options[:reuse] == true
      ole_xl ||= WIN32OLE.new('Excel.Application')
      hwnd = ole_xl.HWnd
      stored = hwnd2excel(hwnd)
      if stored && stored.alive?
        result = stored
      else
        result = super(options)
        result.instance_variable_set(:@ole_excel, ole_xl)
        WIN32OLE.const_load(ole_xl, RobustExcelOle) unless RobustExcelOle.const_defined?(:CONSTANTS)
        @@hwnd2excel[hwnd] = WeakRef.new(result)
      end

      unless options.is_a? WIN32OLE
        begin
          reused = options[:reuse] && stored && stored.alive?
          unless reused
            options = { :displayalerts => :if_visible, :visible => false, :screenupdating => true }.merge(options)
          end
          result.visible = options[:visible] unless options[:visible].nil?
          result.displayalerts = options[:displayalerts] unless options[:displayalerts].nil?
          result.calculation = options[:calculation] unless options[:calculation].nil?
          result.screenupdating = options[:screenupdating] unless options[:screenupdating].nil?
          result.created = !reused
        end
      end
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
        opts = {
          :visible => @visible || false,
          :displayalerts => @displayalerts || :if_visible
        }.merge(opts)
        @ole_excel = WIN32OLE.new('Excel.Application')
        self.visible = opts[:visible]
        self.displayalerts = opts[:displayalerts]
        self.calculation = opts[:calculation]
        if opts[:reopen_workbooks]
          books = workbook_class.books
          books.each do |book|
            book.reopen if !book.alive? && book.excel.alive? && book.excel == self
          end
        end
      end
      self
    end

  private

    # returns a Win32OLE object that represents a Excel instance to which Excel connects
    # connects to the first opened Excel instance
    # if this Excel instance is being closed, then Excel creates a new Excel instance
    # @private
    def self.current_excel 
      result = begin
                 WIN32OLE.connect('Excel.Application')
               rescue
                 nil
               end
      if result
        begin
          result.Visible # send any method, just to see if it responds
        rescue
          trace 'dead excel ' + (begin
                                   "Window-handle = #{result.HWnd}"
                                 rescue
                                   'without window handle'
                                 end)
          return nil
        end
      end
      result
    end

  public

    # retain the saved status of all workbooks
    def retain_saved_workbooks
      saved = []
      @ole_excel.Workbooks.each { |w| saved << w.Saved }
      begin
        yield self
      ensure
        i = 0
        @ole_excel.Workbooks.each do |w|
          w.Saved = saved[i] if saved[i]
          i += 1
        end
      end
    end

    def self.contains_unsaved_workbooks?
      !Excel.current.unsaved_workbooks.empty?
    end

    # returns unsaved workbooks (win32ole objects)
    def unsaved_workbooks
      unsaved_workbooks = []
      begin
        @ole_excel.Workbooks.each { |w| unsaved_workbooks << w unless w.Saved || w.ReadOnly }
      rescue RuntimeError => msg
        raise ExcelDamaged, 'Excel instance not alive or damaged' if msg.message =~ /failed to get Dispatch Interface/
      end
      unsaved_workbooks
    end

    # closes workbooks
    # @option options [Symbol] :if_unsaved :raise, :save, :forget, :alert, Proc
    #  :if_unsaved    if unsaved workbooks are open in an Excel instance
    #                      :raise (default) -> raises an exception
    #                      :save            -> saves the workbooks before closing
    #                      :forget          -> closes the Excel instance without saving the workbooks
    #                      :alert           -> let Excel do it
    def close_workbooks(options = { :if_unsaved => :raise })
      return unless alive?

      weak_wkbks = @ole_excel.Workbooks
      unless unsaved_workbooks.empty?
        case options[:if_unsaved]
        when Proc
          options[:if_unsaved].call(self, unsaved_workbooks)
        when :raise
          raise UnsavedWorkbooks, 'Excel contains unsaved workbooks'
        when :alert
          # nothing
        when :forget
          unsaved_workbooks.each { |m| m.Saved = true }
        when :save
          unsaved_workbooks.each { |m| m.Save }
        else
          raise OptionInvalid, ":if_unsaved: invalid option: #{options[:if_unsaved].inspect}"
        end
      end
      begin
        @ole_excel.Workbooks.Close
      rescue WIN32OLERuntimeError => msg
        if msg.message =~ /800A03EC/
          raise ExcelREOError, 'user canceled or runtime error'
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
    def self.close_all(options = { :if_unsaved => :raise }, &blk)
      options[:if_unsaved] = blk if blk
      finished_number = error_number = overall_number = 0
      first_error = nil
      finishing_action = proc do |excel|
        if excel
          begin
            overall_number += 1
            finished_number += excel.close(:if_unsaved => options[:if_unsaved])
          rescue
            first_error = $!
            #trace "error when finishing #{$!}"
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
        excel = begin
                  new(WIN32OLE.connect('Excel.Application'))
                rescue
                  nil
                end
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
    def close(options = { :if_unsaved => :raise })
      finishing_living_excel = alive?
      if finishing_living_excel
        hwnd = (begin
                  @ole_excel.HWnd
                rescue
                  nil
                end)
        close_workbooks(:if_unsaved => options[:if_unsaved])
        @ole_excel.Quit
        if false && defined?(weak_wkbks) && weak_wkbks.weakref_alive?
          weak_wkbks.ole_free
        end
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
          begin
            Process.kill('KILL', pid)
          rescue
            # trace "kill_error: #{$!}"
          end
        end
        @@hwnd2excel.delete(hwnd)
        if weak_xl.weakref_alive?
          # if WIN32OLE.ole_reference_count(weak_xlapp) > 0
          begin
            weak_xl.ole_free
          rescue
            # trace "weakref_probl_olefree"
          end
        end
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
          if p.name == 'EXCEL.EXE'
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
      processes = WIN32OLE.connect('winmgmts:\\\\.').InstancesOf('win32_process')
      processes.select { |p| p.name == 'EXCEL.EXE' }.size
    end

    # returns all Excel objects for all Excel instances opened with RobustExcelOle,
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
      # excel_processes = processes.select{ |p| p.name == "EXCEL.EXE" && pid2excel.include?(p.processid)}
      # excel_processes.map{ |p| Excel.new(pid2excel[p.processid]) }
      processes.select { |p| Excel.new(pid2excel[p.processid]) if p.name == 'EXCEL.EXE' && pid2excel.include?(p.processid) }
      result = []
      processes.each do |p|
        next unless p.name == 'EXCEL.EXE'

        if pid2excel.include?(p.processid)
          excel = pid2excel[p.processid]
          result << excel
        end
        # how to connect to an (interactively opened) Excel instance and get a WIN32OLE object?
        # after that, lift it to an Excel object
      end
      result
    end

    # @private
    def excel
      self
    end

    # @private
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
      result = []
      @@hwnd2excel.each do |_hwnd,wr_excel|
        excel = wr_excel.__getobj__ if wr_excel.weakref_alive?
        result << excel.unsaved_workbooks
      end
      result
    end

    # @private
    def print_workbooks
      self.Workbooks.each { |w| trace "#{w.Name} #{w}" }
    end

    # generates, saves, and closes empty workbook
    # @private
    def generate_workbook file_name
      raise FileNameNotGiven, 'filename is nil' if file_name.nil?

      self.Workbooks.Add
      empty_ole_workbook = self.Workbooks.Item(self.Workbooks.Count)
      filename = General.absolute_path(file_name).tr('/','\\')
      unless File.exist?(filename)
        begin
          empty_ole_workbook.SaveAs(filename)
        rescue WIN32OLERuntimeError => msg
          # if msg.message =~ /SaveAs/ and msg.message =~ /Workbook/ then
          raise FileNotFound, "could not save workbook with filename #{file_name.inspect}"
          # else
          #  # todo some time: find out when this occurs :
          #  raise UnexpectedREOError, "unknown WIN32OLERuntimeError with filename #{file_name.inspect}: \n#{msg.message}"
          # end
        end
      end
      empty_ole_workbook
    end

    # sets DisplayAlerts in a block
    def with_displayalerts displayalerts_value
      old_displayalerts = displayalerts
      self.displayalerts = displayalerts_value
      begin
        yield self
      ensure
        self.displayalerts = old_displayalerts if alive?
      end
    end

    # makes the current Excel instance visible or invisible
    def visible= visible_value
      @ole_excel.Visible = @visible = visible_value
      @ole_excel.DisplayAlerts = @visible if @displayalerts == :if_visible
    end

    # enables DisplayAlerts in the current Excel instance
    def displayalerts= displayalerts_value
      @displayalerts = displayalerts_value
      @ole_excel.DisplayAlerts = @displayalerts == :if_visible ? @ole_excel.Visible : displayalerts_value
    end

    # sets ScreenUpdating
    def screenupdating= screenupdating_value
      @ole_excel.ScreenUpdating = @screenupdating = screenupdating_value
    end

    # sets calculation mode
    # retains the saved-status of the workbooks when set to manual
    def calculation= calculation_mode
      return if calculation_mode.nil?
      @calculation = calculation_mode
      calc_mode_changable = @ole_excel.Workbooks.Count > 0 && @ole_excel.Calculation.is_a?(Integer)
      if calc_mode_changable
        begin
        #if calculation_mode == :manual
          saved = []
          (1..@ole_excel.Workbooks.Count).each { |i| saved << @ole_excel.Workbooks(i).Saved }
        #end
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
        @ole_excel.Calculation =
          calculation_mode == :automatic ? XlCalculationAutomatic : XlCalculationManual
        #if calculation_mode == :manual
          (1..@ole_excel.Workbooks.Count).each { |i| @ole_excel.Workbooks(i).Saved = true if saved[i - 1] }
        #end
      end
    end

    # VBA method overwritten
    def Calculation= calculation_vba_mode
      case calculation_vba_mode
      when XlCalculationManual
        @calculation = :manual
      when XlCalculationAutomatic
        @calculation = :automatic
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
      self.class.new(@ole_excel, options)
    end

    def set_options(options)
      for_this_instance(options)
    end    

    # set options in all workbooks
    def for_all_workbooks(options)
      ole_workbooks = begin
        @ole_excel.Workbooks
      rescue WIN32OLERuntimeError => msg
        if msg.message =~ /failed to get Dispatch Interface/
          raise ExcelDamaged, 'Excel instance not alive or damaged'
        else
          raise ExcelREOError, 'workbooks could not be determined'
        end
      end
      ole_workbooks.each do |ole_workbook|
        workbook_class.open(ole_workbook).for_this_workbook(options)
      end
    end

    def workbooks
      @ole_excel.Workbooks.map {|ole_workbook| ole_workbook.to_reo }
    end

    def each_workbook
      @ole_excel.Workbooks.each { |ole_workbook| yield ole_workbook.to_reo } 
    end

    def each_workbook_with_index(offset = 0)
      i = offset
      @ole_excel.Workbooks.each do |ole_workbook|
        yield ole_workbook.to_reo, i
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

    # returns the value of a range
    # @param [String] name the name of a range
    # @returns [Variant] the value of the range
    def [] name
      namevalue_glob(name)
    end

    # sets the value of a range
    # @param [String]  name  the name of the range
    # @param [Variant] value the contents of the range
    def []=(name, value)
      set_namevalue_glob(name,value, :color => 42) # 42 - aqua-marin, 7-green
    end

    # @private
    def to_s            
      '#<Excel: ' + hwnd.to_s + ('not alive' unless alive?).to_s + '>'
    end

    # @private
    def inspect         
      to_s
    end

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

    include MethodHelpers

  private

    # @private
    def method_missing(name, *args) 
      if name.to_s[0,1] =~ /[A-Z]/
        begin
          raise ObjectNotAlive, 'method missing: Excel not alive' unless alive?

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

  public

  Application = Excel

end

class WIN32OLE
  include Enumerable
end
