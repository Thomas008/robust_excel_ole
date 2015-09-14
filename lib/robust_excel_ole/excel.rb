# -*- coding: utf-8 -*-

module RobustExcelOle

  class Excel

    @@hwnd2excel = {}

    # creates a new Excel instance
    def self.create
      new(:reuse => false)
    end

    # uses the current Excel instance (connects), if such a running Excel instance exists
    # creates a new one, otherwise 
    def self.current
      new(:reuse => true)
    end

    # returns an Excel instance  
    # options:
    #  :reuse          uses an already running Excel instance (default: true)
    #  :displayalerts  allows display alerts in Excel         (default: false)
    #  :visible        makes the Excel visible                (default: false)
    #  if :reuse => true, then DisplayAlerts and Visible are set only if they are given
    def self.new(options= {})
      options = {:reuse => true}.merge(options)
      if options[:reuse] then
        excel = current_excel
      end
      if not (excel)
        excel = WIN32OLE.new('Excel.Application')
        options = {
          :displayalerts => false,
          :visible => false,
        }.merge(options)
      end
      excel.DisplayAlerts = options[:displayalerts] unless options[:displayalerts].nil?
      excel.Visible = options[:visible] unless options[:visible].nil?

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

    def self.get_excel_processes
      procs = WIN32OLE.connect("winmgmts:\\\\.")
      procs.InstancesOf("win32_process").each do |p|
        puts "name:#{p.name.to_s} process_id:#{p.processid}"  if p.name == "EXCEL.EXE"       
      end
    end

    def self.kill_excel_processes
      procs = WIN32OLE.connect("winmgmts:\\\\.")
      procs.InstancesOf("win32_process").each do |p|
        puts "name:#{p.name.to_s} process_id:#{p.processid}"  if p.name == "EXCEL.EXE" 
        Process.kill('KILL', p.processid) if p.name == "EXCEL.EXE"        
      end
      #wmi = WIN32OLE.connect("winmgmts://")
      #processes = wmi.ExecQuery("select * from win32_process where commandline like '%excel.exe\"% /automation %'")
      #for process in processes 
      #  Process.kill('KILL', process.ProcessID.to_i)
      #end
    end

    def reanimate 
      # - generated Excel instance differs from all other Excel Instances
      #   (but this is done anyway with Excel.create?!)
      # - keep the old properties: visible, dispayalerts
      # - necessary or even possible? 
      #   traverse hwnd2excel:
      #   find all Excel objects with the old hwnd
      #   (but for each hwnd I have not all Excel objects that refere to the Excel instance with this hwnd)
      #   assign them to the new Excel object

      #excel = self.class.create
      #new(:reuse => false, :visible => @ole_excel.Visible, :displayalerts => @ole_excel.Displayalerts)
      #self
      #excel = new(:reuse => false, :visible => @ole_excel.Visible, :displayalerts => @ole_excel.Displayalerts)
      #@excel = self
    end

    def self.print_hwnd2excel
      @@hwnd2excel.each do |hwnd,wr_excel|
        excel_string = (wr_excel.weakref_alive? ? wr_excel.__getobj__.to_s : "not alive") 
        puts "hwnd: #{hwnd} => excel: #{excel_string}"
      end
    end

    # closes all Excel instances
    # options:
    #  :if_unsaved    if unsaved workbooks are open in an Excel instance
    #                      :raise (default) -> raises an exception       
    #                      :save            -> saves the workbooks before closing
    #                      :forget          -> closes the excel instance without saving the workbooks 
    #                      :alert           -> gives control to Excel
    #  :hard          kills the Excel instances hard (default: false) 
    def self.close_all(options={})
      options = {
        :if_unsaved => :raise,
        :hard => false
      }.merge(options)
      if options[:hard]
        kill_excel_processes
      else
        while current_excel do
          #current_excel.close(options)
          close_one_excel
          GC.start
          sleep 0.3
          # free_all_ole_objects if options[:hard] ???
        end
      end
    end

    # close the Excel
    #  :if_unsaved    if unsaved workbooks are open in an Excel instance
    #                      :raise (default) -> raises an exception       
    #                      :save            -> saves the workbooks before closing
    #                      :forget          -> closes the Excel instance without saving the workbooks 
    #                      :alert           -> gives control to Excel
    #  :hard          kill the Excel instance hard (default: false) 
    def close(options = {})
      options = {
        :if_unsaved => :raise,
        :hard => false
      }.merge(options)
      unsaved_books = self.unsaved_workbooks
      unless unsaved_books.empty? 
        case options[:if_unsaved]
        when :raise
          raise ExcelErrorClose, "Excel contains unsaved workbooks"
        when :save
          unsaved_workbooks.each do |workbook|
            workbook.Save
          end
          close_excel(:hard => options[:hard])
        when :forget
          close_excel(:hard => options[:hard])
        when :alert
          with_displayalerts true do
            unsaved_workbooks.each do |workbook|
              workbook.Save
            end
            close_excel(:hard => options[:hard])
          end
        else
          raise ExcelErrorClose, ":if_unsaved: invalid option: #{options[:if_unsaved].inpect}"
        end
      else
        close_excel(:hard => options[:hard])
      end
      raise ExcelUserCanceled, "close: canceled by user" if options[:if_unsaved] == :alert && self.unsaved_workbooks
    end

  private

    def close_excel(options)
      excel = @ole_excel
      excel.Workbooks.Close
      excel_hwnd = excel.HWnd
      excel.Quit
      weak_excel_ref = WeakRef.new(excel)
      excel = nil
      GC.start
      sleep 0.2
      if weak_excel_ref.weakref_alive? then
        #if WIN32OLE.ole_reference_count(weak_xlapp) > 0
        begin
          weak_excel_ref.ole_free
          puts "successfully ole_freed #{weak_excel_ref}"
        rescue
          puts "could not do ole_free on #{weak_excel_ref}"
        end
      end
      hwnd2excel(excel_hwnd).die rescue nil
      #@@hwnd2excel[excel_hwnd] = nil
      #Excel.free_all_ole_objects
      if options[:hard] then
        process_id = Win32API.new("user32", "GetWindowThreadProcessId", ["I","P"], "I")
        pid_puffer = " " * 32
        process_id.call(excel_hwnd, pid_puffer)
        pid = pid_puffer.unpack("L")[0]
        Process.kill("KILL", pid)    
      end
    end

  public

    def excel
      self
    end

    # empty workbook is generated, saved and closed 
    def generate_workbook file_name                  
      self.Workbooks.Add                           
      empty_workbook = self.Workbooks.Item(self.Workbooks.Count)          
      filename = RobustExcelOle::absolute_path(file_name).gsub("/","\\")
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

    def self.hwnd2excel(hwnd)
      excel_weakref = @@hwnd2excel[hwnd]
      if excel_weakref
        if excel_weakref.weakref_alive?
          excel_weakref.__getobj__
        else
          puts "dead reference to an Excel"
          begin 
            @@hwnd2excel.delete(hwnd)
          rescue
            puts "Warning: deleting dead reference failed! (hwnd: #{hwnd.inspect})"
          end
        end
      end
    end

    def hwnd
      self.Hwnd
    end

    # returns true, if the Excel instances are alive and identical, false otherwise
    def == other_excel
      self.Hwnd == other_excel.Hwnd    if other_excel.is_a?(Excel) && self.alive? && other_excel.alive?
    end

    # returns true, if the Excel instances responds to VBA methods, false otherwise
    def alive?
      @ole_excel.Name
      true
    rescue
      #puts $!.message
      false
    end

    def print_workbooks
      self.Workbooks.each {|w| puts "#{w.Name} #{w}"}
    end

    def unsaved_workbooks
      result = []
      begin
        self.Workbooks.each {|w| result << w unless (w.Saved || w.ReadOnly)}
      rescue RuntimeError => msg
        puts "RuntimeError: #{msg.message}" 
        raise ExcelErrorOpen, "Excel instance not alive or damaged" if msg.message =~ /failed to get Dispatch Interface/
      end
      result
    end
    # yields different WIN32OLE objects than book.workbook
    #self.class.extend Enumerable
    #self.class.map {|w| (not w.Saved)}

    # sets DisplayAlerts in a block
    def with_displayalerts displayalerts_value
      old_displayalerts = @ole_excel.DisplayAlerts
      @ole_excel.DisplayAlerts = displayalerts_value
      begin
         yield self
      ensure
        @ole_excel.DisplayAlerts = old_displayalerts
      end
    end

    # enables DisplayAlerts in the current Excel instance
    def displayalerts= displayalerts_value
      @ole_excel.DisplayAlerts = displayalerts_value
    end

    # return if in the current Excel instance DisplayAlerts is enabled
    def displayalerts 
      @ole_excel.DisplayAlerts
    end

    # makes the current Excel instance visible or invisible
    def visible= visible_value
      @ole_excel.Visible = visible_value
    end

    # returns whether the current Excel instance is visible
    def visible 
      @ole_excel.Visible
    end

    def to_s
      "#<Excel: " + "#{hwnd_xxxx}" + ("#{"not alive" unless self.alive?}") + ">"
    end

    def inspect
      self.to_s
    end

  private

    # closes one Excel instance
    def self.close_one_excel(options={})
      excel = current_excel
      if excel then
        weak_ole_excel = WeakRef.new(excel)
        excel = nil
        close_excel_ole_instance(weak_ole_excel.__getobj__)
      end
    end

    def self.close_excel_ole_instance(ole_excel)
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
        #sleep 0.1
        if weak_excel_ref.weakref_alive? then
          #if WIN32OLE.ole_reference_count(weak_xlapp) > 0
          begin
            weak_excel_ref.ole_free
            puts "successfully ole_freed #{weak_excel_ref}"
          rescue
            puts "could not do ole_free on #{weak_excel_ref}"
          end
        end

        hwnd2excel(excel_hwnd).die rescue nil
        #@@hwnd2excel[excel_hwnd] = nil

      rescue => e
        puts "Error when closing Excel: " + e.message
        #puts e.backtrace
      end


      free_all_ole_objects

      return
      process_id = Win32API.new("user32", "GetWindowThreadProcessId", ["I","P"], "I")
      pid_puffer = " " * 32
      process_id.call(excel_hwnd, pid_puffer)
      pid = pid_puffer.unpack("L")[0]
      Process.kill("KILL", pid)
    end

    # frees all OLE objects in the object space
    def self.free_all_ole_objects   
      anz_objekte = 0
      ObjectSpace.each_object(WIN32OLE) do |o|
        anz_objekte += 1
        #p [:Name, (o.Name rescue (o.Count rescue "no_name"))]
        #p [:ole_object_name, (o.ole_object_name rescue nil)]
        #p [:methods, (o.ole_methods rescue nil)] unless (o.Name rescue false)
        #puts o.ole_type rescue nil
        #trc_info :obj_hwnd, o.HWnd rescue   nil
        #trc_info :obj_Parent, o.Parent rescue nil
        begin
          o.ole_free
          #puts "olefree OK"
        rescue
          #puts "olefree_error: #{$!}"
          #puts $!.backtrace.first(9).join "\n"
        end
      end
      puts "went through #{anz_objekte} OLE objects"
    end

    # returns the current Excel instance
    def self.current_excel   # :nodoc: #
      result = WIN32OLE.connect('Excel.Application') rescue nil
      if result
        begin
          result.Visible    # send any method, just to see if it responds
        rescue 
          puts "dead excel app " + ("Window-handle = #{result.HWnd}" rescue "without window handle")
          return nil
        end
      end
      result
    end

    def hwnd_xxxx  
      self.HWnd rescue nil
    end

    # sets this Excel instance to nil
    def die 
      @ole_excel = nil
    end

    def method_missing(name, *args) 
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
