# -*- coding: utf-8 -*-

module RobustExcelOle

  class Excel

    @@hwnd2excel = {}

    # closes all Excel instances
    def self.close_all
      while current_excel do
        close_one_excel
        GC.start
        sleep 0.3
        #free_all_ole_objects
      end
    end

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
    #  :reuse          use an already running Excel instance (default: true)
    #  :displayalerts  allow display alerts in Excel         (default: false)
    #  :visible        make visible in Excel                 (default: false)
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
        result.instance_variable_set(:@excel, excel)
        WIN32OLE.const_load(excel, RobustExcelOle) unless RobustExcelOle.const_defined?(:CONSTANTS)
        @@hwnd2excel[hwnd] = result        
      end
      result
    end

    def initialize(options= {}) # :nodoc:
    end

    # generate, save and close an empty workbook
    def self.generate_workbook file_name
      excel = create                   
      excel.Workbooks.Add                           
      empty_workbook = excel.Workbooks.Item(1)          
      empty_workbook.SaveAs(file_name, XlExcel8)      
      empty_workbook.Close                             
    end

    def self.hwnd2excel(hwnd)
      @@hwnd2excel[hwnd]
    end

    def hwnd
      self.Hwnd
    end

    # returns true, if the Excel instances are alive and identical, false otherwise
    def == other_excel
      self.Hwnd == other_excel.Hwnd    if other_excel.is_a?(Excel) && self.alive? && other_excel.alive?
    end

    # returns true, if the Excel instances responds to VVA methods, false otherwise
    def alive?
      @excel.Name
      true
    rescue
      puts $!.message
      false
    end

    # set DisplayAlerts in a block
    def with_displayalerts displayalerts_value
      old_displayalerts = @excel.DisplayAlerts
      @excel.DisplayAlerts = displayalerts_value
      begin
         yield self
      ensure
        @excel.DisplayAlerts = old_displayalerts
      end
    end

    # enable DisplayAlerts in the current Excel instance
    def displayalerts= displayalerts_value
      @excel.DisplayAlerts = displayalerts_value
    end

    # return if in the current Excel instance DisplayAlerts is enabled
    def displayalerts 
      @excel.DisplayAlerts
    end

    # make the current Excel instance visible or invisible
    def visible= visible_value
      @excel.Visible = visible_value
    end

    # return if the current Excel instance is visible
    def visible 
      @excel.Visible
    end


  private

    # closes one Excel instance
    def self.close_one_excel  
      excel = current_excel
      if excel then
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
      self.HWnd #rescue Win32 nil
    end

    # set this Excel instance to nil
    def die 
      @excel = nil
    end


    def method_missing(name, *args) 
      if name.to_s[0,1] =~ /[A-Z]/ 
        begin
          @excel.send(name, *args)
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




end
