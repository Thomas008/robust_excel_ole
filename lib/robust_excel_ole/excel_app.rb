# -*- coding: utf-8 -*-

module RobustExcelOle

  class ExcelApp

    def self.create
      new(:reuse => false)
    end

    def self.reuse_if_possible
      new(:reuse => true)
    end

    
    def self.close_all
      while running_app do
        close_one_app
        GC.start
      end
    end

    def self.close_one_app
      excel = running_app
      if excel then 
        excel.Workbooks.Close
        excel_hwnd = excel.HWnd
        excel.Quit
        weak_excel_ref = WeakRef.new(excel)
        excel = nil
        GC.start
        if weak_excel_ref.weakref_alive? then
          #if WIN32OLE.ole_reference_count(weak_xlapp) > 0
          begin
            weak_xlapp.ole_free
          rescue
            #puts "could not do ole_free on #{weak_excel_ref}"
          end
        end
      end
      process_id = Win32API.new("user32", "GetWindowThreadProcessId", ["I","P"], "I")
      pid_puffer = " " * 32
      process_id.call(excel_hwnd, pid_puffer)
      pid = pid_puffer.unpack("L")[0]
      Process.kill("KILL", pid)
      anz_objekte = 0
      ObjectSpace.each_object(WIN32OLE) do |o|
        anz_objekte += 1
        #p [:ole_object_name, o, (o.Name rescue nil)]
        #trc_info :ole_type, o.ole_obj_help rescue nil
        #trc_info :obj_hwnd, o.HWnd rescue   nil
        #trc_info :obj_Parent, o.Parent rescue nil
        begin
          o.ole_free
        rescue
          #puts "olefree_error: #{$!}"
        end
      end
    end
    

    # returns nil, if no excel is running or connected to a dead Excel app
    def self.running_app
      result = WIN32OLE.connect('Excel.Application') rescue nil 
      if result 
        begin
          result.Visible    # send any method, just to see if it responds
        rescue 
          p "Windows-handle = #{result}"
          #puts "Window-handle = #{result.HWnd}"
          # dead!!!
          return nil
        end
      end
      result
    end

    def initialize(options= {})
      options = {:reuse => true}.merge(options)
      if options[:reuse] then
        @winapp = self.class.running_app
        if @winapp
          #p "bestehende Applikation wird wieder benutzt"
          @winapp.DisplayAlerts = options[:displayalerts] unless options[:displayalerts]==nil
          @winapp.Visible = options[:visible] unless options[:visible]==nil
          WIN32OLE.const_load(@winapp, RobustExcelOle) unless RobustExcelOle.const_defined?(:CONSTANTS)
          return
        end
      end

      options = {
        :displayalerts => false,
        :visible => false,
      }.merge(options)
      #p "kreiere neue application"
      @winapp = WIN32OLE.new('Excel.application')
      @winapp.DisplayAlerts = options[:displayalerts]
      @winapp.Visible = options[:visible]
      WIN32OLE.const_load(@winapp, RobustExcelOle) unless RobustExcelOle.const_defined?(:CONSTANTS)
    end

    def == other_app      
      self.HWnd == other_app.HWnd    if other_app.is_a?(ExcelApp)
    end

    def alive?
      @winapp.Name
      true
    rescue 
      puts $!.message
      false
    end
    
    def method_missing(name, *args)
      @winapp.send(name, *args)
    end

  end
end