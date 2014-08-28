# -*- coding: utf-8 -*-

module RobustExcelOle

  class ExcelApp

    attr_writer :ole_app

    @@hwnd2app = {}

    def self.close_one_app
      excel = running_app
      if excel then
        excel.Workbooks.Close
        excel_hwnd = excel.HWnd
        #excel.Quit
        weak_excel_ref = WeakRef.new(excel)
        excel = nil
        GC.start
        #sleep 0.1
        if weak_excel_ref.weakref_alive? then
          #if WIN32OLE.ole_reference_count(weak_xlapp) > 0
          begin
            #weak_xlapp.ole_free
            #puts "successfully ole_freed #{weak_excel_ref}"
          rescue
            puts "could not do ole_free on #{weak_excel_ref}"
          end
        end
      end

      free_all_ole_objects

      return
      process_id = Win32API.new("user32", "GetWindowThreadProcessId", ["I","P"], "I")
      pid_puffer = " " * 32
      process_id.call(excel_hwnd, pid_puffer)
      pid = pid_puffer.unpack("L")[0]
      Process.kill("KILL", pid)
    end

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
          puts "olefree_error: #{$!}"
          #puts $!.backtrace.first(9).join "\n"
        end
      end
      #puts "went through #{anz_objekte} OLE objects"
    end




    # returns nil, if no excel is running or connected to a dead Excel app
    def self.running_app
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

    def self.close_all
      while running_app do
        close_one_app
        GC.start
        sleep 0.3
        #free_all_ole_objects
      end
    end

    def self.create
      new(:reuse => false)
    end

    def self.reuse_if_possible
      new(:reuse => true)
    end

    def self.new(options= {})
      options = {:reuse => true}.merge(options)

      ole_app = nil
      if options[:reuse] then
        ole_app = running_app
        if ole_app
          #p "bestehende Applikation wird wieder benutzt"
          ole_app.DisplayAlerts = options[:displayalerts] unless options[:displayalerts]==nil
          ole_app.Visible = options[:visible] unless options[:visible]==nil
        end
      end

      options = {
        :displayalerts => false,
        :visible => false,
      }.merge(options)
      #p "kreiere neue application"
      unless ole_app
        ole_app = WIN32OLE.new('Excel.application')
        ole_app.DisplayAlerts = options[:displayalerts]
        ole_app.Visible = options[:visible]
      end

      hwnd = ole_app.HWnd
      stored = @@hwnd2app[hwnd]

      if stored 
        result = stored
      else
        WIN32OLE.const_load(ole_app, RobustExcelOle) unless RobustExcelOle.const_defined?(:CONSTANTS)
        result = super(options)
        result.ole_app = ole_app
        @@hwnd2app[hwnd] = result
      end
      result
    end


    def initialize(options= {})
    end


    def hwnd_xxxx
      self.HWnd #rescue Win32 nil
    end

    def == other_app
      self.hwnd == other_app.hwnd    if other_app.is_a?(ExcelApp)
    end

    def alive?
      @ole_app.Name
      true
    rescue
      puts $!.message
      false
    end

    def method_missing(name, *args)
      @ole_app.send(name, *args)
    end

  end
end
