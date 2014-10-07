# -*- coding: utf-8 -*-

module RobustExcelOle

  class Excel

    attr_writer :ole_app

    @@hwnd2app = {}

    # closes all Excel applications
    def self.close_all
      while current_app do
        close_one_app
        GC.start
        sleep 0.3
        #free_all_ole_objects
      end
    end

    # creates a new Excel application
    def self.create
      new(:reuse => false)
    end

    # uses the current Excel application (connects), if such a running Excel application exists
    # creates a new one, otherwise 
    def self.current
      new(:reuse => true)
    end

    # returns an Excel application  
    # options:
    #  :reuse          use an already running Excel application (default: true)
    #  :displayalerts  allow display alerts in Excel            (default: false)
    #  :visible        make visible in Excel                    (default: false)
    def self.new(options= {})
      options = {:reuse => true}.merge(options)

      ole_app = nil
      if options[:reuse] then
        ole_app = options[:excel] ? options[:excel] : current_app
        if ole_app
          ole_app.DisplayAlerts = options[:displayalerts] unless options[:displayalerts]==nil
          ole_app.Visible = options[:visible] unless options[:visible]==nil
        end
      end

      options = {
        :displayalerts => false,
        :visible => false,
      }.merge(options)
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

    def initialize(options= {}) # :nodoc:
    end

    # returns true, if the Excel applications are identical, false otherwise
    def == other_app
      self.hwnd == other_app.hwnd    if other_app.is_a?(Excel)
    end

    # returns true, if the Excel application is alive, false otherwise
    def alive?
      @ole_app.Name
      true
    rescue
      puts $!.message
      false
    end

  private

    # closes one Excel application
    def self.close_one_app   # :nodoc: #
      excel = current_app
      if excel then
        excel.Workbooks.Close
        excel_hwnd = excel.HWnd
        excel.Quit
        #excel.ole_free
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

        @@hwnd2app[excel_hwnd].die rescue nil
        #@@hwnd2app[excel_hwnd] = nil
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
    def self.free_all_ole_objects   # :nodoc: #
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

    # returns the current Excel application, if a running, working Excel appication exists, nil otherwise
    def self.current_app   # :nodoc: #
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

    def hwnd_xxxx  # :nodoc: #
      self.HWnd #rescue Win32 nil
    end

    # set this Excel application to nil
    def die  # :nodoc:
      @ole_app = nil
    end

    def method_missing(name, *args)  # :nodoc: #
      @ole_app.send(name, *args)
    end

  end

  def absolute_path(file)
    file = File.expand_path(file)
    file = RobustExcelOle::Cygwin.cygpath('-w', file) if RUBY_PLATFORM =~ /cygwin/
    WIN32OLE.new('Scripting.FileSystemObject').GetAbsolutePathName(file)
  end
  module_function :absolute_path

end
