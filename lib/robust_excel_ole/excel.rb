# -*- coding: utf-8 -*-

module RobustExcelOle

  class Excel

    attr_writer :excel_app

    @@hwnd2excel = {}

    # closes all Excel applications
    def self.close_all
      while current_excel do
        close_one_excel
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

      excel_app = nil
      if options[:reuse] then
        excel_app = options[:excel] ? options[:excel] : current_excel
        if excel_app
          excel_app.DisplayAlerts = options[:displayalerts] unless options[:displayalerts]==nil
          excel_app.Visible = options[:visible] unless options[:visible]==nil
        end
      end

      options = {
        :displayalerts => false,
        :visible => false,
      }.merge(options)
      unless excel_app
        excel_app = WIN32OLE.new('Excel.application')
        excel_app.DisplayAlerts = options[:displayalerts]
        excel_app.Visible = options[:visible]
      end

      hwnd = excel_app.HWnd
      stored = @@hwnd2excel[hwnd]

      if stored 
        result = stored
      else
        WIN32OLE.const_load(excel_app, RobustExcelOle) unless RobustExcelOle.const_defined?(:CONSTANTS)
        result = super(options)
        result.excel_app = excel_app
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


    # returns true, if the Excel applications are identical, false otherwise
    def == other_excel
      self.Hwnd == other_excel.Hwnd    if other_excel.is_a?(Excel)
    end

    # returns true, if the Excel application is alive, false otherwise
    def alive?
      @excel_app.Name
      true
    rescue
      puts $!.message
      false
    end

    # set DisplayAlerts
    def with_displayalerts displayalerts_value
      old_displayalerts = @excel_app.DisplayAlerts
      @excel_app.DisplayAlerts = displayalerts_value
      begin
         yield self
      ensure
        @excel_app.DisplayAlerts = old_displayalerts
      end
    end

    # make the current Excel application visible or invisible
    def visible= visible_value
      @excel_app.Visible = visible_value
    end

    # return if the current Excel application is visible
    def visible 
      @excel_app.Visible
    end


  private

    # closes one Excel application
    def self.close_one_excel   # :nodoc: #
      excel = current_excel
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

        @@hwnd2excel[excel_hwnd].die rescue nil
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

    def hwnd_xxxx  # :nodoc: #
      self.HWnd #rescue Win32 nil
    end

    # set this Excel application to nil
    def die  # :nodoc:
      @excel_app = nil
    end


    def method_missing(name, *args)  # :nodoc: #
      if name.to_s[0,1] =~ /[A-Z]/ 
        begin
          @excel_app.send(name, *args)
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

  def absolute_path(file)
    file = File.expand_path(file)
    file = RobustExcelOle::Cygwin.cygpath('-w', file) if RUBY_PLATFORM =~ /cygwin/
    WIN32OLE.new('Scripting.FileSystemObject').GetAbsolutePathName(file)
  end
  module_function :absolute_path

 

 class VBAMethodMissingError < RuntimeError
 end


end
