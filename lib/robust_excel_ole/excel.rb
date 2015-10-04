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

  private
    
    # returns an Excel instance to which one 'connect' was possible
    def self.current_excel   # :nodoc: #
      result = WIN32OLE.connect('Excel.Application') rescue nil
      if result
        begin
          result.Visible    # send any method, just to see if it responds
        rescue 
          t "dead excel " + ("Window-handle = #{result.HWnd}" rescue "without window handle")
          return nil
        end
      end
      result
    end

  public

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
      init
      n = 0
      # kill all after 20 repetitions
      while current_excel && n<20 do
        close_one_excel(options)
        GC.start
        sleep 0.3
        n = n+1
      end
      kill_all if options[:hard]
    end

    def self.init
      @@hwnd2excel = {}
    end

  private

    def self.manage_unsaved_workbooks(this_excel, ole_excel, options)
      #this_excel = excel == 0 ? self : excel
      unsaved_workbooks = []
      begin
        this_excel.Workbooks.each {|w| unsaved_workbooks << w unless (w.Saved || w.ReadOnly)}
      rescue RuntimeError => msg
        t "RuntimeError: #{msg.message}" 
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
        when :alert
          ole_excel.DisplayAlerts = true
        else
          raise ExcelErrorClose, ":if_unsaved: invalid option: #{options[:if_unsaved].inspect}"
        end
      end
    end

    # closes one Excel instance to which one was connected
    def self.close_one_excel(options={})
      excel = current_excel
      return unless excel
      manage_unsaved_workbooks(excel, excel, options)
      weak_ole_excel = WeakRef.new(excel)
      excel = nil
      close_excel_ole_instance(weak_ole_excel.__getobj__)
    end

    def self.close_excel_ole_instance(ole_excel)
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
          #if WIN32OLE.ole_reference_count(weak_xlapp) > 0
          begin
            weak_excel_ref.ole_free
            t "successfully ole_freed #{weak_excel_ref}"
          rescue
            t "could not do ole_free on #{weak_excel_ref}"
          end
        end
        hwnd2excel(excel_hwnd).die rescue nil
        @@hwnd2excel.delete(excel_hwnd)
      rescue => e
        t "Error when closing Excel: #{e.message}"
        #t e.backtrace
      end
      free_all_ole_objects
    end

    # frees all OLE objects in the object space
    def self.free_all_ole_objects   
      anz_objekte = 0
      ObjectSpace.each_object(WIN32OLE) do |o|
        anz_objekte += 1
        #t [:Name, (o.Name rescue (o.Count rescue "no_name"))]
        #t [:ole_object_name, (o.ole_object_name rescue nil)]
        #t [:methods, (o.ole_methods rescue nil)] unless (o.Name rescue false)
        #t o.ole_type rescue nil
        begin
          o.ole_free
          #t "olefree OK"
        rescue
          #t "olefree_error: #{$!}"
          #t $!.backtrace.first(9).join "\n"
        end
      end
      t "went through #{anz_objekte} OLE objects"
    end

    def hwnd_xxxx  
      self.HWnd rescue nil
    end

    # sets this Excel instance to nil
    def die 
      @ole_excel = nil
    end

  public

    # closes the Excel
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
      self.class.manage_unsaved_workbooks(self, @ole_excel, options)
      close_excel(options)
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
        #if WIN32OLE.ole_reference_count(weak_xlapp) > 0
        begin
          weak_excel_ref.ole_free
          t "successfully ole_freed #{weak_excel_ref}"
        rescue
          t "could not do ole_free on #{weak_excel_ref}"
        end
      end
      hwnd2excel(excel_hwnd).die rescue nil
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

    def recreate
      unless self.alive?
        # - generated Excel instance differs from all other Excel Instances
        #   (but this is done anyway with Excel.create?!)
        # - keep the old properties: visible, dispayalerts
        puts "not alive"
        new_excel = Excel.new(:reuse => false)
        # how to get the old visible and displayalerts values?: record them in book, or in Excel as attr_reader
        # new_excel = Excel.new(:reuse => false, :visible => @ole_excel.Visible, :displayalerts => @ole_excel.DisplayAlerts)      
        # - find all workbooks that were open in the old Excel instance
        # (- what about the workbooks opened interactively by the user?)
        #   in filename2books: go through all books:        
        #    if in the book the book.excel is identical with the old Excel instance (self), then
        #      open the workbook with this filename in the new Excel  
        #      record in book.excel the new_excel (this is done by Book.open anyway)
        # does not work:  book.excel is nil, self is nil
        #  need another data structure ???
        # but with close we want to delete everything
        bookstore = Book.bookstore
        books = bookstore.books
        puts "books: #{books}"
        if books
          books.each do |book|
            puts "workbook: #{book.workbook}"
            puts "filename: #{book.stored_filename}"
            puts "true" if book.excel == self
            puts "book.excel: #{book.excel}"
            puts "self: #{self}"
            new_book = Book.open(:force_excel => new_excel) # if this workbook was open in this Excel ???
          end
        end
        new_bookstore = Book.bookstore
        new_books = new_bookstore.books
        puts "new_books: #{new_books}"
        puts "workbooks:"
        new_excel.print_workbooks
        @ole_excel = new_excel
        #self ?
        #new_excel ?
      end
    end

    def self.kill_all
      procs = WIN32OLE.connect("winmgmts:\\\\.")
      processes = procs.InstancesOf("win32_process")
      number = processes.select{|p| (p.name == "EXCEL.EXE")}.size
      procs.InstancesOf("win32_process").each do |p|
        Process.kill('KILL', p.processid) if p.name == "EXCEL.EXE"        
      end
      number
    end

    def self.excel_processes
      pid2excel = {}
      @@hwnd2excel.each do |hwnd,wr_excel|
        process_id = Win32API.new("user32", "GetWindowThreadProcessId", ["I","P"], "I")
        pid_puffer = " " * 32
        process_id.call(hwnd, pid_puffer)
        pid = pid_puffer.unpack("L")[0]
        pid2excel[pid] = wr_excel
      end
      procs = WIN32OLE.connect("winmgmts:\\\\.")
      processes = procs.InstancesOf("win32_process")     
      result = []
      processes.each do |p|
        if p.name == "EXCEL.EXE"
          if pid2excel.include?(p.processid)
            excel = pid2excel[p.processid].__getobj__ 
            result << excel
          end
          # how to connect with an Excel instance that is not in hwnd2excel (uplifting Excel)?
          #excel = Excel.uplift unless (excel && excel.alive?)
        end
      end
      result
    end

    def excel
      self
    end

    def self.hwnd2excel(hwnd)
      excel_weakref = @@hwnd2excel[hwnd]
      if excel_weakref
        if excel_weakref.weakref_alive?
          excel_weakref.__getobj__
        else
          t "dead reference to an Excel"
          begin 
            @@hwnd2excel.delete(hwnd)
          rescue
            t "Warning: deleting dead reference failed! (hwnd: #{hwnd.inspect})"
          end
        end
      end
    end

    def hwnd
      self.Hwnd
    end

    def self.print_hwnd2excel
      @@hwnd2excel.each do |hwnd,wr_excel|
        excel_string = (wr_excel.weakref_alive? ? wr_excel.__getobj__.to_s : "weakref not alive") 
        printf("hwnd: %8i => excel: %s\n", hwnd, excel_string)
      end
      @@hwnd2excel.size
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
      #t $!.message
      false
    end

    def print_workbooks
      self.Workbooks.each {|w| t "#{w.Name} #{w}"}
    end

    def unsaved_workbooks
      result = []
      begin
        self.Workbooks.each {|w| result << w unless (w.Saved || w.ReadOnly)}
      rescue RuntimeError => msg
        t "RuntimeError: #{msg.message}" 
        raise ExcelErrorOpen, "Excel instance not alive or damaged" if msg.message =~ /failed to get Dispatch Interface/
      end
      result      
    end
    # yields different WIN32OLE objects than book.workbook
    #self.class.map {|w| (not w.Saved)}

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
