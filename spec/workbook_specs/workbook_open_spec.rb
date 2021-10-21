# -*- coding: utf-8 -*-

require_relative '../spec_helper'


$VERBOSE = nil

include RobustExcelOle
include General

describe Workbook do

  before(:all) do
    excel = Excel.new(:reuse => true)
    open_books = excel == nil ? 0 : excel.Workbooks.Count
    puts "*** open books *** : #{open_books}" if open_books > 0
    Excel.kill_all
  end

  before do
    @dir = create_tmpdir
    @simple_file = @dir + '/workbook.xls'
    @simple_save_file = @dir + '/workbook_save.xls'
    @different_file = @dir + '/different_workbook.xls'
    @simple_file_other_path = @dir + '/more_data/workbook.xls'
    @another_simple_file = @dir + '/another_workbook.xls'
    @simple_file_xlsm = @dir + '/workbook.xlsm'
    @simple_file_xlsx = @dir + '/workbook.xlsx'
    @simple_file1 = @simple_file
    @different_file1 = @different_file
    @simple_file_other_path1 = @simple_file_other_path
    @another_simple_file1 = @another_simple_file
    @simple_file_direct = File.join(File.dirname(__FILE__), 'data') + '/workbook.xls'
    #@simple_file_via_network = File.join('N:/', 'data') + '/workbook.xls'
    @simple_file_network_path = "N:/data/workbook.xls"
    network = WIN32OLE.new('WScript.Network')
    computer_name = network.ComputerName
    @simple_file_hostname_share_path = "//#{computer_name}/#{absolute_path('gems/robust_excel_ole/spec/data/workbook.xls').tr('\\','/').gsub('C:','c$')}"
    @simple_file_network_path_other_path = "N:/data/more_data/workbook.xls"
    @simple_file_hostname_share_path_other_path = "//#{computer_name}/#{absolute_path('gems/robust_excel_ole/spec/more_data/workbook.xls').tr('\\','/').gsub('C:','c$')}"
    @simple_file_network_path1 = @simple_file_network_path
    @simple_file_hostname_share_path1 = @simple_file_hostname_share_path
    @simple_file_network_path_other_path1 = @simple_file_network_path_other_path
    @simple_file_hostname_share_path_other_path1 = @simple_file_hostname_share_path_other_path
    @simple_file_xlsm1 = @simple_file_xlsm
    @simple_file_xlsx1 = @simple_file_xlsx
    #@linked_file = @dir + '/workbook_linked.xlsm'
    #@sub_file = @dir + '/workbook_sub.xlsm'
    @main_file = @dir + '/workbook_linked3.xlsm'
    @sub_file = @dir + '/workbook_linked_sub.xlsm'
    @error_message_excel = "provided Excel option value is neither an Excel object nor a valid option"
  end

  after do
    Excel.kill_all
    rm_tmp(@dir)
  end

  describe "changing ReadOnly mode" do

    it "should change from writable to readonly back to writable" do
      book = Workbook.open(@simple_file1)
      book.ReadOnly.should be false
      book2 = Workbook.open(@simple_file1, read_only: true)
      book2.should == book
      book2.ReadOnly.should be true
      book3 = Workbook.open(@simple_file1, read_only: false)
      book3.should == book
      book3.ReadOnly.should be false
      book3.close
    end

    it "should change from readonly to writable back to readonly" do
      book = Workbook.open(@simple_file1, read_only: true)
      book.ReadOnly.should be true
      book2 = Workbook.open(@simple_file1, read_only: false)
      book2.should == book
      book2.ReadOnly.should be false
      book3 = Workbook.open(@simple_file1, read_only: true)
      book3.should == book
      book3.ReadOnly.should be true
      book3.close
    end

    it "should raise error when read-only workbook unsaved and trying to reopen workbook writable by default" do
      book = Workbook.open(@simple_file1, read_only: true)
      book.ReadOnly.should be true
      sheet = book.sheet(1)
      sheet[1,1] = (sheet[1,1] == "foo" ? "bar" : "foo")
      expect{
        Workbook.open(@simple_file1, read_only: false)
      }.to raise_error(WorkbookNotSaved)
    end

    it "should raise error when read-only workbook unsaved and trying to reopen workbook writable with option :raise" do
      book = Workbook.open(@simple_file1, read_only: true)
      book.ReadOnly.should be true
      sheet = book.sheet(1)
      sheet[1,1] = (sheet[1,1] == "foo" ? "bar" : "foo")
      expect{
        Workbook.open(@simple_file1, read_only: false, if_unsaved: :raise)
      }.to raise_error(WorkbookNotSaved)
    end

    it "should save changes and change from read-only to writable with option :save" do
      book = Workbook.open(@simple_file1, read_only: true)
      book.ReadOnly.should be true
      sheet = book.sheet(1)
      old_value = sheet[1,1]
      sheet[1,1] = (sheet[1,1] == "foo" ? "bar" : "foo")
      new_value = sheet[1,1]
      book2 = Workbook.open(@simple_file1, if_unsaved: :save, read_only: false)
      book2.should == book
      book2.ReadOnly.should be false
      sheet[1,1].should_not == old_value
      sheet[1,1].should == new_value
    end

    it "should discard changes and change from read-only to writable with options :forget" do
      book = Workbook.open(@simple_file1, read_only: true)
      book.ReadOnly.should be true
      sheet = book.sheet(1)
      old_value = sheet[1,1]
      sheet[1,1] = (sheet[1,1] == "foo" ? "bar" : "foo")
      new_value = sheet[1,1]
      book2 = Workbook.open(@simple_file1, if_unsaved: :forget, read_only: false)
      book2.ReadOnly.should be false
      book2.close
      book3 = Workbook.open(@simple_file1)
      sheet3 = book3.sheet(1)
      sheet3[1,1].should == old_value
      sheet3[1,1].should_not == new_value
    end


    it "should raise error when writable workbook unsaved and trying to reopen workbook read-only by default" do
      book = Workbook.open(@simple_file1, read_only: false)
      book.ReadOnly.should be false
      sheet = book.sheet(1)
      sheet[1,1] = (sheet[1,1] == "foo" ? "bar" : "foo")
      expect{
        Workbook.open(@simple_file1, read_only: true)
      }.to raise_error(WorkbookNotSaved)
    end

    it "should raise error when writable workbook unsaved and trying to reopen workbook read-only with option :raise" do
      book = Workbook.open(@simple_file1, read_only: false)
      book.ReadOnly.should be false
      sheet = book.sheet(1)
      sheet[1,1] = (sheet[1,1] == "foo" ? "bar" : "foo")
      expect{
        Workbook.open(@simple_file1, read_only: true, if_unsaved: :raise)
      }.to raise_error(WorkbookNotSaved)
    end

    it "should save changes and change from writable to read-only with option :save" do
      book = Workbook.open(@simple_file1, read_only: false)
      book.ReadOnly.should be false
      sheet = book.sheet(1)
      old_value = sheet[1,1]
      sheet[1,1] = (sheet[1,1] == "foo" ? "bar" : "foo")
      new_value = sheet[1,1]
      book2 = Workbook.open(@simple_file1, if_unsaved: :save, read_only: true)
      book2.should == book
      book2.ReadOnly.should be true
      sheet[1,1].should_not == old_value
      sheet[1,1].should == new_value
    end

    it "should discard changes and change from writable to read-only with options :forget" do
      book = Workbook.open(@simple_file1, read_only: false)
      book.ReadOnly.should be false
      sheet = book.sheet(1)
      old_value = sheet[1,1]
      sheet[1,1] = (sheet[1,1] == "foo" ? "bar" : "foo")
      new_value = sheet[1,1]
      book2 = Workbook.open(@simple_file1, if_unsaved: :forget, read_only: true)
      book2.should == book
      book2.ReadOnly.should be true
      book2.close
      book3 = Workbook.open(@simple_file1)
      sheet3 = book3.sheet(1)
      sheet3[1,1].should == old_value
      sheet3[1,1].should_not == new_value
    end

    context "with :if_unsaved => :excel or :alert and from read-only to writable" do
     
      before do
        @book = Workbook.open(@simple_file1, v: true, readonly: true)
        @book.ReadOnly.should be false
        @sheet = @book.sheet(1)
        @old_value = @sheet[1,1]
        @sheet[1,1] = (@sheet[1,1] == "foo" ? "bar" : "foo")
        @new_value = @sheet[1,1] 
        @key_sender = IO.popen  'ruby "' + File.join(File.dirname(__FILE__), '../helpers/key_sender.rb') + '" "Microsoft Excel" '  , "w"
      end

      after do
        @key_sender.close
      end

      # question to the user, whether the workbook shall be reopened and discard any changes
      it "should discard changes and reopen the workbook, if user answers 'yes'" do
        @key_sender.puts "{enter}"
        book2 = Workbook.open(@simple_file1, read_only: false, if_unsaved: :excel)
        book2.ReadOnly.should be false
        book2.Saved.should be true
        @sheet[1,1].should == @old_value
        book2.close
        book3 = Workbook.open(@simple_file1)
        sheet3 = book3.sheet(1)
        sheet3[1,1].should == @old_value
      end

      it "should not discard changes and reopen the workbook, if user answers 'no'" do
        # "No" is right to "Yes" (the  default). --> language independent
        # strangely, in the "no" case, the question will sometimes be repeated three times
        #@book.excel.Visible = true
        @key_sender.puts "{right}{enter}"
        @key_sender.puts "{right}{enter}"
        @key_sender.puts "{right}{enter}"
        book2 = Workbook.open(@simple_file1, read_only: false, if_unsaved: :excel)
        book2.ReadOnly.should be false
        book2.Saved.should be false
        @sheet[1,1].should == @new_value
        book2.close(if_unsaved: :forget)
        book3 = Workbook.open(@simple_file1)
        sheet3 = book3.sheet(1)
        sheet3[1,1].should == @old_value
      end

      context "with :if_unsaved => :excel or :alert and from writable to read-only" do
     
        before do
          Excel.kill_all
          sleep 3
          @book = Workbook.open(@simple_file1, v: true, readonly: false)
          @book.ReadOnly.should be false
          @sheet = @book.sheet(1)
          @old_value = @sheet[1,1]
          @sheet[1,1] = (@sheet[1,1] == "foo" ? "bar" : "foo")
          @new_value = @sheet[1,1] 
          @key_sender = IO.popen  'ruby "' + File.join(File.dirname(__FILE__), '../helpers/key_sender.rb') + '" "Microsoft Excel" '  , "w"
        end

        after do
          @key_sender.close
        end

        # question to the user, whether to save the changes before changing read-only mode
        it "should save the workbook, if user answers 'no' and 'yes'" do
          # "No" is right to "Yes" (the  default). --> language independent
          @key_sender.puts "{right}{enter}"
          # 2nd question to the user: whether the workbook shall be reopened and discard any changes 
          # "No" is right to "Yes" (the  default). --> language independent       
          @key_sender.puts "{right}{enter}"
          @key_sender.puts "{right}{enter}"
          @key_sender.puts "{right}{enter}"
          # 3rd question whether the workbook shall be saved
          # "Yes"
          @key_sender.puts "{enter}"
          book2 = Workbook.open(@simple_file1, read_only: true, if_unsaved: :excel)
          book2.ReadOnly.should be true
          book2.Saved.should be false
          @sheet[1,1].should == @new_value
          book2.close(if_unsaved: :forget)
          book3 = Workbook.open(@simple_file1)
          sheet3 = book3.sheet(1)
          sheet3[1,1].should == @old_value
        end

        it "should discard (not save) the workbook, if user answers 'no' and 'no'" do
          # "No" is right to "Yes" (the  default). --> language independent
          @key_sender.puts "{right}{enter}"
          # 2nd question to the user: whether the workbook shall be reopened and discard any changes 
          # "No" is right to "Yes" (the  default). --> language independent       
          @key_sender.puts "{right}{enter}"
          @key_sender.puts "{right}{enter}"
          @key_sender.puts "{right}{enter}"
          # 3rd question whether the workbook shall be saved
          # "No"
          @key_sender.puts "{right}{enter}"
          @key_sender.puts "{right}{enter}"
          book2 = Workbook.open(@simple_file1, read_only: true, if_unsaved: :excel)
          book2.ReadOnly.should be true
          book2.Saved.should be false
          @sheet[1,1].should == @new_value
          book2.close(if_unsaved: :forget)
          book3 = Workbook.open(@simple_file1)
          sheet3 = book3.sheet(1)
          sheet3[1,1].should == @old_value
        end

        it "should not save (discard) changes and not reopen the workbook, if user answers 'no' and 'cancel'" do
          # "No" is right to "Yes" (the  default). --> language independent
          @key_sender.puts "{right}{enter}"
          # 2nd question to the user: whether the workbook shall be reopened and discard any changes 
          # "No" is right to "Yes" (the  default). --> language independent       
          @key_sender.puts "{right}{enter}"
          @key_sender.puts "{right}{enter}"
          @key_sender.puts "{right}{enter}"
          # 3rd question whether the workbook shall be saved
          # "Cancel"
          @key_sender.puts "{right}{right}{enter}"
          @key_sender.puts "{right}{right}{enter}"
          book2 = Workbook.open(@simple_file1, read_only: true, if_unsaved: :excel)
          book2.ReadOnly.should be true
          book2.Saved.should be false
          @sheet[1,1].should == @new_value
          book2.close(if_unsaved: :forget)
          book3 = Workbook.open(@simple_file1, if_unsaved: :forget)
          sheet3 = book3.sheet(1)
          sheet3[1,1].should == @old_value
        end

      end
     
    end

  end

  describe "linked workbooks" do

    context "standard" do

      before do
        @book1 = Workbook.open(@main_file)
      end

      it "should open the main workbook and the linked workbook" do
        @book1.should be_alive
        @book1.should be_a Workbook
        @book1.filename.should == @main_file
        Excel.current.workbooks.map{|b| b.filename}.should == [@main_file, @sub_file]
        book2 = Workbook.open(@sub_file)
        book2.should be_alive
        book2.should be_a Workbook
        book2.filename.should == @sub_file
      end

      it "should close the main workbook" do
        @book1.close
        Excel.current.workbooks.map{|b| b.filename}.should == [@sub_file]
      end

      it "should raise error when trying to close the linked workbook" do
        book2 = Workbook.open(@sub_file)
        expect{
         book2.close
        }.to raise_error(WorkbookLinked)
      end

      it "should raise error when trying to change the read-only mode of the linked workbook" do
        book2 = Workbook.open(@sub_file, :read_only => true)
        book2.ReadOnly.should be true
      end
    end
  end

  describe "basic tests with xlsx-workbooks" do

    context "with simple file" do

      it "should simply create a new workbook given a file" do
        book = Workbook.new(@simple_file_xlsx1)
        book.should be_alive
        book.should be_a Workbook
        book.filename.should == @simple_file_xlsx1
      end

    end

    context "with transparency identity" do

      before do
        @book = Workbook.open(@simple_file_xlsx1)        
      end

      after do
        @book.close
      end

      it "should yield identical Workbook objects referring to identical WIN32OLE objects" do
        book2 = Workbook.new(@book.ole_workbook)
        book2.equal?(@book).should be true
      end

    end

    context "with connecting to one unknown workbook" do

      before do
        ole_e1 = WIN32OLE.new('Excel.Application')
        ws = ole_e1.Workbooks
        abs_filename = General.absolute_path(@simple_file_xlsx1)
        @ole_wb = ws.Open(abs_filename)
      end

      it "should connect to an unknown workbook" do
        Workbook.open(@simple_file_xlsx1) do |book|
          book.filename.should == @simple_file_xlsx1
          book.should be_alive
          book.should be_a Workbook
          book.excel.ole_excel.Hwnd.should == @ole_wb.Application.Hwnd
          Excel.instance_count.should == 1
        end
      end
    end

    context "with :force => excel" do

      before do
        @book = Workbook.open(@simple_file_xlsx1)
      end

      it "should open in a new Excel" do
        book2 = Workbook.open(@simple_file_xlsx1, :force => {:excel => :new})
        book2.should be_alive
        book2.should be_a Workbook
        book2.excel.should_not == @book.excel 
        book2.should_not == @book
        @book.Readonly.should be false
        book2.Readonly.should be true
        book2.close
      end

    end

    context "with :if_unsaved" do

      before do
        @book = Workbook.open(@simple_file_xlsx1)
        @sheet = @book.sheet(1)
        @book.add_sheet(@sheet, :as => 'a_name')
        @book.visible = true
      end

      after do
        @book.close(:if_unsaved => :forget)
        @new_book.close rescue nil
      end

      it "should let the book open, if :if_unsaved is :accept" do
        expect {
          @new_book = Workbook.open(@simple_file_xlsx1, :if_unsaved => :accept)
          }.to_not raise_error
        @book.should be_alive
        @new_book.should be_alive
        @new_book.should == @book
      end
    
    end
  
  end

  describe "basic tests with xlsm-workbooks" do

    context "with simple file" do

      it "should simply create a new workbook given a file" do
        book = Workbook.new(@simple_file_xlsm1)
        book.should be_alive
        book.should be_a Workbook
        book.filename.should == @simple_file_xlsm1
      end

    end

    context "with transparency identity" do

      before do
        @book = Workbook.open(@simple_file_xlsm1)        
      end

      after do
        @book.close
      end

      it "should yield identical Workbook objects referring to identical WIN32OLE objects" do
        book2 = Workbook.new(@book.ole_workbook)
        book2.equal?(@book).should be true
      end

    end

    context "connecting to one unknown workbook" do

      before do
        ole_e1 = WIN32OLE.new('Excel.Application')
        ws = ole_e1.Workbooks
        abs_filename = General.absolute_path(@simple_file_xlsm1)
        @ole_wb = ws.Open(abs_filename)
      end

      it "should connect to an unknown workbook" do
        Workbook.open(@simple_file_xlsm1) do |book|
          book.filename.should == @simple_file_xlsm1
          book.should be_alive
          book.should be_a Workbook
          book.excel.ole_excel.Hwnd.should == @ole_wb.Application.Hwnd
          Excel.instance_count.should == 1
        end
      end
    end

    context "with :force => excel" do

      before do
        @book = Workbook.open(@simple_file_xlsm1)
      end

      it "should open in a new Excel" do
        book2 = Workbook.open(@simple_file_xlsm1, :force => {:excel => :new})
        book2.should be_alive
        book2.should be_a Workbook
        book2.excel.should_not == @book.excel 
        book2.should_not == @book
        @book.Readonly.should be false
        book2.Readonly.should be true
        book2.close
      end

    end

    context "with :if_unsaved" do

      before do
        @book = Workbook.open(@simple_file_xlsm1)
        @sheet = @book.sheet(1)
        @book.add_sheet(@sheet, :as => 'a_name')
        @book.visible = true
      end

      after do
        @book.close(:if_unsaved => :forget)
        @new_book.close rescue nil
      end

      it "should let the book open, if :if_unsaved is :accept" do
        expect {
          @new_book = Workbook.open(@simple_file_xlsm1, :if_unsaved => :accept)
          }.to_not raise_error
        @book.should be_alive
        @new_book.should be_alive
        @new_book.should == @book
      end

    end

  end

  describe "open and new" do

    context "with standard" do

      before do
        @book1 = Workbook.new(@simple_file1)
      end

      it "should yield identical Workbook objects for a file name" do
        book2 = Workbook.open(@simple_file1)
        book2.equal?(@book1).should be true
      end

      it "should yield identical Workbook objects for a win32ole-workbook" do
        ole_workbook1 = @book1.ole_workbook
        book2 = Workbook.new(ole_workbook1)
        book3 = Workbook.open(ole_workbook1)
        book3.equal?(book2).should be true
      end

    end

  end

  describe "fetching workbooks with network and hostname share paths" do

    before do
      bookstore = Bookstore.new
    end

    it "should fetch a network path file given a not via Reo opened hostname share file" do
      ole_e1 = WIN32OLE.new('Excel.Application')
      ws = ole_e1.Workbooks
      abs_filename = General.absolute_path(@simple_file_hostname_share_path)
      @ole_wb = ws.Open(abs_filename)
      book2 = Workbook.open(@simple_file_network_path)
      #book2.should === @ole_wb.to_reo
      book2.Fullname.should == @ole_wb.Fullname
      book2.excel.Workbooks.Count.should == 1
    end

    it "should fetch a network path file given a hostname share file" do
      book1 = Workbook.open(@simple_file_hostname_share_path)
      book2 = Workbook.open(@simple_file_network_path)
      book2.should === book1
      book2.Fullname.should == book1.Fullname
      book1.excel.Workbooks.Count.should == 1
    end

    it "should fetch a hostname share file given a network path file" do
      book1 = Workbook.open(@simple_file_network_path)
      book2 = Workbook.open(@simple_file_hostname_share_path)
      book2.should === book1
      book2.Fullname.should == book1.Fullname
      book1.excel.Workbooks.Count.should == 1
    end

    it "should raise WorkbookBlocked" do
      book1 = Workbook.open(@simple_file_hostname_share_path)
      expect{
        book2 = Workbook.open(@simple_file)
        }.to raise_error(WorkbookBlocked)
    end

    it "should raise an error fetching an hostname share file having opened a local path file" do
      book1 = Workbook.open(@simple_file)
      expect{
        Workbook.open(@simple_file_hostname_share_path)
        }.to raise_error(WorkbookBlocked)
    end

    it "should raise an error fetching a local path file having opened a network path file" do
      book1 = Workbook.open(@simple_file_network_path)
      expect{
        Workbook.open(@simple_file)
        }.to raise_error(WorkbookBlocked)
    end

    it "should raise an error fetching a network path file having opened a local path file" do
      book1 = Workbook.open(@simple_file)
      expect{
        Workbook.open(@simple_file_network_path)
        }.to raise_error(WorkbookBlocked)
    end

    it "should raise an error fetching a local path file having opened a hostname share path file" do
      book1 = Workbook.open(@simple_file_hostname_share_path)
      expect{
        Workbook.open(@simple_file)
        }.to raise_error(WorkbookBlocked)
    end

    it "should raise an WorkbookBlockederror" do
      book1 = Workbook.open(@simple_file_network_path1)
      expect{
        Workbook.open(@simple_file_network_path_other_path1)
      }.to raise_error(WorkbookBlocked)
    end

    it "should raise an WorkbookBlockederror" do
      book1 = Workbook.open(@simple_file_network_path_other_path1)
      expect{
        Workbook.open(@simple_file_network_path1)
      }.to raise_error(WorkbookBlocked)
    end

    it "should raise an WorkbookBlockederror" do
      book1 = Workbook.open(@simple_file_hostname_share_path1)
      expect{
        Workbook.open(@simple_file_hostname_share_path_other_path1)
      }.to raise_error(WorkbookBlocked)
    end

    it "should raise an WorkbookBlockederror" do
      book1 = Workbook.open(@simple_file_hostname_share_path_other_path1)
      expect{
        Workbook.open(@simple_file_hostname_share_path1)
      }.to raise_error(WorkbookBlocked)
    end

    it "should raise an WorkbookBlockederror" do
      book1 = Workbook.open(@simple_file_hostname_share_path1)
      expect{
        Workbook.open(@simple_file_network_path_other_path1)
      }.to raise_error(WorkbookBlocked)
    end

    it "should raise an WorkbookBlockederror" do
      book1 = Workbook.open(@simple_file_hostname_share_path_other_path1)
      expect{
        Workbook.open(@simple_file_network_path1)
      }.to raise_error(WorkbookBlocked)
    end

    it "should raise an WorkbookBlockederror" do
      book1 = Workbook.open(@simple_file_network_path1)
      expect{
        Workbook.open(@simple_file_hostname_share_path_other_path1)
      }.to raise_error(WorkbookBlocked)
    end

    it "should raise an WorkbookBlockederror" do
      book1 = Workbook.open(@simple_file_network_path_other_path1)
      expect{
        Workbook.open(@simple_file_hostname_share_path1)
      }.to raise_error(WorkbookBlocked)
    end

  end


  describe "connecting to unknown workbooks" do

    context "with one unknown network path or hostname share file" do

      it "should connect to a network path workbook from a network path file" do
        ole_e1 = WIN32OLE.new('Excel.Application')
        ws = ole_e1.Workbooks
        abs_filename = General.absolute_path(@simple_file_network_path1)
        @ole_wb = ws.Open(abs_filename)
        Workbook.open(@simple_file_network_path1) do |book|
          book.should be_alive
          book.should be_a Workbook
          book.filename.should == @simple_file_network_path1
          book.Fullname.should == @ole_wb.Fullname
          book.excel.ole_excel.Hwnd.should == @ole_wb.Application.Hwnd
          Excel.instance_count.should == 1
          book.excel.Workbooks.Count.should == 1
        end
      end

      it "should connect to a hostname share workbook from a network path file" do
        ole_e1 = WIN32OLE.new('Excel.Application')
        ws = ole_e1.Workbooks
        abs_filename = General.absolute_path(@simple_file_hostname_share_path1)
        @ole_wb = ws.Open(abs_filename)
        Workbook.open(@simple_file_network_path1) do |book|
          book.should be_alive
          book.should be_a Workbook
          book.filename.should == @simple_file_network_path1 #@simple_file_hostname_share_path1.downcase
          book.Fullname.should == @ole_wb.Fullname
          book.excel.ole_excel.Hwnd.should == @ole_wb.Application.Hwnd
          Excel.instance_count.should == 1
          book.excel.Workbooks.Count.should == 1
        end
      end

      it "should raise WorkbookBlocked trying to connect to a local path file from a network path file" do
        ole_e1 = WIN32OLE.new('Excel.Application')
        ws = ole_e1.Workbooks
        abs_filename = General.absolute_path(@simple_file1)
        @ole_wb = ws.Open(abs_filename)
        expect{
          Workbook.open(@simple_file_network_path1)
          }.to raise_error(WorkbookBlocked)
      end

      it "should raise WorkbookBlocked trying to connect to a network path file from a local path file" do
        ole_e1 = WIN32OLE.new('Excel.Application')
        ws = ole_e1.Workbooks
        abs_filename = General.absolute_path(@simple_file_network_path1)
        @ole_wb = ws.Open(abs_filename)
        expect{
          Workbook.open(@simple_file1)
          }.to raise_error(WorkbookBlocked)
      end

      it "should raise WorkbookBlocked trying to connect a hostname share file from a local path file" do
        ole_e1 = WIN32OLE.new('Excel.Application')
        ws = ole_e1.Workbooks
        abs_filename = General.absolute_path(@simple_file_hostname_share_path1)
        @ole_wb = ws.Open(abs_filename)
        expect{
          Workbook.open(@simple_file1)
          }.to raise_error(WorkbookBlocked)
      end

      it "should raise WorkbookBlocked trying to connect to a local path workbook from a hostname share file" do
        ole_e1 = WIN32OLE.new('Excel.Application')
        ws = ole_e1.Workbooks
        abs_filename = General.absolute_path(@simple_file1)
        @ole_wb = ws.Open(abs_filename)
        expect{
          Workbook.open(@simple_file_hostname_share_path1)
          }.to raise_error(WorkbookBlocked)
      end

      it "should connect to a network path workbook from a hostname share file" do
        ole_e1 = WIN32OLE.new('Excel.Application')
        ws = ole_e1.Workbooks
        abs_filename = General.absolute_path(@simple_file_network_path1)
        @ole_wb = ws.Open(abs_filename)
        Workbook.open(@simple_file_hostname_share_path1) do |book|
          book.should be_alive
          book.should be_a Workbook
          book.filename.should == @simple_file_network_path1
          book.Fullname.should == @ole_wb.Fullname
          book.excel.ole_excel.Hwnd.should == @ole_wb.Application.Hwnd
          Excel.instance_count.should == 1
          book.excel.Workbooks.Count.should == 1
        end
      end

      it "should connect to a hostname share workbook from a hostname share file" do
        ole_e1 = WIN32OLE.new('Excel.Application')
        ws = ole_e1.Workbooks
        abs_filename = General.absolute_path(@simple_file_hostname_share_path1)
        @ole_wb = ws.Open(abs_filename)
        Workbook.open(@simple_file_hostname_share_path1) do |book|
          book.should be_alive
          book.should be_a Workbook
          book.filename.should == @simple_file_network_path1 # @simple_file_hostname_share_path1.downcase
          book.Fullname.should == @ole_wb.Fullname
          book.excel.ole_excel.Hwnd.should == @ole_wb.Application.Hwnd
          Excel.instance_count.should == 1
          book.excel.Workbooks.Count.should == 1
        end
      end      

      it "should raise WorkbookBlocked error" do
        ole_e1 = WIN32OLE.new('Excel.Application')
        ws = ole_e1.Workbooks
        abs_filename = General.absolute_path(@simple_file_network_path1)
        @ole_wb = ws.Open(abs_filename)
        expect{
          Workbook.open(@simple_file_network_path_other_path1)
        }.to raise_error(WorkbookBlocked)
      end

      it "should raise WorkbookBlocked error" do
        ole_e1 = WIN32OLE.new('Excel.Application')
        ws = ole_e1.Workbooks
        abs_filename = General.absolute_path(@simple_file_network_path_other_path1)
        @ole_wb = ws.Open(abs_filename)
        expect{
          Workbook.open(@simple_file_network_path1)
        }.to raise_error(WorkbookBlocked)
      end

      it "should raise WorkbookBlocked error" do
        ole_e1 = WIN32OLE.new('Excel.Application')
        ws = ole_e1.Workbooks
        abs_filename = General.absolute_path(@simple_file_hostname_share_path1)
        @ole_wb = ws.Open(abs_filename)
        expect{
          Workbook.open(@simple_file_hostname_share_path_other_path1)
        }.to raise_error(WorkbookBlocked)
      end

      it "should raise WorkbookBlocked error" do
        ole_e1 = WIN32OLE.new('Excel.Application')
        ws = ole_e1.Workbooks
        abs_filename = General.absolute_path(@simple_file_hostname_share_path_other_path1)
        @ole_wb = ws.Open(abs_filename)
        expect{
          Workbook.open(@simple_file_hostname_share_path1)
        }.to raise_error(WorkbookBlocked)
      end

      it "should raise WorkbookBlocked error" do
        ole_e1 = WIN32OLE.new('Excel.Application')
        ws = ole_e1.Workbooks
        abs_filename = General.absolute_path(@simple_file_hostname_share_path1)
        @ole_wb = ws.Open(abs_filename)
        expect{
          Workbook.open(@simple_file_network_path_other_path1)
        }.to raise_error(WorkbookBlocked)
      end

      it "should raise WorkbookBlocked error" do
        ole_e1 = WIN32OLE.new('Excel.Application')
        ws = ole_e1.Workbooks
        abs_filename = General.absolute_path(@simple_file_network_path_other_path1)
        @ole_wb = ws.Open(abs_filename)
        expect{
          Workbook.open(@simple_file_hostname_share_path1)
        }.to raise_error(WorkbookBlocked)
      end

      it "should raise WorkbookBlocked error" do
        ole_e1 = WIN32OLE.new('Excel.Application')
        ws = ole_e1.Workbooks
        abs_filename = General.absolute_path(@simple_file_network_path1)
        @ole_wb = ws.Open(abs_filename)
        expect{
          Workbook.open(@simple_file_hostname_share_path_other_path1)
        }.to raise_error(WorkbookBlocked)
      end

      it "should raise WorkbookBlocked error" do
        ole_e1 = WIN32OLE.new('Excel.Application')
        ws = ole_e1.Workbooks
        abs_filename = General.absolute_path(@simple_file_hostname_share_path_other_path1)
        @ole_wb = ws.Open(abs_filename)
        expect{
          Workbook.open(@simple_file_network_path1)
        }.to raise_error(WorkbookBlocked)
      end

    end

    context "with one unknown hostname share path file" do

      before do
        ole_e1 = WIN32OLE.new('Excel.Application')
        ws = ole_e1.Workbooks
        abs_filename = General.absolute_path( @simple_file_hostname_share_path1)
        @ole_wb = ws.Open(abs_filename)
      end

      it "should connect to an unknown hostname share path workbook" do
        Workbook.open(@simple_file_hostname_share_path1) do |book|
          book.filename.should == @simple_file_network_path1
          book.should be_alive
          book.should be_a Workbook
          book.excel.ole_excel.Hwnd.should == @ole_wb.Application.Hwnd
          Excel.instance_count.should == 1
          book.excel.Workbooks.Count.should == 1
        end
      end

      it "should raise error because blocking" do
        expect{
          Workbook.open(@simple_file1)
          }.to raise_error(WorkbookBlocked)
      end

    end
      
    context "with none workbook" do

      it "should open one new Excel with the worbook" do
        book1 = Workbook.open(@simple_file1)
        book1.should be_alive
        book1.should be_a Workbook
        Excel.instance_count.should == 1
        book1.ReadOnly.should be false
        book1.excel.Visible.should be false
        book1.CheckCompatibility.should be true
        book1.Saved.should be true
      end

      it "should set the options" do
        book1 = Workbook.open(@simple_file1, :force => {:visible => true}, :check_compatibility => true)
        book1.visible.should be true
        book1.CheckCompatibility.should be true
      end

      it "should open in the given known Excel" do
        excel1 = Excel.create
        book1 = Workbook.open(@simple_file1)
        book1.should be_alive
        book1.should be_a Workbook
        book1.excel.should == excel1
        Excel.instance_count.should == 1        
        book1.excel.Visible.should be false
      end

      it "should open in the given known visible Excel" do
        excel1 = Excel.create(:visible => true)
        book1 = Workbook.open(@simple_file1)
        book1.should be_alive
        book1.should be_a Workbook
        book1.excel.should == excel1
        Excel.instance_count.should == 1        
        book1.excel.Visible.should be true
      end

      it "should open in the given known Excel" do
        excel1 = Excel.create
        book1 = Workbook.open(@simple_file1, :force => {:excel => :new})
        book1.should be_alive
        book1.should be_a Workbook
        book1.excel.should_not == excel1
        Excel.instance_count.should == 2        
      end

      it "should set the options in a given known Excel" do
        excel1 = Excel.create
        book1 = Workbook.open(@simple_file1, :force => {:visible => true}, :check_compatibility => true)
        book1.visible.should be true
        book1.CheckCompatibility.should be true
      end

      it "should open the workbook in the given Excel if there are only unknown Excels" do
        ole_excel1 = WIN32OLE.new('Excel.Application')
        book1 = Workbook.open(@simple_file1)
        book1.should be_alive
        book1.should be_a Workbook
        Excel.instance_count.should == 1
        book1.excel.ole_excel.Hwnd.should == ole_excel1.Hwnd
      end

    end

    context "with one unknown workbook" do

      before do
        ole_e1 = WIN32OLE.new('Excel.Application')
        ws = ole_e1.Workbooks
        abs_filename = General.absolute_path(@simple_file1)
        @ole_wb = ws.Open(abs_filename)
      end

      it "should connect to an unknown workbook" do
        Workbook.open(@simple_file1) do |book|
          book.filename.should == @simple_file1
          book.should be_alive
          book.should be_a Workbook
          book.excel.ole_excel.Hwnd.should == @ole_wb.Application.Hwnd
          Excel.instance_count.should == 1
        end
      end

      it "should raise error when connecting to a blocking workbook with :if_blocked => :raise" do
        expect{
          Workbook.open(@simple_file_other_path1) 
          }.to raise_error(WorkbookBlocked, /blocked by/)
      end

      it "should close the blocking workbook and open the new workbook with :if_blocked => :forget" do
        new_book = Workbook.open(@simple_file_other_path1, :if_blocked => :forget)
        expect{
          @ole_wb.Name
        }.to raise_error 
        new_book.should be_alive
        new_book.should be_a Workbook
        new_book.Fullname.should == General.absolute_path(@simple_file_other_path1)
      end      

      it "should let the workbook open, if :if_unsaved is :raise" do        
        @ole_wb.Worksheets.Add
        expect{
          new_book = Workbook.open(@simple_file1, :if_unsaved => :raise)
          }.to raise_error(WorkbookNotSaved, /workbook is already open but not saved: "workbook.xls"/)
      end

      it "should let the workbook open, if :if_unsaved is :save" do        
        @ole_wb.Worksheets.Add
        sheet_num = @ole_wb.Worksheets.Count
        new_book = Workbook.open(@simple_file1, :if_unsaved => :save)
        new_book.should be_alive
        new_book.should be_a Workbook
        new_book.Worksheets.Count.should == sheet_num
        new_book.close
        new_book2 = Workbook.open(@simple_file1)
        new_book2.Worksheets.Count.should == sheet_num
      end

      it "should let the workbook open, if :if_unsaved is :accept" do        
        @ole_wb.Worksheets.Add
        sheet_num = @ole_wb.Worksheets.Count
        new_book = Workbook.open(@simple_file1, :if_unsaved => :accept)
        new_book.should be_alive
        new_book.should be_a Workbook
        new_book.Worksheets.Count.should == sheet_num
        new_book.Saved.should be false
        new_book.close(:if_unsaved => :forget)
        new_book2 = Workbook.open(@simple_file1)
        new_book2.Worksheets.Count.should == sheet_num - 1
      end

      it "should close the workbook, if :if_unsaved is :forget" do        
        @ole_wb.Worksheets.Add
        sheet_num = @ole_wb.Worksheets.Count
        new_book = Workbook.open(@simple_file1, :if_unsaved => :forget)
        new_book.should be_alive
        new_book.should be_a Workbook
        new_book.Worksheets.Count.should == sheet_num - 1
      end

    end

    context "with several unknown workbooks" do

      before do
        ole_e1 = WIN32OLE.new('Excel.Application')
        ws1 = ole_e1.Workbooks
        abs_filename1 = General.absolute_path(@simple_file1)
        @ole_wb1 = ws1.Open(abs_filename1)
        ole_e2 = WIN32OLE.new('Excel.Application')
        ws2 = ole_e2.Workbooks
        abs_filename2 = General.absolute_path(@different_file1)
        @ole_wb2 = ws2.Open(abs_filename2)
      end

      it "should connect to the 1st unknown workbook in the 1st Excel instance" do
        Workbook.open(@simple_file1) do |book|
          book.filename.should == @simple_file1
          book.excel.ole_excel.Hwnd.should == @ole_wb1.Application.Hwnd
          Excel.instance_count.should == 2
        end
      end

      it "should connect to the 2nd unknown workbook in the 2nd Excel instance" do
        Workbook.open(@different_file1) do |book|
          book.filename.should == @different_file1
          book.excel.ole_excel.Hwnd.should == @ole_wb2.Application.Hwnd
          Excel.instance_count.should == 2
        end
      end

    end

  end

  describe "with already open Excel instances and an open unsaved workbook" do

    before do
      @ole_excel1 = WIN32OLE.new('Excel.Application')
      @ole_excel2 = WIN32OLE.new('Excel.Application')
      #@ole_workbook1 = @ole_excel1.Workbooks.Open(@simple_file1, { 'ReadOnly' => false })
      abs_filename = General.absolute_path(@simple_file1)
      @ole_workbook1 = @ole_excel1.Workbooks.Open(abs_filename, nil, false)
      @ole_workbook1.Worksheets.Add
    end

    context "with simple general situations" do
      
      it "should simply open" do
        book = Workbook.open(@simple_file1, :if_unsaved => :accept)
        book.should be_alive
        book.should be_a Workbook
      end

      it "should open in a new Excel" do
        book2 = Workbook.open(@simple_file1, :force => {:excel => :new})
        book2.should be_alive
        book2.should be_a Workbook
        book2.excel.should_not == @ole_excel1 
        book2.Readonly.should be true
      end

      it "should fetch the workbook" do
        new_book = Workbook.new(@ole_workbook1, :if_unsaved => :forget)
        new_book.should be_a Workbook
        new_book.should be_alive
        new_book.ole_workbook.should == @ole_workbook1
        new_book.excel.ole_excel.Hwnd.should == @ole_excel1.Hwnd
      end

      it "should fetch a closed workbook" do
        new_book = Workbook.new(@ole_workbook1)
        new_book.close(:if_unsaved => :forget)
        new_book.should_not be_alive
        book2 = Workbook.open(@simple_file1)
        book2.should === new_book
        book2.should be_alive
        book2.close
      end

      it "should force_excel with :reuse" do
        book2 = Workbook.open(@different_file, :force => {:excel => :current})
        book2.should be_alive
        book2.should be_a Workbook
        book2.excel.ole_excel.Hwnd.should == @ole_excel1.Hwnd 
      end

      it "should force_excel with :reuse when reopening and the Excel is not alive even if :default_excel says sth. else" do
        book1 = Workbook.open(@simple_file1, :if_unsaved => :forget)
        excel2 = Excel.new(:reuse => false)
        book1.excel.close(:if_unsaved => :forget)
        sleep 1
        book2 = Workbook.open(@simple_file1, :force => {:excel => :current}, :default => {:excel => :new})
        book2.should be_alive
        book2.should be_a Workbook
        book2.excel.ole_excel.Hwnd.should == @ole_excel2.Hwnd
      end

      it "should reopen the closed book" do
        book1 = Workbook.open(@simple_file1, :if_unsaved => :accept)
        book1.should be_alive
        book2 = book1
        book1.close(:if_unsaved => :forget)
        book1.should_not be_alive
        book1.open
        book1.should be_a Workbook
        book1.should be_alive
        book1.should === book2
      end
    end

    context "with :if_unsaved" do

      before do
        @book = Workbook.open(@simple_file1, :if_unsaved => :accept)
        sheet = @book.sheet(1)
        @old_value = sheet[1,1]
        sheet[1,1] = (sheet[1,1] == "foo" ? "bar" : "foo")
        @new_value = sheet[1,1]
        @book.Saved.should be false
      end

      after do
        @book.close(:if_unsaved => :forget)
      end

      it "should let the book open, if :if_unsaved is :accept" do
        new_book = Workbook.open(@simple_file1, :if_unsaved => :accept)
        @book.should be_alive
        new_book.should be_alive
        new_book.Saved.should be false      
        @book.Saved.should be false  
        new_book.sheet(1)[1,1].should == @new_value
        new_book.should == @book
      end

      it "should open book and close old book, if :if_unsaved is :forget" do
        new_book = Workbook.open(@simple_file1, :if_unsaved => :forget)
        @book.should_not be_alive
        new_book.should be_alive
        new_book.Saved.should be true
      end
    end

    context "with :if_blocked" do

      it "should raise an error, if :if_blocked is :raise" do
        expect {
          new_book = Workbook.open(@simple_file_other_path1)
        }.to raise_error(WorkbookBlocked, /blocked by/)
      end

      it "should close the other book and open the new book, if :if_blocked is :forget" do
        new_book = Workbook.open(@simple_file_other_path1, :if_blocked => :forget)
        expect{
          @ole_workbook1.Name
        }.to raise_error 
        new_book.should be_alive
      end

    end

    context "with :force => {:excel}" do

      it "should raise if excel is not alive" do
        excel1 = Excel.create
        excel1.close
        expect{
          book1 = Workbook.open(@simple_file1, :force => {:excel => excel1})
          }.to raise_error(ExcelREOError, "Excel is not alive")
      end

      it "should open in a provided Excel" do
        book1 = Workbook.open(@simple_file1, :force => {:excel => :new})
        book2 = Workbook.open(@simple_file1, :force => {:excel => :new})
        book3 = Workbook.open(@simple_file1, :force => {:excel => book2.excel})
        book3.should be_alive
        book3.should be_a Workbook
        book3.excel.should == book2.excel 
        book3.Readonly.should be true
      end 
    end
  end

  describe "simple open" do

    it "should simply open" do
      book = Workbook.open(@simple_file, :v => true, :f => {:e => :new})
    end

  end
  
  describe "new" do

    context "with transparency identity" do

      before do
        @book = Workbook.open(@simple_file1)        
        abs_filename = General.absolute_path(@simple_file1)
        @ole_book = WIN32OLE.connect(abs_filename)
      end

      after do
        @book.close
      end

      it "should yield identical Workbook objects referring to identical WIN32OLE objects" do
        book2 = Workbook.new(@book.ole_workbook)
        book2.equal?(@book).should be true
      end

      it "should yield identical Workbook objects referring to identical WIN32OLE objects with open" do
        book2 = Workbook.open(@book.ole_workbook)
        book2.equal?(@book).should be true
      end

      it "should yield identical Workbook objects created with help of their filenames" do
        book2 = Workbook.open(@simple_file1)
        book2.equal?(@book).should be true
      end

      it "should yield identical Workbook objects created with help of their WIN32OLE objects" do
        book2 = Workbook.new(@book.ole_workbook)
        book3 = Workbook.open(@book.ole_workbook)
        book3.equal?(book2).should be true
      end


      it "should yield identical Workbook objects for identical Excel books after prmoting" do
        book2 = Workbook.new(@ole_book)
        book2.should === @book
        book2.equal?(@book).should be true
        book2.close
      end

      it "should yield different Workbook objects for different Excel books" do
        book3 = Workbook.open(@different_file1)
        abs_filename2 = General.absolute_path(@different_file1)
        ole_book2 = WIN32OLE.connect(abs_filename2)
        book2 = Workbook.new(ole_book2)
        book2.should_not === @book
        book2.equal?(@book).should be false
        book2.close
        book3.close
      end

    end

    it "should simply create a new workbook given a file" do
      book = Workbook.new(@simple_file)
      book.should be_alive
      book.should be_a Workbook
    end

    it "should create a new workbook given a file and set it visible" do
      book = Workbook.new(@simple_file, :visible => true)
      book.should be_alive
      book.should be_a Workbook
      book.excel.Visible.should be true
      book.Windows(book.Name).Visible.should be true
    end

    it "should create a new workbook given a file and set it visible and readonly" do
      book = Workbook.new(@simple_file, :visible => true, :read_only => true)
      book.should be_alive
      book.should be_a Workbook
      book.excel.Visible.should be true
      book.Windows(book.Name).Visible.should be true
      book.ReadOnly.should be true
    end

    it "should create a new workbook given a file and set options" do
      book = Workbook.new(@simple_file, :visible => true, :read_only => true, :force => {:excel => :new})
      book.should be_alive
      book.should be_a Workbook
      book.excel.Visible.should be true
      book.Windows(book.Name).
      Visible.should be true
      book.ReadOnly.should be true
      book2 = Workbook.new(@different_file, :force => {:excel => :new}, :v => true)
      book2.should be_alive
      book2.should be_a Workbook
      book2.excel.Visible.should be true
      book2.Windows(book2.Name).Visible.should be true
      book2.ReadOnly.should be false
      book2.excel.should_not == book.excel
    end

    it "should type-lift an workbook" do
      book = Workbook.open(@simple_file)
      new_book = Workbook.new(book)
      new_book.should == book
      new_book.equal?(book).should be true
      new_book.Fullname.should == book.Fullname
      new_book.excel.should == book.excel
    end

    it "should type-lift an workbook and supply option" do
      book = Workbook.open(@simple_file)
      new_book = Workbook.new(book, :visible => true)
      new_book.should == book
      new_book.equal?(book).should be true
      new_book.Fullname.should == book.Fullname
      new_book.excel.should == book.excel
      new_book.visible.should be true
    end

    it "should type-lift an open known win32ole workbook" do
      book = Workbook.open(@simple_file)
      ole_workbook = book.ole_workbook
      new_book = Workbook.new(ole_workbook)
      new_book.should == book
      new_book.equal?(book).should be true
      new_book.Fullname.should == book.Fullname
      new_book.excel.should == book.excel
    end

    it "should type-lift an open known win32ole workbook and let it be visible" do
      book = Workbook.open(@simple_file, :visible => true)
      ole_workbook = book.ole_workbook
      new_book = Workbook.new(ole_workbook)
      new_book.should == book
      new_book.equal?(book).should be true
      new_book.Fullname.should == book.Fullname
      new_book.excel.should == book.excel
      new_book.excel.Visible.should == true
      new_book.Windows(new_book.ole_workbook.Name).Visible.should == true
    end

    it "should type-lift an open known win32ole workbook and let it be visible and readonly" do
      book = Workbook.open(@simple_file, :visible => true, :read_only => true)
      ole_workbook = book.ole_workbook
      new_book = Workbook.new(ole_workbook)
      new_book.should == book
      new_book.equal?(book).should be true
      new_book.Fullname.should == book.Fullname
      new_book.excel.should == book.excel
      new_book.excel.Visible.should == true
      new_book.Windows(new_book.ole_workbook.Name).Visible.should == true
      new_book.ReadOnly.should == true
    end

    it "should type-lift an open known win32ole workbook and make it visible" do
      book = Workbook.open(@simple_file)
      ole_workbook = book.ole_workbook
      new_book = Workbook.new(ole_workbook, :visible => true)
      new_book.should == book
      new_book.equal?(book).should be true
      new_book.Fullname.should == book.Fullname
      new_book.excel.should == book.excel
      new_book.excel.Visible.should == true
      new_book.Windows(new_book.ole_workbook.Name).Visible.should == true
    end

    it "should type-lift an open unknown win32ole workbook" do
      ole_excel = WIN32OLE.new('Excel.Application')
      ws = ole_excel.Workbooks
      abs_filename = General.absolute_path(@simple_file1)
      ole_workbook = ws.Open(abs_filename)
      new_book = Workbook.new(ole_workbook)
      new_book.Fullname.should == ole_workbook.Fullname
      new_book.excel.Hwnd.should == ole_excel.Hwnd
    end

    it "should type-lift an open unknown win32ole workbook and make it visible" do
      ole_excel = WIN32OLE.new('Excel.Application')
      ws = ole_excel.Workbooks
      abs_filename = General.absolute_path(@simple_file1)
      ole_workbook = ws.Open(abs_filename)
      new_book = Workbook.new(ole_workbook, :visible => true)
      new_book.Fullname.should == ole_workbook.Fullname
      new_book.excel.Hwnd.should == ole_excel.Hwnd
      new_book.excel.Visible.should == true
      new_book.Windows(new_book.ole_workbook.Name).Visible.should == true
    end

    it "should type-lift an open unknown win32ole workbook and make it visible and readonly" do
      ole_excel = WIN32OLE.new('Excel.Application')
      ws = ole_excel.Workbooks
      abs_filename = General.absolute_path(@simple_file1)
      ole_workbook = ws.Open(abs_filename)
      new_book = Workbook.new(ole_workbook, :visible => true)
      new_book.Fullname.should == ole_workbook.Fullname
      new_book.excel.Hwnd.should == ole_excel.Hwnd
      new_book.excel.Visible.should == true
      new_book.Windows(new_book.ole_workbook.Name).Visible.should == true
    end

  end

  describe "open" do

    context "with calculation mode" do

      it "should set calculation mode" do
        book1 = Workbook.open(@simple_file1, :visible => true)
        book1.excel.calculation = :manual
        book1.excel.Calculation.should == XlCalculationManual
        book1.save
        book1.excel.close
        book2 = Workbook.open(@simple_file1, :visible => true)
        book2.excel.calculation = :automatic
        book2.excel.Calculation.should == XlCalculationAutomatic
        book2.save
        book2.excel.close
      end

      it "should not set the default value" do
        book1 = Workbook.open(@simple_file)
        book1.excel.properties[:calculation].should == nil
      end

      it "should set the calculation mode to automatic" do
        book1 = Workbook.open(@simple_file)
        book1.excel.calculation = :automatic
        book1.excel.properties[:calculation].should == :automatic
        book1.excel.Calculation.should == XlCalculationAutomatic
      end

      it "should set the calculation mode to manual" do
        book1 = Workbook.open(@simple_file)
        book1.excel.calculation = :manual
        book1.excel.properties[:calculation].should == :manual
        book1.excel.Calculation.should == XlCalculationManual
      end

      it "should change the calculation mode from manual to automatic" do
        book1 = Workbook.open(@simple_file, :visible => true)
        excel1 = Excel.current(:calculation => :automatic)        
        book2 = Workbook.open(@different_file, :visible => true)
        book2.excel.Calculation.should == XlCalculationAutomatic
        book1.excel.Calculation.should == XlCalculationAutomatic
      end
    end

   
    context "with causing warning dead excel without window handle" do

      it "combined" do
        book1 = Workbook.open(@simple_file1)
        book2 = Workbook.open(@different_file)
        Excel.kill_all
        #sleep 1 #then no warning
        Excel.current # or book3 = Workbook.open(@another_simple_file)
      end

    end

    context "with class identifier 'Workbook'" do

      before do
        @book = Workbook.open(@simple_file)
      end

      after do
        @book.close rescue nil
      end

      it "should open in a new Excel" do
        book2 = Workbook.open(@simple_file, :force => {:excel => :new})
        book2.should be_alive
        book2.should be_a Workbook
        book2.excel.should_not == @book.excel 
        book2.should_not == @book
        @book.Readonly.should be false
        book2.Readonly.should be true
        book2.close
      end
    end

    context "lift a workbook to a Workbook object" do

      before do
        @book = Workbook.open(@simple_file)
      end

      after do
        @book.close
      end

      it "should fetch the workbook" do
        ole_workbook = @book.ole_workbook
        new_book = Workbook.new(ole_workbook)
        new_book.should be_a Workbook
        new_book.should be_alive
        new_book.should == @book
        new_book.filename.should == @book.filename
        new_book.excel.should == @book.excel
        new_book.excel.Visible.should be false
        #new_book.excel.DisplayAlerts.should be false
        new_book.should === @book
        new_book.close
      end

      it "should fetch the workbook" do
        ole_workbook = @book.ole_workbook
        new_book = Workbook.new(ole_workbook, :visible => true)
        new_book.should be_a Workbook
        new_book.should be_alive
        new_book.should == @book
        new_book.filename.should == @book.filename
        new_book.excel.should == @book.excel
        new_book.excel.Visible.should be true
        new_book.excel.DisplayAlerts.should be true
        new_book.should === @book
        new_book.close
      end

      it "should yield an identical Workbook and set visible value" do
        ole_workbook = @book.ole_workbook
        new_book = Workbook.new(ole_workbook, :visible => true)
        new_book.excel.displayalerts = true
        new_book.should be_a Workbook
        new_book.should be_alive
        new_book.should == @book
        new_book.filename.should == @book.filename
        new_book.excel.should == @book.excel
        new_book.should === @book
        new_book.excel.Visible.should be true
        new_book.excel.DisplayAlerts.should be true
        new_book.close
      end

    end

    context "with standard options" do
      before do
        @book = Workbook.open(@simple_file)
      end

      after do
        @book.close
      end

      it "should say that it lives" do
        @book.should be_alive
      end
    end

    context "with identity transperence" do

      before do
        @book = Workbook.open(@simple_file1)
      end

      after do
        @book.close
      end

      it "should yield identical Workbook objects for identical Excel books" do
        book2 = Workbook.open(@simple_file1)
        book2.should === @book
        book2.close
      end

      it "should yield different Workbook objects for different Excel books" do
        book2 = Workbook.open(@different_file)
        book2.should_not === @book
        book2.close
      end

      it "should yield different Workbook objects when opened the same file in different Excel instances" do
        book2 = Workbook.open(@simple_file, :force => {:excel => :new})
        book2.should_not === @book
        book2.close
      end

      it "should yield identical Workbook objects for identical Excel books when reopening" do
        @book.should be_alive
        @book.close
        @book.should_not be_alive
        book2 = Workbook.open(@simple_file1)
        book2.should === @book
        book2.should be_alive
        book2.close
      end

      it "should yield identical Workbook objects for identical Excel books when reopening with current excel" do
        @book.should be_alive
        @book.close
        @book.should_not be_alive
        book2 = Workbook.open(@simple_file1, :default => {:excel => :current})
        book2.should === @book
        book2.should be_alive
        book2.close
      end

      it "should yield identical Workbook objects for identical Excel books when reopening with current excel" do
        @book.should be_alive
        @book.close
        @book.should_not be_alive
        book2 = Workbook.open(@simple_file1, :force => {:excel => :current})
        book2.should === @book
        book2.should be_alive
        book2.close
      end

      it "should yield identical Workbook objects when reopening and the Excel is closed" do
        @book.should be_alive
        @book.close
        Excel.kill_all
        book2 = Workbook.open(@simple_file1)
        book2.should be_alive
        book2.should === @book
        book2.close
      end

      it "should yield different Workbook objects when reopening in a new Excel" do
        @book.should be_alive
        old_excel = @book.excel
        @book.close
        @book.should_not be_alive
        book2 = Workbook.open(@simple_file1, :force => {:excel => :new})
        book2.should_not === @book
        book2.should be_alive
        book2.excel.should_not == old_excel
        book2.close
      end

      it "should yield different Workbook objects when reopening in a new given Excel instance" do
        old_excel = @book.excel
        new_excel = Excel.new(:reuse => false)
        @book.close
        @book.should_not be_alive
        book2 = Workbook.open(@simple_file1, :force => {:excel => new_excel})
        #@book.should be_alive
        #book2.should === @book
        book2.should be_alive
        book2.excel.should == new_excel
        book2.excel.should_not == old_excel
        book2.close
      end

      it "should yield identical Workbook objects when reopening in the old excel" do
        old_excel = @book.excel
        new_excel = Excel.new(:reuse => false)
        @book.close
        @book.should_not be_alive
        book2 = Workbook.open(@simple_file1, :force => {:excel => old_excel})
        book2.should === @book
        book2.should be_alive
        book2.excel.should == old_excel
        @book.should be_alive
        book2.close
      end

    end

    context "with abbrevations" do

      before do
        @book = Workbook.open(@simple_file1)
      end

      after do
        @book.close rescue nil
      end

      it "should work as force" do
        book2 = Workbook.open(@another_simple_file, :excel => :new)
        book2.excel.should_not == @book.excel
        book3 = Workbook.open(@different_file, :excel => book2.excel)
        book3.excel.should == book2.excel
      end

      it "should work with abbrevation of force and excel" do
        book2 = Workbook.open(@another_simple_file, :f => {:e => :new})
        book2.excel.should_not == @book.excel
        book3 = Workbook.open(@different_file, :f => {:e => book2.excel})
        book3.excel.should == book2.excel
      end

      it "should work with abbrevation of force" do
        book2 = Workbook.open(@another_simple_file, :f => {:excel => :new})
        book2.excel.should_not == @book.excel
        book3 = Workbook.open(@different_file, :f => {:excel => book2.excel})
        book3.excel.should == book2.excel
      end

      it "should work with abbrevation of force" do
        book2 = Workbook.open(@another_simple_file, :force => {:e => :new})
        book2.excel.should_not == @book.excel
        book3 = Workbook.open(@different_file, :force => {:e => book2.excel})
        book3.excel.should == book2.excel
      end

      it "should open in a given Excel provided as Excel, Workbook, or WIN32OLE representing an Excel or Workbook" do
        book2 = Workbook.open(@another_simple_file)
        book3 = Workbook.open(@different_file)
        book3 = Workbook.open(@simple_file1, :excel => book2.excel)
        book3.excel.should === book2.excel
        book4 = Workbook.open(@simple_file1, :excel => @book) 
        book4.excel.should === @book.excel
        book3.close
        book4.close
        book5 = Workbook.open(@simple_file1, :excel => book2.ole_workbook)
        book5.excel.should ===  book2.excel
        win32ole_excel1 = WIN32OLE.connect(@book.ole_workbook.Fullname).Application
        book6 = Workbook.open(@simple_file1, :excel => win32ole_excel1)
        book6.excel.should === @book.excel
      end

      it "should use abbreviations of default" do
        book2 = Workbook.open(@simple_file1, :d => {:excel => :current})
        book2.excel.should == @book.excel
      end

      it "should use abbreviations of default" do
        book2 = Workbook.open(@simple_file1, :d => {:e => :current})
        book2.excel.should == @book.excel
      end

      it "should use abbreviations of default" do
        book2 = Workbook.open(@simple_file1, :default => {:e => :current})
        book2.excel.should == @book.excel
      end

      it "should reopen the book in the Excel where it was opened most recently" do
        excel1 = @book.excel
        excel2 = Excel.new(:reuse => false)
        @book.close
        book2 = Workbook.open(@simple_file1, :d => {:e => :current})
        book2.excel.should == excel1
        book2.close
        book3 = Workbook.open(@simple_file1, :e => excel2)
        book3.close
        book3 = Workbook.open(@simple_file1, :d => {:e => :current})
        book3.excel.should == excel2
        book3.close
      end

    end

    context "with :force => {:excel}" do

      before do
        @book = Workbook.open(@simple_file1)
      end

      after do
        @book.close rescue nil
      end

      it "should open in a given Excel provided as Excel, Workbook, or WIN32OLE representing an Excel or Workbook" do
        book2 = Workbook.open(@another_simple_file)
        book3 = Workbook.open(@different_file)
        book3 = Workbook.open(@simple_file1, :force => {:excel => book2.excel})
        book3.excel.should === book2.excel
        book4 = Workbook.open(@simple_file1, :force => {:excel => @book}) 
        book4.excel.should === @book.excel
        book3.close
        book4.close
        book5 = Workbook.open(@simple_file1, :force => {:excel => book2.ole_workbook})
        book5.excel.should ===  book2.excel
        win32ole_excel1 = WIN32OLE.connect(@book.ole_workbook.Fullname).Application
        book6 = Workbook.open(@simple_file1, :force => {:excel => win32ole_excel1})
        book6.excel.should === @book.excel
      end

      it "should open in a new Excel" do
        book2 = Workbook.open(@simple_file1, :force => {:excel => :new})
        book2.should be_alive
        book2.should be_a Workbook
        book2.excel.should_not == @book.excel 
        book2.should_not == @book
        @book.Readonly.should be false
        book2.Readonly.should be true
        book2.close
      end

      it "should open in a given Excel, not provide identity transparency, because old book readonly, new book writable" do
        book2 = Workbook.open(@simple_file1, :force => {:excel => :new})
        book2.excel.should_not == @book.excel
        book3 = Workbook.open(@simple_file1, :force => {:excel => :new})
        book3.excel.should_not == book2.excel
        book3.excel.should_not == @book.excel
        book2.close
        book4 = Workbook.open(@simple_file1, :force => {:excel => book2.excel})
        book4.should be_alive
        book4.should be_a Workbook
        book4.excel.should == book2.excel
        book4.Readonly.should == true
        book4.should_not == book2 
        book4.close
        book5 = Workbook.open(@simple_file1, :force => {:excel => book2})
        book5.should be_alive
        book5.should be_a Workbook
        book5.excel.should == book2.excel
        book5.Readonly.should == true
        book5.should_not == book2 
        book5.close
        book3.close
      end

      it "should open in a given Excel, provide identity transparency, because book can be readonly, such that the old and the new book are readonly" do
        book2 = Workbook.open(@simple_file1, :force => {:excel => :new})
        book2.excel.should_not == @book.excel
        book3 = Workbook.open(@simple_file1, :force => {:excel => :new})
        book3.excel.should_not == book2.excel
        book3.excel.should_not == @book.excel
        book2.close
        book3.close
        @book.close
        book4 = Workbook.open(@simple_file1, :force => {:excel => book2.excel}, :read_only => true)
        book4.should be_alive
        book4.should be_a Workbook
        book4.excel.should == book2.excel
        book4.ReadOnly.should be true
        book4.should == book2
        book4.close
        book5 = Workbook.open(@simple_file1, :force => {:excel => book2}, :read_only => true)
        book5.should be_alive
        book5.should be_a Workbook
        book5.excel.should == book2.excel
        book5.ReadOnly.should be true
        book5.should == book2
        book5.close
        book3.close
      end

      it "should open in a given Excel, provide identity transparency, because book can be readonly, such that the old and the new book are readonly" do
        book2 = Workbook.open(@simple_file1, :force => {:excel => :new})
        book2.excel.should_not == @book.excel
        book2.close
        @book.close
        book4 = Workbook.open(@simple_file1, :force => {:excel => book2}, :read_only => true)
        book4.should be_alive
        book4.should be_a Workbook
        book4.excel.should == book2.excel
        book4.ReadOnly.should be true
        book4.should == book2
        book4.close
      end

      it "should raise an error if no Excel or Workbook is given" do
        expect{
          Workbook.open(@simple_file1, :force => {:excel => :b})
          }.to raise_error(TypeREOError, @error_message_excel)
      end

      it "should do force_excel even if both force_ and default_excel is given" do
        book2 = Workbook.open(@simple_file1, :default => {:excel => @book.excel}, :force => {:excel => :new})
        book2.should be_alive
        book2.should be_a Workbook
        book2.excel.should_not == @book.excel 
        book2.should_not == @book
      end

      it "should do default_excel if force_excel is nil" do
        book2 = Workbook.open(@another_simple_file, :force => {:excel => nil})
        book2.should be_alive
        book2.should be_a Workbook
        book2.excel.should == @book.excel
      end

      it "should force_excel with :reuse" do
        book2 = Workbook.open(@different_file, :force => {:excel => :current})
        book2.should be_alive
        book2.should be_a Workbook
        book2.excel.should == @book.excel
      end

      it "should force_excel with :reuse even if :default_excel says sth. else" do
        book2 = Workbook.open(@different_file, :force => {:excel => :current}, :default => {:excel => :new})
        book2.should be_alive
        book2.should be_a Workbook
        book2.excel.should == @book.excel
      end

      it "should force_excel with :reuse when reopening and the Excel is not alive even if :default_excel says sth. else" do
        excel2 = Excel.new(:reuse => false)
        @book.excel.close
        book2 = Workbook.open(@simple_file1, :force => {:excel => :current}, :default => {:excel => :new})
        book2.should be_alive
        book2.should be_a Workbook
        book2.excel.should === excel2
      end

      it "should force_excel with :reuse when reopening and the Excel is not alive even if :default_excel says sth. else" do
        book2 = Workbook.open(@different_file1, :force => {:excel => :new})
        book2.excel.close
        book3 = Workbook.open(@different_file1, :force => {:excel => :current}, :default => {:excel => :new})
        book3.should be_alive
        book3.should be_a Workbook
        book3.excel.should == @book.excel
      end

    end

    context "with leaving out :force => {:excel}" do

      before do
        @book = Workbook.open(@simple_file1)
      end

      after do
        @book.close rescue nil
      end

      it "should open in a given Excel provided as Excel, Workbook, or WIN32OLE representing an Excel or Workbook" do
        book2 = Workbook.open(@another_simple_file)
        book3 = Workbook.open(@different_file)
        book3 = Workbook.open(@simple_file1, :excel => book2.excel)
        book3.excel.should === book2.excel
        book4 = Workbook.open(@simple_file1, :excel => @book) 
        book4.excel.should === @book.excel
        book3.close
        book4.close
        book5 = Workbook.open(@simple_file1, :excel => book2.ole_workbook)
        book5.excel.should ===  book2.excel
        win32ole_excel1 = WIN32OLE.connect(@book.ole_workbook.Fullname).Application
        book6 = Workbook.open(@simple_file1, :excel => win32ole_excel1)
        book6.excel.should === @book.excel
      end

      it "should open in a new Excel" do
        book2 = Workbook.open(@simple_file1, :excel => :new)
        book2.should be_alive
        book2.should be_a Workbook
        book2.excel.should_not == @book.excel 
        book2.should_not == @book
        @book.Readonly.should be false
        book2.Readonly.should be true
        book2.close
      end

      it "should open in a given Excel, not provide identity transparency, because old book readonly, new book writable" do
        book2 = Workbook.open(@simple_file1, :excel => :new)
        book2.excel.should_not == @book.excel
        book3 = Workbook.open(@simple_file1, :excel => :new)
        book3.excel.should_not == book2.excel
        book3.excel.should_not == @book.excel
        book2.close
        book4 = Workbook.open(@simple_file1, :excel => book2.excel)
        book4.should be_alive
        book4.should be_a Workbook
        book4.excel.should == book2.excel
        book4.Readonly.should == true
        book4.should_not == book2 
        book4.close
        book5 = Workbook.open(@simple_file1, :excel => book2)
        book5.should be_alive
        book5.should be_a Workbook
        book5.excel.should == book2.excel
        book5.Readonly.should == true
        book5.should_not == book2 
        book5.close
        book3.close
      end

      it "should open in a given Excel, provide identity transparency, because book can be readonly, such that the old and the new book are readonly" do
        book2 = Workbook.open(@simple_file1, :excel => :new)
        book2.excel.should_not == @book.excel
        book3 = Workbook.open(@simple_file1, :excel => :new)
        book3.excel.should_not == book2.excel
        book3.excel.should_not == @book.excel
        book2.close
        book3.close
        @book.close
        book4 = Workbook.open(@simple_file1, :excel => book2.excel, :read_only => true)
        book4.should be_alive
        book4.should be_a Workbook
        book4.excel.should == book2.excel
        book4.ReadOnly.should be true
        book4.should == book2
        book4.close
        book5 = Workbook.open(@simple_file1, :excel => book2, :read_only => true)
        book5.should be_alive
        book5.should be_a Workbook
        book5.excel.should == book2.excel
        book5.ReadOnly.should be true
        book5.should == book2
        book5.close
        book3.close
      end

      it "should open in a given Excel, provide identity transparency, because book can be readonly, such that the old and the new book are readonly" do
        book2 = Workbook.open(@simple_file1, :excel => :new)
        book2.excel.should_not == @book.excel
        book2.close
        @book.close
        book4 = Workbook.open(@simple_file1, :excel => book2, :read_only => true)
        book4.should be_alive
        book4.should be_a Workbook
        book4.excel.should == book2.excel
        book4.ReadOnly.should be true
        book4.should == book2
        book4.close
      end

      it "should raise an error if no Excel or Workbook is given" do
        expect{
          Workbook.open(@simple_file1, :excel => :b)
          }.to raise_error(TypeREOError, @error_message_excel)
      end

      it "should do force_excel even if both force_ and default_excel is given" do
        book2 = Workbook.open(@simple_file1, :default => {:excel => @book.excel}, :force => {:excel => :new})
        book2.should be_alive
        book2.should be_a Workbook
        book2.excel.should_not == @book.excel 
        book2.should_not == @book
      end

      it "should do default_excel if force_excel is nil" do
        book2 = Workbook.open(@another_simple_file, :excel => nil)
        book2.should be_alive
        book2.should be_a Workbook
        book2.excel.should == @book.excel
      end

      it "should force_excel with :reuse" do
        book2 = Workbook.open(@different_file, :excel => :current)
        book2.should be_alive
        book2.should be_a Workbook
        book2.excel.should == @book.excel
      end

      it "should force_excel with :reuse even if :default_excel says sth. else" do
        book2 = Workbook.open(@different_file, :excel => :current, :default => {:excel => :new})
        book2.should be_alive
        book2.should be_a Workbook
        book2.excel.should == @book.excel
      end

      it "should force_excel with :reuse when reopening and the Excel is not alive even if :default_excel says sth. else" do
        excel2 = Excel.new(:reuse => false)
        @book.excel.close
        book2 = Workbook.open(@simple_file1, :excel => :current, :default => {:excel => :new})
        book2.should be_alive
        book2.should be_a Workbook
        book2.excel.should === excel2
      end

      it "should force_excel with :reuse when reopening and the Excel is not alive even if :default_excel says sth. else" do
        book2 = Workbook.open(@different_file1, :excel => :new)
        book2.excel.close
        book3 = Workbook.open(@different_file1, :excel => :current, :default => {:excel => :new})
        book3.should be_alive
        book3.should be_a Workbook
        book3.excel.should == @book.excel
      end
   
    end


    context "with :force_excel" do

      before do
        @book = Workbook.open(@simple_file1)
      end

      after do
        @book.close rescue nil
      end

      it "should open in a given Excel provided as Excel, Workbook, or WIN32OLE representing an Excel or Workbook" do
        book2 = Workbook.open(@another_simple_file)
        book3 = Workbook.open(@different_file)
        book3 = Workbook.open(@simple_file1, :force_excel => book2.excel)
        book3.excel.should === book2.excel
        book4 = Workbook.open(@simple_file1, :force_excel => @book) 
        book4.excel.should === @book.excel
        book3.close
        book4.close
        book5 = Workbook.open(@simple_file1, :force_excel => book2.ole_workbook)
        book5.excel.should ===  book2.excel
        win32ole_excel1 = WIN32OLE.connect(@book.ole_workbook.Fullname).Application
        book6 = Workbook.open(@simple_file1, :force_excel => win32ole_excel1)
        book6.excel.should === @book.excel
      end


      it "should open in a new Excel" do
        book2 = Workbook.open(@simple_file1, :force_excel => :new)
        book2.should be_alive
        book2.should be_a Workbook
        book2.excel.should_not == @book.excel 
        book2.should_not == @book
        @book.Readonly.should be false
        book2.Readonly.should be true
        book2.close
      end

      it "should open in a given Excel, not provide identity transparency, because old book readonly, new book writable" do
        book2 = Workbook.open(@simple_file1, :force_excel => :new)
        book2.excel.should_not == @book.excel
        book3 = Workbook.open(@simple_file1, :force_excel => :new)
        book3.excel.should_not == book2.excel
        book3.excel.should_not == @book.excel
        book2.close
        book4 = Workbook.open(@simple_file1, :force_excel => book2.excel)
        book4.should be_alive
        book4.should be_a Workbook
        book4.excel.should == book2.excel
        book4.Readonly.should == true
        book4.should_not == book2 
        book4.close
        book5 = Workbook.open(@simple_file1, :force_excel => book2)
        book5.should be_alive
        book5.should be_a Workbook
        book5.excel.should == book2.excel
        book5.Readonly.should == true
        book5.should_not == book2 
        book5.close
        book3.close
      end

      it "should open in a given Excel, provide identity transparency, because book can be readonly, such that the old and the new book are readonly" do
        book2 = Workbook.open(@simple_file1, :force_excel => :new)
        book2.excel.should_not == @book.excel
        book3 = Workbook.open(@simple_file1, :force_excel => :new)
        book3.excel.should_not == book2.excel
        book3.excel.should_not == @book.excel
        book2.close
        book3.close
        @book.close
        book4 = Workbook.open(@simple_file1, :force_excel => book2.excel, :read_only => true)
        book4.should be_alive
        book4.should be_a Workbook
        book4.excel.should == book2.excel
        book4.ReadOnly.should be true
        book4.should == book2
        book4.close
        book5 = Workbook.open(@simple_file1, :force_excel => book2, :read_only => true)
        book5.should be_alive
        book5.should be_a Workbook
        book5.excel.should == book2.excel
        book5.ReadOnly.should be true
        book5.should == book2
        book5.close
        book3.close
      end

      it "should open in a given Excel, provide identity transparency, because book can be readonly, such that the old and the new book are readonly" do
        book2 = Workbook.open(@simple_file1, :force_excel => :new)
        book2.excel.should_not == @book.excel
        book2.close
        @book.close
        book4 = Workbook.open(@simple_file1, :force_excel => book2, :read_only => true)
        book4.should be_alive
        book4.should be_a Workbook
        book4.excel.should == book2.excel
        book4.ReadOnly.should be true
        book4.should == book2
        book4.close
      end

      it "should raise an error if no Excel or Workbook is given" do
        expect{
          Workbook.open(@simple_file1, :force_excel => :b)
          }.to raise_error(TypeREOError, @error_message_excel)
      end

      it "should do force_excel even if both force_ and default_excel is given" do
        book2 = Workbook.open(@simple_file1, :default_excel => @book.excel, :force_excel => :new)
        book2.should be_alive
        book2.should be_a Workbook
        book2.excel.should_not == @book.excel 
        book2.should_not == @book
      end

      it "should do default_excel if force_excel is nil" do
        book2 = Workbook.open(@another_simple_file, :force_excel => nil)
        book2.should be_alive
        book2.should be_a Workbook
        book2.excel.should == @book.excel
      end

      it "should force_excel with :reuse" do
        book2 = Workbook.open(@different_file, :force_excel => :current)
        book2.should be_alive
        book2.should be_a Workbook
        book2.excel.should == @book.excel
      end

      it "should force_excel with :reuse even if :default_excel says sth. else" do
        book2 = Workbook.open(@different_file, :force_excel => :current, :default_excel => :new)
        book2.should be_alive
        book2.should be_a Workbook
        book2.excel.should == @book.excel
      end

      it "should force_excel with :reuse when reopening and the Excel is not alive even if :default_excel says sth. else" do
        excel2 = Excel.new(:reuse => false)
        @book.excel.close
        book2 = Workbook.open(@simple_file1, :force_excel => :current, :default_excel => :new)
        book2.should be_alive
        book2.should be_a Workbook
        book2.excel.should === excel2
      end

      it "should force_excel with :reuse when reopening and the Excel is not alive even if :default_excel says sth. else" do
        book2 = Workbook.open(@different_file1, :force_excel => :new)
        book2.excel.close
        book3 = Workbook.open(@different_file1, :force_excel => :current, :default_excel => :new)
        book3.should be_alive
        book3.should be_a Workbook
        book3.excel.should == @book.excel
      end
   
    end

    context "with :default => {:excel}" do

      before do
        @book = Workbook.open(@simple_file1, :visible => true)
      end

      after do
        @book.close rescue nil
      end

      context "with :default => {:excel => :current}" do

        it "should use the open book" do
          book2 = Workbook.open(@simple_file1, :default => {:excel => :current})
          book2.excel.should == @book.excel
          book2.should be_alive
          book2.should be_a Workbook
          book2.should == @book
          book2.close
        end

        it "should reopen the book in the excel instance where it was opened before" do
          excel = Excel.new(:reuse => false)
          @book.close
          book2 = Workbook.open(@simple_file1)
          book2.should be_alive
          book2.should be_a Workbook
          book2.excel.should == @book.excel
          book2.excel.should_not == excel
          book2.filename.should == @book.filename
          @book.should be_alive
          book2.should == @book
          book2.close
        end

        it "should reopen a book in a new Excel if all Excel instances are closed" do
          excel = Excel.new(:reuse => false)
          excel2 = @book.excel
          fn = @book.filename
          @book.close
          Excel.kill_all
          book2 = Workbook.open(@simple_file1, :default => {:excel => :current})
          book2.should be_alive
          book2.should be_a Workbook
          book2.filename.should == fn
          @book.should be_alive
          book2.should == @book
          book2.close
        end

        it "should reopen a book in the first opened Excel if the old Excel is closed" do
          excel = @book.excel
          Excel.kill_all
          new_excel = Excel.new(:reuse => false)
          new_excel2 = Excel.new(:reuse => false)
          book2 = Workbook.open(@simple_file1, :default => {:excel => :current})
          book2.should be_alive
          book2.should be_a Workbook
          book2.excel.should_not == excel
          book2.excel.should_not == new_excel2
          book2.excel.should == new_excel
          @book.should be_alive
          book2.should == @book
          book2.close
        end

        it "should reopen a book in the first opened excel, if the book cannot be reopened" do
          @book.close
          Excel.kill_all
          excel1 = Excel.new(:reuse => false)
          excel2 = Excel.new(:reuse => false)
          book2 = Workbook.open(@different_file, :default => {:excel => :current})
          book2.should be_alive
          book2.should be_a Workbook
          book2.excel.should == excel1
          book2.excel.should_not == excel2
          book2.close
        end

        it "should reopen the book in the Excel where it was opened most recently" do
          excel1 = @book.excel
          excel2 = Excel.new(:reuse => false)
          @book.close
          book2 = Workbook.open(@simple_file1, :default => {:excel => :current})
          book2.excel.should == excel1
          book2.close
          book3 = Workbook.open(@simple_file1, :force => {:excel => excel2})
          book3.close
          book3 = Workbook.open(@simple_file1, :default => {:excel => :current})
          book3.excel.should == excel2
          book3.close
        end

      end

      context "with :default => {:excel => :new}" do

        it "should reopen a book in the excel instance where it was opened most recently" do
          book2 = Workbook.open(@simple_file, :force => {:excel => :new})
          @book.close
          book2.close
          book3 = Workbook.open(@simple_file1)
          book2.should be_alive
          book2.should be_a Workbook
          book3.excel.should == book2.excel
          book3.excel.should_not == @book.excel
          book3.should == book2
          book3.should_not == @book
        end

        it "should open the book in a new excel if the book was not opened before" do
          book2 = Workbook.open(@different_file, :default => {:excel => :current})
          book2.excel.should == @book.excel
          book3 = Workbook.open(@another_simple_file, :default => {:excel => :new})
          book3.excel.should_not == @book.excel
        end

        it "should open the book in a new excel if the book was opened before but the excel has been closed" do
          excel = @book.excel
          excel2 = Excel.new(:reuse => false)
          excel.close
          book2 = Workbook.open(@simple_file1, :default => {:excel => :new})
          book2.excel.should_not == excel2
          book2.close
        end

      end

      context "with :default => {:excel => <excel-instance>}" do

        it "should open the book in a given excel if the book was not opened before" do
          book2 = Workbook.open(@different_file, :default => {:excel => :current})
          book2.excel.should == @book.excel
          excel = Excel.new(:reuse => false)
          book3 = Workbook.open(@another_simple_file, :default => {:excel => excel})
          book3.excel.should == excel
        end

        it "should open the book in a given excel if the book was opened before but the excel has been closed" do
          excel2 = Excel.new(:reuse => false, :visible => true)
          @book.excel.close        
          book2 = Workbook.open(@simple_file1, :default => {:excel => excel2, :visible => true})
          book2.excel.should == excel2
        end

        it "should open a new excel, if the book cannot be reopened" do
          @book.close
          new_excel = Excel.new(:reuse => false)
          book2 = Workbook.open(@different_file, :default => {:excel => :new})
          book2.should be_alive
          book2.should be_a Workbook
          book2.excel.should_not == new_excel
          book2.excel.should_not == @book.excel
          book2.close
        end

        it "should open a given excel, if the book cannot be reopened" do
          @book.close
          new_excel = Excel.new(:reuse => false)
          book2 = Workbook.open(@different_file, :default => {:excel => @book.excel})
          book2.should be_alive
          book2.should be_a Workbook
          book2.excel.should_not == new_excel
          book2.excel.should == @book.excel
          book2.close
        end

        it "should open a given excel, if the book cannot be reopened" do
          @book.close
          new_excel = Excel.new(:reuse => false)
          book2 = Workbook.open(@different_file, :default => {:excel => @book})
          book2.should be_alive
          book2.should be_a Workbook
          book2.excel.should_not == new_excel
          book2.excel.should == @book.excel
          book2.close
        end

      end

      it "should reuse an open book by default" do
        book2 = Workbook.open(@simple_file1)
        book2.excel.should == @book.excel
        book2.should == @book
      end

      it "should raise an error if no Excel or Workbook is given" do
        expect{
          Workbook.open(@different_file, :default => {:excel => :a})
          }.to raise_error(TypeREOError, @error_message_excel)
      end
      
    end

    context "with :default_excel" do

      before do
        @book = Workbook.open(@simple_file1, :visible => true)
      end

      after do
        @book.close rescue nil
      end

      it "should use the open book" do
        book2 = Workbook.open(@simple_file1, :default_excel => :current)
        book2.excel.should == @book.excel
        book2.should be_alive
        book2.should be_a Workbook
        book2.should == @book
        book2.close
      end

      it "should reopen the book in the excel instance where it was opened before" do
        excel = Excel.new(:reuse => false)
        @book.close
        book2 = Workbook.open(@simple_file1)
        book2.should be_alive
        book2.should be_a Workbook
        book2.excel.should == @book.excel
        book2.excel.should_not == excel
        book2.filename.should == @book.filename
        @book.should be_alive
        book2.should == @book
        book2.close
      end

      it "should reopen a book in a new Excel if all Excel instances are closed" do
        excel = Excel.new(:reuse => false)
        excel2 = @book.excel
        fn = @book.filename
        @book.close
        Excel.kill_all
        book2 = Workbook.open(@simple_file1, :default_excel => :current)
        book2.should be_alive
        book2.should be_a Workbook
        book2.filename.should == fn
        @book.should be_alive
        book2.should == @book
        book2.close
      end

      it "should reopen a book in the first opened Excel if the old Excel is closed" do
        excel = @book.excel
        Excel.kill_all
        new_excel = Excel.new(:reuse => false)
        new_excel2 = Excel.new(:reuse => false)
        book2 = Workbook.open(@simple_file1, :default_excel => :current)
        book2.should be_alive
        book2.should be_a Workbook
        book2.excel.should_not == excel
        book2.excel.should_not == new_excel2
        book2.excel.should == new_excel
        @book.should be_alive
        book2.should == @book
        book2.close
      end

      it "should reopen a book in the first opened excel, if the book cannot be reopened" do
        @book.close
        Excel.kill_all
        excel1 = Excel.new(:reuse => false)
        excel2 = Excel.new(:reuse => false)
        book2 = Workbook.open(@different_file, :default_excel => :current)
        book2.should be_alive
        book2.should be_a Workbook
        book2.excel.should == excel1
        book2.excel.should_not == excel2
        book2.close
      end

      it "should reopen the book in the Excel where it was opened most recently" do
        excel1 = @book.excel
        excel2 = Excel.new(:reuse => false)
        @book.close
        book2 = Workbook.open(@simple_file1, :default_excel => :current)
        book2.excel.should == excel1
        book2.close
        book3 = Workbook.open(@simple_file1, :force_excel => excel2)
        book3.close
        book3 = Workbook.open(@simple_file1, :default_excel => :current)
        book3.excel.should == excel2
        book3.close
      end

      it "should reopen a book in the excel instance where it was opened most recently" do
        book2 = Workbook.open(@simple_file, :force_excel => :new)
        @book.close
        book2.close
        book3 = Workbook.open(@simple_file1)
        book2.should be_alive
        book2.should be_a Workbook
        book3.excel.should == book2.excel
        book3.excel.should_not == @book.excel
        book3.should == book2
        book3.should_not == @book
      end

      it "should open the book in a new excel if the book was not opened before" do
        book2 = Workbook.open(@different_file, :default_excel => :current)
        book2.excel.should == @book.excel
        book3 = Workbook.open(@another_simple_file, :default_excel => :new)
        book3.excel.should_not == @book.excel
      end

      it "should open the book in a new excel if the book was opened before but the excel has been closed" do
        excel = @book.excel
        excel2 = Excel.new(:reuse => false)
        excel.close
        book2 = Workbook.open(@simple_file1, :default_excel => :new)
        book2.excel.should_not == excel2
        book2.close
      end

      it "should open the book in a given excel if the book was not opened before" do
        book2 = Workbook.open(@different_file, :default_excel => :current)
        book2.excel.should == @book.excel
        excel = Excel.new(:reuse => false)
        book3 = Workbook.open(@another_simple_file, :default_excel => excel)
        book3.excel.should == excel
      end

      it "should open the book in a given excel if the book was opened before but the excel has been closed" do
        excel2 = Excel.new(:reuse => false, :visible => true)
        @book.excel.close        
        book2 = Workbook.open(@simple_file1, :visible => true, :default_excel => excel2)
        book2.excel.should == excel2
      end

      it "should open a new excel, if the book cannot be reopened" do
        @book.close
        new_excel = Excel.new(:reuse => false)
        book2 = Workbook.open(@different_file, :default_excel => :new)
        book2.should be_alive
        book2.should be_a Workbook
        book2.excel.should_not == new_excel
        book2.excel.should_not == @book.excel
        book2.close
      end

      it "should open a given excel, if the book cannot be reopened" do
        @book.close
        new_excel = Excel.new(:reuse => false)
        book2 = Workbook.open(@different_file, :default_excel => @book.excel)
        book2.should be_alive
        book2.should be_a Workbook
        book2.excel.should_not == new_excel
        book2.excel.should == @book.excel
        book2.close
      end

      it "should open a given excel, if the book cannot be reopened" do
        @book.close
        new_excel = Excel.new(:reuse => false)
        book2 = Workbook.open(@different_file, :default_excel => @book)
        book2.should be_alive
        book2.should be_a Workbook
        book2.excel.should_not == new_excel
        book2.excel.should == @book.excel
        book2.close
      end

      it "should reuse an open book by default" do
        book2 = Workbook.open(@simple_file1)
        book2.excel.should == @book.excel
        book2.should == @book
      end

      it "should raise an error if no Excel or Workbook is given" do
        expect{
          Workbook.open(@different_file, :default_excel => :a)
          }.to raise_error(TypeREOError, @error_message_excel)
      end
      
    end

    context "with :active instead of :current" do
      
      before do
        @book = Workbook.open(@simple_file1)
      end

      after do
        @book.close rescue nil
      end

      it "should force_excel with :active" do
        book2 = Workbook.open(@different_file, :force => {:excel => :active})
        book2.should be_alive
        book2.should be_a Workbook
        book2.excel.should == @book.excel
      end

      it "should force_excel with :reuse even if :default_excel says sth. else" do
        book2 = Workbook.open(@different_file, :force => {:excel => :active}, :default => {:excel => :new})
        book2.should be_alive
        book2.should be_a Workbook
        book2.excel.should == @book.excel
      end

      it "should open force_excel with :reuse when reopening and the Excel is not alive even if :default_excel says sth. else" do
        excel2 = Excel.new(:reuse => false)
        @book.excel.close
        book2 = Workbook.open(@simple_file1, :force => {:excel => :active}, :default => {:excel => :new})
        book2.should be_alive
        book2.should be_a Workbook
        book2.excel.should === excel2
      end

      it "should force_excel with :reuse when reopening and the Excel is not alive even if :default_excel says sth. else" do
        book2 = Workbook.open(@different_file1, :force => {:excel => :new})
        book2.excel.close
        book3 = Workbook.open(@different_file1, :force => {:excel => :active}, :default => {:excel => :new})
        book3.should be_alive
        book3.should be_a Workbook
        book3.excel.should == @book.excel
      end

      it "should use the open book" do
        book2 = Workbook.open(@simple_file1, :default => {:excel => :active})
        book2.excel.should == @book.excel
        book2.should be_alive
        book2.should be_a Workbook
        book2.should == @book
        book2.close
      end

      it "should reopen a book in a new Excel if all Excel instances are closed" do
        excel = Excel.new(:reuse => false)
        excel2 = @book.excel
        fn = @book.filename
        @book.close
        Excel.kill_all
        book2 = Workbook.open(@simple_file1, :default => {:excel => :active})
        book2.should be_alive
        book2.should be_a Workbook
        book2.filename.should == fn
        @book.should be_alive
        book2.should == @book
        book2.close
      end

      it "should reopen a book in the first opened Excel if the old Excel is closed" do
        excel = @book.excel
        Excel.kill_all
        new_excel = Excel.new(:reuse => false)
        new_excel2 = Excel.new(:reuse => false)
        book2 = Workbook.open(@simple_file1, :default => {:excel => :active})
        book2.should be_alive
        book2.should be_a Workbook
        book2.excel.should_not == excel
        book2.excel.should_not == new_excel2
        book2.excel.should == new_excel
        @book.should be_alive
        book2.should == @book
        book2.close
      end

      it "should reopen a book in the first opened excel, if the book cannot be reopened" do
        @book.close
        Excel.kill_all
        excel1 = Excel.new(:reuse => false)
        excel2 = Excel.new(:reuse => false)
        book2 = Workbook.open(@different_file, :default => {:excel => :active})
        book2.should be_alive
        book2.should be_a Workbook
        book2.excel.should == excel1
        book2.excel.should_not == excel2
        book2.close
      end

      it "should reopen the book in the Excel where it was opened most recently" do
        excel1 = @book.excel
        excel2 = Excel.new(:reuse => false)
        @book.close
        book2 = Workbook.open(@simple_file1, :default => {:excel => :active})
        book2.excel.should == excel1
        book2.close
        book3 = Workbook.open(@simple_file1, :force => {:excel => excel2})
        book3.close
        book3 = Workbook.open(@simple_file1, :default => {:excel => :active})
        book3.excel.should == excel2
        book3.close
      end

    end

    context "with :reuse instead of :current" do
      
      before do
        @book = Workbook.open(@simple_file1)
      end

      after do
        @book.close rescue nil
      end

      it "should force_excel with :reuse" do
        book2 = Workbook.open(@different_file, :force => {:excel => :reuse})
        book2.should be_alive
        book2.should be_a Workbook
        book2.excel.should == @book.excel
      end

      it "should force_excel with :reuse even if :default_excel says sth. else" do
        book2 = Workbook.open(@different_file, :force => {:excel => :reuse}, :default => {:excel => :new})
        book2.should be_alive
        book2.should be_a Workbook
        book2.excel.should == @book.excel
      end

      it "should open force_excel with :reuse when reopening and the Excel is not alive even if :default_excel says sth. else" do
        excel2 = Excel.new(:reuse => false)
        @book.excel.close
        book2 = Workbook.open(@simple_file1, :force => {:excel => :reuse}, :default => {:excel => :new})
        book2.should be_alive
        book2.should be_a Workbook
        book2.excel.should === excel2
      end

      it "should force_excel with :reuse when reopening and the Excel is not alive even if :default_excel says sth. else" do
        book2 = Workbook.open(@different_file1, :force => {:excel => :new})
        book2.excel.close
        book3 = Workbook.open(@different_file1, :force => {:excel => :reuse}, :default => {:excel => :new})
        book3.should be_alive
        book3.should be_a Workbook
        book3.excel.should == @book.excel
      end

      it "should use the open book" do
        book2 = Workbook.open(@simple_file1, :default => {:excel => :reuse})
        book2.excel.should == @book.excel
        book2.should be_alive
        book2.should be_a Workbook
        book2.should == @book
        book2.close
      end

      it "should reopen a book in a new Excel if all Excel instances are closed" do
        excel = Excel.new(:reuse => false)
        excel2 = @book.excel
        fn = @book.filename
        @book.close
        Excel.kill_all
        book2 = Workbook.open(@simple_file1, :default => {:excel => :reuse})
        book2.should be_alive
        book2.should be_a Workbook
        book2.filename.should == fn
        @book.should be_alive
        book2.should == @book
        book2.close
      end

      it "should reopen a book in the first opened Excel if the old Excel is closed" do
        excel = @book.excel
        Excel.kill_all
        new_excel = Excel.new(:reuse => false)
        new_excel2 = Excel.new(:reuse => false)
        book2 = Workbook.open(@simple_file1, :default => {:excel => :reuse})
        book2.should be_alive
        book2.should be_a Workbook
        book2.excel.should_not == excel
        book2.excel.should_not == new_excel2
        book2.excel.should == new_excel
        @book.should be_alive
        book2.should == @book
        book2.close
      end

      it "should reopen a book in the first opened excel, if the book cannot be reopened" do
        @book.close
        Excel.kill_all
        excel1 = Excel.new(:reuse => false)
        excel2 = Excel.new(:reuse => false)
        book2 = Workbook.open(@different_file, :default => {:excel => :reuse})
        book2.should be_alive
        book2.should be_a Workbook
        book2.excel.should == excel1
        book2.excel.should_not == excel2
        book2.close
      end

      it "should reopen the book in the Excel where it was opened most recently" do
        excel1 = @book.excel
        excel2 = Excel.new(:reuse => false)
        @book.close
        book2 = Workbook.open(@simple_file1, :default => {:excel => :reuse})
        book2.excel.should == excel1
        book2.close
        book3 = Workbook.open(@simple_file1, :force => {:excel => excel2})
        book3.close
        book3 = Workbook.open(@simple_file1, :default => {:excel => :reuse})
        book3.excel.should == excel2
        book3.close
      end

    end

    it "should new_excel" do
      book = Workbook.open(@simple_file1)
      book.sheet(1)[1,1] = "foo"
      book.Saved.should be false
      book2 = Workbook.open(@simple_file1, :if_unsaved => :new_excel)
    end

    context "with :if_unsaved" do

      before do
        @book = Workbook.open(@simple_file)
        @sheet = @book.sheet(1)
        @book.add_sheet(@sheet, :as => 'a_name')
        @book.visible = true
      end

      after do
        @book.close(:if_unsaved => :forget)
        @new_book.close rescue nil
      end

      it "should raise an error, if :if_unsaved is :raise" do
        expect {
          @new_book = Workbook.open(@simple_file, :if_unsaved => :raise)
        }.to raise_error(WorkbookNotSaved, /workbook is already open but not saved: "workbook.xls"/)
      end

      it "should let the book open, if :if_unsaved is :accept" do
        expect {
          @new_book = Workbook.open(@simple_file, :if_unsaved => :accept)
          }.to_not raise_error
        @book.should be_alive
        @new_book.should be_alive
        @new_book.should == @book
      end

      context "with :if_unsaved => :alert or :if_unsaved => :excel" do
        before do
         @key_sender = IO.popen  'ruby "' + File.join(File.dirname(__FILE__), '../helpers/key_sender.rb') + '" "Microsoft Excel" '  , "w"
        end

        after do
          @key_sender.close
        end

        it "should open the new book and close the unsaved book, if user answers 'yes'" do
          # "Yes" is the  default. --> language independent
          @key_sender.puts "{enter}"
          @new_book = Workbook.open(@simple_file1, :if_unsaved => :alert)
          @new_book.should be_alive
          @new_book.filename.downcase.should == @simple_file1.downcase
          @book.should_not be_alive
        end

        it "should open the new book and close the unsaved book, if user answers 'yes'" do
          # "Yes" is the  default. --> language independent
          @key_sender.puts "{enter}"
          @new_book = Workbook.open(@simple_file1, :if_unsaved => :excel)
          @new_book.should be_alive
          @new_book.filename.downcase.should == @simple_file1.downcase
          @book.should_not be_alive
        end

        it "should not open the new book and not close the unsaved book, if user answers 'no'" do
          # "No" is right to "Yes" (the  default). --> language independent
          # strangely, in the "no" case, the question will sometimes be repeated three times
          #@book.excel.Visible = true
          @key_sender.puts "{right}{enter}"
          @key_sender.puts "{right}{enter}"
          @key_sender.puts "{right}{enter}"
          #expect{
            Workbook.open(@simple_file, :if_unsaved => :alert)    
          #  }.to raise_error(UnexpectedREOError)
          @book.should be_alive
          @book.Saved.should be false
        end

        it "should not open the new book and not close the unsaved book, if user answers 'no'" do
          # "No" is right to "Yes" (the  default). --> language independent
          # strangely, in the "no" case, the question will sometimes be repeated three time
          @key_sender.puts "{right}{enter}"
          @key_sender.puts "{right}{enter}"
          @key_sender.puts "{right}{enter}"
          #expect{
            Workbook.open(@simple_file, :if_unsaved => :excel)
          #}.to raise_error(UnexpectedREOError)
          @book.should be_alive
          @book.Saved.should be false
        end

      end
    end

    context "with :if_obstructed" do

      for i in 1..2 do

        context "with and without reopen" do

          before do        
            if i == 1 then 
              book_before = Workbook.open(@simple_file1)
              book_before.close
            end
            @book = Workbook.open(@simple_file_other_path1)
            #@book.Windows(@book.Name).Visible = true
            #@sheet_count = @book.ole_workbook.Worksheets.Count
            sheet = @book.sheet(1)
            #@book.add_sheet(@sheet, :as => 'a_name')
            @old_value = sheet[1,1]
            sheet[1,1] = (sheet[1,1] == "foo" ? "bar" : "foo")
            @new_value = sheet[1,1]
            @book.Saved.should be false
          end

          after do
            @book.close(:if_unsaved => :forget)
            #@new_book.close rescue nil
          end

          it "should raise an error, if :if_obstructed is :raise" do
            expect {
              new_book = Workbook.open(@simple_file1, :if_obstructed => :raise)
            }.to raise_error(WorkbookBlocked, /blocked by/)
          end

          it "should close the other book and open the new book, if :if_obstructed is :forget" do
            new_book = Workbook.open(@simple_file1, :if_obstructed => :forget)
            @book.should_not be_alive
            new_book.should be_alive
            new_book.filename.downcase.should == @simple_file.downcase
            old_book = Workbook.open(@simple_file_other_path1, :if_obstructed => :forget)
            old_book.sheet(1)[1,1].should == @old_value
          end

          it "should let the old book open, if :if_obstructed is :accept" do
            new_book = Workbook.open(@simple_file1, :if_obstructed => :accept)
            @book.should be_alive
            new_book.should be_alive
            new_book.filename.downcase.should == @simple_file_other_path1.downcase
            old_book = Workbook.open(@simple_file_other_path1, :if_unsaved => :forget)
            old_book.sheet(1)[1,1].should == @old_value
          end

          it "should save the old book, close it, and open the new book, if :if_obstructed is :save" do
            new_book = Workbook.open(@simple_file1, :if_obstructed => :save)
            @book.should_not be_alive
            new_book.should be_alive
            new_book.filename.downcase.should == @simple_file1.downcase
            old_book = Workbook.open(@simple_file_other_path1, :if_obstructed => :forget)
            old_book.sheet(1)[1,1].should == @new_value
            #old_book.ole_workbook.Worksheets.Count.should ==  @sheet_count + 1
            old_book.close
          end

          it "should raise an error, if the old book is unsaved, and close the old book and open the new book, 
              if :if_obstructed is :close_if_saved" do
            expect{
              new_book = Workbook.open(@simple_file1, :if_obstructed => :close_if_saved)
            }.to raise_error(WorkbookBlocked, /same name in a different path/)
            @book.save
            new_book = Workbook.open(@simple_file1, :if_obstructed => :close_if_saved)
            @book.should_not be_alive
            new_book.should be_alive
            new_book.filename.downcase.should == @simple_file1.downcase
            old_book = Workbook.open(@simple_file_other_path1, :if_obstructed => :forget)
            old_book.sheet(1)[1,1].should == @new_value
            #old_book.ole_workbook.Worksheets.Count.should ==  @sheet_count + 1
            old_book.close
          end

          it "should close the old book and open the new book, if :if_obstructed is :close_if_saved" do
            @book.close(:if_unsaved => :forget)
            book = Workbook.open(@simple_file_other_path)
            book2 = Workbook.open(@simple_file1, :if_obstructed => :close_if_saved)
          end

          it "should open the book in a new excel instance, if :if_obstructed is :new_excel" do
            new_book = Workbook.open(@simple_file1, :if_obstructed => :new_excel)
            @book.should be_alive
            @book.Saved.should be false
            @book.sheet(1)[1,1].should == @new_value
            new_book.should be_alive
            new_book.filename.should_not == @book.filename
            new_book.excel.should_not == @book.excel
            new_book.sheet(1)[1,1].should == @old_value
          end

          it "should raise an error, if :if_obstructed is default" do
            expect {
              new_book = Workbook.open(@simple_file1)              
            }.to raise_error(WorkbookBlocked, /blocked by/)
          end         

          it "should raise an error, if :if_obstructed is invalid option" do
            expect {
              new_book = Workbook.open(@simple_file1, :if_obstructed => :invalid_option)
            }.to raise_error(OptionInvalid)  
            #}.to raise_error(OptionInvalid, ":if_obstructed: invalid option: :invalid_option" +
            #  "\nHint: Use the option :if_obstructed with values :forget or :save,
            # to close the old workbook, without or with saving before, respectively,
            # and to open the new workbook")
          end
        end
      end
    end

    context "with an already saved book" do
      before do
        @book = Workbook.open(@simple_file)
      end

      after do
        @book.close
      end

      possible_options = [:read_only, :raise, :accept, :forget, nil]
      possible_options.each do |options_value|
        context "with :if_unsaved => #{options_value} and in the same and different path" do
          before do
            @new_book = Workbook.open(@simple_file1, :reuse => true, :if_unsaved => options_value)
            @different_book = Workbook.open(@different_file, :reuse => true, :if_unsaved => options_value)
          end
          after do
            @new_book.close
            @different_book.close
          end
          it "should open without problems " do
            @new_book.should be_a Workbook
            @different_book.should be_a Workbook
          end
          it "should belong to the same Excel instance" do
            @new_book.excel.should == @book.excel
            @different_book.excel.should == @book.excel
          end
        end
      end
    end      
    
    context "with non-existing file" do

      it "should raise error if filename is nil" do
        expect{
          Workbook.open(@nonexisting)
          }.to raise_error(FileNameNotGiven, "filename is nil")
      end

      it "should raise error if file is a directory" do
        expect{
          Workbook.open(@dir)
          }.to raise_error(FileNotFound, "file #{General::absolute_path(@dir).gsub("/","\\").inspect} is a directory")
      end

      it "should raise error if file does not exist" do
        File.delete @simple_save_file rescue nil
        expect {
          Workbook.open(@simple_save_file, :if_absent => :raise)
        }.to raise_error(FileNotFound, "file #{General::absolute_path(@simple_save_file).gsub("/","\\").inspect} not found" +
          "\nHint: If you want to create a new file, use option :if_absent => :create or Workbook::create")
      end

      it "should create a workbook" do
        File.delete @simple_save_file rescue nil
        book = Workbook.open(@simple_save_file, :if_absent => :create)
        book.should be_a Workbook
        book.close
        File.exist?(@simple_save_file).should be true
      end

      it "should raise an exception by default" do
        File.delete @simple_save_file rescue nil
        expect {
          Workbook.open(@simple_save_file)
        }.to raise_error(FileNotFound, "file #{General::absolute_path(@simple_save_file).gsub("/","\\").inspect} not found" +
          "\nHint: If you want to create a new file, use option :if_absent => :create or Workbook::create")
      end

    end

    context "with attr_reader excel" do
     
      before do
        @new_book = Workbook.open(@simple_file)
      end
      after do
        @new_book.close
      end
      it "should provide the excel instance of the book" do
        excel = @new_book.excel
        excel.class.should == Excel
        excel.should be_a Excel
      end
    end

=begin
    # work in progress
    context "with :update_links" do
      
      it "should set update_links to :alert" do
        book = Workbook.open(@simple_file, :update_links => :alert)
        book.UpdateLinks.should == XlUpdateLinksUserSetting
        book.Saved.should be true
      end

      it "should set update_links to :never" do
        book = Workbook.open(@simple_file, :update_links => :never)
        book.UpdateLinks.should == XlUpdateLinksNever
        book = Workbook.open(@simple_file, :update_links => :foo)
        book.UpdateLinks.should == XlUpdateLinksNever
      end

      it "should set update_links to :always" do
        book = Workbook.open(@simple_file, :update_links => :always)
        book.UpdateLinks.should == XlUpdateLinksAlways
      end

      it "should set update_links to :never per default" do
        book = Workbook.open(@simple_file)
        book.UpdateLinks.should == XlUpdateLinksNever
      end

    end
=end

    context "with :read_only" do
      
      it "should raise error, when :if_unsaved => :accept and change readonly to false" do
        book = Workbook.open(@simple_file1, :read_only => true)
        book.ReadOnly.should be true
        book.should be_alive
        sheet = book.sheet(1)
        old_cell_value = sheet[1,1]
        sheet[1,1] = sheet[1,1] == "foo" ? "bar" : "foo"
        book.Saved.should be false
        expect{
          new_book = Workbook.open(@simple_file1, :read_only => false, :if_unsaved => :accept)
        }.to raise_error(OptionInvalid)
      end

      it "should raise error, when :if_unsaved => :accept and change readonly to false" do
        book = Workbook.open(@simple_file1, :read_only => false)
        book.ReadOnly.should be false
        book.should be_alive
        sheet = book.sheet(1)
        old_cell_value = sheet[1,1]
        sheet[1,1] = sheet[1,1] == "foo" ? "bar" : "foo"
        book.Saved.should be false
        expect{
          new_book = Workbook.open(@simple_file1, :read_only => true, :if_unsaved => :accept)
        }.to raise_error(OptionInvalid)
      end

      it "should not reopen the book with writable" do
        book = Workbook.open(@simple_file1, :read_only => true)
        book.ReadOnly.should be true
        book.should be_alive
        sheet = book.sheet(1)
        old_cell_value = sheet[1,1]
        sheet[1,1] = sheet[1,1] == "foo" ? "bar" : "foo"
        book.Saved.should be false
        new_book = Workbook.open(@simple_file1, :read_only => false, :if_unsaved => :forget)
        new_book.ReadOnly.should be false 
        new_book.should be_alive
        book.should_not be_alive   
        new_book.should_not == book 
        new_sheet = new_book.sheet(1)
        new_cell_value = new_sheet[1,1]
        new_cell_value.should == old_cell_value
      end

      it "should raise an error when trying to reopen the book as read_only while the writable book had unsaved changes" do
        book = Workbook.open(@simple_file1, :read_only => false)
        book.ReadOnly.should be false
        book.should be_alive
        sheet = book.sheet(1)
        old_cell_value = sheet[1,1]        
        sheet[1,1] = sheet[1,1] == "foo" ? "bar" : "foo"
        book.Saved.should be false
        expect{
          Workbook.open(@simple_file1, :read_only => true, :if_unsaved => :accept)
        }.to raise_error(OptionInvalid)
      end

      it "should not raise an error when trying to reopen the book as read_only while the writable book had unsaved changes" do
        book = Workbook.open(@simple_file1, :read_only => false)
        book.ReadOnly.should be false
        book.should be_alive
        sheet = book.sheet(1)
        old_cell_value = sheet[1,1]        
        sheet[1,1] = sheet[1,1] == "foo" ? "bar" : "foo"
        book.Saved.should be false
        new_book = Workbook.open(@simple_file1, :read_only => true, :if_unsaved => :save)
        new_book.ReadOnly.should be true
        new_sheet = new_book.sheet(1)
        new_cell_value = new_sheet[1,1]
        new_cell_value.should_not == old_cell_value
      end

      it "should open the second book in another Excel as writable" do
        book = Workbook.open(@simple_file1, :read_only => true)
        book.ReadOnly.should be true
        new_book = Workbook.open(@simple_file1, :force => {:excel => :new}, :read_only => false)
        new_book.ReadOnly.should be false
        new_book.close
        book.close
      end

      it "should be able to save, if :read_only => false" do
        book = Workbook.open(@simple_file1, :read_only => false)
        book.should be_a Workbook
        expect {
          book.save_as(@simple_save_file, :if_exists => :overwrite)
        }.to_not raise_error
        book.close
      end

      it "should be able to save, if :read_only is default" do
        book = Workbook.open(@simple_file1)
        book.should be_a Workbook
        expect {
          book.save_as(@simple_save_file, :if_exists => :overwrite)
        }.to_not raise_error
        book.close
      end

      it "should raise an error, if :read_only => true" do
        book = Workbook.open(@simple_file, :read_only => true)
        book.should be_a Workbook
        expect {
          book.save_as(@simple_save_file, :if_exists => :overwrite)
        }.to raise_error
        book.close
      end
    end

    context "with various file formats" do

      it "should open linked workbook" do
        book = Workbook.open(@main_file, :visible => true)
        book.close
      end

      it "should open xlsm file" do
        book = Workbook.open(@simple_file_xlsm, :visible => true)
        book.close
      end

      it "should open xlsx file" do
        book = Workbook.open(@simple_file_xlsx, :visible => true)
        book.close
      end
      
    end


    context "with block" do
      it 'block parameter should be instance of Workbook' do
        Workbook.open(@simple_file) do |book|
          book.should be_a Workbook
        end
      end
    end

    context "with WIN32OLE#GetAbsolutePathName" do
      it "'~' should be HOME directory" do
        path = '~/Abrakadabra.xlsx'
        expected_path = Regexp.new(File.expand_path(path).gsub(/\//, "."))
        expect {
          Workbook.open(path)
        }.to raise_error(FileNotFound, "file #{General::absolute_path(path).gsub("/","\\").inspect} not found" +
          "\nHint: If you want to create a new file, use option :if_absent => :create or Workbook::create")
      end
    end
  end

  describe "reopen" do

    context "with standard" do
      
      before do
        @book = Workbook.open(@simple_file)
      end

      after do
        @book.close
      end

      it "should reopen the closed book" do
        @book.should be_alive
        book1 = @book
        @book.close
        @book.should_not be_alive
        @book.open
        @book.should be_a Workbook
        @book.should be_alive
        @book.should === book1
      end
    end
  end  
end