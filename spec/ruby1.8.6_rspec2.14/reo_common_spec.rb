# -*- coding: utf-8 -*-

require File.join(File.dirname(__FILE__), './spec_helper')
require File.expand_path( '../../lib/robust_excel_ole/reo_common', __FILE__)

$VERBOSE = nil

include General
include RobustExcelOle

module RobustExcelOle

  describe REOCommon do

    it "should simple test" do
      puts "a"
    end

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
      @linked_file = @dir + '/workbook_linked.xlsm'
      @simple_file_xlsm = @dir + '/workbook.xls'
      @simple_file_xlsx = @dir + '/workbook.xlsx'
    end

    after do
      Excel.kill_all
      #rm_tmp(@dir)
    end

    describe "trace" do

      it "should put some number" do
        a = 4
        REOCommon::trace "some text #{a}"
      end

      it "should put another text" do
        b = Book.open(@simple_file)
        REOCommon::trace "book: #{b}"
      end
    end

    describe "own_methods" do

      before do
        @book1 = Book.open(@simple_file)
        @ole_workbook_methods = 
          ["Activate", "ActiveSheet", "Application", "Close", "FullName", "HasPassword", "Name", "Names", 
            "Password", "Protect", "ProtectSharing", "ProtectStructure", "Protect", "ReadOnly", "Save", 
            "SaveAs", "Saved", "Sheets", "Unprotect"]
        @book_methods = ["focus", "add_sheet", "alive?", "close", "filename", "nameval", "ole_object", 
                         "ole_workbook", "reopen", "save", "save_as", "saved", "set_nameval"]
        @ole_excel_methods = 
          ["ActiveCell", "ActiveSheet", "ActiveWorkbook", "Application",  "Calculate", "Cells", "Columns",
            "DisplayAlerts", "Evaluate", "Hwnd", "Name", "Names", "Quit", "Range", "Ready", "Save", 
            "Sheets", "UserName", "Value", "Visible", "Workbooks", "Worksheets"]
        @excel_methods = ["alive?", "book_class", "close", "displayalerts", "recreate", "visible", "with_displayalerts"] 
      end

      after do
        @book1.close
      end

      it "should do own_methods with popular ole_workbook and workbook methods" do
        ((@ole_workbook_methods + @book_methods) - @book1.own_methods).should be_empty
        (Object.instance_methods - @book1.own_methods).should == Object.instance_methods 
      end

      it "should do own_methods with popular ole_excel and excel methods" do
        ((@ole_excel_methods + @excel_methods) - @book1.excel.own_methods).should be_empty
         (Object.instance_methods - @book1.excel.own_methods).should == Object.instance_methods
      end

    end

    describe "Object methods" do

      before do
        @book = Book.open(@simple_file)
        @sheet = @book.sheet(1)
      end

      before do
        @book.close
      end

      it "should raise an error when asking excel of a sheet" do
        expect{
          @sheet.excel
          }.to raise_error(TypeREOError, "receiver instance is neither an Excel nor a Book")
      end

    end

    describe "misc" do

      LOG_TO_STDOUT = true
      REOCommon::trace "foo"

      LOG_TO_STDOUT = false
      REOCommon::trace "foo"

      REO_LOG_DIR = ""
      REOCommon::trace "foo"

      REO_LOG_DIR = "C:"
      REOCommon::trace "foo"

      REOCommon::tr1 "foo"

      h = {:a => {:c => 4}, :b => 2}
      REOCommon::puts_hash(h)
 
    end

  end
end
