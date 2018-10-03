# -*- coding: utf-8 -*-

require File.join(File.dirname(__FILE__), './spec_helper')
require File.expand_path( '../../lib/robust_excel_ole/reo_common', __FILE__)

$VERBOSE = nil

include General
include RobustExcelOle

module RobustExcelOle

  describe REOCommon do

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
      rm_tmp(@dir)
    end

    describe "Address" do

      it "should read a1-format" do
        address = Address.new("A1")
        address.rows.should == (1..1)
        address.columns.should == (1..1)
      end

      it "should read a1-format with brackets" do
        address = Address.new(["A1"])
        address.rows.should == (1..1)
        address.columns.should == (1..1)
      end

      it "should read a1-format" do
        address = Address.new("ABO15")
        address.columns.should == (743..743)
        address.rows.should == (15..15)
      end

      it "should read a1-format when row and column are given separated" do
        address = Address.new([1,"A"])
        address.rows.should == (1..1)
        address.columns.should == (1..1)
      end

      it "should read a1-format with rows as integer range" do
        address = Address.new([2..4, "AB"])
        address.rows.should == (2..4)
        address.columns.should == (28..28)
      end

      it "should read a1-format with columns as string range" do
        address = Address.new([2, "A".."C"])
        address.rows.should == (2..2)
        address.columns.should == (1..3)
      end

      it "should read a1-format with rows and columns as string range" do
        address = Address.new([2..6, "A" .. "C"])
        address.rows.should == (2..6)
        address.columns.should == (1..3)
      end

      it "should read r1c1-format" do
        address = Address.new([1,2])
        address.rows.should == (1..1)
        address.columns.should == (2..2)
      end

      it "should read r1c1-format with rows as integer range" do
        address = Address.new([1..2,3])
        address.rows.should == (1..2)
        address.columns.should == (3..3)
      end
     
      it "should read r1c1-format with columns as integer range" do
        address = Address.new([1,3..5])
        address.rows.should == (1..1)
        address.columns.should == (3..5)
      end

      it "should read r1c1-format with rows and columns as integer range" do
        address = Address.new([1..4,3..5])
        address.rows.should == (1..4)
        address.columns.should == (3..5)
      end

      it "should read a1-format for a rectangular range" do
        address = Address.new(["A1:B3"])
        address.rows.should == (1..3)
        address.columns.should == (1..2)
      end

      it "should read a1-format for a rectangular range without brackets" do
        address = Address.new("A1:B3")
        address.rows == (1..3)
        address.columns == (1..2)
      end

      it "should raise an error" do
        expect{
          Address.new("1A")
        }.to raise_error(AddressInvalid, /not in A1/)
        expect{
          Address.new("A1B")
        }.to raise_error(AddressInvalid, /not in A1/)
        expect{
          Address.new(["A".."B","C".."D"])
        }.to raise_error(AddressInvalid, /not in A1/)
        expect{
          Address.new(["A",1,2])
        }.to raise_error(AddressInvalid, /more than two components/)
      end

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
        @book_methods = ["focus", "add_sheet", "alive?", "close", "filename", "namevalue", "ole_object", 
                         "ole_workbook", "reopen", "save", "save_as", "saved", "set_namevalue"]
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
