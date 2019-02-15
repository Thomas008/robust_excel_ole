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

      it "should transform a1-format" do
        Address.a1("A2").should == "A2"
        Address.r1c1("A2").should == "Z2S1:Z2S1"
        Address.int_range("A2").should == [2..2,1..1]
      end

      it "should transform several-letter-a1-format" do
        Address.a1("ABO15").should == "ABO15"
        Address.r1c1("ABO15").should == "Z15S743:Z15S743"
        Address.int_range("ABO15").should == [15..15,743..743]
      end

      it "should transform complex a1-format" do
        Address.a1("A2:B3").should == "A2:B3"
        Address.r1c1("A2:B3").should == "Z2S1:Z3S2"
        Address.int_range("A2:B3").should == [2..3,1..2]
        Address.a1("S1:DP2").should == "S1:DP2"
        Address.r1c1("S1:DP2").should == "Z1S19:Z2S120"
        Address.int_range("S1:DP2").should == [1..2,19..120]
      end

      it "should transform infinite a1-format" do
        Address.a1("A:B").should == "A:B"
        Address.r1c1("A:B").should == "S1:S2"
        Address.int_range("A:B").should == [nil,1..2]
        Address.a1("1:3").should == "1:3"
        Address.r1c1("1:3").should == "Z1:Z3"
        Address.int_range("1:3").should == [1..3,nil]
        Address.a1("B").should == "B"
        Address.r1c1("B").should == "S2:S2"
        Address.int_range("B").should == [nil,2..2]
        Address.a1("3").should == "3"
        Address.r1c1("3").should == "Z3:Z3"
        Address.int_range("3").should == [3..3,nil]
      end

      it "should transform r1c1-format" do
        Address.r1c1("Z2S1").should == "Z2S1"
        Address.int_range("Z2S1").should == [2..2,1..1]
        expect{
          Address.a1("Z2S1")
        }.to raise_error(NotImplementedREOError)
      end

      it "should transform complex r1c1-format" do
        Address.r1c1("Z2S1:Z3S2").should == "Z2S1:Z3S2"
        Address.int_range("Z2S1:Z3S2").should == [2..3,1..2]
      end

      it "should transform int_range format" do
        Address.int_range([2..2,1..1]).should == [2..2,1..1]
        Address.r1c1([2..2,1..1]).should == "Z2S1:Z2S1"
        expect{
          Address.a1([2..2,1..1])
        }.to raise_error(NotImplementedREOError)
      end

      it "should transform simple int_range format" do
        Address.int_range([2,1]).should == [2..2,1..1]
        Address.r1c1([2,1]).should == "Z2S1:Z2S1"
      end

      it "should transform complex int_range format" do
        Address.int_range([2,"A"]).should == [2..2,1..1]
        Address.r1c1([2,"A"]).should == "Z2S1:Z2S1"
        Address.int_range([2,"A".."B"]).should == [2..2,1..2]
        Address.r1c1([2,"A".."B"]).should == "Z2S1:Z2S2"
        Address.int_range([1..2,"C"]).should == [1..2,3..3]
        Address.r1c1([1..2,"C"]).should == "Z1S3:Z2S3"
        Address.int_range([1..2,"C".."E"]).should == [1..2,3..5]
        Address.r1c1([1..2,"C".."E"]).should == "Z1S3:Z2S5"
        Address.int_range([2,3..5]).should == [2..2,3..5]
        Address.r1c1([2,3..5]).should == "Z2S3:Z2S5"
        Address.int_range([1..2,3..5]).should == [1..2,3..5]
        Address.r1c1([1..2,3..5]).should == "Z1S3:Z2S5"
      end

      it "should transform infinite int_range format" do
        Address.int_range([nil,1..2]).should == [nil,1..2]
        Address.r1c1([nil,1..2]).should == "S1:S2"
        Address.int_range([1..3,nil]).should == [1..3,nil]
        Address.r1c1([1..3,nil]).should == "Z1:Z3"
        Address.int_range([nil,2]).should == [nil,2..2]
        Address.r1c1([nil,2]).should == "S2:S2"
        Address.int_range([3,nil]).should == [3..3,nil]
        Address.r1c1([3,nil]).should == "Z3:Z3"
      end
    
      it "should raise an error" do
        expect{
          Address.a1("1A")
        }.to raise_error(AddressInvalid, /format not correct/)
        expect{
          Address.r1c1("A1B")
        }.to raise_error(AddressInvalid, /format not correct/)
        expect{
          Address.int_range(["A".."B","C".."D"])
        }.to raise_error(AddressInvalid, /format not correct/)
        expect{
          Address.int_range(["A".."B",1..2])
        }.to raise_error(AddressInvalid, /format not correct/)
        expect{
          Address.int_range(["A".."B",nil])
        }.to raise_error(AddressInvalid, /format not correct/)
        expect{
          Address.int_range(["A",1,2])
        }.to raise_error(AddressInvalid, /more than two components/)
      end

    end

    describe "trace" do

      it "should put some number" do
        a = 4
        REOCommon::trace "some text #{a}"
      end

      it "should put another text" do
        b = Workbook.open(@simple_file)
        REOCommon::trace "book: #{b}"
      end
    end

    describe "own_methods" do

      before do
        @book1 = Workbook.open(@simple_file)
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
        @excel_methods = ["alive?", "workbook_class", "close", "displayalerts", "recreate", "visible", "with_displayalerts"] 
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
        @book = Workbook.open(@simple_file)
        @sheet = @book.sheet(1)
      end

      before do
        @book.close
      end

      it "should raise an error when asking excel of a sheet" do
        expect{
          @sheet.excel
          }.to raise_error(TypeREOError, "receiver instance is neither an Excel nor a Workbook")
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
