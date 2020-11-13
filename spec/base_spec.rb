# -*- coding: utf-8 -*-

require File.join(File.dirname(__FILE__), './spec_helper')
require File.expand_path( '../../lib/robust_excel_ole/base', __FILE__)

$VERBOSE = nil

include General
include RobustExcelOle

module RobustExcelOle

  describe Base do

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

    describe "own_methods" do

      before do
        @book1 = Workbook.open(@simple_file)
        @ole_workbook_methods = 
          ["Activate", "ActiveSheet", "Application", "Close", "FullName", "HasPassword", "Name", "Names", 
            "Password", "Protect", "ProtectSharing", "ProtectStructure", "Protect", "ReadOnly", "Save", 
            "SaveAs", "Saved", "Sheets", "Unprotect"]
        @book_methods = ["focus", "add_sheet", "alive?", "close", "filename", "ole_object", 
                         "ole_workbook", "reopen", "save", "save_as", "saved"]
        @ole_excel_methods = 
          ["ActiveCell", "ActiveSheet", "ActiveWorkbook", "Application",  "Calculate", "Cells", "Columns",
            "DisplayAlerts", "Evaluate", "Hwnd", "Name", "Names", "Quit", "Range", "Ready", "Save", 
            "Sheets", "UserName", "Value", "Visible", "Workbooks", "Worksheets"]
        @excel_methods = ["alive?", "workbook_class", "close", "properties", "recreate", "with_displayalerts"] 
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

    describe "trace" do

      it "should put some number" do
        a = 4
        Base::trace "some text #{a}"
      end

      it "should put another text" do
        b = Workbook.open(@simple_file)
        Base::trace "book: #{b}"
      end
    end

    

    describe "misc" do

      it "should" do

        LOG_TO_STDOUT = true
        Base::trace "foo"

        LOG_TO_STDOUT = false
        Base::trace "foo"

        REO_LOG_DIR = ""
        Base::trace "foo"

        #REO_LOG_DIR = "C:"
        #Base::trace "foo"

        Base::tr1 "foo"

        h = {:a => {:c => 4}, :b => 2}
        Base::puts_hash(h)

      end
 
    end
  end
end
