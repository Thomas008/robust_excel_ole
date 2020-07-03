# -*- coding: utf-8 -*-

require File.join(File.dirname(__FILE__), './spec_helper')

$VERBOSE = nil

include General
include RobustExcelOle

module RobustExcelOle

  describe General do

    before(:all) do
      excel = Excel.new(:reuse => true)
      open_books = excel == nil ? 0 : excel.Workbooks.Count
      puts "*** open books *** : #{open_books}" if open_books > 0
      Excel.kill_all
    end

    before do
      @dir = create_tmpdir
      @listobject_file = @dir + '/workbook_listobjects.xlsx'
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

    describe "to_reo" do

      before do
        @book1 = Workbook.open(@simple_file)
        @book2 = Workbook.open(@listobject_file)        
      end

      it "should type-lift an ListObject" do
        worksheet = @book1.sheet(3)
        ole_table = worksheet.ListObjects.Item(1)
        table = Table.new(ole_table)
        table.Name.should == "table3"
        table.HeaHeaderRowRange.Value.first.should == ["Number","Person","Amount","Time","Date"]
        table.ListRows.Count.should == 6
        worksheet[3,4].Value.should == "Number"
      end

      it "should type-lift an Excel" do
        excel = @book1.excel.ole_excel.to_reo
        excel.class.should == RobustExcelOle::Excel
        excel.should be_alive
      end

      it "should type-lift a workbook" do
        workbook = @book1.ole_workbook.to_reo
        workbook.should be_a Workbook
        workbook.should be_alive
      end

      it "should type-lift a worksheet" do
        worksheet = @book1.sheet(1).ole_worksheet.to_reo
        worksheet.should be_kind_of Worksheet
        worksheet.name.should == "Sheet1"
      end

      it "should type-lift a range" do
        range = @book1.sheet(1).range([1..2,1]).ole_range.to_reo
        range.should be_kind_of Range
        range.Value.should == [["foo"],["foo"]]
      end

      it "should type-lift a cell" do
        cell = @book1.sheet(1).range([1,1]).ole_range.to_reo
        cell.should be_kind_of Cell
        cell.Value.should == "foo"
      end

      it "should not do anything with a REO object" do
        @book1.to_reo.should == @book1 
        @book1.sheet(1).to_reo.should == @book1.sheet(1)
        @book1.excel.to_reo.should == @book1.excel
        @book1.sheet(1).range([1,1]).to_reo.should == @book1.sheet(1).range([1,1])
      end

    end

    describe "methods, own_methods, respond_to?" do

      before do
        @book1 = Workbook.open(@simple_file)
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
        @excel_methods = ["alive?", "workbook_class", "close", "properties", "recreate", "with_displayalerts"] 
        @ole_sheet_methods = []
         # ["Activate", "Calculate", "Copy", "Name", "Select", "Evaluate", "Protect", "Unprotect"]
        @sheet_methods = ["workbook_class", "col_range", "each", "each_column", "each_column_with_index",
                          "each_row", "each_row_with_index", "nameval", "namevalue", 
                          "set_namevalue", "row_range", "set_nameval"]
      end

      after do
        @book1.close
      end

      it "should do methods for book" do
        ((@ole_workbook_methods + @book_methods) - @book1.methods).should be_empty
        # (Object.instance_methods.select{|m| m =~ /^(?!\_)/} - @book1.methods).should be_empty
      end

      it "should do own_methods with popular ole_workbook and workbook methods" do
        ((@ole_workbook_methods + @book_methods) - @book1.own_methods).should be_empty
        (Object.instance_methods - @book1.own_methods).should == Object.instance_methods 
      end

      it "should respond to popular workbook methods" do
        @book_methods.each{|m| @book1.respond_to?(m).should be true}
      end

      it "should do methods for excel" do
        ((@ole_excel_methods + @excel_methods) - @book1.excel.methods).should be_empty
        #(Object.instance_methods.select{|m| m =~ /^(?!\_)/}  - @book1.excel.methods).sort.should be_empty       
      end

      it "should do own_methods with popular ole_excel and excel methods" do
        ((@ole_excel_methods + @excel_methods) - @book1.excel.own_methods).should be_empty
         (Object.instance_methods - @book1.excel.own_methods).should == Object.instance_methods
      end

      it "should respond to popular excel methods" do
        @excel_methods.each{|m| @book1.excel.respond_to?(m).should be true}
      end

      it "should do methods for sheet" do
        # ((@ole_sheet_methods + @sheet_methods) - @book1.sheet(1).methods).should be_empty
        (Object.instance_methods.select{|m| m =~ /^(?!\_)/}  - @book1.sheet(1).methods).sort.should be_empty       
      end

      it "should do own_methods with popular ole_excel and excel methods" do
        # ((@ole_sheet_methods + @sheet_methods) - @book1.sheet(1).own_methods).should == [] #be_empty
         (Object.instance_methods - @book1.sheet(1).own_methods).should == Object.instance_methods
      end

      it "should respond to popular sheet methods" do
        @sheet_methods.each{|m| @book1.sheet(1).respond_to?(m).should be true}
      end

    end

    describe "#absolute_path" do
      
      context "with standard" do
        
        before do
          @previous_dir = Dir.getwd
        end
        
        after do
          Dir.chdir @previous_dir
        end

        it "should return the right absolute paths" do
          absolute_path("C:/abc").should == "C:\\abc"
          #absolute_path("C:\\abc").should == "C:\\abc"
          Dir.chdir "C:/windows"
          absolute_path("C:abc").downcase.should == Dir.pwd.gsub("/","\\").downcase + "\\abc"
          absolute_path("C:abc").upcase.should   == File.expand_path("abc").gsub("/","\\").upcase
        end

        it "should return right absolute path name" do
          filename = 'C:/Dokumente und Einstellungen/Zauberthomas/Eigene Dateien/robust_excel_ole/spec/book_spec.rb'
          absolute_path(filename).gsub("\\","/").should == filename
        end
      end
    end

    describe "canonize" do

      context "with standard" do
        
        it "should reduce slash at the end" do
          normalize("hallo/").should == "hallo"
          normalize("/this/is/the/Path/").should == "/this/is/the/Path"
        end

        it "should save capital letters" do
          normalize("HALLO/").should == "HALLO"
          normalize("/This/IS/tHe/patH/").should == "/This/IS/tHe/patH"
        end

        it "should reduce multiple shlashes" do
          normalize("/this/is//the/path").should == "/this/is/the/path"
          normalize("///this/////////is//the/path/////").should == "/this/is/the/path"
        end

        it "should reduce dots in the paths" do
          canonize("/this/is/./the/path").should == "/this/is/the/path"
          canonize("this/.is/./the/pa.th/").should == "this/.is/the/pa.th"
          canonize("this//.///.//.is/the/pa.th/").should == "this/.is/the/pa.th"
        end

        it "should change to the upper directory with two dots" do
          canonize("/this/is/../the/path").should == "/this/the/path"
          canonize("this../.i.s/.../..the/..../pa.th/").should == "this../.i.s/.../..the/..../pa.th"
        end

        it "should downcase" do
          canonize("/This/IS/tHe/path").should == "/this/is/the/path"
          canonize("///THIS/.///./////iS//the/../PatH/////").should == "/this/is/path"
        end

        it "should raise an error for no strings" do
          expect{
            canonize(1)
          }.to raise_error(TypeREOError, "No string given to canonize, but 1")
        end

      end
    end

    describe "path" do

      it "should create a path" do
        path1 = "this" / "is" / "a" / "path"
        path1.should == "this/is/a/path"
        path2 = "this" / "is" / "a" / "path" / 
        #path2.should == "this/is/a/path/"
        path3 = "this" / 
        #path3.should == "this/"
        path4 = "this" / nil
        path4.should == "this"
      end
    end

  end
end
