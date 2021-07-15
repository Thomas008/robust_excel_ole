# -*- coding: utf-8 -*-

require 'win32ole'

class WIN32OLE
  def save
    "save"
  end
end

NETWORK = WIN32OLE.new('WScript.Network')

require_relative 'spec_helper'

$VERBOSE = nil

include General
include RobustExcelOle

module RobustExcelOle

  using StringRefinement
  using ToReoRefinement
  using FindAllIndicesRefinement

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
      @network_path = "N:/data/workbook.xls"
      computer_name = NETWORK.ComputerName
      #@hostname_share_path = "//#{computer_name}/spec/data/workbook.xls"
      @hostname_share_path = "//#{computer_name}/#{absolute_path('gems/robust_excel_ole/spec/data/workbook.xls').tr('\\','/').gsub('C:','c$')}"
      @network_path_downcase = @network_path.downcase
      @hostname_share_path_downcase = @hostname_share_path.downcase
      @simple_file_extern = "D:/data/workbook.xls"
    end

    after do
      Excel.kill_all
      rm_tmp(@dir)
    end

    describe "General.init_reo_for_win32ole" do

      before do
        @book1 = Workbook.open(@simple_file, :visible => true)
      end

      it "should preserve the instance method of a win32ole object via calling an aliased method" do
        NETWORK.save.should == "save"
      end

      it "should preserve the instance methods of a win32ole object " do
        RobustExcelOle::Excel.define_method(:ComputerName){ "computer" }
        network = WIN32OLE.new('WScript.Network')
        computername = network.ComputerName
        General.init_reo_for_win32ole
        network.ComputerName.should == computername
      end

      #it "should preserve the lower-case instance methods of a win32ole object " do
      #  RobustExcelOle::Excel.define_method(:computername){ "computer" }
      #  network = WIN32OLE.new('WScript.Network')
      #  computername = network.computername
      #  General.init_reo_for_win32ole
      #  network.computername.should == computername
      #end

      it "should call a capitalized method for an instance method occurring in one classes" do
        expect{
          NETWORK.delete_empty_columns
        }.to raise_error(NoMethodError, /Delete_empty_columns/)
      end


      it "should call a capitalized method for an instance method occurring in several classes" do
        expect{
          NETWORK.focus
        }.to raise_error(NoMethodError, /Focus/)
      end

      it "should apply reo-methods to win32ole objects" do
        ole_book1 = @book1.ole_workbook
        sheet1 = ole_book1.sheet(1)
        sheet1.should be_a Worksheet
        sheet1.name.should == "Sheet1"
        ole_sheet1 = sheet1.ole_worksheet
        range1 = ole_sheet1.range([1..2,3..4])
        range1.should be_a RobustExcelOle::Range
        range1.value.should == [["sheet1"],["foobaaa"]]
        ole_range1 = range1.ole_range
        ole_range1.copy([6,6])
        range2 = sheet1.range([6..7,6..7])
        range2.value.should == [["sheet1"],["foobaaa"]]
        excel1 = @book1.excel
        ole_excel1 = excel1.ole_excel
        ole_excel1.close(:if_unsaved => :forget)
        Excel.kill_all
      end

    end

    describe "find_all_indices" do

      it "should find all occurrences" do
        [1,2,3,1].find_all_indices(1).should == [0,3]
        [1,2,3,1].find_all_indices(4).should be_empty
        ["a","b","c","a"].find_all_indices("a").should == [0,3]
        ["a","b","c","a"].find_all_indices("d").should be_empty
        ["a","ö","ß","a"].find_all_indices("a").should == [0,3]
        ["a","b","c","d"].find_all_indices("ä").should be_empty
        ["ä","ö","ß","ä"].find_all_indices("ä").should == [0,3]
        ["stück","öl","straße","stück"].find_all_indices("stück").should == [0,3]
      end

    end

    describe "relace_umlauts, underscore" do

      it "should replace umlauts" do
        "BeforeÄÖÜäöüß²³After".replace_umlauts.should == "BeforeAeOeUeaeoeuess23After"
      end

      it "should underscore" do
        "BeforeAfter".underscore.should == "before_after"
      end

    end

    describe "to_reo" do

      before do
        @book1 = Workbook.open(@simple_file)
        @book2 = Workbook.open(@listobject_file)        
      end

      it "should type-lift an ListRow" do
        worksheet = @book2.sheet(3)
        ole_table = worksheet.ListObjects.Item(1)
        table = Table.new(ole_table)
        listrow = table[1]
        listrow.values.should == [3.0, "John", 50.0, 0.5, 30.0]
        type_lifted_listrow = listrow.ole_tablerow.to_reo
        type_lifted_listrow.should be_a ListRow
        type_lifted_listrow.values.should == [3.0, "John", 50.0, 0.5, 30.0]
      end

      it "should type-lift an ListObject" do
        worksheet = @book2.sheet(3)
        ole_table = worksheet.ListObjects.Item(1)
        table = Table.new(ole_table)
        table.Name.should == "table3"
        table.HeaderRowRange.Value.first.should == ["Number","Person","Amount","Time","Price"]
        table.ListRows.Count.should == 13
        worksheet[3,4].should == "Number"
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
        @book2.sheet(3).table(1).to_reo.should == @book2.sheet(3).table(1)
        @book2.sheet(3).table(1).should == @book2.sheet(3).table(1)
        @book2.sheet(3).table(1)[1].to_reo.should == @book2.sheet(3).table(1)[1]
      end

      it "should raise error" do
        expect{
          WIN32OLE.new('WScript.Network').to_reo
        }.to raise_error(TypeREOError)
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
        ((@ole_sheet_methods + @sheet_methods) - @book1.sheet(1).methods).should be_empty
      end

      it "should do own_methods with popular ole_excel and excel methods" do
        ((@ole_sheet_methods + @sheet_methods) - @book1.sheet(1).own_methods).should == [] #be_empty
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

        it "should raise an error for no strings" do
          expect{
            canonize(1)
          }.to raise_error(TypeREOError, "No string given to canonize, but 1")
        end

        it "should yield the network path" do
          General.canonize(@hostname_share_path).should == @network_path
          General.canonize(@network_path).should == @network_path
          General.canonize(@simple_file).should == @simple_file
          General.canonize(@simple_file_extern).should == @simple_file_extern
          General.canonize(@hostname_share_path_downcase).should == @network_path
          General.canonize(@network_path_downcase).should == @network_path_downcase
        end

      end
    end

    describe "path" do

      it "should create a path" do
        path1 = "this" / "is" / "a" / "path"
        path1.should == "this/is/a/path"
        path2 = "this" / nil
        path2.should == "this"
        path3 = "N:/E2" / "C:/gim/E2/workbook.xls"
        path3.should == "C:/gim/E2/workbook.xls"
        path4 = "N:/E2/" / "/gim/E2/workbook.xls"
        path4.should == "/gim/E2/workbook.xls"
        path5 = "N:/E2/" / "gim/E2/workbook.xls"
        path5.should == "N:/E2/gim/E2/workbook.xls"
        path6 = "N:/E2" / "spec/data/workbook.xls"
        path6.should == "N:/E2/spec/data/workbook.xls"
        path7 = "N:/E2" / "c:/gim/E2/workbook.xls"
        path7.should == "c:/gim/E2/workbook.xls"
      end
    end

  end
end
