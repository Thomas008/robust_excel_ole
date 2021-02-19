# -*- coding: utf-8 -*-

require_relative 'spec_helper'

$VERBOSE = nil

include RobustExcelOle
include General

describe Worksheet do
 
  before(:all) do
    excel = Excel.new(:reuse => true)
    open_books = excel == nil ? 0 : excel.Workbooks.Count
    puts "*** open books *** : #{open_books}" if open_books > 0
    Excel.kill_all
  end 

  before do
    @dir = create_tmpdir
    @simple_file = @dir + '/workbook.xls'
    @another_workbook = @dir + '/another_workbook.xls'
    @protected_file = @dir + '/protected_sheet.xls'
    @blank_file = @dir + '/book_with_blank.xls'
    @merge_file = @dir + '/merge_cells.xls'
    @listobject_file = @dir + '/workbook_listobjects.xlsx'    
    @book = Workbook.open(@simple_file)
    @sheet = @book.sheet(1)
  end

  after do
    @book.close(:if_unsaved => :forget)
    Excel.kill_all
    rm_tmp(@dir)
  end

  describe "Sheet" do

    describe 'access first and last sheet' do

      it "should access the first sheet" do
        first_sheet = @book.first_sheet
        first_sheet.name.should == Sheet.new(@book.Sheets.Item(1)).Name
        first_sheet.name.should == @book.sheet(1).Name
      end

      it "should access the last sheet" do
        last_sheet = @book.last_sheet
        last_sheet.name.should == Sheet.new(@book.Sheets.Item(3)).Name
        last_sheet.name.should == @book.sheet(3).Name
      end
    end
  end

  describe ".initialize" do
    context "when open sheet protected(with password is 'protect')" do
      before do
        @key_sender = IO.popen  'ruby "' + File.join(File.dirname(__FILE__), '/helpers/key_sender.rb') + '" "Microsoft Excel" '  , "w"
        @book_protect = Workbook.open(@protected_file, :visible => true, :read_only => true, :force_excel => :new)
        @book_protect.excel.displayalerts = false
        @key_sender.puts "{p}{r}{o}{t}{e}{c}{t}{enter}"
        @protected_sheet = @book_protect.sheet('protect')
      end

      after do
        @book_protect.close
        @key_sender.close
      end

      it "should be a protected sheet" do
        @protected_sheet.ProtectContents.should be true
      end

      it "protected sheet can't be write" do
        expect { @protected_sheet[1,1] = 'write' }.to raise_error
      end
    end

  end

  shared_context "sheet 'open book with blank'" do
    before do
      @book_with_blank = Workbook.open(@blank_file, :read_only => true)
      @sheet_with_blank = @book_with_blank.sheet(1)
    end

    after do
      @book_with_blank.close
    end
  end

  describe "workbook" do
    before do
      @book = Workbook.open(@simple_file)
      @sheet = @book.sheet(1)
    end

    after do
      @book.close
    end

    it "should return workbook" do
      @sheet.workbook.should === @book
    end

  end

  describe "==" do

    it "should return true for identical worksheets" do
      sheet2 = @book.sheet(1)
      @sheet.should == sheet2
    end

     it "should return false for non-identical worksheets" do
      sheet2 = @book.sheet(2)
      @sheet.should_not == sheet2
    end

  end
    
  describe "get and access sheet name" do
    
    it 'get sheet1 name' do
      @sheet.name.should eq 'Sheet1'
    end
    
    it 'change sheet1 name to foo' do
      @sheet.name = 'foo'
      @sheet.name.should eq 'foo'
    end

    it "should raise error when adding the same name" do
      @sheet.name = 'foo'
      @sheet.name.should eq 'foo'
      new_sheet = @book.add_sheet @sheet
      expect{
        new_sheet.name = 'foo'
      }.to raise_error(NameAlreadyExists, /sheet name "foo" already exists/)
    end

    it "should get and set name with umlaut" do
      @sheet.name = "Straße"
      @sheet.name.should == "Straße"
    end

    it "should set and get numbers as name" do
      @sheet.name = 1
      @sheet.name.should == "1"
    end

  end

  describe 'access cell' do

    describe "#[,]" do      

      context "access [1,1]" do

        it { @sheet[1, 1].should be_kind_of Cell }
        it { @sheet[1, 1].Value.should eq 'foo' }
      end

      context "access [1, 1], [1, 2], [3, 1]" do
        it "should get every values" do
          @sheet[1, 1].Value.should eq 'foo'
          @sheet[1, 2].Value.should eq 'workbook'
          @sheet[3, 1].Value.should eq 'matz'
        end
      end

     # context "supplying nil as parameter" do
     #   it "should access [1,1]" do
     #     @sheet[1, nil].Value.should eq 'foo'
     #     @sheet[nil, 1].Value.should eq 'foo'
     #   end
     # end

    end

    describe "cellval" do

      it "should return value" do
        @sheet.cellval(1,1).should == "foo"
      end
    end

    it "change a cell to 'bar'" do
      @sheet[1, 1] = 'bar'
      @sheet[1, 1].Value.should eq 'bar'
    end

    it "should change a cell to nil" do
      @sheet[1, 1] = nil
      @sheet[1, 1].Value.should eq nil
    end

    it "should raise error for bad ranges" do
      expect{
        @sheet[0,0]
      }.to raise_error(RangeNotEvaluatable, /cannot read cell/)
      expect{
        @sheet[0,0] = "foo"
      }.to raise_error(RangeNotEvaluatable, /cannot assign value/)
    end

    describe "set_cellval" do

      it "should set color" do
        @sheet.set_cellval(1,1,"foo",:color => 42)
        @sheet.cellval(1,1).should == "foo"
        @sheet[1,1].Interior.ColorIndex.should == 42
      end
    end

    describe "range" do

      it "should a range with relative r1c1-reference" do
        @sheet.range([1,1]).Select
        @sheet.range(["Z1S[3]:Z[2]S8"]).Address.should == "$D$1:$H$3"
        @sheet.range(["Z1S3:Z2S8"]).Address.should == "$C$1:$H$2"
      end

      it "should a range with relative integer-range-reference" do
        @sheet.range([1,1]).Select
        @sheet.range([1..[2],[3]..8]).Address.should == "$D$1:$H$3"
      end

      it "should create a range of one cell" do
        @sheet.range([1,2]).values.should == ["workbook"]
        @sheet.range(["B1"]).values.should == ["workbook"]
        @sheet.range("B1").values.should == ["workbook"]
        @sheet.range(["Z1S2"]).values.should == ["workbook"]
        @sheet.range("Z1S2").values.should == ["workbook"]
      end

      it "should create a rectangular range" do
        @sheet.range([1..3,2..4]).values.should == ["workbook", "sheet1", nil, nil, "foobaaa", nil, "is", "nice", nil]
        @sheet.range([1..3, "B".."D"]).values.should == ["workbook", "sheet1", nil, nil, "foobaaa", nil, "is", "nice", nil]     
        @sheet.range(["B1:D3"]).values.should == ["workbook", "sheet1", nil, nil, "foobaaa", nil, "is", "nice", nil]
        @sheet.range("B1:D3").values.should == ["workbook", "sheet1", nil, nil, "foobaaa", nil, "is", "nice", nil]
        @sheet.range(["Z1S2:Z3S4"]).values.should == ["workbook", "sheet1", nil, nil, "foobaaa", nil, "is", "nice", nil]
        @sheet.range("Z1S2:Z3S4").values.should == ["workbook", "sheet1", nil, nil, "foobaaa", nil, "is", "nice", nil]
      end

      it "should accept old interface" do
        @sheet.range(1..3,2..4).values.should == ["workbook", "sheet1", nil, nil, "foobaaa", nil, "is", "nice", nil]
        @sheet.range(1..3, "B".."D").values.should == ["workbook", "sheet1", nil, nil, "foobaaa", nil, "is", "nice", nil]     
      end

      it "should create infinite ranges" do
        @sheet.range([1..3,nil]).Address.should == "$1:$3"
        @sheet.range(nil,"B".."D").Address.should == "$B:$D"
        @sheet.range("1:3").Address.should == "$1:$3"
        @sheet.range("B:D").Address.should == "$B:$D"
      end

      it "should raise an error" do
        expect{
          @sheet.range([0,0])
          }.to raise_error(RangeNotCreated, /cannot create/)
      end

    end

    describe "table" do

      before do
        @book = Workbook.open(@listobject_file, :visible => true)
        @sheet = @book.sheet(3)
      end

      it "should yield table given number" do
        table = @sheet.table(1)
        table.Name.should == "table3"
        table.HeaderRowRange.Value.first.should == ["Number","Person","Amount","Time","Price"]
        table.ListRows.Count.should == 13
        @sheet[3,4].Value.should == "Number"
      end

      it "should yield table given name" do
        table = @sheet.table("table3")
        table.Name.should == "table3"
        table.HeaderRowRange.Value.first.should == ["Number","Person","Amount","Time","Price"]
        table.ListRows.Count.should == 13
        @sheet[3,4].Value.should == "Number"
      end

      it "should yield table given name with umlauts" do
        table = Table.new(@sheet, "lösung", [1,1], 3, ["Verkäufer","Straße"])
        table2 = @sheet.table("lösung")
        table2.Name.encode_value.should == "lösung"
        table2.HeaderRowRange.Value.first.encode_value.should == ["Verkäufer","Straße"]
      end

      it "should raise error" do
        expect{
          @sheet.table("table4")
          }.to raise_error(WorksheetREOError)
      end

    end

    describe '#each' do
      it "should sort line in order of column" do
        @sheet.each_with_index do |cell, i|
          case i
          when 0
            cell.Value.should eq 'foo'
          when 1
            cell.Value.should eq 'workbook'
          when 2
            cell.Value.should eq 'sheet1'
          when 3
            cell.Value.should eq 'foo'
          when 4
            cell.Value.should be_nil
          when 5
            cell.Value.should eq 'foobaaa'
          end
        end
      end

      context "read sheet with blank" do
        include_context "sheet 'open book with blank'"

        it 'should get from ["A1"]' do
          @sheet_with_blank.each_with_index do |cell, i|
            case i
            when 5
              cell.Value.should be_nil
            when 6
              cell.Value.should eq 'simple'
            when 7
              cell.Value.should be_nil
            when 8
              cell.Value.should eq 'workbook'
            when 9
              cell.Value.should eq 'sheet1'
            end
          end
        end
      end

    end

    describe "#values" do

      it "should yield cell values of the used range" do
        @sheet.values.should == [["foo", "workbook", "sheet1"], ["foo", nil, "foobaaa"], ["matz", "is", "nice"]]
      end

    end

    describe "#each_rowvalue" do

      it "should yield arrays" do
        @sheet.each_rowvalue do |row_value|
          row_value.should be_kind_of Array
        end
      end

      it "should read the rows" do
        i = 0
        @sheet.each_rowvalue do |row_values|
          case i
          when 0
            row_values.should == ['foo', 'workbook', 'sheet1']
          when 1
            row_values.should == ['foo', nil, 'foobaaa']
          end
          i += 1
        end
      end

      it "should read the rows with index" do
        @sheet.each_rowvalue_with_index do |row_values, i|
          case i
          when 0
            row_values.should == ['foo', 'workbook', 'sheet1']
          when 1
            row_values.should == ['foo', nil, 'foobaaa']
          end
        end
      end

    end

    describe "#each_row" do
      it "items should RobustExcelOle::Range" do
        @sheet.each_row do |rows|
          rows.should be_kind_of RobustExcelOle::Range
        end
      end

      context "with argument 1" do
        it 'should read from second row' do
          @sheet.each_row(1) do |rows|
            case rows.Row
            when 2
              rows.values.should eq ['foo', nil, 'foobaaa']
            when 3
              rows.values.should eq ['matz', 'is', 'nice']
            end
          end
        end
      end

      context "read sheet with blank" do
        include_context "sheet 'open book with blank'"

        it 'should get from ["A1"]' do
          @sheet_with_blank.each_row do |rows|
            case rows.Row - 1
            when 0
              rows.values.should eq [nil, nil, nil, nil, nil]
            when 1
              rows.values.should eq [nil, 'simple', nil, 'workbook', 'sheet1']
            when 2
              rows.values.should eq [nil, 'foo', nil, nil, 'foobaaa']
            when 3
              rows.values.should eq [nil, nil, nil, nil, nil]
            when 4
              rows.values.should eq [nil, 'matz', nil, 'is', 'nice']
            end
          end
        end
      end

    end

    describe "#each_row_with_index" do
      it "should read with index" do
        @sheet.each_row_with_index do |rows, idx|
          case idx
          when 0
            rows.values.should eq ['foo', 'workbook', 'sheet1']
          when 1
            rows.values.should eq ['foo', nil, 'foobaaa']
          when 2
            rows.values.should eq ['matz', 'is', 'nice']
          end
        end
      end

      context "with argument 1" do
        it "should read from second row, index is started 0" do
          @sheet.each_row_with_index(1) do |rows, idx|
            case idx
            when 0
              rows.values.should eq ['foo', nil, 'foobaaa']
            when 1
              rows.values.should eq ['matz', 'is', 'nice']
            end
          end
        end
      end

    end

    describe "#each_column" do
      it "items should RobustExcelOle::Range" do
        @sheet.each_column do |columns|
          columns.should be_kind_of RobustExcelOle::Range
        end
      end

      context "with argument 1" do
        it "should read from second column" do
          @sheet.each_column(1) do |columns|
            case columns.Column
            when 2
              columns.values.should eq ['workbook', nil, 'is']
            when 3
              columns.values.should eq ['sheet1', 'foobaaa', 'nice']
            end
          end
        end
      end

      context "read sheet with blank" do
        include_context "sheet 'open book with blank'"

        it 'should get from ["A1"]' do
          @sheet_with_blank.each_column do |columns|
            case columns.Column- 1
            when 0
              columns.values.should eq [nil, nil, nil, nil, nil]
            when 1
              columns.values.should eq [nil, 'simple', 'foo', nil, 'matz']
            when 2
              columns.values.should eq [nil, nil, nil, nil, nil]
            when 3
              columns.values.should eq [nil, 'workbook', nil, nil, 'is']
            when 4
              columns.values.should eq [nil, 'sheet1', 'foobaaa', nil, 'nice']
            end
          end
        end
      end

      context "read sheet which last cell is merged" do
        before do
          @book_merge_cells = Workbook.open(@merge_file)
          @sheet_merge_cell = @book_merge_cells.sheet(1)
        end

        after do
          @book_merge_cells.close
        end

        it "should get from ['A1'] to ['C2']" do
          columns_values = []
          @sheet_merge_cell.each_column do |columns|
            columns_values << columns.values
          end
          columns_values.should eq [
                                [nil, 'first merged', nil, 'merged'],
                                [nil, 'first merged', 'first', 'merged'],
                                [nil, 'first merged', 'second', 'merged'],
                                [nil, nil, 'third', 'merged']
                           ]
        end
      end
    end

    describe "#each_column_with_index" do
      it "should read with index" do
        @sheet.each_column_with_index do |columns, idx|
          case idx
          when 0
            columns.values.should eq ['foo', 'foo', 'matz']
          when 1
            columns.values.should eq ['workbook', nil, 'is']
          when 2
            columns.values.should eq ['sheet1', 'foobaaa', 'nice']
          end
        end
      end

      context "with argument 1" do
        it "should read from second column, index is started 0" do
          @sheet.each_column_with_index(1) do |column_range, idx|
            case idx
            when 0
              column_range.values.should eq ['workbook', nil, 'is']
            when 1
              column_range.values.should eq ['sheet1', 'foobaaa', 'nice']
            end
          end
        end
      end
    end

    describe "#row_range" do
      context "with second argument" do
        before do
          @row_range = @sheet.row_range(1, 2..3)
        end

        it { @row_range.should be_kind_of RobustExcelOle::Range }

        it "should get range cells of second argument" do
          @row_range.values.should eq ['workbook', 'sheet1']
        end
      end

      context "without second argument" do
        before do
          @row_range = @sheet.row_range(3)
        end

        it "should get all cells" do
          @row_range.values.should eq ['matz', 'is', 'nice']
        end
      end

    end

    describe "#col_range" do
      context "with second argument" do
        before do
          @col_range = @sheet.col_range(1, 2..3)
        end

        it { @col_range.should be_kind_of RobustExcelOle::Range }

        it "should get range cells of second argument" do
          @col_range.values.should eq ['foo', 'matz']
        end
      end

      context "without second argument" do
        before do
          @col_range = @sheet.col_range(2)
        end

        it "should get all cells" do
          @col_range.values.should eq ['workbook', nil, 'is']
        end
      end
    end

    describe "[], []=" do
      before do
        @book1 = Workbook.open(@dir + '/another_workbook.xls')
        @sheet1 = @book1.sheet(1)
      end

      after do
        @book1.close(:if_unsaved => :forget)
      end   

      it "should return value of a defined name" do
        @sheet1["firstcell"].should == "foo"
      end

      it "should return value of a defined name" do
        @sheet1["new"].should == "foo"         
        @sheet1["one"].should == 1.0    
        @sheet1["four"].should == [[1,2],[3,4]]
        @sheet1["firstrow"].should == [[1,2]]
      end  

      it "should return value of a name with coordinates" do
        @sheet1["A1"].should == "foo"         
      end  

      it "should return nil for a range with empty contents" do
        @sheet1["another"].should == nil
      end 

      #it "should evaluate named formula" do
      #  @sheet1["named_formula"].should == 4
      #end                

      #it "should evaluate a formula" do
      #  @sheet1["another_formula"].should == 5
      #end      

      it "should raise an error if name not defined" do
        expect {
          @sheet1["foo"]
        }.to raise_error(NameNotFound, /name "foo" not in #<Worksheet: Sheet1/)        
      end

      it "should set a range to a value" do
        @sheet1[1,1].Value.should == "foo"
        @sheet1["firstcell"] = "bar"
        @sheet1[1,1].Value.should == "bar"
        @sheet1["new"] = "bar"
        @sheet1["new"].should == "bar"
      end

      it "should raise an error if name cannot be evaluated" do
        expect{
          @sheet1["foo"] = 1
          }.to raise_error(NameNotFound, /name "foo" not in #<Worksheet: Sheet1/)
      end
    end

    describe "namevalue_global, set_namevalue_global" do

      before do
        @book1 = Workbook.open(@dir + '/another_workbook.xls')
        @sheet1 = @book1.sheet(1)
      end

      after do
        @book1.close(:if_unsaved => :forget)
      end   

      it "should return value of a defined name" do
        @sheet1.namevalue_global("firstcell").should == "foo"
      end

      #it "should evaluate a formula" do
      #  @sheet1.namevalue_global("another_formula").should == 5
      #end      

      it "should raise an error if name not defined" do
        expect {
          @sheet1.namevalue_global("foo")
        }.to raise_error(NameNotFound, /name "foo" not in/)
      end

      it "should raise an error of coordinates are given instead of a defined name" do
        expect {
          @sheet1.namevalue_global("A1")
        }.to raise_error(NameNotFound, /name "A1" not in/)
      end

      it "should return default value for a range with empty contents" do
        @sheet1.namevalue_global("another", :default => 2) == 2
      end 

      it "should set a range to a value" do
        @sheet1.namevalue_global("firstcell").should == "foo"
        @sheet1[1,1].Value.should == "foo"
        @sheet1.set_namevalue_global("firstcell","bar")
        @sheet1.namevalue_global("firstcell").should == "bar"
        @sheet1[1,1].Value.should == "bar"
      end

      it "should raise an error if name cannot be evaluated" do
        expect{
          @sheet1.set_namevalue_global("foo", 1)
        }.to raise_error(RangeNotEvaluatable, /cannot assign value/)
      end

      it "should color the cell (deprecated)" do
        @sheet1.set_namevalue_global("new", "bar")
        @book1.Names.Item("new").RefersToRange.Interior.ColorIndex.should == -4142
        @sheet1.set_namevalue_global("new", "bar", :color => 4)
        @book1.Names.Item("new").RefersToRange.Interior.ColorIndex.should == 4
      end

      it "should color the cell" do
        @sheet1.set_namevalue_global("new", "bar")
        @book1.Names.Item("new").RefersToRange.Interior.ColorIndex.should == -4142
        @sheet1.set_namevalue_global("new", "bar", :color => 4)
        @book1.Names.Item("new").RefersToRange.Interior.ColorIndex.should == 4
      end

      it "should set a range to a value with umlauts" do
        @sheet1.add_name("lösung", [1,1])
        @sheet1.namevalue_global("lösung").should == "foo"
        @sheet1[1,1].Value.should == "foo"
        @sheet1.set_namevalue_global("lösung","bar")
        @sheet1.namevalue_global("lösung").should == "bar"
        @sheet1[1,1].Value.should == "bar"  
      end

    end

    describe "namevalue, set_namevalue" do
      
      before do
        @book1 = Workbook.open(@dir + '/another_workbook.xls')
        @sheet1 = @book1.sheet(1)
        @sheet2 = @book1.sheet(2)
      end

      after do
        @book1.close(:if_unsaved => :forget)
      end   

      it "should return value of a locally defined name" do
        @sheet1.namevalue("firstcell").should == "foo"          
      end        

      it "should return value of a name with coordinates" do
        @sheet1.namevalue("A1").should == "foo"         
      end  

      it "should return nil for a range with empty contents" do
        @sheet1.namevalue("another").should == nil
      end 

      it "should return value of a defined name" do
        @sheet1.namevalue("new").should == "foo"         
        @sheet1.namevalue("one").should == 1.0    
        @sheet1.namevalue("four").should == [[1,2],[3,4]]
        @sheet1.namevalue("firstrow").should == [[1,2]]
      end    

      it "should return default value if name not defined and default value is given" do
        @sheet1.namevalue("foo", :default => 2).should == 2
      end

      it "should raise an error if name not defined for the sheet" do
        expect {
          @sheet1.namevalue("foo")
          }.to raise_error(NameNotFound, /name "foo" not in #<Worksheet: Sheet1/)
        expect {
          @sheet1.namevalue("named_formula")
          }.to raise_error(NameNotFound, /name "named_formula" not in #<Worksheet: Sheet1/)
        expect {
          @sheet2.namevalue("firstcell")
          }.to raise_error(NameNotFound, /name "firstcell" not in #<Worksheet: Sheet2/)
      end
    
      it "should set a range to a value" do
        @sheet1.namevalue("firstcell").should == "foo"
        @sheet1[1,1].Value.should == "foo"
        @sheet1.set_namevalue("firstcell","bar")
        @sheet1.namevalue("firstcell").should == "bar"
        @sheet1[1,1].Value.should == "bar"          
      end

      it "should set a range to a value with umlauts" do
        @sheet1.add_name("lösung", [1,1])
        @sheet1.namevalue("lösung").should == "foo"
        @sheet1[1,1].Value.should == "foo"
        @sheet1.set_namevalue("lösung","bar")
        @sheet1.namevalue("lösung").should == "bar"
        @sheet1[1,1].Value.should == "bar"  
      end

      it "should raise an error if name cannot be evaluated" do
        expect{
          @sheet1.set_namevalue_global("foo", 1)
        }.to raise_error(RangeNotEvaluatable, /cannot assign value/)
      end

      it "should raise an error if name not defined and default value is not provided" do
        expect {
          @sheet1.namevalue("foo", :default => nil)
        }.to_not raise_error
        expect {
          @sheet1.namevalue("foo", :default => :__not_provided)
        }.to raise_error(NameNotFound, /name "foo" not in #<Worksheet: Sheet1/)
        expect {
          @sheet1.namevalue("foo")
        }.to raise_error(NameNotFound, /name "foo" not in #<Worksheet: Sheet1/)
        @sheet1.namevalue("foo", :default => nil).should be_nil
        @sheet1.namevalue("foo", :default => 1).should == 1
        @sheet1.namevalue_global("empty", :default => 1).should be_nil
      end

      it "should color the cell (depracated)" do
        @sheet1.set_namevalue("new", "bar")
        @book1.Names.Item("new").RefersToRange.Interior.ColorIndex.should == -4142
        @sheet1.set_namevalue("new", "bar", :color => 4)
        @book1.Names.Item("new").RefersToRange.Interior.ColorIndex.should == 4
      end

      it "should color the cell" do
        @sheet1.set_namevalue("new", "bar")
        @book1.Names.Item("new").RefersToRange.Interior.ColorIndex.should == -4142
        @sheet1.set_namevalue("new", "bar", :color => 4)
        @book1.Names.Item("new").RefersToRange.Interior.ColorIndex.should == 4
      end

    end

    describe "add_name, delete_name, rename_range" do

      context "adding, renaming, deleting the name of a range" do

        before do
          @book1 = Workbook.open(@dir + '/another_workbook.xls', :visible => true)
          @sheet1 = @book1.sheet(1)
        end

        after do
          @book1.close(:if_unsaved => :forget)
        end

        it "should add a name of a rectangular range using relative int-range-reference" do
          @sheet1.add_name("foo",[[1]..3,1..[2]])
          @sheet1.range("foo").Address.should == "$A$3:$D$5"
        end   

        it "should add a name of a rectangular range using relative r1c1-reference" do
          @sheet1.add_name("foo","Z[1]S3:Z1S[2]")
          @sheet1.range("foo").Address.should == "$C$1:$D$5"
          @sheet1.add_name("bar","Z[-3]S[-2]")
          @sheet1.range("bar").Address.should == "$IV$1"
        end  

        it "should name an unnamed range with a giving address" do
          @sheet1.add_name("foo",[1,2])
          @sheet1.Range("foo").Address.should == "$B$1"
        end

        it "should rename an already named range with a giving address" do
          @sheet1[1,1].Name.Name.should == "Sheet1!firstcell"
          @sheet1.add_name("foo",[1..2,3..4])
          @sheet1.Range("foo").Address.should == "$C$1:$D$2"
        end

        it "should raise an error" do
          expect{
            @sheet1.add_name("foo", [-2, 1])
          }.to raise_error(RangeNotEvaluatable, /cannot add name "foo" to range/)
        end

        it "should rename a range" do
          @sheet1.add_name("foo",[1,1])
          @sheet1.rename_range("foo","bar")
          @sheet1.namevalue_global("bar").should == "foo"
        end

        it "should rename a range with umlauts" do
          @sheet1.add_name("lösung",[1,1])
          @sheet1.rename_range("lösung","bär")
          @sheet1.namevalue_global("bär").should == "foo"
        end


        it "should delete a name of a range" do
          @sheet1.add_name("foo",[1,1])
          @sheet1.delete_name("foo")
          expect{
            @sheet1.namevalue_global("foo")
         }.to raise_error(NameNotFound, /name "foo"/)
        end

        it "should delete a name of a range with umlauts" do
          @sheet1.add_name("lösung",[1,1])
          @sheet1.delete_name("lösung")
          expect{
            @sheet1.namevalue_global("lösung")
         }.to raise_error(NameNotFound, /name/)
        end

        it "should add a name of a rectangular range" do
          @sheet1.add_name("foo",[1..3,1..4])
          @sheet1["foo"].should == [["foo", "workbook", "sheet1", nil], ["foo", 1.0, 2.0, 4.0], ["matz", 3.0, 4.0, 4.0]] 
        end

        it "should add a name of a rectangular range" do
          @sheet1.add_name("foo","A1:D3")
          @sheet1["foo"].should == [["foo", "workbook", "sheet1", nil], ["foo", 1.0, 2.0, 4.0], ["matz", 3.0, 4.0, 4.0]] 
        end

        it "should add a name of a rectangular range" do
          @sheet1.add_name("foo",["A1:D3"])
          @sheet1["foo"].should == [["foo", "workbook", "sheet1", nil], ["foo", 1.0, 2.0, 4.0], ["matz", 3.0, 4.0, 4.0]] 
        end

        it "should use the old interface" do
          @sheet1.add_name("foo",1..3,"A".."D")
          @sheet1["foo"].should == [["foo", "workbook", "sheet1", nil], ["foo", 1.0, 2.0, 4.0], ["matz", 3.0, 4.0, 4.0]] 
        end

        it "should add a name of a rectangular range" do
          @sheet1.add_name("foo",[1..3, "A".."D"])
          @sheet1["foo"].should == [["foo", "workbook", "sheet1", nil], ["foo", 1.0, 2.0, 4.0], ["matz", 3.0, 4.0, 4.0]] 
        end

        it "should add a name of another rectangular range" do
          @sheet1.add_name("foo",[1..3, "A"])
          @sheet1["foo"].should == [["foo"], ["foo"],["matz"]]
          @sheet1.Range("foo").Address.should == "$A$1:$A$3"
        end

        it "should add a name of an infinite row range" do
          @sheet1.add_name("foo",[1..3, nil])
          @sheet1.Range("foo").Address.should == "$1:$3"
        end

        it "should add a name of an infinite column range" do
          @sheet1.add_name("foo",[nil, "A".."C"])
          @sheet1.Range("foo").Address.should == "$A:$C"
        end

        it "should add a name of an infinite column range" do
          @sheet1.add_name("foo",[nil, 1..3])
          @sheet1.Range("foo").Address.should == "$A:$C"
        end

        it "should add a name of an infinite column range" do
          @sheet1.add_name("foo","A:C")
          @sheet1.Range("foo").Address.should == "$A:$C"
        end

        it "should add a name of an infinite column range" do
          @sheet1.add_name("foo","1:2")
          @sheet1.Range("foo").Address.should == "$1:$2"
        end

        it "should name an range with a relative columns" do
          @sheet1.add_name("foo",[1,2])
          @sheet1.Range("foo").Address.should == "$B$1"
          @sheet1.add_name("bar","Z3S[4]")
          @sheet1.Range("bar").Address.should == "$E$3"
        end

        it "should name an range with a relative row" do
          @sheet1.add_name("foo",[1,2])
          @sheet1.Range("foo").Address.should == "$B$1"
          @sheet1.add_name("bar","Z[3]S4")
          @sheet1.Range("bar").Address.should == "$D$4"
        end

      end
    end

    describe "send methods to worksheet" do

      it "should send methods to worksheet" do
        @sheet.Cells(1,1).Value.should eq 'foo'
      end

      it "should raise an error for unknown methods or properties" do
        expect{
          @sheet.Foo
        }.to raise_error
      end

    end

    describe "#method_missing" do
      it "can access COM method" do
        @sheet.Cells(1,1).Value.should eq 'foo'
      end

      context "unknown method" do
        it { expect { @sheet.hogehogefoo }.to raise_error }
      end
    end

  end
end
