# -*- coding: utf-8 -*-

require File.join(File.dirname(__FILE__), './spec_helper')

$VERBOSE = nil

include RobustExcelOle
include General

describe ListObject do
 
  before(:all) do
    excel = Excel.new(:reuse => true)
    open_books = excel == nil ? 0 : excel.Workbooks.Count
    puts "*** open books *** : #{open_books}" if open_books > 0
    Excel.kill_all
  end 

  before do
    @dir = create_tmpdir
    @listobject_file = @dir + '/workbook_listobjects.xlsx'
    @book = Workbook.open(@listobject_file, :visible => true)
    @sheet = @book.sheet(3)
  end

  after do
    @book.close(:if_unsaved => :forget)
    Excel.kill_all
    rm_tmp(@dir)
  end

  describe "creating" do

    context "with standard" do

      it "should simply create a new table" do
        table = Table.new(@sheet, "table_name", [1,1], 3, ["Person","Amount"])
        table.Name.should == "table_name"
        table.HeaderRowRange.Value.first.should == ["Person","Amount"]
        table.ListRows.Count.should == 3
        @sheet[1,1].Value.should == "Person"
      end

      it "should type-lift a Win32ole list object into a RobustExcelOle list object" do
        ole_table = @sheet.ListObjects.Item(1)
        table = Table.new(ole_table)
        table.Name.should == "table3"
        table.HeaderRowRange.Value.first.should == ["Number","Person","Amount","Time","Price"]
        table.ListRows.Count.should == 6
        @sheet[3,4].Value.should == "Number"
      end

      it "should type-lift a Win32ole list object into a RobustExcelOle list object with table name" do
        ole_table = @sheet.ListObjects.Item(1)
        table = Table.new(@sheet, "table3")
        table.Name.should == "table3"
        table.HeaderRowRange.Value.first.should == ["Number","Person","Amount","Time","Price"]
        table.ListRows.Count.should == 6
        @sheet[3,4].Value.should == "Number"
      end

      it "should type-lift a Win32ole list object into a RobustExcelOle list object with item number" do
        ole_table = @sheet.ListObjects.Item(1)
        table = Table.new(@sheet, 1)
        table.Name.should == "table3"
        table.HeaderRowRange.Value.first.should == ["Number","Person","Amount","Time","Price"]
        table.ListRows.Count.should == 6
        @sheet[3,4].Value.should == "Number"
      end

      it "should simply create a new table from a ole-worksheet" do
        table = Table.new(@sheet.ole_worksheet, "table_name", [1,1], 3, ["Person","Amount"])
        table.Name.should == "table_name"
        table.HeaderRowRange.Value.first.should == ["Person","Amount"]
        table.ListRows.Count.should == 3
        @sheet[1,1].Value.should == "Person"
      end

      it "should type-lift a Win32ole list object into a RobustExcelOle list object with table name" do
        ole_table = @sheet.ListObjects.Item(1)
        table = Table.new(@sheet.ole_worksheet, "table3")
        table.Name.should == "table3"
        table.HeaderRowRange.Value.first.should == ["Number","Person","Amount","Time","Price"]
        table.ListRows.Count.should == 6
        @sheet[3,4].Value.should == "Number"
      end

      it "should type-lift a Win32ole list object into a RobustExcelOle list object with item number" do
        ole_table = @sheet.ListObjects.Item(1)
        table = Table.new(@sheet.ole_worksheet, 1)
        table.Name.should == "table3"
        table.HeaderRowRange.Value.first.should == ["Number","Person","Amount","Time","Price"]
        table.ListRows.Count.should == 6
        @sheet[3,4].Value.should == "Number"
      end

    end

  end

  describe "getting and setting values" do

    context "with new table" do

      before do
        @table = Table.new(@sheet, "table_name", [1,1], 3, ["Person","Amount"])
        @table_row1 = @table[1]
      end

      it "should set and read values" do
        @table_row1.person.should be nil
        @table_row1.person = "John"
        @table_row1.person.should == "John"
        @sheet[2,1].Value.should == "John"
        @table_row1.amount.should be nil
        @table_row1.amount = 42
        @table_row1.amount.should == 42
        @sheet[2,2].Value.should == 42
      end

    end

    context "with type-lifted ole list object" do

      before do
        ole_table = @sheet.ListObjects.Item(1)
        @table = Table.new(ole_table)
        @table_row1 = @table[1]
      end

      it "should set and read values" do
        @table_row1.number.should == 3
        @table_row1.number = 1
        @table_row1.number.should == 1
        @sheet[4,4].Value.should == 1
        @table_row1.person.should == "John"
        @table_row1.person = "Herbert"
        @table_row1.person.should == "Herbert"
        @sheet[4,5].Value.should == "Herbert"
      end
    end

  end

  describe "reading and setting contents of rows and columns" do

    before do
      ole_table = @sheet.ListObjects.Item(1)
      @table = Table.new(ole_table)
    end

    it "should read contents of a column" do
      @table.column_values("Person").should == ["John","Fred",nil,"Angel",nil,"Werner"]
      expect{
        @table.column_values("P")
      }.to raise_error(TableError)
    end

    it "should set contents of a column" do
      @table.set_column_values("Person",["H",nil,nil,nil,"G","W"])
      @table.ListColumns.Item(2).Range.Value.should == [["Person"],["H"],[nil],[nil],[nil],["G"],["W"]]
      @table.set_column_values("Person",["T","Z"])
      @table.ListColumns.Item(2).Range.Value.should == [["Person"],["T"],["Z"],[nil],[nil],["G"],["W"]]
      expect{
        @table.set_column_values("P",["H",nil,nil,nil,"G","W"])
      }.to raise_error(TableError)
    end

    it "should read contents of a row" do
      @table.row_values(1).should == [3.0, "John", 50.0, 0.5, 30]
      @table[1].values.should == [3.0, "John", 50.0, 0.5, 30]
      expect{
        @table.row_values(9)
      }.to raise_error(TableError)
    end

    it "should set contents of a row" do
      @table.set_row_values(1, [5, "George", 30.0, 0.2, 50])
      @table.ListRows.Item(1).Range.Value.first.should == [5, "George", 30.0, 0.2, 50]
      @table.set_row_values(1, [6, "Martin"])
      @table.ListRows.Item(1).Range.Value.first.should == [6, "Martin", 30.0, 0.2, 50]
      @table[1].set_values([2, "Merlin", 20.0, 0.1, 40])
      @table[1].set_values([4, "John"])
      @table.ListRows.Item(1).Range.Value.first.should == [4, "John", 20.0, 0.1, 40]
      expect{
        @table.set_row_values(9, [5, "George", 30.0, 0.2, 50])
      }.to raise_error(TableError)
    end

  end

  describe "renaming, adding and deleting columns and rows" do

    before do
      ole_table = @sheet.ListObjects.Item(1)
      @table = Table.new(ole_table)
    end

    it "should list column names" do
      @table.column_names.should == @table.HeaderRowRange.Value.first
    end

    it "should rename a column name" do
      @table.rename_column("Person", "P")
      @table.HeaderRowRange.Value.first.should == ["Number","P","Amount","Time","Price"]
    end

    it "should append a column" do
      @table.add_column("column_name")
      column_names = @table.HeaderRowRange.Value.first.should == ["Number","Person", "Amount","Time","Price", "column_name"]
    end

    it "should add a column" do
      @table.add_column("column_name", 3)
      column_names = @table.HeaderRowRange.Value.first.should == ["Number","Person","column_name","Amount","Time","Price"]
      expect{
        @table.add_column(8, "column_name")
      }.to raise_error(TableError)
    end

    it "should add a column with contents" do
      @table.add_column("column_name", 3, ["a","b","c","d","e","f","g"])
      column_names = @table.HeaderRowRange.Value.first.should == ["Number","Person","column_name","Amount","Time","Price"]
      @table.ListColumns.Item(3).Range.Value.should == [["column_name"],["a"],["b"],["c"],["d"],["e"],["f"],["g"]]
    end

    it "should delete a column" do
      @table.delete_column(4)
      @table.HeaderRowRange.Value.first.should == ["Number","Person", "Amount","Price"]
      expect{
        @table.delete_column(6)
      }.to raise_error(TableError)
    end

    it "should append a row" do
      @table.add_row
      listrows = @table.ListRows
      listrows.Item(listrows.Count).Range.Value.first.should == [nil,nil,nil,nil,nil]
    end

    it "should add a row" do
      @table.add_row(2)
      listrows = @table.ListRows
      listrows.Item(1).Range.Value.first.should == [3.0, "John", 50.0, 0.5, 30]
      listrows.Item(2).Range.Value.first.should == [nil,nil,nil,nil,nil]
      listrows.Item(3).Range.Value.first.should == [2.0, "Fred", nil, 0.5416666666666666, 40]
      expect{
        @table.add_row(9)
      }.to raise_error(TableError)
    end

    it "should add a row with contents" do
      @table.add_row(2, [2.0, "Herbert", 30.0, 0.25, 40])
      listrows = @table.ListRows
      listrows.Item(1).Range.Value.first.should == [3.0, "John", 50.0, 0.5, 30]
      listrows.Item(2).Range.Value.first.should == [2.0, "Herbert", 30.0, 0.25, 40]
      listrows.Item(3).Range.Value.first.should == [2.0, "Fred", nil, 0.5416666666666666, 40]
    end

    it "should delete a row" do
      @table.delete_row(4)
      listrows = @table.ListRows
      listrows.Item(5).Range.Value.first.should == [1,"Werner",40,0.5, 80]
      listrows.Item(4).Range.Value.first.should == [nil,nil,nil,nil,nil]
      expect{
        @table.delete_row(8)
      }.to raise_error(TableError)
    end

    it "should delete the contents of a column" do
      @table.ListColumns.Item(3).Range.Value.should == [["Amount"],[50],[nil],[nil],[100],[nil],[40]]
      @table.delete_column_values(3)
      @table.HeaderRowRange.Value.first.should == ["Number","Person", "Amount", "Time","Price"]
      @table.ListColumns.Item(3).Range.Value.should == [["Amount"],[nil],[nil],[nil],[nil],[nil],[nil]]
      @table.ListColumns.Item(1).Range.Value.should == [["Number"],[3],[2],[nil],[3],[nil],[1]]
      @table.delete_column_values("Number")
      @table.ListColumns.Item(1).Range.Value.should == [["Number"],[nil],[nil],[nil],[nil],[nil],[nil]]
      expect{
        @table.delete_column_values("N")
      }.to raise_error(TableError)
    end

    it "should delete the contents of a row" do
      @table.ListRows.Item(2).Range.Value.first.should == [2.0, "Fred", nil, 0.5416666666666666, 40]
      @table.delete_row_values(2)
      @table.ListRows.Item(2).Range.Value.first.should == [nil,nil,nil,nil,nil]
      @table.ListRows.Item(1).Range.Value.first.should == [3.0, "John", 50.0, 0.5, 30]
      @table[1].delete_values
      @table.ListRows.Item(1).Range.Value.first.should == [nil,nil,nil,nil,nil]
      expect{
        @table.delete_row_values(9)
      }.to raise_error(TableError)
    end

    it "should delete empty rows" do
      @table.delete_empty_rows
      @table.ListRows.Count.should == 4
      @table.ListRows.Item(1).Range.Value.first.should == [3.0, "John", 50.0, 0.5, 30]
      @table.ListRows.Item(2).Range.Value.first.should == [2.0, "Fred", nil, 0.5416666666666666, 40]
      @table.ListRows.Item(3).Range.Value.first.should == [3, "Angel", 100, 0.6666666666666666, 60]
      @table.ListRows.Item(4).Range.Value.first.should == [1,"Werner",40,0.5, 80]
    end

    it "should delete empty columns" do
      @table.delete_columns_values(4)
      @table.ListColumns.Count.should == 5
      @table.HeaderRowRange.Value.first.should == ["Number","Person", "Amount", "Time","Price"]
      @table.delete_empty_columns
      @table.ListColumns.Count.should == 4
      @table.HeaderRowRange.Value.first.should == ["Number","Person", "Amount","Price"]
    end
  end

  describe "find all occurrences of a value" do

    it "should find all occurrences" do
      @table.find(40).should == [[2,5],[6,3]]
      @table.find("Herbert").should == []
    end
  
  end

  describe "sort the table" do

    it "should sort the table according to first table" do
      @table.sort("Number")
      @table.ListRows.Item(1).Range.Value.first.should == [1,"Werner",40,0.5, 80]
      @table.ListRows.Item(2).Range.Value.first.should == [2, "Fred", nil, 0.5416666666666666, 40]     
      @table.ListRows.Item(1).Range.Value.first.should == [3, "John", 50.0, 0.5, 30]
      @table.ListRows.Item(3).Range.Value.first.should == [3, "Angel", 100, 0.6666666666666666, 60]
    end

  end

end
