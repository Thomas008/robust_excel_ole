# -*- coding: utf-8 -*-

require_relative 'spec_helper'

$VERBOSE = nil

include RobustExcelOle
include General

describe ListObject do
 
  before(:all) do
    excel = Excel.new(reuse: true)
    open_books = excel == nil ? 0 : excel.Workbooks.Count
    puts "*** open books *** : #{open_books}" if open_books > 0
    Excel.kill_all
  end 

  before do
    @dir = create_tmpdir
    @listobject_file = @dir + '/workbook_listobjects.xlsx'
    @book = Workbook.open(@listobject_file, visible: true)
    @sheet = @book.sheet(3)
  end

  after do
    @book.close(if_unsaved: :forget)
    Excel.kill_all
    rm_tmp(@dir)
  end

  describe "accessing a table" do

    it "should access a table via its number" do
      table = @sheet.table(1)
      table.Name.should == "table3"
    end

    it "should access a table via its name" do
      table = @sheet.table("table3")
      table.Name.should == "table3"
    end

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

      it "should create a new table with umlauts" do
        table = Table.new(@sheet, "lösung", [1,1], 3, ["Verkäufer","Straße"])
        table.Name.encode_value.should == "lösung"
        table.HeaderRowRange.Value.first.encode_value.should == ["Verkäufer","Straße"]
        table.ListRows.Count.should == 3
        @sheet[1,1].Value.encode_value.should == "Verkäufer"
      end

      it "should do the idempotence" do
        ole_table = @sheet.ListObjects.Item(1)
        table = Table.new(ole_table)
        table2 = Table.new(table)
        table2.ole_table.should be_a WIN32OLE
      end

      it "should type-lift a Win32ole list object into a RobustExcelOle list object" do
        ole_table = @sheet.ListObjects.Item(1)
        table = Table.new(ole_table)
        table.Name.should == "table3"
        table.HeaderRowRange.Value.first.should == ["Number","Person","Amount","Time","Price"]
        table.ListRows.Count.should == 13
        @sheet[3,4].Value.should == "Number"
        table.position.should == [3,4]
      end

      it "should type-lift a Win32ole list object into a RobustExcelOle list object with table name" do
        ole_table = @sheet.ListObjects.Item(1)
        table = Table.new(@sheet, "table3")
        table.Name.should == "table3"
        table.HeaderRowRange.Value.first.should == ["Number","Person","Amount","Time","Price"]
        table.ListRows.Count.should == 13
        @sheet[3,4].Value.should == "Number"
      end

      it "should type-lift a Win32ole list object into a RobustExcelOle list object with item number" do
        ole_table = @sheet.ListObjects.Item(1)
        table = Table.new(@sheet, 1)
        table.Name.should == "table3"
        table.HeaderRowRange.Value.first.should == ["Number","Person","Amount","Time","Price"]
        table.ListRows.Count.should == 13
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
        table.ListRows.Count.should == 13
        @sheet[3,4].Value.should == "Number"
      end

      it "should type-lift a Win32ole list object into a RobustExcelOle list object with item number" do
        ole_table = @sheet.ListObjects.Item(1)
        table = Table.new(@sheet.ole_worksheet, 1)
        table.Name.should == "table3"
        table.HeaderRowRange.Value.first.should == ["Number","Person","Amount","Time","Price"]
        table.ListRows.Count.should == 13
        @sheet[3,4].Value.should == "Number"
      end

    end

  end

  describe "benchmarking for accessing a listrow" do

    it "should access the last row" do
      rows =  150
      table = Table.new(@sheet.ole_worksheet, "table_name", [20,1], rows, ["Index","Person", "Time", "Price", "Sales", "Length", "Size", "Width", "Weight", "Income", "Outcome", "Holiday", "Gender", "Sex", "Tallness", "Kindness", "Music", "Activity", "Goal", "Need"])
      (1..rows).each do |row|
        table[row].values = [12345678, "Johnason", 12345678, "Johnason", 12345678, "Johnason", 12345678, "Johnason", 12345678, "Johnason", 12345678, "Johnason", 12345678, "Johnason", 12345678, "Johnason", 12345678, "Johnason", 12345678, "Johnason"]
      end
      table[rows].values = [12345123, "Peterson", 12345678, "Johnason", 12345678, "Johnason", 12345678, "Johnason", 12345678, "Johnason", 12345678, "Johnason", 12345678, "Johnason", 12345678, "Johnason", 12345678, "Johnason", 12345678, "Johnason"]
      sleep 1
      start_time = Time.now
      listrow = table[{"Index" => 12345123, "Person" => "Peterson"}]
      end_time = Time.now
      duration = end_time - start_time
      puts "duration: #{duration}"
      puts "listrow.values: #{listrow.values}"
    end

  end

  describe "accessing a listrow" do

    before do
      @table1 = @sheet.table(1)
    end

    it "should access a listrow given its number" do
      @table1[2].values.should == [2.0, "Fred", nil, 0.5416666666666666, 40]
    end   

    it "should access a listrow via a multiple-column key" do
      @table1[{"Number" => 3, "Person" => "Angel"}].values.should == [3.0, "Angel", 100, 0.6666666666666666, 60]
    end

    it "should yield nil if there is no match" do
      @table1[{"Number" => 5, "Person" => "Angel"}].should == nil
      @table1[{"Number" => 5, "Person" => "Angel"}, limit: :first].should == nil
    end

    it "should yield nil if there is no match" do
      @table1[{"Number" => 5, "Person" => "Angel"}, limit: 1].should == []
    end

    #it "should raise an error if the key contains no existing columns" do
    #  expect{
     #   @table1[{"Number" => 3, "Persona" => "Angel"}]
     #   }.to raise_error(TableError)
     #end

    it "should access one matching listrow" do
      @table1[{"Number" => 3}, limit: :first].values.should == [3.0, "John", 50.0, 0.5, 30]
    end

    it "should access one matching listrow" do
      @table1[{"Number" => 3}, limit: 1].map{|l| l.values}.should == [[3.0, "John", 50.0, 0.5, 30]]
    end
 
    it "should access two matching listrows" do
      @table1[{"Number" => 3}, limit: 2].map{|l| l.values}.should == [[3.0, "John", 50.0, 0.5, 30],[3.0, "Angel", 100, 0.6666666666666666, 60]]
    end

    it "should access four matching listrows" do
      @table1[{"Number" => 3}, limit: 4].map{|l| l.values}.should == [[3.0, "John", 50.0, 0.5, 30],[3.0, "Angel", 100, 0.6666666666666666, 60],
                                                               [3.0, "Eiffel", 50.0, 0.5, 30], [3.0, "Berta", nil, 0.5416666666666666, 40]]
    end

    it "should access all matching listrows" do
      @table1[{"Number" => 3}, limit: nil].map{|l| l.values}.should == [[3.0, "John", 50.0, 0.5, 30],
                                                                 [3.0, "Angel", 100, 0.6666666666666666, 60],
                                                                 [3.0, "Eiffel", 50.0, 0.5, 30], 
                                                                 [3.0, "Berta", nil, 0.5416666666666666, 40],
                                                                 [3.0, "Martha", nil, nil, nil],
                                                                 [3.0, "Paul", 40.0, 0.5, 80]]
    end

    it "should access listrows containing umlauts" do
      #@table1[1].values = [1, "Sören", 20.0, 0.1, 40]
      #@table1[{"Number" => 1, "Person" => "Sören"}].values.encode_value.should == [1, "Sören", 20.0, 0.1, 40]
      @table1.add_column("Straße", 3, ["ä","ö","ü","ß","²","³","g","h","i","j","k","l","m"])
      @table1[1].values = [1, "Sören", "Lösung", 20.0, 0.1, 40]
      @table1[1].values.should == [1, "Sören", "Lösung", 20.0, 0.1, 40]
      @table1[{"Number" => 1, "Person" => "Sören"}].values.encode_value.should == [1, "Sören", "Lösung", 20.0, 0.1, 40]
      @table1[{"Number" => 1, "Straße" => "Lösung"}].values.encode_value.should == [1, "Sören", "Lösung", 20.0, 0.1, 40]
    end

  end

  describe "reading and setting contents of rows and columns" do

    before do
      ole_table = @sheet.ListObjects.Item(1)
      @table = Table.new(ole_table)
    end

    it "should read contents of a column" do
      @table.column_values("Person").should == ["John","Fred",nil,"Angel",nil,"Werner","Eiffel","Berta",nil,nil,"Martha","Paul","Napoli"]
      expect{
        @table.column_values("P")
      }.to raise_error(TableError)
    end

    it "should set contents of a column" do
      @table.set_column_values("Person",["H",nil,nil,nil,"G","W",nil,nil,nil,nil,nil,nil,nil])
      @table.ListColumns.Item(2).Range.Value.should == [["Person"],["H"],[nil],[nil],[nil],["G"],["W"],[nil],[nil],[nil],[nil],[nil],[nil],[nil]]
      @table.set_column_values("Person",["T","Z"])
      @table.ListColumns.Item(2).Range.Value.should == [["Person"],["T"],["Z"],[nil],[nil],["G"],["W"],[nil],[nil],[nil],[nil],[nil],[nil],[nil]]
      expect{
        @table.set_column_values("P",["H",nil,nil,nil,"G","W"])
      }.to raise_error(TableError)
    end

    it "should read contents of a row" do
      @table.row_values(1).should == [3.0, "John", 50.0, 0.5, 30]
      @table[1].values.should == [3.0, "John", 50.0, 0.5, 30]
      expect{
        @table.row_values(14)
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
        @table.set_row_values(14, [5, "George", 30.0, 0.2, 50])
      }.to raise_error(TableError)
    end

    it "should set contents of an incomplete row " do
      @table[1].values = [2, "Merlin", 20.0, 0.1, 40]
      @table[1].values = [4, "John"]
      @table.ListRows.Item(1).Range.Value.first.should == [4, "John", 20.0, 0.1, 40]
    end

    it "should set contents of a row with umlauts" do
      @table[1].values = [1, "Sören", 20.0, 0.1, 40]
      @table.ListRows.Item(1).Range.Value.first.encode_value.should == [1, "Sören", 20.0, 0.1, 40]
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
      @table.add_column("column_name", 3, ["a","b","c","d","e","f","g","h","i","j","k","l","m"])
      column_names = @table.HeaderRowRange.Value.first.should == ["Number","Person","column_name","Amount","Time","Price"]
      @table.ListColumns.Item(3).Range.Value.should == [["column_name"],["a"],["b"],["c"],["d"],["e"],["f"],["g"],["h"],["i"],["j"],["k"],["l"],["m"]]
    end

    it "should add a column with umlauts" do
      @table.add_column("Sören", 3, ["ä","ö","ü","ß","²","³","g","h","i","j","k","l","m"])
      column_names = @table.HeaderRowRange.Value.first.encode_value.should == ["Number","Person","Sören","Amount","Time","Price"]
      @table.ListColumns.Item(3).Range.Value.map{|v| v.encode_value}.should == [["Sören"],["ä"],["ö"],["ü"],["ß"],["²"],["³"],["g"],["h"],["i"],["j"],["k"],["l"],["m"]]
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
        @table.add_row(16)
      }.to raise_error(TableError)
    end

    it "should add a row with contents" do
      @table.add_row(2, [2.0, "Herbert", 30.0, 0.25, 40])
      listrows = @table.ListRows
      listrows.Item(1).Range.Value.first.should == [3.0, "John", 50.0, 0.5, 30]
      listrows.Item(2).Range.Value.first.should == [2.0, "Herbert", 30.0, 0.25, 40]
      listrows.Item(3).Range.Value.first.should == [2.0, "Fred", nil, 0.5416666666666666, 40]
    end

    it "should add a row with contents with umlauts" do
      @table.add_row(1, [2.0, "Sören", 30.0, 0.25, 40])
      @table.ListRows.Item(1).Range.Value.first.encode_value.should == [2.0, "Sören", 30.0, 0.25, 40]
    end

    it "should delete a row" do
      @table.delete_row(4)
      listrows = @table.ListRows
      listrows.Item(5).Range.Value.first.should == [1,"Werner",40,0.5, 80]
      listrows.Item(4).Range.Value.first.should == [nil,nil,nil,nil,nil]
      expect{
        @table.delete_row(15)
      }.to raise_error(TableError)
    end

    it "should delete the contents of a column" do
      @table.ListColumns.Item(3).Range.Value.should == [["Amount"],[50],[nil],[nil],[100],[nil],[40],[50],[nil],[nil],[nil],[nil],[40],[20]]
      @table.delete_column_values(3)
      @table.HeaderRowRange.Value.first.should == ["Number","Person", "Amount", "Time","Price"]
      @table.ListColumns.Item(3).Range.Value.should == [["Amount"],[nil],[nil],[nil],[nil],[nil],[nil],[nil],[nil],[nil],[nil],[nil],[nil],[nil]]
      @table.ListColumns.Item(1).Range.Value.should == [["Number"],[3],[2],[nil],[3],[nil],[1],[3],[3],[nil],[nil],[3],[3],[1]]
      @table.delete_column_values("Number")
      @table.ListColumns.Item(1).Range.Value.should == [["Number"],[nil],[nil],[nil],[nil],[nil],[nil],[nil],[nil],[nil],[nil],[nil],[nil],[nil]]
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
        @table.delete_row_values(14)
      }.to raise_error(TableError)
    end

    it "should delete empty rows" do
      @table.delete_empty_rows
      @table.ListRows.Count.should == 9
      @table.ListRows.Item(1).Range.Value.first.should == [3.0, "John", 50.0, 0.5, 30]
      @table.ListRows.Item(2).Range.Value.first.should == [2.0, "Fred", nil, 0.5416666666666666, 40]
      @table.ListRows.Item(3).Range.Value.first.should == [3, "Angel", 100, 0.6666666666666666, 60]
      @table.ListRows.Item(4).Range.Value.first.should == [1,"Werner",40,0.5, 80]
    end

    it "should delete empty columns" do
      @table.delete_column_values(4)
      @table.ListColumns.Count.should == 5
      @table.HeaderRowRange.Value.first.should == ["Number","Person", "Amount", "Time","Price"]
      @table.delete_empty_columns
      @table.ListColumns.Count.should == 4
      @table.HeaderRowRange.Value.first.should == ["Number","Person", "Amount","Price"]
    end
  end

  describe "find all cells of a given value" do

    context "with standard" do

      before do
        ole_table = @sheet.ListObjects.Item(1)
        @table = Table.new(ole_table)
      end

      it "should find all cells" do
        cells = @table.find_cells(40)
        cells[0].Row.should == 5
        cells[0].Column.should == 8
        cells[1].Row.should == 9
        cells[1].Column.should == 6
      end

    end

    context "with umlauts" do

      before do
        @table = Table.new(@sheet, "lösung", [1,1], 3, ["Verkäufer","Straße"])
        @table[1].values = ["sören", "stück"]
        @table[2].values = ["stück", "glück"]
        @table[3].values = ["soße",  "stück"]
      end

      it "should find all cells" do
        cells = @table.find_cells("stück")
        cells[0].Row.should == 2
        cells[0].Column.should == 2
        cells[1].Row.should == 3
        cells[1].Column.should == 1
        cells[2].Row.should == 4
        cells[2].Column.should == 2
      end

    end
  
  end

  describe "sort the table" do

    before do
      ole_table = @sheet.ListObjects.Item(1)
      @table = Table.new(ole_table)
    end

    it "should sort the table according to first table" do
      @table.sort("Number")
      @table.ListRows.Item(1).Range.Value.first.should == [1, "Werner",40,0.5, 80]
      @table.ListRows.Item(2).Range.Value.first.should == [1, "Napoli", 20.0, 0.4166666666666667, 70.0]
      @table.ListRows.Item(3).Range.Value.first.should == [2, "Fred", nil, 0.5416666666666666, 40]     
      @table.ListRows.Item(4).Range.Value.first.should == [3, "John", 50.0, 0.5, 30]
      @table.ListRows.Item(5).Range.Value.first.should == [3, "Angel", 100, 0.6666666666666666, 60]
      @table.ListRows.Item(6).Range.Value.first.should == [3, "Eiffel", 50.0, 0.5, 30]
      @table.ListRows.Item(7).Range.Value.first.should == [3, "Berta", nil, 0.5416666666666666, 40]
      @table.ListRows.Item(8).Range.Value.first.should == [3, "Martha", nil, nil, nil]
      @table.ListRows.Item(9).Range.Value.first.should == [3, "Paul", 40.0, 0.5, 80]
    end

  end

end
