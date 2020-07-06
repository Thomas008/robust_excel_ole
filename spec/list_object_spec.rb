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
        table.HeaderRowRange.Value.first.should == ["Number","Person","Amount","Time","Date"]
        table.ListRows.Count.should == 6
        @sheet[3,4].Value.should == "Number"
      end

      it "should type-lift a Win32ole list object into a RobustExcelOle list object with table name" do
        ole_table = @sheet.ListObjects.Item(1)
        table = Table.new(@sheet, "table3")
        table.Name.should == "table3"
        table.HeaderRowRange.Value.first.should == ["Number","Person","Amount","Time","Date"]
        table.ListRows.Count.should == 6
        @sheet[3,4].Value.should == "Number"
      end

      it "should type-lift a Win32ole list object into a RobustExcelOle list object with item number" do
        ole_table = @sheet.ListObjects.Item(1)
        table = Table.new(@sheet, 1)
        table.Name.should == "table3"
        table.HeaderRowRange.Value.first.should == ["Number","Person","Amount","Time","Date"]
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
        table.HeaderRowRange.Value.first.should == ["Number","Person","Amount","Time","Date"]
        table.ListRows.Count.should == 6
        @sheet[3,4].Value.should == "Number"
      end

      it "should type-lift a Win32ole list object into a RobustExcelOle list object with item number" do
        ole_table = @sheet.ListObjects.Item(1)
        table = Table.new(@sheet.ole_worksheet, 1)
        table.Name.should == "table3"
        table.HeaderRowRange.Value.first.should == ["Number","Person","Amount","Time","Date"]
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
        @table_row1.person.should == "Herbert"
        @table_row1.person = "John"
        @table_row1.person.should == "John"
        @sheet[4,5].Value.should == "John"
      end
    end

  end
end
