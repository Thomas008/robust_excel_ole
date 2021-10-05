# -*- coding: utf-8 -*-

require_relative 'spec_helper'

$VERBOSE = nil

include RobustExcelOle
include General

describe ListRow do
 
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

  describe "accessing several tables" do

    it "should preserve ole_table" do
      table1 = @sheet.table(1)
      table1[1].values.should == [3.0, "John", 50.0, 0.5, 30.0]
      table2 = Table.new(@sheet, "table_name", [1,1], 3, ["Person","Amount"])
      table2[1].values.should == [nil, nil]
      table1[1].values.should == [3.0, "John", 50.0, 0.5, 30.0]
    end

  end

  describe "#methods, #respond_to?" do

    before do
      @table1 = @sheet.table(1)
      @tablerow1 = @table1[2]
    end

    it "should contain column name as methods" do
      column_names_both_cases = @table1.column_names + @table1.column_names.map{|c| c.downcase}
      column_names_both_cases.map{|c| c.to_sym}.each do |column_name_method|
        @tablerow1.methods.include?(column_name_method).should be true
        @tablerow1.respond_to?(column_name_method)
      end
    end
  end

  describe "==" do

    before do
      @table1 = @sheet.table(1)
    end

    it "should yield true" do
      (@table1[1] == @table1[1]).should be true
    end

    it "should yield true" do
      (@table1[1] == @table1[2]).should be false
    end

  end

  describe "promote" do

     it "should promote a win32ole tablerow" do
      table1 = @sheet.table(1)
      tablerow1 = table1[2]
      ole_tablerow1 = tablerow1.ole_tablerow
      ListRow.new(ole_tablerow1).values.should == [2.0, "Fred", 0, 0.5416666666666666, 40]
      ListRow.new(tablerow1).values.should == [2.0, "Fred", 0, 0.5416666666666666, 40]
    end 
  end

  describe "to_a, to_h" do

    before do
      @table1 = @sheet.table(1)
    end

    it "should yield values of a row" do
      @table1[2].to_a.should == [2.0, "Fred", 0, 0.5416666666666666, 40]
      @table1[2].values.should == [2.0, "Fred", 0, 0.5416666666666666, 40]
    end

    it "should yield key-value pairs of a row" do
      @table1[2].to_h.should == {"Number" => 2.0, "Person" => "Fred", "Amount" => 0, "Time" => 0.5416666666666666, "Price" => 40}
      @table1[2].keys_values.should == {"Number" => 2.0, "Person" => "Fred", "Amount" => 0, "Time" => 0.5416666666666666, "Price" => 40}
    end

    it "should yield values and key-value pairs of a row with umlauts" do
      @table1[1].values = [1, "Sören", 20.0, 0.1, 40]
      @table1[1].values.should == [1.0, "Sören", 20.0, 0.1, 40]
      @table1[1].to_a.should == [1.0, "Sören", 20.0, 0.1, 40]
      @table1[1].to_h.should == {"Number" => 1.0, "Person" => "Sören", "Amount" => 20.0, "Time" => 0.1, "Price" => 40}
    end

  end

  describe "[], []=" do

    before do
      @table1 = @sheet.table(1)
      @table_row1 = @table1[1]
    end

    it "should read value in column" do
      @table_row1[2].should == "John"
      @table_row1["Person"].should == "John"
    end

    it "should read value in column with umlauts" do
      @table1.add_column("Straße", 1, ["Sören","ö","ü","ß","²","³","g","h","i","j","k","l","m"])
      table_row2 = @table1[1]
      @table_row1[1].should == "Sören"
      @table_row1["Straße"].should == "Sören"
    end

  end

  describe "getting and setting values" do

    context "with various column names" do

      context "with standard" do

        before do
          @table = Table.new(@sheet, "table_name", [22,1], 3, ["Person1","Win/Sales", "xiq-Xs", "OrderID", "YEAR", "length in m", "Amo%untSal___es"])
          @table_row1 = @table[1]
        end

        it "should read and set values via alternative column names" do
          @table_row1.person1.should be nil
          @table_row1.person1 = "John"
          @table_row1.person1.should == "John"
          @sheet[23,1].should == "John"
          @table_row1.Person1 = "Herbert"
          @table_row1.Person1.should == "Herbert"
          @sheet[23,1].should == "Herbert"
          @table_row1.win_sales.should be nil
          @table_row1.win_sales = 42
          @table_row1.win_sales.should == 42
          @sheet[23,2].should == 42
          @table_row1.Win_Sales = 80
          @table_row1.Win_Sales.should == 80
          @sheet[23,2].should == 80
          @table_row1.xiq_xs.should == nil
          @table_row1.xiq_xs = 90
          @table_row1.xiq_xs.should == 90
          @sheet[23,3].should == 90
          @table_row1.xiq_Xs = 100
          @table_row1.xiq_Xs.should == 100
          @sheet[23,3].should == 100
          @table_row1.order_id.should == nil
          @table_row1.order_id = 1
          @table_row1.order_id.should == 1
          @sheet[23,4].should == 1
          @table_row1.OrderID = 2
          @table_row1.OrderID.should == 2
          @sheet[23,4].should == 2
          @table_row1.year = 1984
          @table_row1.year.should == 1984
          @sheet[23,5].should == 1984
          @table_row1.YEAR = 2020
          @table_row1.YEAR.should == 2020
          @sheet[23,5].should == 2020
          @table_row1.length_in_m.should == nil
          @table_row1.length_in_m = 20
          @table_row1.length_in_m.should == 20
          @sheet[23,6].should == 20
          @table_row1.length_in_m = 40
          @table_row1.length_in_m.should == 40
          @sheet[23,6].should == 40
          @table_row1.amo_unt_sal___es.should == nil
          @table_row1.amo_unt_sal___es = 80
          @table_row1.amo_unt_sal___es.should == 80
          @sheet[23,7].should == 80
        end

      end

      context "with umlauts" do

        before do
          @table = Table.new(@sheet, "table_name", [1,1], 3, ["Verkäufer", "Straße", "area in m²"])
          @table_row1 = @table[1]         
        end

        it "should read and set values via alternative column names" do
          @table_row1.verkaeufer.should be nil
          @table_row1.verkaeufer = "John"
          @table_row1.verkaeufer.should == "John"
          @sheet[2,1].should == "John"
          @table_row1.Verkaeufer = "Herbert"
          @table_row1.Verkaeufer.should == "Herbert"
          @sheet[2,1].should == "Herbert"
          @table_row1.strasse.should be nil
          @table_row1.strasse = 42
          @table_row1.strasse.should == 42
          @sheet[2,2].should == 42
          @table_row1.Strasse = 80
          @table_row1.Strasse.should == 80
          @sheet[2,2].should == 80
          @table_row1.area_in_m2.should be nil
          @table_row1.area_in_m2 = 10
          @table_row1.area_in_m2.should == 10
          @sheet[2,3].should == 10
        end

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
        @sheet[4,4].should == 1
        @table_row1.person.should == "John"
        @table_row1.person = "Herbert"
        @table_row1.person.should == "Herbert"
        @sheet[4,5].should == "Herbert"
      end
    end

  end

end
