# -*- coding: utf-8 -*-
require_relative 'spec_helper'

include RobustExcelOle
include General

describe Cell do

  before(:all) do
    excel = Excel.new(:reuse => true)
    open_books = excel == nil ? 0 : excel.Workbooks.Count
    puts "*** open books *** : #{open_books}" if open_books > 0
    Excel.kill_all
  end

  before do
    @dir = create_tmpdir
    @simple_file = @dir + '/workbook.xls'
    @merge_cells_file = @dir + '/merge_cells.xls'
    @book = Workbook.open(@simple_file)
    @sheet = @book.sheet(1)
    @cell = @sheet[1, 1]
  end

  after do
    Excel.kill_all
    rm_tmp(@dir)
  end


  describe "values" do

    it "should yield one element values" do
      @cell.values.should == ["foo"]
    end

  end

  describe "#[]" do

    it "should access to the cell itself" do
      @cell[0].should be_kind_of RobustExcelOle::Cell
      @cell[0].v.should == "foo"
    end

    it "should access to the cell itself" do
      @cell[1].should be_kind_of RobustExcelOle::Cell
      @cell[1].v.should == 'foo'
    end
    
  end

  describe "#copy" do
  
    before do
      @book1 = Workbook.open(@dir + '/workbook.xls')
      @sheet1 = @book1.sheet(1)
      @cell1 = @sheet1[1,1]
    end

    after do
      @book1.close(:if_unsaved => :forget)
    end

    it "should copy range" do
      @cell1.copy([2,3])
      @sheet1.range([1..2,1..3]).v.should == [["foo", "workbook", "sheet1"],["foo", nil, "foo"]]
    end
  end

  context "open merge_cells.xls" do
 
    before do
      @book = Workbook.open(@merge_cells_file, :read_only => true)
      @sheet = @book.sheet(1)
    end

    it "merged cell get same value" do
      @sheet[1, 1].Value.should be_nil
      @sheet[2, 1].Value.should eq 'first merged'
    end

    it "set merged cell" do
      @sheet[2, 1].Value = "set merge cell"
      @sheet[2, 1].Value.should eq "set merge cell"
      @sheet[2, 2].Value.should eq "set merge cell"
    end
  end

  describe "==" do

    it "should check equality of cells" do
      @cell.should == @sheet[1,1]
      @cell.should_not == @sheet[1,2]
    end

  end

  describe "#Value" do
    it "get cell's value" do
      @cell.Value.should eq 'foo'
    end
  end

  describe "#Value=" do
    it "change cell data to 'fooooo'" do
      @cell.Value = 'fooooo'
      @cell.Value.should eq 'fooooo'
    end
  end

  describe "#method_missing" do
    context "unknown method" do
      it { expect { @cell.hogehogefoo }.to raise_error(NoMethodError) }
    end
  end

end

