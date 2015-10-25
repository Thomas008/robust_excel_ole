# -*- coding: utf-8 -*-
require File.join(File.dirname(__FILE__), './spec_helper')

include RobustExcelOle

describe Cell do

  before(:all) do
    excel = Excel.new(:reuse => true)
    open_books = excel == nil ? 0 : excel.Workbooks.Count
    puts "*** open books *** : #{open_books}" if open_books > 0
    Excel.close_all
  end

  before do
    @dir = create_tmpdir
  end

  after do
    rm_tmp(@dir)
  end

  context "open simple.xls" do
    before do
      @book = Book.open(@dir + '/workbook.xls', :read_only => true)
      @sheet = @book[1]
      @cell = @sheet[1, 1]
    end

    after do
      @book.close
    end

    describe "#value" do
      it "get cell's value" do
        @cell.value.should eq 'simple'
      end
    end

    describe "#value=" do
      it "change cell data to 'fooooo'" do
        @cell.value = 'fooooo'
        @cell.value.should eq 'fooooo'
      end
    end

    describe "#method_missing" do
      context "unknown method" do
        it { expect { @cell.hogehogefoo }.to raise_error }
      end
    end

  end

  context "open merge_cells.xls" do
    before do
      @book = Book.open(@dir + '/merge_cells.xls', :read_only => true)
      @sheet = @book[0]
    end

    after do
      @book.close
    end

    it "merged cell get same value" do
      @sheet[1, 1].value.should be_nil
      @sheet[2, 1].value.should eq 'first merged'
    end

    it "set merged cell" do
      @sheet[2, 1].value = "set merge cell"
      @sheet[2, 1].value.should eq "set merge cell"
      @sheet[2, 2].value.should eq "set merge cell"
    end
  end
end
