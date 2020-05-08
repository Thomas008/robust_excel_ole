# -*- coding: utf-8 -*-
require File.join(File.dirname(__FILE__), './spec_helper')

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
  end

  after do
    rm_tmp(@dir)
  end

  context "open simple.xls" do
    before do
      @book = Workbook.open(@dir + '/workbook.xls', :read_only => true)
      @sheet = @book.sheet(2)
      @cell = @sheet[1, 1]
    end

    after do
      @book.close
    end

    describe "#value" do
      it "get cell's value" do
        @cell.Value.should eq 'simple'
      end
    end

    describe "#value=" do
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

  context "open merge_cells.xls" do
    before do
      @book = Workbook.open(@dir + '/merge_cells.xls', :read_only => true)
      @sheet = @book.sheet(1)
    end

    after do
      @book.close
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
end
