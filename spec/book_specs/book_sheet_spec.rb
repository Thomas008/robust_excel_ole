# -*- coding: utf-8 -*-

require File.join(File.dirname(__FILE__), './../spec_helper')


$VERBOSE = nil

include RobustExcelOle
include General

describe Book do

  before(:all) do
    excel = Excel.new(:reuse => true)
    open_books = excel == nil ? 0 : excel.Workbooks.Count
    puts "*** open books *** : #{open_books}" if open_books > 0
    Excel.close_all
  end

  before do
    @dir = create_tmpdir
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

  describe "#add_sheet" do
    
    context "with no given sheet" do

      before do
      @book = Book.open(@simple_file)
      @sheet = @book[0]
    end

    after do
      @book.close(:if_unsaved => :forget)
    end

      it "should add empty sheet" do
        @book.ole_workbook.Worksheets.Count.should == 3
        @book.add_sheet
        @book.ole_workbook.Worksheets.Count.should == 4
      end

      it "should add an empty sheet and return this added sheet" do
        sheet = @book.add_sheet
        copyed_sheet = @book.ole_workbook.Worksheets.Item(@book.ole_workbook.Worksheets.Count)
        sheet.name.should eq copyed_sheet.name
      end

      it "should return new sheet" do
        @book.add_sheet(:as => 'new sheet').name.should eq 'new sheet'
      end

      it "should add the first sheet" do
        @book.add_sheet(:before => @sheet).name.should eq @book[0].name
      end

      it "should add the second sheet" do
        @book.add_sheet(:after => @sheet).name.should eq @book[1].name
      end

    end

    context "with a given sheet" do

      before do
        @book = Book.open(@simple_file)
        @sheet = @book[0]
        @another_book = Book.open(@another_simple_file)
      end

      after do
        @book.close(:if_unsaved => :forget)
        @another_book.close(:if_unsaved => :forget)
      end

      it "should copy and append a given sheet" do
        @book.ole_workbook.Worksheets.Count.should == 3
        @book.add_sheet @sheet
        @book.ole_workbook.Worksheets.Count.should == 4
        @book.ole_workbook.Worksheets(4).Name.should == @sheet.Name + " (2)"
      end

      it "should copy sheet from another book " do
        @book.ole_workbook.Worksheets.Count.should == 3
        @another_book.add_sheet @sheet
        @another_book.ole_workbook.Worksheets.Count.should == 4
        @another_book.ole_workbook.Worksheets(4).Name.should == @sheet.Name + " (2)"
      end

      it "should return copyed sheet" do
        sheet = @book.add_sheet @sheet
        copyed_sheet = @book.ole_workbook.Worksheets.Item(@book.ole_workbook.Worksheets.Count)
        sheet.name.should eq copyed_sheet.name
      end

      it "should copy a given sheet and name the copyed sheet to 'copyed_name'" do
        @book.add_sheet(@sheet, :as => 'copyed_name').name.should eq 'copyed_name'
      end
    
      it "should copy the first sheet and insert it before the first sheet" do
        @book.add_sheet(@sheet, :before => @sheet).name.should eq @book[0].name
      end
   
      it "should copy the first sheet and insert it after the first sheet" do
        @book.add_sheet(@sheet, :after => @sheet).name.should eq @book[1].name
      end
    
      it "should copy the first sheet before the third sheet and give 'before' the highest priority" do
        @book.add_sheet(@sheet, :after => @sheet, :before => @book[2]).name.should eq @book[2].name
      end

      it "should copy the first sheet before the third sheet and give 'before' the highest priority" do
        @book.add_sheet(@sheet, :before => @book[2], :after => @sheet).name.should eq @book[2].name
      end
        
      it "should raise error with giving a name that already exists" do
        @book.add_sheet(@sheet, :as => 'new_sheet')
        expect{
          @book.add_sheet(@sheet, :as => 'new_sheet')
          }.to raise_error(ExcelErrorSheet, /sheet name "new_sheet" already exists/)
      end
    end
  end

  describe 'access sheet' do
    before do
      @book = Book.open(@simple_file)
    end

    after do
      @book.close
    end

    it 'with sheet name' do
      @book["Sheet1"].should be_kind_of Sheet
      @book["Sheet1"].name.should == "Sheet1"
    end

    it 'with integer' do
      @book[0].should be_kind_of Sheet
    end

    it 'with block' do
      @book.each do |sheet|
        sheet.should be_kind_of Sheet
      end
    end

    context 'open with block' do
      it {
        Book.open(@simple_file) do |book|
          book['Sheet1'].should be_a Sheet
        end
      }
    end
  end

  describe 'access first and last sheet' do
    before do
      @book = Book.open(@simple_file)
    end

    it "should access the first sheet" do
      first_sheet = @book.first_sheet
      first_sheet.name.should == Sheet.new(@book.Worksheets.Item(1)).Name
      first_sheet.name.should == @book[0].Name
    end

    it "should access the last sheet" do
      last_sheet = @book.last_sheet
      last_sheet.name.should == Sheet.new(@book.Worksheets.Item(3)).Name
      last_sheet.name.should == @book[2].Name
    end
  end

end

