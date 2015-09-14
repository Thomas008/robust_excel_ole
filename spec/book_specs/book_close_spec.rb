# -*- coding: utf-8 -*-

require File.join(File.dirname(__FILE__), './../spec_helper')


$VERBOSE = nil

include RobustExcelOle

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
    Excel.close_all
    rm_tmp(@dir)
  end

  describe "close" do

    context "with saved book" do
      before do
        @book = Book.open(@simple_file)
      end

      it "should close book" do
        expect{
          @book.close
        }.to_not raise_error
        @book.should_not be_alive
      end
    end

    context "with unsaved read_only book" do
      before do
        @book = Book.open(@simple_file, :read_only => true)
        @sheet_count = @book.workbook.Worksheets.Count
        @book.add_sheet(@sheet, :as => 'a_name')
      end

      it "should close the unsaved book without error and without saving" do
        expect{
          @book.close
          }.to_not raise_error
        new_book = Book.open(@simple_file)
        new_book.workbook.Worksheets.Count.should ==  @sheet_count
        new_book.close
      end
    end

    context "with unsaved book" do
      before do
        @book = Book.open(@simple_file)
        @sheet_count = @book.workbook.Worksheets.Count
        @book.add_sheet(@sheet, :as => 'a_name')
        @sheet = @book[0]
      end

      after do
        @book.close(:if_unsaved => :forget) rescue nil
      end

      it "should raise error with option :raise" do
        expect{
          @book.close(:if_unsaved => :raise)
        }.to raise_error(ExcelErrorClose, /workbook is unsaved: "workbook.xls"/)
      end

      it "should raise error by default" do
        expect{
          @book.close(:if_unsaved => :raise)
        }.to raise_error(ExcelErrorClose, /workbook is unsaved: "workbook.xls"/)
      end

      it "should close the book and leave its file untouched with option :forget" do
        ole_workbook = @book.workbook
        excel = @book.excel
        expect {
          @book.close(:if_unsaved => :forget)
        }.to change {excel.Workbooks.Count }.by(-1)
        @book.workbook.should == nil
        @book.should_not be_alive
        expect{
          ole_workbook.Name}.to raise_error(WIN32OLERuntimeError)
        new_book = Book.open(@simple_file)
        begin
          new_book.workbook.Worksheets.Count.should ==  @sheet_count
        ensure
          new_book.close
        end
      end

      it "should raise an error for invalid option" do
        expect {
          @book.close(:if_unsaved => :invalid_option)
        }.to raise_error(ExcelErrorClose, ":if_unsaved: invalid option: :invalid_option") 
      end


      it "should save the book before close with option :save" do
        ole_workbook = @book.workbook
        excel = @book.excel
        expect {
          @book.close(:if_unsaved => :save)
        }.to change {excel.Workbooks.Count }.by(-1)
        @book.workbook.should == nil
        @book.should_not be_alive
        expect{
          ole_workbook.Name}.to raise_error(WIN32OLERuntimeError)
        new_book = Book.open(@simple_file)
        begin
          new_book.workbook.Worksheets.Count.should == @sheet_count + 1
        ensure
          new_book.close
        end
      end

      context "with :if_unsaved => :alert" do
        before do
          @key_sender = IO.popen  'ruby "' + File.join(File.dirname(__FILE__), '/helpers/key_sender.rb') + '" "Microsoft Excel" '  , "w"
        end

        after do
          @key_sender.close
        end

        possible_answers = [:yes, :no, :cancel]
        possible_answers.each_with_index do |answer, position|
          it "should" + (answer == :yes ? "" : " not") + " the unsaved book and" + (answer == :cancel ? " not" : "") + " close it" + "if user answers '#{answer}'" do
            # "Yes" is the  default. "No" is right of "Yes", "Cancel" is right of "No" --> language independent
            @key_sender.puts  "{right}" * position + "{enter}"
            ole_workbook = @book.workbook
            excel = @book.excel
            displayalert_value = @book.excel.DisplayAlerts
            if answer == :cancel then
              expect {
              @book.close(:if_unsaved => :alert)
              }.to raise_error(ExcelUserCanceled, "close: canceled by user")
              @book.workbook.Saved.should be_false
              @book.workbook.should_not == nil
              @book.should be_alive
            else
              expect {
                @book.close(:if_unsaved => :alert)
              }.to change {@book.excel.Workbooks.Count }.by(-1)
              @book.workbook.should == nil
              @book.should_not be_alive
              expect{ole_workbook.Name}.to raise_error(WIN32OLERuntimeError)
            end
            new_book = Book.open(@simple_file, :if_unsaved => :forget)
            begin
              new_book.workbook.Worksheets.Count.should == @sheet_count + (answer==:yes ? 1 : 0)
              new_book.excel.DisplayAlerts.should == displayalert_value
            ensure
              new_book.close
            end
          end
        end
      end
    end
  end
end
