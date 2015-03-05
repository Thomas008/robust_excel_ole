# -*- coding: utf-8 -*-

require File.join(File.dirname(__FILE__), './spec_helper')


$VERBOSE = nil

include RobustExcelOle

describe BookStore do

  before(:all) do
    excel = Excel.new(:reuse => true)
    open_books = excel == nil ? 0 : excel.Workbooks.Count
    puts "*** open books *** : #{open_books}" if open_books > 0
    Excel.close_all
  end

  before do
    @dir = create_tmpdir
    @simple_file = @dir + '/simple.xls'
    @simple_save_file = @dir + '/simple_save.xls'
    @different_file = @dir + '/different_simple.xls'
    @simple_file_other_path = @dir + '/more_data/simple.xls'
  end

  after do
    #Excel.close_all
    rm_tmp(@dir)
  end


  describe "create bookstore" do
    context "with standard" do
      it "should create book store" do
        expect {
          @bookstore = BookStore.new
        }.to_not raise_error
        @bookstore.should be_a BookStore
      end
    end
  end

  describe "simple store and fetch" do
    
    context "with standard" do
      before do
        BookStore.new
        @book = Book.open(@simple_file)
      end

      after do
        @book.close
      end

      it "should do simple store and fetch" do        
        BookStore.store(@book)
        new_book = BookStore.fetch(@simple_file)
        new_book.should be_a Book
        new_book.should be_alive
        new_book.should == @book
        new_book.close
      end
    end
  end

end
