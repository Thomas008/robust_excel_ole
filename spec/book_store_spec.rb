# -*- coding: utf-8 -*-

require File.join(File.dirname(__FILE__), './spec_helper')

RSpec.configure do |config|

  config.mock_with :rspec do |mocks|
    mocks.syntax = :should
  end
end


$VERBOSE = nil

include RobustExcelOle

module RobustExcelOle
  class MockBookstore
    def fetch(filename, options = { })
      nil
    end
    def store(book)
    end
    def print
      puts "MockBookstore is always empty"
    end
  end
end


$mock_bookstore = MockBookstore.new

class Book
  @@bookstore = $mock_bookstore
end


describe BookStore do

  before(:all) do
    excel = Excel.new(:reuse => true)
    open_books = excel == nil ? 0 : excel.Workbooks.Count
    puts "*** open books *** : #{open_books}" if open_books > 0
    Excel.close_all
  end

  before do
    @bookstore = BookStore.new
    @dir = create_tmpdir
    @simple_file = @dir + '/simple.xls'
    @simple_save_file = @dir + '/simple_save.xls'
    @different_file = @dir + '/different_simple.xls'
    @simple_file_other_path = @dir + '/more_data/simple.xls'
  end

  after do
    Excel.close_all
    rm_tmp(@dir)
  end

  describe "create bookstore" do
    context "with standard" do
      it "should create book store" do
        expect {
          @book_store = BookStore.new
        }.to_not raise_error
        @book_store.should be_a BookStore
      end
    end
  end

  describe "Mock-Test" do
    it "should never store any book" do
      b1 = Book.open(@simple_file)
      b2 = Book.open(@simple_file)
      b2.object_id.should_not == b1.object_id
    end
  end


  describe "fetch" do
    
    context "with one open book" do
      
      before do
        @book = Book.open(@simple_file)
      end

      after do
        @book.close rescue nil
      end

      it "should do simple store and fetch" do        
        @bookstore.store(@book)
        new_book = @bookstore.fetch(@simple_file)
        new_book.should be_a Book
        new_book.should be_alive
        new_book.should == @book
        new_book.close
      end

      it "should fetch one book several times" do        
        @bookstore.store(@book)
        book1 = @bookstore.fetch(@simple_file)
        book2 = @bookstore.fetch(@simple_file)
        expect(book1).to be_a Book
        book1.should be_alive
        book1.should == @book
        book2.should be_a Book
        book2.should be_alive
        book2.should == @book
        book1.should == book2
        book1.close
        book2.close
      end

      it "should fetch nothing without stóring before" do
        new_book = @bookstore.fetch(@simple_file)
        new_book.should == nil
      end

      it "should fetch a closed book" do
        @bookstore.store(@book)
        @book.close
        book1 = @bookstore.fetch(@simple_file)
        book1.should be_a Book
        book1.should_not be_alive
      end

      it "should fetch nothing when fetching a different book" do
        @bookstore.store(@book)
        new_book = @bookstore.fetch(@different_file)
        new_book.should == nil
      end

      it "should fetch nothing when fetching a non-existing book" do
        @bookstore.store(@book)
        new_book = @bookstore.fetch("foo")
        new_book.should == nil
      end

    end

    context "with several books" do

      before do
        @book = Book.open(@simple_file)
        @bookstore.store(@book)
      end

      after do
        @book.close
        @book2.close(:if_unsaved => :forget)
      end

      it "should store and open two different books" do
        @book2 = Book.open(@different_file)
        @bookstore.store(@book2)
        new_book = @bookstore.fetch(@simple_file)
        new_book2 = @bookstore.fetch(@different_file)
        new_book.should be_a Book
        new_book.should be_alive
        new_book.should == @book
        new_book2.should be_a Book
        new_book2.should be_alive
        new_book2.should == @book2
        new_book.should_not == new_book2
        new_book.close
        new_book2.close
      end

      it "should fetch the first, writable book" do
        @book2 = Book.open(@simple_file, :force_excel => :new)
        @bookstore.store(@book2)
        @book.ReadOnly.should be_false
        @book2.ReadOnly.should be_true
        new_book = @bookstore.fetch(@simple_file)
        new_book.should == @book
        new_book.should_not == @book2
        new_book.close
      end

      it "should fetch the writable book even if the readonly book has unsaved changes" do
        @book2 = Book.open(@simple_file, :force_excel => :new)
        sheet = @book2[0]
        @bookstore.store(@book2)
        sheet[0,0] = sheet[0,0].value == "simple" ? "complex" : "simple"
        @book.ReadOnly.should be_false
        @book2.ReadOnly.should be_true
        @book2.Saved. should be_false
        new_book = @bookstore.fetch(@simple_file)
        new_book.should == @book
        new_book.should_not == @book2
        new_book.close
      end

    end

    context "with readonly book" do

      before do
        @book = Book.open(@simple_file, :read_only => true)
        @bookstore.store(@book)
      end

      after do
        @book.close
        @book2.close(:if_unsaved => :forget)
      end

      it "should fetch the second, writable book" do
        @book2 = Book.open(@simple_file, :force_excel => :new)
        @bookstore.store(@book2)
        @book.ReadOnly.should be_true
        @book2.ReadOnly.should be_false
        new_book = @bookstore.fetch(@simple_file)
        new_book.should == @book2
        new_book.should_not == @book
        new_book.close
      end

      it "should fetch the recent readonly book when there are only readonly books" do
        @book2 = Book.open(@simple_file, :force_excel => :new, :read_only => true)
        @bookstore.store(@book2)
        @book.ReadOnly.should be_true
        @book2.ReadOnly.should be_true
        new_book = @bookstore.fetch(@simple_file)
        new_book.should == @book2
        new_book.should_not == @book
        new_book.close
      end

      it "should fetch the second readonly book with unsaved changes" do
        @book2 = Book.open(@simple_file, :force_excel => :new, :read_only => true)
        sheet = @book2[0]
        @bookstore.store(@book2)
        sheet[0,0] = sheet[0,0].value == "simple" ? "complex" : "simple"
        @book.ReadOnly.should be_true
        @book2.ReadOnly.should be_true
        @book2.Saved.should be_false
        new_book = @bookstore.fetch(@simple_file)
        new_book.should == @book2
        new_book.should_not == @book        
        new_book.close
      end

      it "should fetch the second, writable book, if a writable, a readonly and an unsaved readonly book exist" do
        @book2 = Book.open(@simple_file, :force_excel => :new)
        @book3 = Book.open(@simple_file, :force_excel => :new)
        @bookstore.store(@book2)
        @bookstore.store(@book3)
        sheet = @book3[0]
        sheet[0,0] = sheet[0,0].value == "simple" ? "complex" : "simple"
        @book.ReadOnly.should be_true
        @book2.ReadOnly.should be_false
        @book3.ReadOnly.should be_true
        @book3.Saved.should be_false
        new_book = @bookstore.fetch(@simple_file)
        new_book.should == @book2
        new_book.should_not == @book        
        new_book.should_not == @book3  
        new_book.close
      end
    end
   
    context "with several closed books" do
      
      before do
        @book = Book.open(@simple_file)
        @bookstore.store(@book)
        @book2 = Book.open(@simple_file, :force_excel => :new)
        @bookstore.store(@book2)
        @book.close
        @book2.close
      end

      it "should fetch the recent closed book" do
        new_book = @bookstore.fetch(@simple_file)
        new_book.should == @book2
        new_book.should_not == @book
      end
      
    end

    context "with changing file name" do

      before do
        @book = Book.open(@simple_file)
        @book.save_as(@simple_save_file, :if_exists => :overwrite)      
        @bookstore.store(@book)
        #@bookstore = @book.book_store
      end

      after do
        @book.close
      end

      it "should return only book with correct file name" do
        book1 = @bookstore.fetch(@simple_save_file)
        book1.should == @book
      end

      it "should return only book with correct file name" do
        book1 = @bookstore.fetch(@simple_file)
        book1.should == nil
      end
    end


    context "with given excel instance and fetching readonly" do
      
      before do
        @book = Book.open(@simple_file)
        @bookstore.store(@book)
        @book2 = Book.open(@simple_file, :force_excel => :new)
        @bookstore.store(@book2)        
      end

      after do
        @book.close
      end

      it "should fetch the book in the given excel instance" do
        @book.ReadOnly.should be_false
        @book2.ReadOnly.should be_true
        book_new = @bookstore.fetch(@simple_file, :prefer_excel => @book2.excel)
        book_new.should be_a Book
        book_new.should be_alive
        book_new.should == @book2
      end
    end
  end
  
  describe "book life cycle" do
    
    context "with an open book" do

      before do
        @book = Book.open(@simple_file)
        @bookstore.store(@book)
      end

      after do
        @book.close rescue nil
      end

      it "should find the book if the book has astill got a reference" do
        GC.start
        @bookstore.fetch(@simple_file).should == @book
      end

      it "should have forgotten the book if there is no reference anymore" do
        @book = nil
        GC.start
        @bookstore.fetch(@simple_file).should == nil
      end

      it "should have forgotten some books if they have no reference anymore" do
        book_new = Book.open(@different_file)
        @bookstore.store(book_new)
        @book = nil
        GC.start
        @bookstore.fetch(@simple_file).should == nil
        @bookstore.fetch(@different_file).should == book_new
      end
    end
  end

  describe "excel_list" do

    context "with no books" do
      
      it "should yield nil" do
        @bookstore.excel_list.should == {}
      end
    
    end

    context "with open books" do
    
      before do
        @book = Book.open(@simple_file)
        @bookstore.store(@book)
      end

      after do
        @book.close
      end

      it "should yield an excel and the workbook" do
        excels = @bookstore.excel_list
        excels.size.should == 1
        excels.each do |excel,workbooks|
          excel.should be_a Excel
          workbooks.size.should == 1
          workbooks.each do |workbook|
            workbook.should == @book.workbook
          end
        end
      end

      it "should yield an excel with two books" do
        book1 = Book.open(@different_file)
        @bookstore.store(book1)
        excels = @bookstore.excel_list
        excels.size.should == 1
        excels.each do |excel,workbooks|
          excel.should be_a Excel
          workbooks.size.should == 2
          workbooks[0].should == @book.workbook
          #workbooks[1].should == book1.workbook
        end
      end

      it "should yield two excels and two book" do
        e = Excel.create
        book1 = Book.open(@simple_file, :force_excel => :new)
        @bookstore.store(book1)
        excels = @bookstore.excel_list
        excels.size.should == 2
        num = 0
        excels.each do |excel,workbooks|
          num = num + 1
          excel.should be_a Excel
          workbooks[0].should == @book.workbook if num == 1
          workbooks.size.should == 1 if num == 2
          workbooks[0].should == book1.workbook if num == 2
        end
      end
    end
  end
end
