# -*- coding: utf-8 -*-

require File.join(File.dirname(__FILE__), './spec_helper')

=begin
RSpec.configure do |config|

  config.mock_with :rspec do |mocks|
    mocks.syntax = :should
  end
end
=end

$VERBOSE = nil

include RobustExcelOle

module RobustExcelOle
  class MockBookstore # :nodoc: #
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


describe Bookstore do

  before(:all) do
    excel = Excel.new(:reuse => true)
    open_books = excel == nil ? 0 : excel.Workbooks.Count
    puts "*** open books *** : #{open_books}" if open_books > 0
    Excel.close_all
  end

  before do
    @bookstore = Bookstore.new
    @dir = create_tmpdir
    @simple_file = @dir + '/workbook.xls'
    @simple_save_file = @dir + '/workbook_save.xls'
    @different_file = @dir + '/different_workbook.xls'
    @simple_file_other_path = @dir + '/more_data/workbook.xls'
  end

  after do
    begin
      Excel.kill_all
    rescue WeakRef::RefError => msg
      puts "#{msg.message}"
      Excel.kill_all
    end
    begin
      rm_tmp(@dir) rescue nil
    end
  end

  describe "create bookstore" do
    context "with standard" do
      it "should create book store" do
        expect {
          @book_store = Bookstore.new
        }.to_not raise_error
        @book_store.should be_a Bookstore
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

      it "should fetch the last book with :prefer_writeable => false" do
        @book2 = Book.open(@simple_file, :force_excel => :new)
        @bookstore.store(@book2)
        @book.ReadOnly.should be_false
        @book2.ReadOnly.should be_true
        new_book = @bookstore.fetch(@simple_file, :prefer_writable => false)
        new_book.should_not == @book
        new_book.should == @book2
        new_book.close
      end

      it "should fetch the second, open book, if the first book is closed" do
        @book2 = Book.open(@simple_file, :force_excel => :new)
        @bookstore.store(@book2)
        @book.ReadOnly.should be_false
        @book2.ReadOnly.should be_true
        @book.close
        new_book = @bookstore.fetch(@simple_file, :prefer_writable => false)
        new_book2 = @bookstore.fetch(@simple_file)
        new_book.should_not == @book
        new_book2.should_not == @book
        new_book.should == @book2
        new_book2.should == @book2
        new_book.close
        new_book2.close
      end

      it "should fetch the first, open book, if the second book is closed, even with :prefer_writeable => false" do
        @book2 = Book.open(@simple_file, :force_excel => :new)
        @bookstore.store(@book2)
        @book.ReadOnly.should be_false
        @book2.ReadOnly.should be_true
        @book2.close
        new_book = @bookstore.fetch(@simple_file, :prefer_writable => false)
        new_book2 = @bookstore.fetch(@simple_file)
        new_book.should_not == @book2
        new_book2.should_not == @book2
        new_book.should == @book
        new_book2.should == @book
        new_book.close
        new_book2.close
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
        new_book2 = @bookstore.fetch(@simple_file, :prefer_writable => true)
        new_book3 = @bookstore.fetch(@simple_file, :prefer_writable => false)
        new_book.should == @book2
        new_book2.should == @book2
        new_book3.should == @book2
        new_book.should_not == @book
        new_book2.should_not == @book
        new_book3.should_not == @book
        new_book.close
        new_book2.close
        new_book3.close
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

      it "should fetch the second, writable book, if a writable, a readonly and an unsaved readonly book exist" do
        @book2 = Book.open(@simple_file, :force_excel => :new)
        @book3 = Book.open(@simple_file, :force_excel => :new)
        @bookstore.store(@book2)
        @bookstore.store(@book3)
        sheet = @book3[0]
        sheet[1,1] = sheet[1,1].Value == "foo" ? "bar" : "foo"
        @book.ReadOnly.should be_true
        @book2.ReadOnly.should be_false
        @book3.ReadOnly.should be_true
        @book3.Saved.should be_false
        new_book = @bookstore.fetch(@simple_file)
        new_book2 = @bookstore.fetch(@simple_file, :prefer_writable => false)
        new_book.should == @book2
        new_book2.should == @book3
        new_book.should_not == @book        
        new_book.should_not == @book3
        new_book2.should_not == @book        
        new_book2.should_not == @book2  
        new_book.close
        new_book2.close
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
        @book = "Bla"
        @book = Book.open(@simple_file)
        @bookstore.store(@book)
        @book = nil
        GC.start
        @bookstore.fetch(@simple_file).should == nil
        @bookstore.fetch(@different_file).should == book_new
      end
    end
  end

  describe "hidden_excel" do
    
    context "with some open book" do

      before do
        @book = Book.open(@simple_file)
      end

      after do
        @book.close
      end

      it "should create and use a hidden Excel instance" do
        h_excel1 = @bookstore.hidden_excel
        h_excel1.should_not == @book.excel
        h_excel1.Visible.should be_false
        h_excel1.DisplayAlerts.should be_false
        book1 = Book.open(@simple_file, :force_excel => @bookstore.hidden_excel)
        book1.excel.should === h_excel1
        book1.excel.should_not === @book.excel
        Excel.close_all    
        h_excel2 = @bookstore.hidden_excel
        h_excel2.should_not == @book.excel
        h_excel2.should_not == book1.excel
        h_excel2.Visible.should be_false
        h_excel2.DisplayAlerts.should be_false
        book2 = Book.open(@simple_file, :force_excel => @bookstore.hidden_excel)
        book2.excel.should === h_excel2
        book2.excel.should_not === @book.excel
        book2.excel.should_not === book1.excel
      end

      it "should exclude hidden excel" do
        book1 = Book.open(@simple_file, :force_excel => @bookstore.hidden_excel)
        @bookstore.store(book1)
        book1.close
        book2 = @bookstore.fetch(@simple_file)
        book2.should == nil
      end  
    end
  end
end
