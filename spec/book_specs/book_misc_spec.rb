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
    #rm_tmp(@dir)
  end

  describe "create file" do
    context "with standard" do
      it "open an existing file" do
        expect {
          @book = Book.new(@simple_file)
        }.to_not raise_error
        @book.should be_a Book
        @book.close
      end
    end
  end

  describe "send methods to workbook" do

    context "with standard" do
      before do
        @book = Book.open(@simple_file)
      end

      after do
        @book.close
      end

      it "should send Saved to workbook" do
        @book.Saved.should be_true
      end

      it "should send Fullname to workbook" do
        @book.Fullname.tr('\\','/').should == @simple_file
      end

      it "should raise an error for unknown methods or properties" do
        expect{
          @book.Foo
        }.to raise_error(VBAMethodMissingError, /unknown VBA property or method :Foo/)
      end

      it "should report that workbook is not alive" do
        @book.close
        expect{ @book.Nonexisting_method }.to raise_error(ExcelError, "method missing: workbook not alive")
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
        book2 = Book.open(@simple_file, :force_excel => @book.bookstore.hidden_excel)
        book2.excel.should_not == @book.excel
        book2.excel.visible.should be_false
        book2.excel.displayalerts.should be_false
        book2.close 
      end
    end
  end

  describe "nameval, set_nameval, [], []=" do
  
    before do
      @book1 = Book.open(@another_simple_file)
    end

    after do
      @book1.close(:if_unsaved => :forget)
    end   

    it "should return value of a range" do
      @book1.nameval("new").should == "foo"
      @book1.nameval("one").should == 1
      @book1.nameval("firstrow").should == [[1,2]]        
      @book1.nameval("four").should == [[1,2],[3,4]]
      @book1.nameval("firstrow").should_not == "12"
      @book1.nameval("firstcell").should == "foo"        
    end

    it "should return value of a range via []" do
      @book1["new"].should == "foo"
      @book1["one"].should == 1
      @book1["firstrow"] == [[1,2]]        
      @book1["four"].should == [[1,2],[3,4]]
      @book1["firstrow"].should_not == "12"
      @book1["firstcell"].should == "foo"        
    end

    it "should set value of a range" do
      @book1.set_nameval("new", "bar")
      @book1.nameval("new").should == "bar"
    end

    it "should set value of a range via []=" do
      @book1["new"] = "bar"
      @book1.nameval("new").should == "bar"
    end

    it "should evaluate a formula" do
      @book1.nameval("named_formula").should == 4      
    end

    it "should evaluate a formula via []" do
      @book1["named_formula"].should == 4      
    end


    it "should return default value if name not defined" do
      @book1.nameval("foo", :default => 2).should == 2
    end

    it "should raise an error if name not defined" do
      expect {
        @book1.nameval("foo")
      }.to raise_error(ExcelError, /name "foo" not in "another_workbook.xls"/)
      expect {
          @book1.set_nameval("foo","bar")
      }.to raise_error(ExcelError, /name "foo" not in "another_workbook.xls"/)
      expect {
          @book1["foo"] = "bar"
      }.to raise_error(ExcelError, /name "foo" not in "another_workbook.xls"/)
    end    

    it "should raise an error if name was defined but contents is calcuated" do
      expect {
        @book1.set_nameval("named_formula","bar")
      }.to raise_error(ExcelError, /cannot assign value to range named "named_formula" in "another_workbook.xls"/)
      expect {
        @book1["named_formula"] = "bar"
      }.to raise_error(ExcelError, /cannot assign value to range named "named_formula" in "another_workbook.xls"/)
    end
  end

  describe "rename_range" do
    
    before do
      @book1 = Book.open(@another_simple_file)
    end

    after do
      @book1.close(:if_unsaved => :forget)
    end

    it "should rename a range" do
      @book1.rename_range("four","five")
      @book1.nameval("five").should == [[1,2],[3,4]]
      expect {
        @book1.rename_range("four","five")
      }.to raise_error(ExcelError, /name "four" not in "another_workbook.xls"/)
    end
  end

  describe "alive?, filename, ==, visible, displayalerts, activate, saved, check_compatibility" do

    context "with alive?" do

      before do
        @book = Book.open(@simple_file)
      end

      after do
        @book.close
      end

      it "should return true, if book is alive" do
        @book.should be_alive
      end

      it "should return false, if book is dead" do
        @book.close
        @book.should_not be_alive
      end

    end

    context "with filename" do

      before do
        @book = Book.open(@simple_file)
      end

      after do
        @book.close
      end

      it "should return full file name" do
        @book.filename.should == @simple_file
      end

      it "should return nil for dead book" do
        @book.close
        @book.filename.should == nil
      end

    end

    context "with ==" do

      before do
        @book = Book.open(@simple_file)
      end

      after do
        @book.close
        @new_book.close rescue nil
      end

      it "should be true with two identical books" do
        @new_book = Book.open(@simple_file)
        @new_book.should == @book
      end

      it "should be false with two different books" do
        @new_book = Book.new(@different_file)
        @new_book.should_not == @book
      end

      it "should be false with same book names but different paths" do       
        @new_book = Book.new(@simple_file_other_path, :force_excel => :new)
        @new_book.should_not == @book
      end

      it "should be false with same book names but different excel instances" do
        @new_book = Book.new(@simple_file, :force_excel => :new)
        @new_book.should_not == @book
      end

      it "should be false with non-Books" do
        @book.should_not == "hallo"
        @book.should_not == 7
        @book.should_not == nil
      end
    end

    context "with saved" do

      before do
        @book = Book.open(@simple_file)
      end

      after do
        @book.close(:if_unsaved => :forget)
      end

      it "should yield true for a saved book" do
        @book.saved.should be_true
      end

      it "should yield false for an unsaved book" do
        sheet = @book.sheet(1)
        sheet[1,1] = sheet[1,1].value == "foo" ? "bar" : "foo"
        @book.saved.should be_false
      end
    end

    context "with visible" do

      before do
        @book = Book.open(@simple_file)
      end

      after do
        @book.close
      end

      it "should make the workbook visible" do
        @book.excel.visible = true
        @book.excel.visible.should be_true
        @book.visible.should be_true
        @book.excel.Windows(@book.ole_workbook.Name).Visible.should be_true
        @book.visible = false
        @book.excel.visible.should be_true
        @book.visible.should be_false
        @book.excel.Windows(@book.ole_workbook.Name).Visible.should be_false
        @book.visible = true
        @book.excel.visible.should be_true
        @book.visible.should be_true
        @book.excel.Windows(@book.ole_workbook.Name).Visible.should be_true
      end

    end

    context "with activate" do

      before do
        @key_sender = IO.popen  'ruby "' + File.join(File.dirname(__FILE__), '../helpers/key_sender.rb') + '" "Microsoft Office Excel" '  , "w"        
        @book = Book.open(@simple_file, :visible => true)
        @book2 = Book.open(@another_simple_file, :force_excel => :new, :visible => true)
      end

      after do
        @book.close(:if_unsaved => :forget)
        @book2.close(:if_unsaved => :forget)
        @key_sender.close
      end

      it "should activate a book" do
        sheet = @book.sheet(2)
        sheet.Activate
        sheet[2,3].Activate
        sheet2 = @book2.sheet(2)
        sheet2.Activate
        sheet2[3,2].Activate
        Excel.current.should == @book.excel
        @book2.activate
        @key_sender.puts "{a}{enter}"
        sleep 1
        sheet2[3,2].Value.should == "a"
        #Excel.current.should == @book2.excel
        @book.activate
        @key_sender.puts "{a}{enter}"
        sleep 1
        sheet[2,3].Value.should == "a"
        Excel.current.should == @book.excel
      end
    end

    context "with compatibility" do      

      it "should open and check compatibility" do
        book = Book.open(@simple_file, :visible => true, :check_compatibility => false)
        book.CheckCompatibility.should be_false
        book.CheckCompatibility = true
        book.CheckCompatibility.should be_true
        Book.unobtrusively(@simple_file, :visible => true, :check_compatibility => false) do |book|
          book.CheckCompatibility.should be_false
        end
        Book.unobtrusively(@simple_file, :visible => true, :check_compatibility => true) do |book|
          book.CheckCompatibility.should be_true
        end

      end

    end

  end
end
