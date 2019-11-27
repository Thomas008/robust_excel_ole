# -*- coding: utf-8 -*-

require File.join(File.dirname(__FILE__), './../spec_helper')


$VERBOSE = nil

include RobustExcelOle
include General

describe Workbook do

  before(:all) do
    excel = Excel.new(:reuse => true)
    open_books = excel == nil ? 0 : excel.Workbooks.Count
    puts "*** open books *** : #{open_books}" if open_books > 0
    Excel.kill_all
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
    @simple_file1 = @simple_file
  end

  after do
    Excel.kill_all
    rm_tmp(@dir)
  end

  describe "create file" do
    context "with standard" do
      it "open an existing file" do
        expect {
          @book = Workbook.new(@simple_file)
        }.to_not raise_error
        @book.should be_a Workbook
        @book.close
      end
    end
  end

  describe "for_this_workbook" do

    before do
      @book = Workbook.open(@simple_file)
    end

    it "should set options" do
      @book.for_this_workbook(:visible => true)
      @book.excel.Visible.should be true
      @book.Windows(@book.Name).Visible.should be true
      @book.visible.should be true
      @book.ReadOnly.should be false
      @book.CheckCompatibility.should be false
      @book.for_this_workbook(:visible => false)
      @book.excel.Visible.should be true
      @book.Windows(@book.Name).Visible.should be false
      @book.visible.should be false
      @book.ReadOnly.should be false
      @book.CheckCompatibility.should be false
      @book.for_this_workbook(:read_only => true)
      @book.excel.Visible.should be true
      @book.Windows(@book.Name).Visible.should be false
      @book.visible.should be false
      @book.ReadOnly.should be true
      @book.CheckCompatibility.should be false
      @book.for_this_workbook(:visible => true)
      @book.excel.Visible.should be true
      @book.Windows(@book.Name).Visible.should be true
      @book.visible.should be true
      @book.ReadOnly.should be true
      @book.CheckCompatibility.should be false
      @book.for_this_workbook(:check_compatibility => true)
      @book.excel.Visible.should be true
      @book.Windows(@book.Name).Visible.should be true
      @book.visible.should be true
      @book.ReadOnly.should be true
      @book.CheckCompatibility.should be true
      @book.for_this_workbook(:visible => false, :check_compatibility => false)
      @book.excel.Visible.should be true
      @book.Windows(@book.Name).Visible.should be false
      @book.visible.should be false
      @book.ReadOnly.should be true
      @book.CheckCompatibility.should be false
      @book.excel.Visible.should be true
      @book.Windows(@book.Name).Visible.should be false
      @book.visible.should be false
      @book.ReadOnly.should be true
      @book.CheckCompatibility.should be false
      #@book.for_this_workbook(:calculation => true)
      #@book.excel.calculation.should be true    
    end

    it "should set options" do
      @book.for_this_workbook(:read_only => true, :check_compatibility => true)
      @book.CheckCompatibility.should be true
    end

  end

  describe "excel_of" do

    before do
      @book = Workbook.open(@simple_file)
    end

    after do
      @book.close
    end

    it "should access the excel" do
      workbook = @book.ole_workbook
      excel = Workbook.excel_of(workbook)
      excel.should be_a Excel
      excel.should == @book.excel
    end

    it "should access the Excel of the ole_excel" do
      excel = Workbook.excel_of(@book.excel.ole_excel)
      excel.should be_a Excel
      excel.should == @book.excel
    end

    it "should access the Excel of a Workbook" do
      excel = Workbook.excel_of(@book)
      excel.should be_a Excel
      excel.should == @book.excel
    end

  end

  describe "with retain_saved" do

    before do
      @book = Workbook.open(@simple_file)
    end

    after do
      @book.close(:if_unsaved => :forget)
    end

    it "should keep the save state 'saved' with empty assignments" do
      @book.Saved.should be true
      @book.retain_saved do
      end
      @book.Saved.should be true
    end

    it "should keep the save state 'saved' with non-affecting assignments" do
      @book.Saved.should be true
      @book.retain_saved do
        sheet = @book.sheet(1)
        a = sheet[1,1]
        b = @book.visible
      end
      @book.Saved.should be true
    end

    it "should keep the save state 'unsaved'" do
      sheet = @book.sheet(1)
      sheet[1,1] = sheet[1,1].Value == "foo" ? "bar" : "foo"
      @book.Saved.should be false
      @book.retain_saved do
        sheet = @book.sheet(1)
        a = sheet[1,1]
        b = @book.visible
      end
      @book.Saved.should be false
    end

    it "should keep the save state 'saved'" do
      @book.Saved.should be true
      @book.retain_saved do
        sheet = @book.sheet(1)
        sheet[1,1] = sheet[1,1].Value == "foo" ? "bar" : "foo"
        @book.Saved.should be false
      end
      @book.Saved.should be true
    end

    it "should keep the save state 'unsaved' even when the workbook was saved before" do
      sheet = @book.sheet(1)
      sheet[1,1] = sheet[1,1].Value == "foo" ? "bar" : "foo"
      @book.Saved.should be false
      @book.retain_saved do
        @book.save
        @book.Saved.should be true
      end
      @book.Saved.should be false
    end
  end

=begin
  describe "default-visible" do

    it "should keep the visibility of the open workbook" do
      book1 = Workbook.open(@simple_file1)
      book1.excel.Visible.should be false
      book1.Windows(book1.Name).Visible.should be true
      book1.visible.should be false
      book2 = Workbook.open(@simple_file1, :default => {:visible => true})
      book2.visible.should be false      
      book2.excel.Visible.should be false
      book2.Windows(book2.Name).Visible.should be true
      book1.visible.should be false
      book2 = Workbook.open(@simple_file1, :default => {:visible => false})
      book2.visible.should be false      
      book2.excel.Visible.should be false
      book2.Windows(book2.Name).Visible.should be true
      book1.visible.should be false      
    end

    it "should keep the visibility of the open workbook per default" do
      book1 = Workbook.open(@simple_file1)
      book1.excel.Visible.should be false
      book1.Windows(book1.Name).Visible.should be true
      book1.visible.should be false
      book2 = Workbook.open(@simple_file1)
      book2.visible.should be false      
      book2.excel.Visible.should be false
      book2.Windows(book2.Name).Visible.should be true
      book1.visible.should be false
    end

    it "should keep the found Excel instance invisible" do
      book1 = Workbook.open(@simple_file1)
      excel1 = book1.excel
      excel1.Visible.should be false
      book1.close
      book2 = Workbook.open(@simple_file1, :default => {:visible => true})
      excel2 = book2.excel
      excel2.should == excel1
      excel2.Visible.should be false
      book2.Windows(book2.Name).Visible.should be true
    end

    it "should keep the found Excel instance invisible with default invisible" do
      book1 = Workbook.open(@simple_file1)
      excel1 = book1.excel
      excel1.Visible.should be false
      book1.close
      book2 = Workbook.open(@simple_file1, :default => {:visible => false})
      excel2 = book1.excel
      excel2.should == excel1
      excel2.Visible.should be false
      book2.Windows(book2.Name).Visible.should be false
    end

    it "should keep the found Excel instance visible" do
      book1 = Workbook.open(@simple_file1, :visible => true)
      excel1 = book1.excel
      book1.Windows(book1.Name).Visible.should be true
      excel1.Visible.should be true
      book1.close
      book2 = Workbook.open(@simple_file1, :default => {:visible => false})
      excel2 = book1.excel
      excel2.should == excel1
      excel2.Visible.should be true
      book2.Windows(book2.Name).Visible.should be false
    end

    it "should keep the found Excel instance visible with default visible true" do
      book1 = Workbook.open(@simple_file1, :visible => true)
      excel1 = book1.excel
      book1.Windows(book1.Name).Visible.should be true
      excel1.Visible.should be true
      book1.close
      book2 = Workbook.open(@simple_file1, :default => {:visible => true})
      excel2 = book1.excel
      excel2.should == excel1
      excel2.Visible.should be true
      book2.Windows(book2.Name).Visible.should be true
    end

    it "should keep the found Excel instance invisible per default" do
      book1 = Workbook.open(@simple_file1)
      excel1 = book1.excel
      excel1.Visible.should be false
      book1.close
      book2 = Workbook.open(@simple_file1)
      excel2 = book1.excel
      excel2.should == excel1
      excel2.Visible.should be false
      book2.Windows(book2.Name).Visible.should be true      
    end

    it "should open the workbook visible if the workbook is new" do
      book1 = Workbook.open(@simple_file1, :default => {:visible => true})
      book1.excel.Visible.should be true
      book1.Windows(book1.Name).Visible.should be true
      book1.visible.should be true      
    end

    it "should open the workbook invisible if the workbook is new" do
      book1 = Workbook.open(@simple_file1, :default => {:visible => false})
      book1.excel.Visible.should be false
      book1.Windows(book1.Name).Visible.should be false
      book1.visible.should be false      
    end

    it "should open the workbook invisible per default if the workbook is new" do
      book1 = Workbook.open(@simple_file1)
      book1.excel.Visible.should be false
      book1.Windows(book1.Name).Visible.should be true
      book1.visible.should be false      
    end

    it "should open the workbook visible if the old Excel is closed" do
      book1 = Workbook.open(@simple_file1)
      book1.visible.should be false
      excel1 = book1.excel
      excel1.Visible.should be false
      book1.Windows(book1.Name).Visible.should be true
      book1.close
      excel1.close
      book2 = Workbook.open(@simple_file1, :default => {:visible => true})
      excel2 = book2.excel
      book1.Windows(book1.Name).Visible.should be true
      excel2.Visible.should be true  
      book2.visible.should be true
    end

    it "should open the workbook invisible if the old Excel is closed" do
      book1 = Workbook.open(@simple_file1, :default => {:visible => true})
      book1.visible.should be true
      excel1 = book1.excel
      excel1.Visible.should be true
      book1.Windows(book1.Name).Visible.should be true
      book1.close
      excel1.close
      book2 = Workbook.open(@simple_file1, :default => {:visible => false})
      excel2 = book2.excel
      excel2.Visible.should be false
      book1.Windows(book1.Name).Visible.should be false
      book2.visible.should be false      
    end

  end

=end

  describe "force-visible" do

    it "should change the visibility of the workbooks" do
      book1 = Workbook.open(@simple_file)
      book1.excel.Visible.should be false
      book1.Windows(book1.Name).Visible.should be true
      book1.visible.should be false
      book2 = Workbook.open(@simple_file, :visible => true)
      book2.visible.should be true      
      book2.excel.Visible.should be true
      book2.Windows(book2.Name).Visible.should be true
      book1.visible.should be true
      book2 = Workbook.open(@simple_file, :visible => false)
      book2.visible.should be false      
      book2.excel.Visible.should be true
      book2.Windows(book2.Name).Visible.should be false
      book1.visible.should be false 
      book1.Windows(book2.Name).Visible.should be false     
    end
  end

  describe "with visible" do

    it "should adapt its default value at the visible value of the Excel" do
      excel1 = Excel.create
      excel1.visible = true
      book1 = Workbook.open(@simple_file)
      excel1.Visible.should be true
      excel1.visible.should be true
      book1.visible.should be true
    end

    it "should preserve :visible if it is not set" do
      book1 = Workbook.open(@simple_file)
      book1.excel.Visible.should be false
      book1.Windows(book1.Name).Visible.should be true
      book1.visible.should be false
    end

    it "should set :visible to false" do
      book1 = Workbook.open(@simple_file, :visible => false)
      book1.excel.Visible.should be false
      book1.Windows(book1.Name).Visible.should be false
      book1.visible.should be false
    end

    it "should set :visible to true" do
      book1 = Workbook.open(@simple_file, :visible => true)
      book1.excel.Visible.should be true
      book1.Windows(book1.Name).Visible.should be true
      book1.visible.should be true
    end

    it "should preserve :visible if they are set to visible" do
      excel1 = Excel.create(:visible => true)
      book1 = Workbook.open(@simple_file)
      book1.excel.Visible.should be true
      book1.Windows(book1.Name).Visible.should be true
      book1.visible.should be true
    end

    it "should preserve :visible" do
      excel1 = Excel.create
      book1 = Workbook.open(@simple_file)
      book1.excel.Visible.should be false
      book1.Windows(book1.Name).Visible.should be true
      book1.visible.should be false
    end


    it "should preserve :visible if it is set to false" do
      excel1 = Excel.create
      book1 = Workbook.open(@simple_file, :visible => false)
      book1.excel.Visible.should be false
      book1.Windows(book1.Name).Visible.should be false
      book1.visible.should be false
    end

    it "should preserve :visible if it is not set" do
      excel1 = Excel.create
      book1 = Workbook.open(@simple_file)
      book1.excel.Visible.should be false
      book1.Windows(book1.Name).Visible.should be true
      book1.visible.should be false
    end

    it "should overwrite :visible to false" do
      excel1 = Excel.create(:visible => true)
      book1 = Workbook.open(@simple_file, :visible => false)
      book1.excel.Visible.should be true
      book1.Windows(book1.Name).Visible.should be false
      book1.visible.should be false
    end

    it "should overwrite :visible to true" do
      excel1 = Excel.create(:visible => false)
      book1 = Workbook.open(@simple_file, :visible => true)
      book1.excel.Visible.should be true
      book1.Windows(book1.Name).Visible.should be true
      book1.visible.should be true
    end

    it "should preserve :visible if it is not set with default_excel" do
      excel1 = Excel.create(:visible => true)
      book1 = Workbook.open(@simple_file)
      book2 = Workbook.open(@different_file, :default_excel => :new)
      book2.excel.Visible.should be false
      book2.Windows(book2.Name).Visible.should be true
      book2.visible.should be false
    end

    it "should set :visible to true with default_excel" do
      excel1 = Excel.create(:visible => true)
      book1 = Workbook.open(@simple_file)
      book2 = Workbook.open(@different_file, :default_excel => :new, :visible => true)
      book2.excel.Visible.should be true
      book2.Windows(book2.Name).Visible.should be true
      book2.visible.should be true
    end

    it "should set :visible to false with default_excel" do
      excel1 = Excel.create(:visible => true)
      book1 = Workbook.open(@simple_file)
      book2 = Workbook.open(@different_file, :default_excel => :new, :visible => false)
      book2.excel.Visible.should be false
      book2.Windows(book2.Name).Visible.should be false
      book2.visible.should be false
    end

    it "should preserve :visible if it is set to true with default_excel" do
      excel1 = Excel.create(:visible => true)
      excel2 = Excel.create(:visible => true)
      book1 = Workbook.open(@different_file, :default_excel => excel2)
      book1.excel.Visible.should be true
      book1.Windows(book1.Name).Visible.should be true
      book1.visible.should be true
    end

    it "should overwrite :visible to false with default_excel" do
      excel1 = Excel.create(:visible => true)
      excel2 = Excel.create(:visible => true)
      book1 = Workbook.open(@different_file, :default_excel => excel2, :visible => false)
      book1.excel.Visible.should be true
      book1.Windows(book1.Name).Visible.should be false
      book1.visible.should be false
    end

    it "should preserve :visible if it is not set with force_excel => new" do
      excel1 = Excel.create(:visible => true)
      book1 = Workbook.open(@different_file, :force_excel => :new)
      book1.excel.Visible.should be false
      book1.Windows(book1.Name).Visible.should be true
      book1.visible.should be false
    end

    it "should set :visible to true with force_excel" do
      excel1 = Excel.create(:visible => true)
      book1 = Workbook.open(@different_file, :force_excel => :new, :visible => true)
      book1.excel.Visible.should be true
      book1.Windows(book1.Name).Visible.should be true
      book1.visible.should be true
    end

    it "should preserve :visible if it is not set with force_excel => excel" do
      excel1 = Excel.create(:visible => true)
      excel2 = Excel.create(:visible => true)
      book1 = Workbook.open(@different_file, :force_excel => excel2)
      book1.excel.Visible.should be true
      book1.Windows(book1.Name).Visible.should be true
      book1.visible.should be true
    end

    it "should set visible to false with force_excel => excel" do
      excel1 = Excel.create(:visible => true)
      excel2 = Excel.create(:visible => true)
      book1 = Workbook.open(@different_file, :force_excel => excel2, :visible => false)
      book1.excel.Visible.should be true
      book1.Windows(book1.Name).Visible.should be false
      book1.visible.should be false
    end

    it "should set visible to true with force_excel => excel" do
      excel1 = Excel.create(:visible => true)
      excel2 = Excel.create(:visible => true)
      book1 = Workbook.open(@different_file, :force_excel => excel2, :visible => true)
      book1.excel.Visible.should be true
      book1.Windows(book1.Name).Visible.should be true
      book1.visible.should be true
    end

    it "should preserve :visible if it is set to true with force_excel => current" do
      excel1 = Excel.create(:visible => true)
      book1 = Workbook.open(@different_file, :force_excel => :current)      
      book1.excel.Visible.should be true
      book1.Windows(book1.Name).Visible.should be true
      book1.visible.should be true
    end

    it "should set :visible to false with force_excel => current" do
      excel1 = Excel.create(:visible => true)
      book1 = Workbook.open(@different_file, :force_excel => :current, :visible => false)
      book1.excel.Visible.should be true
      book1.Windows(book1.Name).Visible.should be false
      book1.visible.should be false
    end

    it "should preserve :visible if it is set to false with force_excel => current" do
      excel1 = Excel.create(:visible => false)
      book1 = Workbook.open(@simple_file, :force_excel => :current)
      book1.excel.Visible.should be false
      book1.Windows(book1.Name).Visible.should be true
      book1.visible.should be false
    end

    it "should set :visible to false with force_excel => current" do
      excel1 = Excel.create(:visible => false)
      book1 = Workbook.open(@simple_file, :force_excel => :current, :visible => true)
      book1.excel.Visible.should be true
      book1.Windows(book1.Name).Visible.should be true
      book1.visible.should be true
    end

    it "should let an open Workbook open" do
      @book = Workbook.open(@simple_file1, :visible => true)
      Workbook.unobtrusively(@simple_file1) do |book|
        book.should be_a Workbook
        book.should be_alive
        book.excel.should == @book.excel
        book.excel.Visible.should be true
      end        
      @book.should be_alive
      @book.should be_a Workbook
      @book.excel.Visible.should be true
      @book.close(:if_unsaved => :forget)
      @book2.close(:if_unsaved => :forget) rescue nil
    end

    it "should set visible and displayalerts if displayalerts => :if_visible" do
      book1 = Workbook.open(@simple_file)
      book1.excel.Visible.should be false
      book1.excel.displayalerts.should == :if_visible
      book1.Windows(book1.Name).Visible.should be true
      book1.visible.should be false
      book2 = Workbook.open(@different_file)
      book2.excel.Visible.should be false
      book2.Windows(book2.Name).Visible.should be true
      book2.visible.should be false
      book2.excel.visible.should be false
      book2.excel.displayalerts.should == :if_visible
      book2.excel.DisplayAlerts.should be false
    end

    it "should set visible and displayalerts if displayalerts => :if_visible" do
      book1 = Workbook.open(@simple_file)
      book1.excel.Visible.should be false
      book1.excel.displayalerts.should == :if_visible
      book1.Windows(book1.Name).Visible.should be true
      book1.visible.should be false
      book2 = Workbook.open(@different_file, :visible => true)
      book2.excel.Visible.should be true
      book2.Windows(book2.Name).Visible.should be true
      book2.visible.should be true
      book2.excel.visible.should be true
      book2.excel.displayalerts.should == :if_visible
      book2.excel.DisplayAlerts.should be true
    end
  end

  describe "send methods to workbook" do

    context "with standard" do
      before do
        @book = Workbook.open(@simple_file)
      end

      after do
        @book.close
      end

      it "should send Saved to workbook" do
        @book.Saved.should be true
      end

      it "should send Fullname to workbook" do
        @book.Fullname.tr('\\','/').should == @simple_file
      end

      it "should raise an error for unknown methods or properties" do
        expect{
          @book.Foo
        }.to raise_error #(VBAMethodMissingError, /unknown VBA property or method :Foo/)
      end

      it "should report that workbook is not alive" do
        @book.close
        expect{ @book.Nonexisting_method }.to raise_error(ObjectNotAlive, "method missing: workbook not alive")
      end
    end

  end

  describe "hidden_excel" do
    
    context "with some open book" do

      before do
        @book = Workbook.open(@simple_file1)
      end

      after do
        @book.close
      end

      it "should create and use a hidden Excel instance" do
        book2 = Workbook.open(@simple_file1, :force_excel => @book.bookstore.hidden_excel)
        book2.excel.should_not == @book.excel
        book2.excel.visible.should be false
        book2.excel.displayalerts.should == :if_visible
        book2.close 
      end
    end
  end

  describe "nameval, set_nameval, [], []=" do
  
    before do
      @book1 = Workbook.open(@another_simple_file)
    end

    after do
      @book1.close(:if_unsaved => :forget)
    end   

    it "should return value of a range" do
      @book1.namevalue_glob("new").should == "foo"
      @book1.namevalue_glob("one").should == 1
      @book1.namevalue_glob("firstrow").should == [[1,2]]        
      @book1.namevalue_glob("four").should == [[1,2],[3,4]]
      @book1.namevalue_glob("firstrow").should_not == "12"
      @book1.namevalue_glob("firstcell").should == "foo"        
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
      @book1.set_namevalue_glob("new", "bar")
      @book1.namevalue_glob("new").should == "bar"
    end

    it "should set value of a range via []=" do
      @book1["new"] = "bar"
      @book1.namevalue_glob("new").should == "bar"
    end

    #it "should evaluate a formula" do
    #  @book1.namevalue_glob("named_formula").should == 4      
    #end

    #it "should evaluate a formula via []" do
    #  @book1["named_formula"].should == 4      
    #end

    it "should raise an error if name not defined and default value is not provided" do
      expect {
        @book1.namevalue_glob("foo", :default => nil)
      }.to_not raise_error
      expect {
        @book1.namevalue_glob("foo", :default => :__not_provided)
      }.to raise_error(NameNotFound, /name "foo" not in #<Workbook: another_workbook/)
      expect {
        @book1.namevalue_glob("foo")
      }.to raise_error(NameNotFound, /name "foo" not in #<Workbook: another_workbook/)
      @book1.namevalue_glob("foo", :default => nil).should be_nil
      @book1.namevalue_glob("foo", :default => 1).should == 1
      expect {
          @book1.set_namevalue_glob("foo","bar")
      }.to raise_error(NameNotFound, /name "foo" not in #<Workbook: another_workbook/)
      expect {
          @book1["foo"] = "bar"
      }.to raise_error(NameNotFound, /name "foo" not in #<Workbook: another_workbook/)
      @book1.namevalue_glob("empty", :default => 1).should be_nil
    end    

    it "should raise an error if name was defined but contents is calcuated" do
      expect {
        @book1.set_namevalue_glob("named_formula","bar")
      }.to raise_error(RangeNotEvaluatable, /cannot assign value to range named "named_formula" in #<Workbook: another_workbook/)
      expect {
        @book1["named_formula"] = "bar"
      }.to raise_error(RangeNotEvaluatable, /cannot assign value to range named "named_formula" in #<Workbook: another_workbook/)
    end

    # Excel Bug: for local names without uqifier: takes the first sheet as default even if another sheet is activated
    it "should take the first sheet as default even if the second sheet is activated" do
      @book1.namevalue_glob("Sheet1!localname").should == "bar"
      @book1.namevalue_glob("Sheet2!localname").should == "simple"
      @book1.namevalue_glob("localname").should == "bar"
      @book1.Worksheets.Item(2).Activate
      @book1.namevalue_glob("localname").should == "bar"
      @book1.Worksheets.Item(1).Delete
      @book1.namevalue_glob("localname").should == "simple"
    end

    it "should color the cell" do
      @book1.set_namevalue_glob("new", "bar")
      @book1.Names.Item("new").RefersToRange.Interior.ColorIndex.should == -4142
      @book1.set_namevalue_glob("new", "bar", :color => 4)
      @book1.Names.Item("new").RefersToRange.Interior.ColorIndex.should == 4
      @book1["new"].should == "bar"
      @book1["new"] = "bar"
      @book1.Names.Item("new").RefersToRange.Interior.ColorIndex.should == 42
      @book1.save
      @book1.close
      #book2 = Workbook.open(@simple_file1, :visible =>  true)
      #book2.Names.Item("new").RefersToRange.Interior.ColorIndex.should == 42
    end

    it "should save without color" do
      @book1.set_namevalue_glob("new", "bar", :color => 4)
      @book1.Names.Item("new").RefersToRange.Interior.ColorIndex.should == 4
      @book1.save(:discoloring => true)
      @book1.close
      #book2 = Workbook.open(@simple_file1, :visible =>  true)
      #book2.Names.Item("new").RefersToRange.Interior.ColorIndex.should == 0
    end

  end

  describe "rename_range" do
    
    before do
      @book1 = Workbook.open(@another_simple_file)
    end

    after do
      @book1.close(:if_unsaved => :forget)
    end

    it "should rename a range" do
      @book1.rename_range("four","five")
      @book1.namevalue_glob("five").should == [[1,2],[3,4]]
      expect {
        @book1.rename_range("four","five")
      }.to raise_error(NameNotFound, /name "four" not in "another_workbook.xls"/)
    end
  end

  describe "alive?, filename, ==, visible, focus, saved, check_compatibility" do

    context "with alive?" do

      before do
        @book = Workbook.open(@simple_file)
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
        @book = Workbook.open(@simple_file)
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
        @book = Workbook.open(@simple_file1)
      end

      after do
        @book.close
        @new_book.close rescue nil
      end

      it "should be true with two identical books" do
        @new_book = Workbook.open(@simple_file1)
        @new_book.should == @book
      end

      it "should be false with two different books" do
        @new_book = Workbook.new(@different_file)
        @new_book.should_not == @book
      end

      it "should be false with same book names but different paths" do       
        @new_book = Workbook.new(@simple_file_other_path, :force_excel => :new)
        @new_book.should_not == @book
      end

      it "should be false with same book names but different excel instances" do
        @new_book = Workbook.new(@simple_file, :force_excel => :new)
        @new_book.should_not == @book
      end

      it "should be false with non-Workbooks" do
        @book.should_not == "hallo"
        @book.should_not == 7
        @book.should_not == nil
      end
    end

    context "with saved" do

      before do
        @book = Workbook.open(@simple_file)
      end

      after do
        @book.close(:if_unsaved => :forget)
      end

      it "should yield true for a saved book" do
        @book.saved.should be true
      end

      it "should yield false for an unsaved book" do
        sheet = @book.sheet(1)
        sheet[1,1] = sheet[1,1].Value == "foo" ? "bar" : "foo"
        @book.saved.should be false
      end
    end


    context "with :visible => " do

      it "should leave the excel invisible when opening with default option" do
        excel1 = Excel.new(:reuse => false, :visible => false)
        book1 = Workbook.open(@simple_file)
        excel1.Visible.should be false
        book1.Windows(book1.Name).Visible.should be true
        book1.visible.should be false
      end

      it "should leave the excel invisible when opening with :visible => false" do
        excel1 = Excel.new(:reuse => false, :visible => false)
        book1 = Workbook.open(@simple_file, :visible => false)
        excel1.Visible.should be false
        book1.Windows(book1.Name).Visible.should be false
        book1.visible.should be false
      end

      it "should leave the excel visible" do
        excel1 = Excel.new(:reuse => false, :visible => false)
        book1 = Workbook.open(@simple_file, :visible => true)
        excel1.Visible.should be true
        book1.Windows(book1.Name).Visible.should be true
        book1.visible.should be true
        book2 = Workbook.open(@another_simple_file)
        excel1.Visible.should be true
        book2.Windows(book2.Name).Visible.should be true
        book2.visible.should be true
        book3 = Workbook.open(@different_file, :visible => false)
        excel1.Visible.should be true
        book3.Windows(book3.Name).Visible.should be false
        book3.visible.should be false
      end

      it "should leave the excel visible when opening with default option" do
        excel1 = Excel.new(:reuse => false, :visible => true)
        book1 = Workbook.open(@simple_file)
        excel1.Visible.should be true
        book1.Windows(book1.Name).Visible.should be true
        book1.visible.should be true
      end

      it "should leave the excel visible when opening with :visible => false" do
        excel1 = Excel.new(:reuse => false, :visible => true)
        book1 = Workbook.open(@simple_file, :visible => false)
        excel1.Visible.should be true
        book1.Windows(book1.Name).Visible.should be false
        book1.visible.should be false
        book2 = Workbook.open(@another_simple_file)
        excel1.Visible.should be true
        book2.Windows(book2.Name).Visible.should be true
        book2.visible.should be true
      end

      it "should leave the excel visible" do
        excel1 = Excel.new(:reuse => false, :visible => true)
        book1 = Workbook.open(@simple_file, :visible => true)
        excel1.Visible.should be true
        book1.Windows(book1.Name).Visible.should be true
        book1.visible.should be true
        book2 = Workbook.open(@different_file, :visible => false)
        excel1.Visible.should be true
        book2.Windows(book2.Name).Visible.should be false
        book2.visible.should be false
      end

      it "should leave the visibility of Excel" do
        excel1 = Excel.new(:reuse => false, :visible => false)
        book1 = Workbook.open(@simple_file, :visible => true)
        excel1.Visible.should be true
        book1.Windows(book1.Name).Visible.should be true
        book1.visible.should be true
        excel1.visible = false
        book2 = Workbook.open(@different_file)
        excel1.Visible.should be false
        book2.Windows(book2.Name).Visible.should be true
        book2.visible.should be false
      end

      it "should leave the visibility of Excel" do
        excel1 = Excel.new(:reuse => false, :visible => false)
        book1 = Workbook.open(@simple_file, :visible => false)
        excel1.Visible.should be false
        book1.Windows(book1.Name).Visible.should be false
        book1.visible.should be false
        excel1.visible = true
        book2 = Workbook.open(@different_file)
        excel1.Visible.should be true
        book2.Windows(book2.Name).Visible.should be true
        book2.visible.should be true
      end
    end

    context "with visible, visible=" do

      before do
        @book1 = Workbook.open(@simple_file)
        @book2 = Workbook.open(@different_file, :force_excel => :new, :visible => true)
      end

      after do
        @book1.close
        @book2.close
      end

      it "should make the invisible workbook visible and invisible" do
        @book1.excel.Visible.should be false
        @book1.Windows(@book1.Name).Visible.should be true
        @book1.visible.should be false
        @book1.visible = true
        @book1.Saved.should be true
        @book1.excel.Visible.should be true
        @book1.Windows(@book1.Name).Visible.should be true
        @book1.visible.should be true
        @book1.visible = false
        @book1.Saved.should be true
        @book1.excel.Visible.should be true
        @book1.Windows(@book1.Name).Visible.should be false
        @book1.visible.should be false
        @book2.excel.Visible.should be true
      end

      it "should make the visible workbook and the invisible workbook invisible" do
        @book2.Windows(@book2.Name).Visible.should be true
        @book2.visible.should be true
        @book2.visible = true
        @book2.Saved.should be true
        @book2.excel.Visible.should be true
        @book2.Windows(@book2.Name).Visible.should be true
        @book2.excel.visible = false
        @book2.visible = false
        @book2.Saved.should be true
        @book2.excel.Visible.should be false
        @book2.Windows(@book2.Name).Visible.should be false
        @book2.visible.should be false
      end

    end

    context "with focus" do

      before do
        @key_sender = IO.popen  'ruby "' + File.join(File.dirname(__FILE__), '../helpers/key_sender.rb') + '" "Microsoft Excel" '  , "w"        
        @book = Workbook.open(@simple_file, :visible => true)
        @book.excel.displayalerts = false
        @book2 = Workbook.open(@another_simple_file, :visible => true)
        @book2.excel.displayalerts = false
      end

      after do
        @book.close(:if_unsaved => :forget)
        @book2.close(:if_unsaved => :forget)
        @key_sender.close
      end

      it "should bring a book to focus" do
        sheet = @book.sheet(2)
        sheet.Activate
        sheet[2,3].Activate
        sheet2 = @book2.sheet(2)
        sheet2.Activate
        sheet2[3,2].Activate
        Excel.current.should == @book.excel
        @book2.focus
        @key_sender.puts "{a}{enter}"
        sleep 1
        #sheet2[3,2].Value.should == "a"
        #Excel.current.should == @book2.excel
        @book.focus
        @key_sender.puts "{a}{enter}"
        sleep 1
        #sheet[2,3].Value.should == "a"
        Excel.current.should == @book.excel
      end
    end

    context "range" do

      before do
        @book1 = Workbook.open(@dir + '/another_workbook.xls', :visible => true)
      end

      it "should create a range from the name" do
        @book1.add_name("foo",[1..3,1..4])
        @book1.save
        range = @book1.range("foo")
        range.Address.should == "$A$1:$D$3"
      end

    end

    context "adding and deleting the name of a range" do

       before do
        @book1 = Workbook.open(@dir + '/another_workbook.xls', :read_only => true, :visible => true)
        @book1.excel.displayalerts = false
      end

      after do
        @book1.close
      end   

      it "should name an unnamed range with a giving address" do
        @book1.add_name("foo",[1,2])
        @book1.Names.Item("foo").Name.should == "foo"
        @book1.Names.Item("foo").Value.should == "=Sheet1!$B$1:$B$1"
        @book1.add_name("foo",[1..2,2..4])
        @book1.Names.Item("foo").Name.should == "foo"
        @book1.Names.Item("foo").Value.should == "=Sheet1!$B$1:$D$2"
        @book1.add_name("foo","B1")
        @book1.Names.Item("foo").Name.should == "foo"
        @book1.Names.Item("foo").Value.should == "=Sheet1!$B$1:$B$1"
        @book1.add_name("foo","B1:D2")
        @book1.Names.Item("foo").Name.should == "foo"
        @book1.Names.Item("foo").Value.should == "=Sheet1!$B$1:$D$2"
        @book1.add_name("foo","Z1S2")
        @book1.Names.Item("foo").Name.should == "foo"
        @book1.Names.Item("foo").Value.should == "=Sheet1!$B$1"
        @book1.add_name("foo","Z1S2:Z2S4")
        @book1.Names.Item("foo").Name.should == "foo"
        @book1.Names.Item("foo").Value.should == "=Sheet1!$B$1:$D$2"
        @book1.add_name("foo","A:B")
        @book1.Names.Item("foo").Name.should == "foo"
        @book1.Names.Item("foo").Value.should == "=Sheet1!$A:$B"
        @book1.add_name("foo","1:2")
        @book1.Names.Item("foo").Name.should == "foo"
        @book1.Names.Item("foo").Value.should == "=Sheet1!$1:$2"
        @book1.add_name("foo",[1..2,nil])
        @book1.Names.Item("foo").Name.should == "foo"
        @book1.Names.Item("foo").Value.should == "=Sheet1!$1:$2"
        @book1.add_name("foo",[nil,1..2])
        @book1.Names.Item("foo").Name.should == "foo"
        @book1.Names.Item("foo").Value.should == "=Sheet1!$A:$B"
        @book1.add_name("foo",[nil,"A".."B"])
        @book1.Names.Item("foo").Name.should == "foo"
        @book1.Names.Item("foo").Value.should == "=Sheet1!$A:$B"
      end

      it "should raise an error" do
        expect{
          @book1.add_name("foo", [-2, 1])
        }.to raise_error(RangeNotEvaluatable, /cannot add name "foo" to range/)
      end

      it "should delete a name of a range" do
        @book1.add_name("foo",[1,1])
        @book1.delete_name("foo")
        expect{
          @book1.namevalue_glob("foo")
        }.to raise_error(NameNotFound, /name "foo"/)
      end

      it "should add a name of a rectangular range" do
        @book1.add_name("foo",[1..3,1..4])
        @book1["foo"].should == [["foo", "workbook", "sheet1", nil], ["foo", 1.0, 2.0, 4.0], ["matz", 3.0, 4.0, 4.0]] 
      end

       it "should accept the old interface" do
        @book1.add_name("foo",1..3,1..4)
        @book1["foo"].should == [["foo", "workbook", "sheet1", nil], ["foo", 1.0, 2.0, 4.0], ["matz", 3.0, 4.0, 4.0]] 
      end

      it "should add a name of an infinite row range" do
        @book1.add_name("foo",[1..3, nil])
        @book1.Names.Item("foo").Value.should == "=Sheet1!$1:$3"
      end

      it "should add a name of an infinite column range" do
        @book1.add_name("foo",[nil, "A".."C"])
        @book1.Names.Item("foo").Value.should == "=Sheet1!$A:$C"
      end

    end
  
    context "with compatibility" do      

      it "should open with checking compatibility" do
        book = Workbook.open(@simple_file, :visible => true, :check_compatibility => true)
        book.CheckCompatibility.should be true
      end

      it "should open without checking compatibility" do
        book = Workbook.open(@simple_file, :visible => true, :check_compatibility => false)
        book.CheckCompatibility.should be false
      end
    end
  end
end
