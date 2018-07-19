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
    @different_file1 = @different_file
    @simple_file_other_path1 = @simple_file_other_path
  end

  after do
    Excel.kill_all
    #rm_tmp(@dir)
  end

  describe "with already open Excel instances and an open unsaved workbook" do

    before do
      @ole_excel1 = WIN32OLE.new('Excel.Application')
      @ole_excel2 = WIN32OLE.new('Excel.Application')
      @ole_workbook1 = @ole_excel1.Workbooks.Open(@simple_file1, { 'ReadOnly' => false })
      @ole_workbook1.Worksheets.Add
    end

    context "with simple general situations" do
      
      it "should simply open" do
        book = Workbook.open(@simple_file1, :if_unsaved => :accept)
        book.should be_alive
        book.should be_a Book
      end

      it "should open in a new Excel" do
        book2 = Workbook.open(@simple_file1, :force => {:excel => :new})
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should_not == @ole_excel1 
        book2.Readonly.should be true
      end

      it "should fetch the workbook" do
        new_book = Book.new(@ole_workbook1)
        new_book.should be_a Book
        new_book.should be_alive
        new_book.ole_workbook.should == @ole_workbook1
        new_book.excel.ole_excel.Hwnd.should == @ole_excel1.Hwnd
      end

      it "should fetch a closed workbook" do
        new_book = Book.new(@ole_workbook1)
        new_book.close(:if_unsaved => :forget)
        new_book.should_not be_alive
        book2 = Book.open(@simple_file1)
        book2.should === new_book
        book2.should be_alive
        book2.close
      end

      it "should force_excel with :reuse" do
        book2 = Book.open(@different_file, :force => {:excel => :current})
        book2.should be_alive
        book2.should be_a Book
        book2.excel.ole_excel.Hwnd.should == @ole_excel1.Hwnd 
      end

      it "should force_excel with :reuse when reopening and the Excel is not alive even if :default_excel says sth. else" do
        book1 = Book.open(@simple_file1, :if_unsaved => :forget)
        excel2 = Excel.new(:reuse => false)
        book1.excel.close(:if_unsaved => :forget)
        sleep 1
        book2 = Book.open(@simple_file1, :force => {:excel => :current}, :default => {:excel => :new})
        book2.should be_alive
        book2.should be_a Book
        book2.excel.ole_excel.Hwnd.should == @ole_excel2.Hwnd
      end

      it "should reopen the closed book" do
        book1 = Book.open(@simple_file1, :if_unsaved => :accept)
        book1.should be_alive
        book2 = book1
        book1.close(:if_unsaved => :forget)
        book1.should_not be_alive
        book1.reopen
        book1.should be_a Book
        book1.should be_alive
        book1.should === book2
      end
    end

    context "with :if_unsaved" do

      before do
        @book = Book.open(@simple_file1, :if_unsaved => :accept)
        sheet = @book.sheet(1)
        @old_value = sheet[1,1].Value
        sheet[1,1] = (sheet[1,1].value == "foo" ? "bar" : "foo")
        @new_value = sheet[1,1].Value
        @book.Saved.should be false
      end

      after do
        @book.close(:if_unsaved => :forget)
      end

      it "should let the book open, if :if_unsaved is :accept" do
        new_book = Book.open(@simple_file1, :if_unsaved => :accept)
        @book.should be_alive
        new_book.should be_alive
        new_book.Saved.should be false      
        @book.Saved.should be false  
        new_book.sheet(1)[1,1].Value.should == @new_value
        new_book.should == @book
      end

      it "should open book and close old book, if :if_unsaved is :forget" do
        new_book = Book.open(@simple_file1, :if_unsaved => :forget)
        @book.should_not be_alive
        new_book.should be_alive
        new_book.Saved.should be true
      end
    end

    context "with :if_obstructed" do

      it "should raise an error, if :if_obstructed is :raise" do
        expect {
          new_book = Book.open(@simple_file_other_path1)
        }.to raise_error(WorkbookBlocked, /blocked by a workbook with the same name in a different path/)
      end

      it "should close the other book and open the new book, if :if_obstructed is :forget" do
        new_book = Book.open(@simple_file_other_path1, :if_obstructed => :forget)
        expect{
          @ole_workbook1.Name
        }.to raise_error 
        new_book.should be_alive
      end

    end

    context "with :force => {:excel}" do

      it "should open in a provided Excel" do
        book1 = Book.open(@simple_file1, :force => {:excel => :new})
        book2 = Book.open(@simple_file1, :force => {:excel => :new})
        book3 = Book.open(@simple_file1, :force => {:excel => book2.excel})
        book3.should be_alive
        book3.should be_a Book
        book3.excel.should == book2.excel 
        book3.Readonly.should be true
      end 
    end
  end

  describe "simple open" do

    it "should simply open" do
      book = Workbook.open(@simple_file, :v => true, :f => {:e => :new})
    end

  end

  describe "new" do

    it "should simply create a new one" do
      book = Workbook.new(@simple_file)
      book.should be_alive
      book.should be_a Book
    end

    it "should set options" do
      book = Workbook.new(@simple_file, :visible => true, :read_only => true, :force => {:excel => :new})
      book.should be_alive
      book.should be_a Book
      book.excel.Visible.should be true
      book.Windows(book.Name).Visible.should be true
      book.ReadOnly.should be true
      book2 = Workbook.new(@different_file, :force => {:excel => :new}, :v => true)
      book2.should be_alive
      book2.should be_a Book
      book2.excel.Visible.should be true
      book2.Windows(book2.Name).Visible.should be true
      book2.ReadOnly.should be false
      book2.excel.should_not == book.excel
    end

  end

  describe "open" do

    context "with calculation mode" do

      it "should set calculation mode" do
        book1 = Book.open(@simple_file1, :visible => true)
        book1.excel.calculation = :manual
        book1.excel.Calculation.should == -4135
        book1.save
        book1.excel.close
        book2 = Book.open(@simple_file1, :visible => true)
        book2.excel.calculation = :automatic
        book2.excel.Calculation.should == -4105
        book2.save
        book2.excel.close
      end

      it "should not set the default value" do
        book1 = Book.open(@simple_file)
        book1.excel.calculation.should == nil
      end

      it "should set the calculation mode to automatic" do
        book1 = Book.open(@simple_file)
        book1.excel.calculation = :automatic
        book1.excel.calculation.should == :automatic
        book1.excel.Calculation.should == -4105
      end

      it "should set the calculation mode to manual" do
        book1 = Book.open(@simple_file)
        book1.excel.calculation = :manual
        book1.excel.calculation.should == :manual
        book1.excel.Calculation.should == -4135
      end

      it "should change the calculation mode from manual to automatic" do
        book1 = Book.open(@simple_file, :visible => true)
        excel1 = Excel.current(:calculation => :automatic)        
        book2 = Book.open(@different_file, :visible => true)
        book2.excel.Calculation.should == -4105
        book1.excel.Calculation.should == -4105
      end
    end

   
    context "with causing warning dead excel without window handle" do

      it "combined" do
        Excel.kill_all
        book1 = Book.open(@simple_file, :visible => true)
        book2 = Book.open(@different_file)
        Excel.kill_all
        #sleep 1
        book3 = Book.open(@another_simple_file)
      end

    end

    context "with class identifier 'Workbook'" do

      before do
        @book = Workbook.open(@simple_file)
      end

      after do
        @book.close rescue nil
      end

      it "should open in a new Excel" do
        book2 = Workbook.open(@simple_file, :force => {:excel => :new})
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should_not == @book.excel 
        book2.should_not == @book
        @book.Readonly.should be false
        book2.Readonly.should be true
        book2.close
      end
    end

    context "lift a workbook to a Book object" do

      before do
        @book = Book.open(@simple_file)
      end

      after do
        @book.close
      end

      it "should fetch the workbook" do
        workbook = @book.ole_workbook
        new_book = Book.new(workbook)
        new_book.should be_a Book
        new_book.should be_alive
        new_book.should == @book
        new_book.filename.should == @book.filename
        new_book.excel.should == @book.excel
        new_book.excel.Visible.should be false
        new_book.excel.DisplayAlerts.should be false
        new_book.should === @book
        new_book.close
      end

      it "should fetch the workbook" do
        workbook = @book.ole_workbook
        new_book = Book.new(workbook, :visible => true)
        new_book.should be_a Book
        new_book.should be_alive
        new_book.should == @book
        new_book.filename.should == @book.filename
        new_book.excel.should == @book.excel
        new_book.excel.Visible.should be true
        new_book.excel.DisplayAlerts.should be true
        new_book.should === @book
        new_book.close
      end

      it "should yield an identical Book and set visible value" do
        workbook = @book.ole_workbook
        new_book = Book.new(workbook, :visible => true)
        new_book.excel.displayalerts = true
        new_book.should be_a Book
        new_book.should be_alive
        new_book.should == @book
        new_book.filename.should == @book.filename
        new_book.excel.should == @book.excel
        new_book.should === @book
        new_book.excel.Visible.should be true
        new_book.excel.DisplayAlerts.should be true
        new_book.close
      end

    end

    context "with standard options" do
      before do
        @book = Book.open(@simple_file)
      end

      after do
        @book.close
      end

      it "should say that it lives" do
        @book.should be_alive
      end
    end

    context "with identity transperence" do

      before do
        @book = Book.open(@simple_file1)
      end

      after do
        @book.close
      end

      it "should yield identical Book objects for identical Excel books" do
        book2 = Book.open(@simple_file1)
        book2.should === @book
        book2.close
      end

      it "should yield different Book objects for different Excel books" do
        book2 = Book.open(@different_file)
        book2.should_not === @book
        book2.close
      end

      it "should yield different Book objects when opened the same file in different Excel instances" do
        book2 = Book.open(@simple_file, :force => {:excel => :new})
        book2.should_not === @book
        book2.close
      end

      it "should yield identical Book objects for identical Excel books when reopening" do
        @book.should be_alive
        @book.close
        @book.should_not be_alive
        book2 = Book.open(@simple_file1)
        book2.should === @book
        book2.should be_alive
        book2.close
      end

      it "should yield identical Book objects when reopening and the Excel is closed" do
        @book.should be_alive
        @book.close
        Excel.close_all
        book2 = Book.open(@simple_file1)
        book2.should be_alive
        book2.should === @book
        book2.close
      end

      it "should yield different Book objects when reopening in a new Excel" do
        @book.should be_alive
        old_excel = @book.excel
        @book.close
        @book.should_not be_alive
        book2 = Book.open(@simple_file1, :force => {:excel => :new})
        book2.should_not === @book
        book2.should be_alive
        book2.excel.should_not == old_excel
        book2.close
      end

      it "should yield different Book objects when reopening in a new given Excel instance" do
        old_excel = @book.excel
        new_excel = Excel.new(:reuse => false)
        @book.close
        @book.should_not be_alive
        book2 = Book.open(@simple_file1, :force => {:excel => new_excel})
        #@book.should be_alive
        #book2.should === @book
        book2.should be_alive
        book2.excel.should == new_excel
        book2.excel.should_not == old_excel
        book2.close
      end

      it "should yield identical Book objects when reopening in the old excel" do
        old_excel = @book.excel
        new_excel = Excel.new(:reuse => false)
        @book.close
        @book.should_not be_alive
        book2 = Book.open(@simple_file1, :force => {:excel => old_excel})
        book2.should === @book
        book2.should be_alive
        book2.excel.should == old_excel
        @book.should be_alive
        book2.close
      end

    end

    context "with abbrevations" do

      before do
        @book = Book.open(@simple_file1)
      end

      after do
        @book.close rescue nil
      end

      it "should work as force" do
        book2 = Book.open(@another_simple_file, :excel => :new)
        book2.excel.should_not == @book.excel
        book3 = Book.open(@different_file, :excel => book2.excel)
        book3.excel.should == book2.excel
      end

      it "should work with abbrevation of force and excel" do
        book2 = Book.open(@another_simple_file, :f => {:e => :new})
        book2.excel.should_not == @book.excel
        book3 = Book.open(@different_file, :f => {:e => book2.excel})
        book3.excel.should == book2.excel
      end

      it "should work with abbrevation of force" do
        book2 = Book.open(@another_simple_file, :f => {:excel => :new})
        book2.excel.should_not == @book.excel
        book3 = Book.open(@different_file, :f => {:excel => book2.excel})
        book3.excel.should == book2.excel
      end

      it "should work with abbrevation of force" do
        book2 = Book.open(@another_simple_file, :force => {:e => :new})
        book2.excel.should_not == @book.excel
        book3 = Book.open(@different_file, :force => {:e => book2.excel})
        book3.excel.should == book2.excel
      end

      it "should open in a given Excel provided as Excel, Book, or WIN32OLE representing an Excel or Workbook" do
        book2 = Book.open(@another_simple_file)
        book3 = Book.open(@different_file)
        book3 = Book.open(@simple_file1, :excel => book2.excel)
        book3.excel.should === book2.excel
        book4 = Book.open(@simple_file1, :excel => @book) 
        book4.excel.should === @book.excel
        book3.close
        book4.close
        book5 = Book.open(@simple_file1, :excel => book2.ole_workbook)
        book5.excel.should ===  book2.excel
        win32ole_excel1 = WIN32OLE.connect(@book.ole_workbook.Fullname).Application
        book6 = Book.open(@simple_file1, :excel => win32ole_excel1)
        book6.excel.should === @book.excel
      end

      it "should use abbreviations of default" do
        book2 = Book.open(@simple_file1, :d => {:excel => :current})
        book2.excel.should == @book.excel
      end

      it "should use abbreviations of default" do
        book2 = Book.open(@simple_file1, :d => {:e => :current})
        book2.excel.should == @book.excel
      end

      it "should use abbreviations of default" do
        book2 = Book.open(@simple_file1, :default => {:e => :current})
        book2.excel.should == @book.excel
      end

      it "should reopen the book in the Excel where it was opened most recently" do
        excel1 = @book.excel
        excel2 = Excel.new(:reuse => false)
        @book.close
        book2 = Book.open(@simple_file1, :d => {:e => :current})
        book2.excel.should == excel1
        book2.close
        book3 = Book.open(@simple_file1, :e => excel2)
        book3.close
        book3 = Book.open(@simple_file1, :d => {:e => :current})
        book3.excel.should == excel2
        book3.close
      end

    end

    context "with :force => {:excel}" do

      before do
        @book = Book.open(@simple_file1)
      end

      after do
        @book.close rescue nil
      end

      it "should open in a given Excel provided as Excel, Book, or WIN32OLE representing an Excel or Workbook" do
        book2 = Book.open(@another_simple_file)
        book3 = Book.open(@different_file)
        book3 = Book.open(@simple_file1, :force => {:excel => book2.excel})
        book3.excel.should === book2.excel
        book4 = Book.open(@simple_file1, :force => {:excel => @book}) 
        book4.excel.should === @book.excel
        book3.close
        book4.close
        book5 = Book.open(@simple_file1, :force => {:excel => book2.ole_workbook})
        book5.excel.should ===  book2.excel
        win32ole_excel1 = WIN32OLE.connect(@book.ole_workbook.Fullname).Application
        book6 = Book.open(@simple_file1, :force => {:excel => win32ole_excel1})
        book6.excel.should === @book.excel
      end

      it "should open in a new Excel" do
        book2 = Book.open(@simple_file1, :force => {:excel => :new})
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should_not == @book.excel 
        book2.should_not == @book
        @book.Readonly.should be false
        book2.Readonly.should be true
        book2.close
      end

      it "should open in a given Excel, not provide identity transparency, because old book readonly, new book writable" do
        book2 = Book.open(@simple_file1, :force => {:excel => :new})
        book2.excel.should_not == @book.excel
        book3 = Book.open(@simple_file1, :force => {:excel => :new})
        book3.excel.should_not == book2.excel
        book3.excel.should_not == @book.excel
        book2.close
        book4 = Book.open(@simple_file1, :force => {:excel => book2.excel})
        book4.should be_alive
        book4.should be_a Book
        book4.excel.should == book2.excel
        #book4.Readonly.should == true
        book4.should_not == book2 
        book4.close
        book5 = Book.open(@simple_file1, :force => {:excel => book2})
        book5.should be_alive
        book5.should be_a Book
        book5.excel.should == book2.excel
        #book5.Readonly.should == true
        book5.should_not == book2 
        book5.close
        book3.close
      end

      it "should open in a given Excel, provide identity transparency, because book can be readonly, such that the old and the new book are readonly" do
        book2 = Book.open(@simple_file1, :force => {:excel => :new})
        book2.excel.should_not == @book.excel
        book3 = Book.open(@simple_file1, :force => {:excel => :new})
        book3.excel.should_not == book2.excel
        book3.excel.should_not == @book.excel
        book2.close
        book3.close
        @book.close
        book4 = Book.open(@simple_file1, :force => {:excel => book2.excel}, :read_only => true)
        book4.should be_alive
        book4.should be_a Book
        book4.excel.should == book2.excel
        book4.ReadOnly.should be true
        book4.should == book2
        book4.close
        book5 = Book.open(@simple_file1, :force => {:excel => book2}, :read_only => true)
        book5.should be_alive
        book5.should be_a Book
        book5.excel.should == book2.excel
        book5.ReadOnly.should be true
        book5.should == book2
        book5.close
        book3.close
      end

      it "should open in a given Excel, provide identity transparency, because book can be readonly, such that the old and the new book are readonly" do
        book2 = Book.open(@simple_file1, :force => {:excel => :new})
        book2.excel.should_not == @book.excel
        book2.close
        @book.close
        book4 = Book.open(@simple_file1, :force => {:excel => book2}, :read_only => true)
        book4.should be_alive
        book4.should be_a Book
        book4.excel.should == book2.excel
        book4.ReadOnly.should be true
        book4.should == book2
        book4.close
      end

      it "should raise an error if no Excel or Book is given" do
        expect{
          Book.open(@simple_file1, :force => {:excel => :b})
          }.to raise_error(TypeREOError, "given object is neither an Excel, a Workbook, nor a Win32ole")
      end

      it "should do force_excel even if both force_ and default_excel is given" do
        book2 = Book.open(@simple_file1, :default => {:excel => @book.excel}, :force => {:excel => :new})
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should_not == @book.excel 
        book2.should_not == @book
      end

      it "should do default_excel if force_excel is nil" do
        book2 = Book.open(@another_simple_file, :force => {:excel => nil})
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should == @book.excel
      end

      it "should force_excel with :reuse" do
        book2 = Book.open(@different_file, :force => {:excel => :current})
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should == @book.excel
      end

      it "should force_excel with :reuse even if :default_excel says sth. else" do
        book2 = Book.open(@different_file, :force => {:excel => :current}, :default => {:excel => :new})
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should == @book.excel
      end

      it "should force_excel with :reuse when reopening and the Excel is not alive even if :default_excel says sth. else" do
        excel2 = Excel.new(:reuse => false)
        @book.excel.close
        book2 = Book.open(@simple_file1, :force => {:excel => :current}, :default => {:excel => :new})
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should === excel2
      end

      it "should force_excel with :reuse when reopening and the Excel is not alive even if :default_excel says sth. else" do
        book2 = Book.open(@different_file1, :force => {:excel => :new})
        book2.excel.close
        book3 = Book.open(@different_file1, :force => {:excel => :current}, :default => {:excel => :new})
        book3.should be_alive
        book3.should be_a Book
        book3.excel.should == @book.excel
      end
   
    end

    context "with leaving out :force => {:excel}" do

      before do
        @book = Book.open(@simple_file1)
      end

      after do
        @book.close rescue nil
      end

      it "should open in a given Excel provided as Excel, Book, or WIN32OLE representing an Excel or Workbook" do
        book2 = Book.open(@another_simple_file)
        book3 = Book.open(@different_file)
        book3 = Book.open(@simple_file1, :excel => book2.excel)
        book3.excel.should === book2.excel
        book4 = Book.open(@simple_file1, :excel => @book) 
        book4.excel.should === @book.excel
        book3.close
        book4.close
        book5 = Book.open(@simple_file1, :excel => book2.ole_workbook)
        book5.excel.should ===  book2.excel
        win32ole_excel1 = WIN32OLE.connect(@book.ole_workbook.Fullname).Application
        book6 = Book.open(@simple_file1, :excel => win32ole_excel1)
        book6.excel.should === @book.excel
      end

      it "should open in a new Excel" do
        book2 = Book.open(@simple_file1, :excel => :new)
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should_not == @book.excel 
        book2.should_not == @book
        @book.Readonly.should be false
        book2.Readonly.should be true
        book2.close
      end

      it "should open in a given Excel, not provide identity transparency, because old book readonly, new book writable" do
        book2 = Book.open(@simple_file1, :excel => :new)
        book2.excel.should_not == @book.excel
        book3 = Book.open(@simple_file1, :excel => :new)
        book3.excel.should_not == book2.excel
        book3.excel.should_not == @book.excel
        book2.close
        book4 = Book.open(@simple_file1, :excel => book2.excel)
        book4.should be_alive
        book4.should be_a Book
        book4.excel.should == book2.excel
        #book4.Readonly.should == true
        book4.should_not == book2 
        book4.close
        book5 = Book.open(@simple_file1, :excel => book2)
        book5.should be_alive
        book5.should be_a Book
        book5.excel.should == book2.excel
        #book5.Readonly.should == true
        book5.should_not == book2 
        book5.close
        book3.close
      end

      it "should open in a given Excel, provide identity transparency, because book can be readonly, such that the old and the new book are readonly" do
        book2 = Book.open(@simple_file1, :excel => :new)
        book2.excel.should_not == @book.excel
        book3 = Book.open(@simple_file1, :excel => :new)
        book3.excel.should_not == book2.excel
        book3.excel.should_not == @book.excel
        book2.close
        book3.close
        @book.close
        book4 = Book.open(@simple_file1, :excel => book2.excel, :read_only => true)
        book4.should be_alive
        book4.should be_a Book
        book4.excel.should == book2.excel
        book4.ReadOnly.should be true
        book4.should == book2
        book4.close
        book5 = Book.open(@simple_file1, :excel => book2, :read_only => true)
        book5.should be_alive
        book5.should be_a Book
        book5.excel.should == book2.excel
        book5.ReadOnly.should be true
        book5.should == book2
        book5.close
        book3.close
      end

      it "should open in a given Excel, provide identity transparency, because book can be readonly, such that the old and the new book are readonly" do
        book2 = Book.open(@simple_file1, :excel => :new)
        book2.excel.should_not == @book.excel
        book2.close
        @book.close
        book4 = Book.open(@simple_file1, :excel => book2, :read_only => true)
        book4.should be_alive
        book4.should be_a Book
        book4.excel.should == book2.excel
        book4.ReadOnly.should be true
        book4.should == book2
        book4.close
      end

      it "should raise an error if no Excel or Book is given" do
        expect{
          Book.open(@simple_file1, :excel => :b)
          }.to raise_error(TypeREOError, "given object is neither an Excel, a Workbook, nor a Win32ole")
      end

      it "should do force_excel even if both force_ and default_excel is given" do
        book2 = Book.open(@simple_file1, :default => {:excel => @book.excel}, :force => {:excel => :new})
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should_not == @book.excel 
        book2.should_not == @book
      end

      it "should do default_excel if force_excel is nil" do
        book2 = Book.open(@another_simple_file, :excel => nil)
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should == @book.excel
      end

      it "should force_excel with :reuse" do
        book2 = Book.open(@different_file, :excel => :current)
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should == @book.excel
      end

      it "should force_excel with :reuse even if :default_excel says sth. else" do
        book2 = Book.open(@different_file, :excel => :current, :default => {:excel => :new})
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should == @book.excel
      end

      it "should force_excel with :reuse when reopening and the Excel is not alive even if :default_excel says sth. else" do
        excel2 = Excel.new(:reuse => false)
        @book.excel.close
        book2 = Book.open(@simple_file1, :excel => :current, :default => {:excel => :new})
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should === excel2
      end

      it "should force_excel with :reuse when reopening and the Excel is not alive even if :default_excel says sth. else" do
        book2 = Book.open(@different_file1, :excel => :new)
        book2.excel.close
        book3 = Book.open(@different_file1, :excel => :current, :default => {:excel => :new})
        book3.should be_alive
        book3.should be_a Book
        book3.excel.should == @book.excel
      end
   
    end


    context "with :force_excel" do

      before do
        @book = Book.open(@simple_file1)
      end

      after do
        @book.close rescue nil
      end

      it "should open in a given Excel provided as Excel, Book, or WIN32OLE representing an Excel or Workbook" do
        book2 = Book.open(@another_simple_file)
        book3 = Book.open(@different_file)
        book3 = Book.open(@simple_file1, :force_excel => book2.excel)
        book3.excel.should === book2.excel
        book4 = Book.open(@simple_file1, :force_excel => @book) 
        book4.excel.should === @book.excel
        book3.close
        book4.close
        book5 = Book.open(@simple_file1, :force_excel => book2.ole_workbook)
        book5.excel.should ===  book2.excel
        win32ole_excel1 = WIN32OLE.connect(@book.ole_workbook.Fullname).Application
        book6 = Book.open(@simple_file1, :force_excel => win32ole_excel1)
        book6.excel.should === @book.excel
      end


      it "should open in a new Excel" do
        book2 = Book.open(@simple_file1, :force_excel => :new)
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should_not == @book.excel 
        book2.should_not == @book
        @book.Readonly.should be false
        book2.Readonly.should be true
        book2.close
      end

      it "should open in a given Excel, not provide identity transparency, because old book readonly, new book writable" do
        book2 = Book.open(@simple_file1, :force_excel => :new)
        book2.excel.should_not == @book.excel
        book3 = Book.open(@simple_file1, :force_excel => :new)
        book3.excel.should_not == book2.excel
        book3.excel.should_not == @book.excel
        book2.close
        book4 = Book.open(@simple_file1, :force_excel => book2.excel)
        book4.should be_alive
        book4.should be_a Book
        book4.excel.should == book2.excel
        #book4.Readonly.should == false
        book4.should_not == book2 
        book4.close
        book5 = Book.open(@simple_file1, :force_excel => book2)
        book5.should be_alive
        book5.should be_a Book
        book5.excel.should == book2.excel
        #book5.Readonly.should == false
        book5.should_not == book2 
        book5.close
        book3.close
      end

      it "should open in a given Excel, provide identity transparency, because book can be readonly, such that the old and the new book are readonly" do
        book2 = Book.open(@simple_file1, :force_excel => :new)
        book2.excel.should_not == @book.excel
        book3 = Book.open(@simple_file1, :force_excel => :new)
        book3.excel.should_not == book2.excel
        book3.excel.should_not == @book.excel
        book2.close
        book3.close
        @book.close
        book4 = Book.open(@simple_file1, :force_excel => book2.excel, :read_only => true)
        book4.should be_alive
        book4.should be_a Book
        book4.excel.should == book2.excel
        book4.ReadOnly.should be true
        book4.should == book2
        book4.close
        book5 = Book.open(@simple_file1, :force_excel => book2, :read_only => true)
        book5.should be_alive
        book5.should be_a Book
        book5.excel.should == book2.excel
        book5.ReadOnly.should be true
        book5.should == book2
        book5.close
        book3.close
      end

      it "should open in a given Excel, provide identity transparency, because book can be readonly, such that the old and the new book are readonly" do
        book2 = Book.open(@simple_file1, :force_excel => :new)
        book2.excel.should_not == @book.excel
        book2.close
        @book.close
        book4 = Book.open(@simple_file1, :force_excel => book2, :read_only => true)
        book4.should be_alive
        book4.should be_a Book
        book4.excel.should == book2.excel
        book4.ReadOnly.should be true
        book4.should == book2
        book4.close
      end

      it "should raise an error if no Excel or Book is given" do
        expect{
          Book.open(@simple_file1, :force_excel => :b)
          }.to raise_error(TypeREOError, "given object is neither an Excel, a Workbook, nor a Win32ole")
      end

      it "should do force_excel even if both force_ and default_excel is given" do
        book2 = Book.open(@simple_file1, :default_excel => @book.excel, :force_excel => :new)
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should_not == @book.excel 
        book2.should_not == @book
      end

      it "should do default_excel if force_excel is nil" do
        book2 = Book.open(@another_simple_file, :force_excel => nil)
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should == @book.excel
      end

      it "should force_excel with :reuse" do
        book2 = Book.open(@different_file, :force_excel => :current)
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should == @book.excel
      end

      it "should force_excel with :reuse even if :default_excel says sth. else" do
        book2 = Book.open(@different_file, :force_excel => :current, :default_excel => :new)
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should == @book.excel
      end

      it "should force_excel with :reuse when reopening and the Excel is not alive even if :default_excel says sth. else" do
        excel2 = Excel.new(:reuse => false)
        @book.excel.close
        book2 = Book.open(@simple_file1, :force_excel => :current, :default_excel => :new)
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should === excel2
      end

      it "should force_excel with :reuse when reopening and the Excel is not alive even if :default_excel says sth. else" do
        book2 = Book.open(@different_file1, :force_excel => :new)
        book2.excel.close
        book3 = Book.open(@different_file1, :force_excel => :current, :default_excel => :new)
        book3.should be_alive
        book3.should be_a Book
        book3.excel.should == @book.excel
      end
   
    end

    context "with :default => {:excel}" do

      before do
        @book = Book.open(@simple_file1, :visible => true)
      end

      after do
        @book.close rescue nil
      end

      it "should use the open book" do
        book2 = Book.open(@simple_file1, :default => {:excel => :current})
        book2.excel.should == @book.excel
        book2.should be_alive
        book2.should be_a Book
        book2.should == @book
        book2.close
      end

      it "should reopen the book in the excel instance where it was opened before" do
        excel = Excel.new(:reuse => false)
        @book.close
        book2 = Book.open(@simple_file1)
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should == @book.excel
        book2.excel.should_not == excel
        book2.filename.should == @book.filename
        @book.should be_alive
        book2.should == @book
        book2.close
      end

      it "should reopen a book in a new Excel if all Excel instances are closed" do
        excel = Excel.new(:reuse => false)
        excel2 = @book.excel
        fn = @book.filename
        @book.close
        Excel.close_all
        book2 = Book.open(@simple_file1, :default => {:excel => :current})
        book2.should be_alive
        book2.should be_a Book
        book2.filename.should == fn
        @book.should be_alive
        book2.should == @book
        book2.close
      end

      it "should reopen a book in the first opened Excel if the old Excel is closed" do
        excel = @book.excel
        Excel.close_all
        new_excel = Excel.new(:reuse => false)
        new_excel2 = Excel.new(:reuse => false)
        book2 = Book.open(@simple_file1, :default => {:excel => :current})
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should_not == excel
        book2.excel.should_not == new_excel2
        book2.excel.should == new_excel
        @book.should be_alive
        book2.should == @book
        book2.close
      end

      it "should reopen a book in the first opened excel, if the book cannot be reopened" do
        @book.close
        Excel.close_all
        excel1 = Excel.new(:reuse => false)
        excel2 = Excel.new(:reuse => false)
        book2 = Book.open(@different_file, :default => {:excel => :current})
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should == excel1
        book2.excel.should_not == excel2
        book2.close
      end

      it "should reopen the book in the Excel where it was opened most recently" do
        excel1 = @book.excel
        excel2 = Excel.new(:reuse => false)
        @book.close
        book2 = Book.open(@simple_file1, :default => {:excel => :current})
        book2.excel.should == excel1
        book2.close
        book3 = Book.open(@simple_file1, :force => {:excel => excel2})
        book3.close
        book3 = Book.open(@simple_file1, :default => {:excel => :current})
        book3.excel.should == excel2
        book3.close
      end

      it "should reopen a book in the excel instance where it was opened most recently" do
        book2 = Book.open(@simple_file, :force => {:excel => :new})
        @book.close
        book2.close
        book3 = Book.open(@simple_file1)
        book2.should be_alive
        book2.should be_a Book
        book3.excel.should == book2.excel
        book3.excel.should_not == @book.excel
        book3.should == book2
        book3.should_not == @book
      end

      it "should open the book in a new excel if the book was not opened before" do
        book2 = Book.open(@different_file, :default => {:excel => :current})
        book2.excel.should == @book.excel
        book3 = Book.open(@another_simple_file, :default => {:excel => :new})
        book3.excel.should_not == @book.excel
      end

      it "should open the book in a new excel if the book was opened before but the excel has been closed" do
        excel = @book.excel
        excel2 = Excel.new(:reuse => false)
        excel.close
        book2 = Book.open(@simple_file1, :default => {:excel => :new})
        book2.excel.should_not == excel2
        book2.close
      end

      it "should open the book in a given excel if the book was not opened before" do
        book2 = Book.open(@different_file, :default => {:excel => :current})
        book2.excel.should == @book.excel
        excel = Excel.new(:reuse => false)
        book3 = Book.open(@another_simple_file, :default => {:excel => excel})
        book3.excel.should == excel
      end

      it "should open the book in a given excel if the book was opened before but the excel has been closed" do
        excel2 = Excel.new(:reuse => false, :visible => true)
        @book.excel.close        
        book2 = Book.open(@simple_file1, :default => {:excel => excel2, :visible => true})
        book2.excel.should == excel2
      end

      it "should open a new excel, if the book cannot be reopened" do
        @book.close
        new_excel = Excel.new(:reuse => false)
        book2 = Book.open(@different_file, :default => {:excel => :new})
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should_not == new_excel
        book2.excel.should_not == @book.excel
        book2.close
      end

      it "should open a given excel, if the book cannot be reopened" do
        @book.close
        new_excel = Excel.new(:reuse => false)
        book2 = Book.open(@different_file, :default => {:excel => @book.excel})
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should_not == new_excel
        book2.excel.should == @book.excel
        book2.close
      end

      it "should open a given excel, if the book cannot be reopened" do
        @book.close
        new_excel = Excel.new(:reuse => false)
        book2 = Book.open(@different_file, :default => {:excel => @book})
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should_not == new_excel
        book2.excel.should == @book.excel
        book2.close
      end

      it "should reuse an open book by default" do
        book2 = Book.open(@simple_file1)
        book2.excel.should == @book.excel
        book2.should == @book
      end

      it "should raise an error if no Excel or Book is given" do
        expect{
          Book.open(@different_file, :default => {:excel => :a})
          }.to raise_error(TypeREOError, "given object is neither an Excel, a Workbook, nor a Win32ole")
      end
      
    end

    context "with :default_excel" do

      before do
        @book = Book.open(@simple_file1, :visible => true)
      end

      after do
        @book.close rescue nil
      end

      it "should use the open book" do
        book2 = Book.open(@simple_file1, :default_excel => :current)
        book2.excel.should == @book.excel
        book2.should be_alive
        book2.should be_a Book
        book2.should == @book
        book2.close
      end

      it "should reopen the book in the excel instance where it was opened before" do
        excel = Excel.new(:reuse => false)
        @book.close
        book2 = Book.open(@simple_file1)
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should == @book.excel
        book2.excel.should_not == excel
        book2.filename.should == @book.filename
        @book.should be_alive
        book2.should == @book
        book2.close
      end

      it "should reopen a book in a new Excel if all Excel instances are closed" do
        excel = Excel.new(:reuse => false)
        excel2 = @book.excel
        fn = @book.filename
        @book.close
        Excel.close_all
        book2 = Book.open(@simple_file1, :default_excel => :current)
        book2.should be_alive
        book2.should be_a Book
        book2.filename.should == fn
        @book.should be_alive
        book2.should == @book
        book2.close
      end

      it "should reopen a book in the first opened Excel if the old Excel is closed" do
        excel = @book.excel
        Excel.close_all
        new_excel = Excel.new(:reuse => false)
        new_excel2 = Excel.new(:reuse => false)
        book2 = Book.open(@simple_file1, :default_excel => :current)
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should_not == excel
        book2.excel.should_not == new_excel2
        book2.excel.should == new_excel
        @book.should be_alive
        book2.should == @book
        book2.close
      end

      it "should reopen a book in the first opened excel, if the book cannot be reopened" do
        @book.close
        Excel.close_all
        excel1 = Excel.new(:reuse => false)
        excel2 = Excel.new(:reuse => false)
        book2 = Book.open(@different_file, :default_excel => :current)
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should == excel1
        book2.excel.should_not == excel2
        book2.close
      end

      it "should reopen the book in the Excel where it was opened most recently" do
        excel1 = @book.excel
        excel2 = Excel.new(:reuse => false)
        @book.close
        book2 = Book.open(@simple_file1, :default_excel => :current)
        book2.excel.should == excel1
        book2.close
        book3 = Book.open(@simple_file1, :force_excel => excel2)
        book3.close
        book3 = Book.open(@simple_file1, :default_excel => :current)
        book3.excel.should == excel2
        book3.close
      end

      it "should reopen a book in the excel instance where it was opened most recently" do
        book2 = Book.open(@simple_file, :force_excel => :new)
        @book.close
        book2.close
        book3 = Book.open(@simple_file1)
        book2.should be_alive
        book2.should be_a Book
        book3.excel.should == book2.excel
        book3.excel.should_not == @book.excel
        book3.should == book2
        book3.should_not == @book
      end

      it "should open the book in a new excel if the book was not opened before" do
        book2 = Book.open(@different_file, :default_excel => :current)
        book2.excel.should == @book.excel
        book3 = Book.open(@another_simple_file, :default_excel => :new)
        book3.excel.should_not == @book.excel
      end

      it "should open the book in a new excel if the book was opened before but the excel has been closed" do
        excel = @book.excel
        excel2 = Excel.new(:reuse => false)
        excel.close
        book2 = Book.open(@simple_file1, :default_excel => :new)
        book2.excel.should_not == excel2
        book2.close
      end

      it "should open the book in a given excel if the book was not opened before" do
        book2 = Book.open(@different_file, :default_excel => :current)
        book2.excel.should == @book.excel
        excel = Excel.new(:reuse => false)
        book3 = Book.open(@another_simple_file, :default_excel => excel)
        book3.excel.should == excel
      end

      it "should open the book in a given excel if the book was opened before but the excel has been closed" do
        excel2 = Excel.new(:reuse => false, :visible => true)
        @book.excel.close        
        book2 = Book.open(@simple_file1, :visible => true, :default_excel => excel2)
        book2.excel.should == excel2
      end

      it "should open a new excel, if the book cannot be reopened" do
        @book.close
        new_excel = Excel.new(:reuse => false)
        book2 = Book.open(@different_file, :default_excel => :new)
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should_not == new_excel
        book2.excel.should_not == @book.excel
        book2.close
      end

      it "should open a given excel, if the book cannot be reopened" do
        @book.close
        new_excel = Excel.new(:reuse => false)
        book2 = Book.open(@different_file, :default_excel => @book.excel)
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should_not == new_excel
        book2.excel.should == @book.excel
        book2.close
      end

      it "should open a given excel, if the book cannot be reopened" do
        @book.close
        new_excel = Excel.new(:reuse => false)
        book2 = Book.open(@different_file, :default_excel => @book)
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should_not == new_excel
        book2.excel.should == @book.excel
        book2.close
      end

      it "should reuse an open book by default" do
        book2 = Book.open(@simple_file1)
        book2.excel.should == @book.excel
        book2.should == @book
      end

      it "should raise an error if no Excel or Book is given" do
        expect{
          Book.open(@different_file, :default_excel => :a)
          }.to raise_error(TypeREOError, "given object is neither an Excel, a Workbook, nor a Win32ole")
      end
      
    end

    context "with :active instead of :current" do
      
      before do
        @book = Book.open(@simple_file1)
      end

      after do
        @book.close rescue nil
      end

      it "should force_excel with :active" do
        book2 = Book.open(@different_file, :force => {:excel => :active})
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should == @book.excel
      end

      it "should force_excel with :reuse even if :default_excel says sth. else" do
        book2 = Book.open(@different_file, :force => {:excel => :active}, :default => {:excel => :new})
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should == @book.excel
      end

      it "should open force_excel with :reuse when reopening and the Excel is not alive even if :default_excel says sth. else" do
        excel2 = Excel.new(:reuse => false)
        @book.excel.close
        book2 = Book.open(@simple_file1, :force => {:excel => :active}, :default => {:excel => :new})
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should === excel2
      end

      it "should force_excel with :reuse when reopening and the Excel is not alive even if :default_excel says sth. else" do
        book2 = Book.open(@different_file1, :force => {:excel => :new})
        book2.excel.close
        book3 = Book.open(@different_file1, :force => {:excel => :active}, :default => {:excel => :new})
        book3.should be_alive
        book3.should be_a Book
        book3.excel.should == @book.excel
      end

      it "should use the open book" do
        book2 = Book.open(@simple_file1, :default => {:excel => :active})
        book2.excel.should == @book.excel
        book2.should be_alive
        book2.should be_a Book
        book2.should == @book
        book2.close
      end

      it "should reopen a book in a new Excel if all Excel instances are closed" do
        excel = Excel.new(:reuse => false)
        excel2 = @book.excel
        fn = @book.filename
        @book.close
        Excel.close_all
        book2 = Book.open(@simple_file1, :default => {:excel => :active})
        book2.should be_alive
        book2.should be_a Book
        book2.filename.should == fn
        @book.should be_alive
        book2.should == @book
        book2.close
      end

      it "should reopen a book in the first opened Excel if the old Excel is closed" do
        excel = @book.excel
        Excel.close_all
        new_excel = Excel.new(:reuse => false)
        new_excel2 = Excel.new(:reuse => false)
        book2 = Book.open(@simple_file1, :default => {:excel => :active})
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should_not == excel
        book2.excel.should_not == new_excel2
        book2.excel.should == new_excel
        @book.should be_alive
        book2.should == @book
        book2.close
      end

      it "should reopen a book in the first opened excel, if the book cannot be reopened" do
        @book.close
        Excel.close_all
        excel1 = Excel.new(:reuse => false)
        excel2 = Excel.new(:reuse => false)
        book2 = Book.open(@different_file, :default => {:excel => :active})
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should == excel1
        book2.excel.should_not == excel2
        book2.close
      end

      it "should reopen the book in the Excel where it was opened most recently" do
        excel1 = @book.excel
        excel2 = Excel.new(:reuse => false)
        @book.close
        book2 = Book.open(@simple_file1, :default => {:excel => :active})
        book2.excel.should == excel1
        book2.close
        book3 = Book.open(@simple_file1, :force => {:excel => excel2})
        book3.close
        book3 = Book.open(@simple_file1, :default => {:excel => :active})
        book3.excel.should == excel2
        book3.close
      end

    end

    context "with :reuse instead of :current" do
      
      before do
        @book = Book.open(@simple_file1)
      end

      after do
        @book.close rescue nil
      end

      it "should force_excel with :reuse" do
        book2 = Book.open(@different_file, :force => {:excel => :reuse})
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should == @book.excel
      end

      it "should force_excel with :reuse even if :default_excel says sth. else" do
        book2 = Book.open(@different_file, :force => {:excel => :reuse}, :default => {:excel => :new})
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should == @book.excel
      end

      it "should open force_excel with :reuse when reopening and the Excel is not alive even if :default_excel says sth. else" do
        excel2 = Excel.new(:reuse => false)
        @book.excel.close
        book2 = Book.open(@simple_file1, :force => {:excel => :reuse}, :default => {:excel => :new})
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should === excel2
      end

      it "should force_excel with :reuse when reopening and the Excel is not alive even if :default_excel says sth. else" do
        book2 = Book.open(@different_file1, :force => {:excel => :new})
        book2.excel.close
        book3 = Book.open(@different_file1, :force => {:excel => :reuse}, :default => {:excel => :new})
        book3.should be_alive
        book3.should be_a Book
        book3.excel.should == @book.excel
      end

      it "should use the open book" do
        book2 = Book.open(@simple_file1, :default => {:excel => :reuse})
        book2.excel.should == @book.excel
        book2.should be_alive
        book2.should be_a Book
        book2.should == @book
        book2.close
      end

      it "should reopen a book in a new Excel if all Excel instances are closed" do
        excel = Excel.new(:reuse => false)
        excel2 = @book.excel
        fn = @book.filename
        @book.close
        Excel.close_all
        book2 = Book.open(@simple_file1, :default => {:excel => :reuse})
        book2.should be_alive
        book2.should be_a Book
        book2.filename.should == fn
        @book.should be_alive
        book2.should == @book
        book2.close
      end

      it "should reopen a book in the first opened Excel if the old Excel is closed" do
        excel = @book.excel
        Excel.close_all
        new_excel = Excel.new(:reuse => false)
        new_excel2 = Excel.new(:reuse => false)
        book2 = Book.open(@simple_file1, :default => {:excel => :reuse})
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should_not == excel
        book2.excel.should_not == new_excel2
        book2.excel.should == new_excel
        @book.should be_alive
        book2.should == @book
        book2.close
      end

      it "should reopen a book in the first opened excel, if the book cannot be reopened" do
        @book.close
        Excel.close_all
        excel1 = Excel.new(:reuse => false)
        excel2 = Excel.new(:reuse => false)
        book2 = Book.open(@different_file, :default => {:excel => :reuse})
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should == excel1
        book2.excel.should_not == excel2
        book2.close
      end

      it "should reopen the book in the Excel where it was opened most recently" do
        excel1 = @book.excel
        excel2 = Excel.new(:reuse => false)
        @book.close
        book2 = Book.open(@simple_file1, :default => {:excel => :reuse})
        book2.excel.should == excel1
        book2.close
        book3 = Book.open(@simple_file1, :force => {:excel => excel2})
        book3.close
        book3 = Book.open(@simple_file1, :default => {:excel => :reuse})
        book3.excel.should == excel2
        book3.close
      end

    end

    it "should new_excel" do
      book = Book.open(@simple_file1)
      book.sheet(1)[1,1].Value = "f"
      book.Saved.should be false
      book2 = Book.open(@simple_file1, :if_unsaved => :new_excel)
    end

    context "with :if_unsaved" do

      before do
        @book = Book.open(@simple_file1)
        sheet = @book.sheet(1)
        #@book.add_sheet(@sheet, :as => 'a_name')
        @old_value = sheet[1,1].Value
        sheet[1,1] = (sheet[1,1].value == "foo" ? "bar" : "foo")
        @new_value = sheet[1,1].Value
        @book.Saved.should be false
      end

      after do
        @book.close(:if_unsaved => :forget)
      end

      it "should open the book in a new excel instance, if :if_unsaved is :new_excel" do
        new_book = Book.open(@simple_file1, :if_unsaved => :new_excel)
        new_book.excel.should_not == @book.excel
        @book.should be_alive
        @book.Saved.should be false
        new_book.should be_alive
        new_book.Saved.should be true
        new_book.sheet(1)[1,1].Value.should == @old_value
        #new_book.filename.should == @book.filename
        new_book.excel.should_not == @book.excel       
        new_book.close
      end

      it "should raise an error, if :if_unsaved is :raise" do
        expect {
          new_book = Book.open(@simple_file1, :if_unsaved => :raise)
        }.to raise_error(WorkbookNotSaved, /workbook is already open but not saved: "workbook.xls"/)
      end

      it "should let the book open, if :if_unsaved is :accept" do
        new_book = Book.open(@simple_file1, :if_unsaved => :accept)
        @book.should be_alive
        new_book.should be_alive
        new_book.Saved.should be false      
        @book.Saved.should be false  
        new_book.sheet(1)[1,1].Value.should == @new_value
        new_book.should == @book
      end

      it "should open book and close old book, if :if_unsaved is :forget" do
        new_book = Book.open(@simple_file1, :if_unsaved => :forget)
        @book.should_not be_alive
        new_book.should be_alive
        new_book.Saved.should be true
        new_book.sheet(1)[1,1].Value.should == @old_value
        #@new_book.filename.downcase.should == @simple_file.downcase
      end

      context "with :if_unsaved => :alert" do
        before do
         @key_sender = IO.popen  'ruby "' + File.join(File.dirname(__FILE__), '../helpers/key_sender.rb') + '" "Microsoft Office Excel" '  , "w"
        end

        after do
          @key_sender.close
        end

        it "should open the new book and close the unsaved book, if user answers 'yes'" do
          # "Yes" is the  default. --> language independent
          @key_sender.puts "{enter}"
          new_book = Book.open(@simple_file1, :if_unsaved => :alert)
          new_book.should be_alive
          #@book.should_not be_alive
          #new_book.filename.downcase.should == @simple_file.downcase
          new_book.Saved.should be true
          new_book.sheet(1)[1,1].Value.should == @old_value
        end
      end
=begin
        # only for Excel2007:
        it "should not open the new book and not close the unsaved book, if user answers 'no'" do
          # "No" is right to "Yes" (the  default). --> language independent
          # strangely, in the "no" case, the question will sometimes be repeated three times
          #@book.excel.Visible = true
          @key_sender.puts "{right}{enter}"
          @key_sender.puts "{right}{enter}"
          @key_sender.puts "{right}{enter}"
          expect{
            Book.open(@simple_file1, :if_unsaved => :alert)
            }.to raise_error(ExcelREOError, "user canceled or runtime error")
          @book.should be_alive
        end
    
=end      

      context "with :if_unsaved => :excel" do
        before do
         @key_sender = IO.popen  'ruby "' + File.join(File.dirname(__FILE__), '../helpers/key_sender.rb') + '" "Microsoft Office Excel" '  , "w"
        end

        after do
          @key_sender.close
        end

        it "should open the new book and close the unsaved book, if user answers 'yes'" do
          # "Yes" is the  default. --> language independent
          @key_sender.puts "{enter}"
          new_book = Book.open(@simple_file1, :if_unsaved => :excel)
          new_book.should be_alive
          #@book.should_not be_alive
          new_book.Saved.should be true
          new_book.sheet(1)[1,1].Value.should == @old_value
        end

=begin
        it "should not open the new book and not close the unsaved book, if user answers 'no'" do
          # "No" is right to "Yes" (the  default). --> language independent
          # strangely, in the "no" case, the question will sometimes be repeated three times
          #@book.excel.Visible = true
          @key_sender.puts "{right}{enter}"
          @key_sender.puts "{right}{enter}"
          @key_sender.puts "{right}{enter}"
          expect{
            Book.open(@simple_file1, :if_unsaved => :excel)
            }.to raise_error(ExcelREOError, "user canceled or runtime error")
          @book.should be_alive
        end
=end        

      end

      it "should raise an error, if :if_unsaved is default" do
        expect {
          new_book = Book.open(@simple_file1, :if_unsaved => :raise)
        }.to raise_error(WorkbookNotSaved, /workbook is already open but not saved: "workbook.xls"/)
      end

      it "should raise an error, if :if_unsaved is invalid option" do
        expect {
          new_book = Book.open(@simple_file1, :if_unsaved => :invalid_option)
        }.to raise_error(OptionInvalid, ":if_unsaved: invalid option: :invalid_option")
      end
    end

    context "with :if_obstructed" do

      for i in 1..2 do

        context "with and without reopen" do

          before do        
            if i == 1 then 
              book_before = Book.open(@simple_file1)
              book_before.close
            end
            @book = Book.open(@simple_file_other_path1)
            #@book.Windows(@book.Name).Visible = true
            #@sheet_count = @book.ole_workbook.Worksheets.Count
            sheet = @book.sheet(1)
            #@book.add_sheet(@sheet, :as => 'a_name')
            @old_value = sheet[1,1].Value
            sheet[1,1] = (sheet[1,1].value == "foo" ? "bar" : "foo")
            @new_value = sheet[1,1].Value
            @book.Saved.should be false
          end

          after do
            @book.close(:if_unsaved => :forget)
            #@new_book.close rescue nil
          end

          it "should raise an error, if :if_obstructed is :raise" do
            expect {
              new_book = Book.open(@simple_file1, :if_obstructed => :raise)
            }.to raise_error(WorkbookBlocked, /blocked by a workbook with the same name in a different path/)
          end

          it "should close the other book and open the new book, if :if_obstructed is :forget" do
            new_book = Book.open(@simple_file1, :if_obstructed => :forget)
            @book.should_not be_alive
            new_book.should be_alive
            new_book.filename.downcase.should == @simple_file.downcase
            old_book = Book.open(@simple_file_other_path1, :if_obstructed => :forget)
            old_book.sheet(1)[1,1].Value.should == @old_value
          end

          it "should save the old book, close it, and open the new book, if :if_obstructed is :save" do
            new_book = Book.open(@simple_file1, :if_obstructed => :save)
            @book.should_not be_alive
            new_book.should be_alive
            new_book.filename.downcase.should == @simple_file1.downcase
            old_book = Book.open(@simple_file_other_path1, :if_obstructed => :forget)
            old_book.sheet(1)[1,1].Value.should == @new_value
            #old_book.ole_workbook.Worksheets.Count.should ==  @sheet_count + 1
            old_book.close
          end

          it "should raise an error, if the old book is unsaved, and close the old book and open the new book, 
              if :if_obstructed is :close_if_saved" do
            expect{
              new_book = Book.open(@simple_file1, :if_obstructed => :close_if_saved)
            }.to raise_error(WorkbookBlocked, /workbook with the same name in a different path is unsaved/)
            @book.save
            new_book = Book.open(@simple_file1, :if_obstructed => :close_if_saved)
            @book.should_not be_alive
            new_book.should be_alive
            new_book.filename.downcase.should == @simple_file1.downcase
            old_book = Book.open(@simple_file_other_path1, :if_obstructed => :forget)
            old_book.sheet(1)[1,1].Value.should == @new_value
            #old_book.ole_workbook.Worksheets.Count.should ==  @sheet_count + 1
            old_book.close
          end

          it "should close the old book and open the new book, if :if_obstructed is :close_if_saved" do
            @book.close(:if_unsaved => :forget)
            book = Book.open(@simple_file_other_path)
            book2 = Book.open(@simple_file1, :if_obstructed => :close_if_saved)
          end

          it "should open the book in a new excel instance, if :if_obstructed is :new_excel" do
            new_book = Book.open(@simple_file1, :if_obstructed => :new_excel)
            @book.should be_alive
            @book.Saved.should be false
            @book.sheet(1)[1,1].Value.should == @new_value
            new_book.should be_alive
            new_book.filename.should_not == @book.filename
            new_book.excel.should_not == @book.excel
            new_book.sheet(1)[1,1].Value.should == @old_value
          end

          it "should raise an error, if :if_obstructed is default" do
            expect {
              new_book = Book.open(@simple_file1)              
            }.to raise_error(WorkbookBlocked, /blocked by a workbook with the same name in a different path/)
          end         

          it "should raise an error, if :if_obstructed is invalid option" do
            expect {
              new_book = Book.open(@simple_file1, :if_obstructed => :invalid_option)
            }.to raise_error(OptionInvalid, ":if_obstructed: invalid option: :invalid_option")
          end
        end
      end
    end

    context "with an already saved book" do
      before do
        @book = Book.open(@simple_file)
      end

      after do
        @book.close
      end

      possible_options = [:read_only, :raise, :accept, :forget, nil]
      possible_options.each do |options_value|
        context "with :if_unsaved => #{options_value} and in the same and different path" do
          before do
            @new_book = Book.open(@simple_file1, :reuse => true, :if_unsaved => options_value)
            @different_book = Book.open(@different_file, :reuse => true, :if_unsaved => options_value)
          end
          after do
            @new_book.close
            @different_book.close
          end
          it "should open without problems " do
            @new_book.should be_a Book
            @different_book.should be_a Book
          end
          it "should belong to the same Excel instance" do
            @new_book.excel.should == @book.excel
            @different_book.excel.should == @book.excel
          end
        end
      end
    end      
    
    context "with non-existing file" do

      it "should raise error if filename is nil" do
        expect{
          Book.open(@nonexisting)
          }.to raise_error(FileNameNotGiven, "filename is nil")
      end

      it "should raise error if file is a directory" do
        expect{
          Book.open(@dir)
          }.to raise_error(FileNotFound, "file #{General::absolute_path(@dir).gsub("/","\\").inspect} is a directory")
      end

      it "should raise error if file does not exist" do
        File.delete @simple_save_file rescue nil
        expect {
          Book.open(@simple_save_file, :if_absent => :raise)
        }.to raise_error(FileNotFound, "file #{General::absolute_path(@simple_save_file).gsub("/","\\").inspect} not found")
      end

      it "should create a workbook" do
        File.delete @simple_save_file rescue nil
        book = Book.open(@simple_save_file, :if_absent => :create)
        book.should be_a Book
        book.close
        File.exist?(@simple_save_file).should be true
      end

      it "should raise an exception by default" do
        File.delete @simple_save_file rescue nil
        expect {
          Book.open(@simple_save_file)
        }.to raise_error(FileNotFound, "file #{General::absolute_path(@simple_save_file).gsub("/","\\").inspect} not found")
      end

    end

    context "with attr_reader excel" do
     
      before do
        @new_book = Book.open(@simple_file)
      end
      after do
        @new_book.close
      end
      it "should provide the excel instance of the book" do
        excel = @new_book.excel
        excel.class.should == Excel
        excel.should be_a Excel
      end
    end

=begin
    # work in progress
    context "with :update_links" do
      
      it "should set update_links to :alert" do
        book = Book.open(@simple_file, :update_links => :alert)
        book.UpdateLinks.should == XlUpdateLinksUserSetting
        book.Saved.should be true
      end

      it "should set update_links to :never" do
        book = Book.open(@simple_file, :update_links => :never)
        book.UpdateLinks.should == XlUpdateLinksNever
        book = Book.open(@simple_file, :update_links => :foo)
        book.UpdateLinks.should == XlUpdateLinksNever
      end

      it "should set update_links to :always" do
        book = Book.open(@simple_file, :update_links => :always)
        book.UpdateLinks.should == XlUpdateLinksAlways
      end

      it "should set update_links to :never per default" do
        book = Book.open(@simple_file)
        book.UpdateLinks.should == XlUpdateLinksNever
      end

    end
=end

    context "with :read_only" do
      
      it "should reopen the book with writable (unsaved changes from readonly will not be saved)" do
        book = Book.open(@simple_file1, :read_only => true)
        book.ReadOnly.should be true
        book.should be_alive
        sheet = book.sheet(1)
        old_cell_value = sheet[1,1].value
        sheet[1,1] = sheet[1,1].value == "foo" ? "bar" : "foo"
        book.Saved.should be false
        new_book = Book.open(@simple_file1, :read_only => false, :if_unsaved => :accept)
        new_book.ReadOnly.should be false 
        new_book.should be_alive
        book.should be_alive   
        new_book.should == book 
        new_sheet = new_book.sheet(1)
        new_cell_value = new_sheet[1,1].value
        new_cell_value.should == old_cell_value
      end

      it "should not raise an error when trying to reopen the book as read_only while the writable book had unsaved changes" do
        book = Book.open(@simple_file1, :read_only => false)
        book.ReadOnly.should be false
        book.should be_alive
        sheet = book.sheet(1)
        old_cell_value = sheet[1,1].value        
        sheet[1,1] = sheet[1,1].value == "foo" ? "bar" : "foo"
        book.Saved.should be false
        new_book = Book.open(@simple_file1, :read_only => true, :if_unsaved => :accept)
        new_book.ReadOnly.should be true
        new_book.Saved.should be true
        new_book.should == book
      end

      it "should reopen the book with writable in the same Excel instance (unsaved changes from readonly will not be saved)" do
        book = Book.open(@simple_file1, :read_only => true)
        book.ReadOnly.should be true
        book.should be_alive
        sheet = book.sheet(1)
        old_cell_value = sheet[1,1].value
        sheet[1,1] = sheet[1,1].value == "foo" ? "bar" : "foo"
        book.Saved.should be false
        new_book = Book.open(@simple_file1, :if_unsaved => :accept, :force => {:excel => book.excel}, :read_only => false)
        new_book.ReadOnly.should be false 
        new_book.should be_alive
        book.should be_alive   
        new_book.should == book 
        new_sheet = new_book.sheet(1)
        new_cell_value = new_sheet[1,1].value
        new_cell_value.should == old_cell_value
      end

      it "should reopen the book with readonly (unsaved changes of the writable should be saved)" do
        book = Book.open(@simple_file1, :force => {:excel => :new}, :read_only => false)
        book.ReadOnly.should be false
        book.should be_alive
        sheet = book.sheet(1)
        old_cell_value = sheet[1,1].value        
        sheet[1,1] = sheet[1,1].value == "foo" ? "bar" : "foo"
        book.Saved.should be false
        new_book = Book.open(@simple_file1, :force => {:excel => book.excel}, :read_only => true, :if_unsaved => :accept)
        new_book.ReadOnly.should be true
        new_book.Saved.should be true
        new_book.should == book
        new_sheet = new_book.sheet(1)
        new_cell_value = new_sheet[1,1].value
        new_cell_value.should_not == old_cell_value
      end

      it "should open the second book in another Excel as writable" do
        book = Book.open(@simple_file1, :read_only => true)
        book.ReadOnly.should be true
        new_book = Book.open(@simple_file1, :force => {:excel => :new}, :read_only => false)
        new_book.ReadOnly.should be false
        new_book.close
        book.close
      end

      it "should be able to save, if :read_only => false" do
        book = Book.open(@simple_file1, :read_only => false)
        book.should be_a Book
        expect {
          book.save_as(@simple_save_file, :if_exists => :overwrite)
        }.to_not raise_error
        book.close
      end

      it "should be able to save, if :read_only is default" do
        book = Book.open(@simple_file1)
        book.should be_a Book
        expect {
          book.save_as(@simple_save_file, :if_exists => :overwrite)
        }.to_not raise_error
        book.close
      end

      it "should raise an error, if :read_only => true" do
        book = Book.open(@simple_file, :read_only => true)
        book.should be_a Book
        expect {
          book.save_as(@simple_save_file, :if_exists => :overwrite)
        }.to raise_error
        book.close
      end
    end

    context "with various file formats" do

      it "should open linked workbook" do
        book = Book.open(@linked_file, :visible => true)
        book.close
      end

      #it "should open xlsm file" do
      #  book = Book.open(@simple_file_xlsm, :visible => true)
      #  book.close
      #end

      it "should open xlsx file" do
        book = Book.open(@simple_file_xlsx, :visible => true)
        book.close
      end
    end


    context "with block" do
      it 'block parameter should be instance of Book' do
        Book.open(@simple_file) do |book|
          book.should be_a Book
        end
      end
    end

    context "with WIN32OLE#GetAbsolutePathName" do
      it "'~' should be HOME directory" do
        path = '~/Abrakadabra.xlsx'
        expected_path = Regexp.new(File.expand_path(path).gsub(/\//, "."))
        expect {
          Book.open(path)
        }.to raise_error(FileNotFound, "file #{General::absolute_path(path).gsub("/","\\").inspect} not found")
      end
    end
  end

  describe "reopen" do

    context "with standard" do
      
      before do
        @book = Book.open(@simple_file)
      end

      after do
        @book.close
      end

      it "should reopen the closed book" do
        @book.should be_alive
        book1 = @book
        @book.close
        @book.should_not be_alive
        @book.reopen
        @book.should be_a Book
        @book.should be_alive
        @book.should === book1
      end
    end
  end

  describe "uplifting" do

    context "with standard" do

      before do
        @book = Book.open(@simple_file)
      end

      after do
        @book.close
      end

      it "should uplift a workbook to a book with an open book" do
        workbook = @book.ole_workbook
        book1 = Book.new(workbook)
        book1.should be_a Book
        book1.should be_alive
        book1.should == @book
      end
    end
  end
end