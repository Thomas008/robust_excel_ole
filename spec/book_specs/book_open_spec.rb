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
    Excel.kill_all
    rm_tmp(@dir)
  end

  describe "open" do

    context "with class identifier 'Workbook'" do

      before do
        @book = Workbook.open(@simple_file)
      end

      after do
        @book.close rescue nil
      end

      it "should open in a new Excel" do
        book2 = Workbook.open(@simple_file, :force_excel => :new)
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should_not == @book.excel 
        book2.should_not == @book
        @book.Readonly.should be_false
        book2.Readonly.should be_true
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
        workbook = @book.workbook
        new_book = Book.new(workbook)
        new_book.should be_a Book
        new_book.should be_alive
        new_book.should == @book
        new_book.filename.should == @book.filename
        new_book.excel.should == @book.excel
        new_book.excel.Visible.should be_false
        new_book.excel.DisplayAlerts.should be_false
        new_book.should === @book
        new_book.close
      end

      it "should fetch the workbook" do
        workbook = @book.workbook
        new_book = Book.new(workbook, :visible => true)
        new_book.should be_a Book
        new_book.should be_alive
        new_book.should == @book
        new_book.filename.should == @book.filename
        new_book.excel.should == @book.excel
        new_book.excel.Visible.should be_true
        new_book.excel.DisplayAlerts.should be_false
        new_book.should === @book
        new_book.close
      end

      it "should yield an identical Book and set visible and displayalerts values" do
        workbook = @book.workbook
        new_book = Book.new(workbook, :visible => true, :displayalerts => true)
        new_book.should be_a Book
        new_book.should be_alive
        new_book.should == @book
        new_book.filename.should == @book.filename
        new_book.excel.should == @book.excel
        new_book.should === @book
        new_book.excel.visible.should be_true
        new_book.excel.displayalerts.should be_true
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
        @book = Book.open(@simple_file)
      end

      after do
        @book.close
      end

      it "should yield identical Book objects for identical Excel books" do
        book2 = Book.open(@simple_file)
        book2.should === @book
        book2.close
      end

      it "should yield different Book objects for different Excel books" do
        book2 = Book.open(@different_file)
        book2.should_not === @book
        book2.close
      end

      it "should yield different Book objects when opened the same file in different Excel instances" do
        book2 = Book.open(@simple_file, :force_excel => :new)
        book2.should_not === @book
        book2.close
      end

      it "should yield identical Book objects for identical Excel books when reopening" do
        @book.should be_alive
        @book.close
        @book.should_not be_alive
        book2 = Book.open(@simple_file)
        book2.should === @book
        book2.should be_alive
        book2.close
      end

      it "should yield identical Book objects when reopening and the Excel is closed" do
        @book.should be_alive
        @book.close
        Excel.close_all
        book2 = Book.open(@simple_file)
        book2.should be_alive
        book2.should === @book
        book2.close
      end

      it "should yield different Book objects when reopening in a new Excel" do
        @book.should be_alive
        old_excel = @book.excel
        @book.close
        @book.should_not be_alive
        book2 = Book.open(@simple_file, :force_excel => :new)
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
        book2 = Book.open(@simple_file, :force_excel => new_excel)
        book2.should_not === @book
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
        book2 = Book.open(@simple_file, :force_excel => old_excel)
        book2.should === @book
        book2.should be_alive
        book2.excel.should == old_excel
        @book.should be_alive
        book2.close
      end

    end

    context "with :force_excel" do

      before do
        @book = Book.open(@simple_file)
      end

      after do
        @book.close rescue nil
      end

      it "should open in a given Excel provided as Excel, Book, or WIN32OLE representing an Excel or Workbook" do
        book2 = Book.open(@another_simple_file)
        book3 = Book.open(@different_file)
        book3 = Book.open(@simple_file, :force_excel => book2.excel)
        book3.excel.should === book2.excel
        book4 = Book.open(@simple_file, :force_excel => @book) 
        book4.excel.should === @book.excel
        book3.close
        book4.close
        book5 = Book.open(@simple_file, :force_excel => book2.workbook)
        book5.excel.should ===  book2.excel
        win32ole_excel1 = WIN32OLE.connect(@book.workbook.Fullname).Application
        book6 = Book.open(@simple_file, :force_excel => win32ole_excel1)
        book6.excel.should === @book.excel
      end


      it "should open in a new Excel" do
        book2 = Book.open(@simple_file, :force_excel => :new)
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should_not == @book.excel 
        book2.should_not == @book
        @book.Readonly.should be_false
        book2.Readonly.should be_true
        book2.close
      end

      it "should open in a given Excel, not provide identity transparency, because old book readonly, new book writable" do
        book2 = Book.open(@simple_file, :force_excel => :new)
        book2.excel.should_not == @book.excel
        book3 = Book.open(@simple_file, :force_excel => :new)
        book3.excel.should_not == book2.excel
        book3.excel.should_not == @book.excel
        book2.close
        book4 = Book.open(@simple_file, :force_excel => book2.excel)
        book4.should be_alive
        book4.should be_a Book
        book4.excel.should == book2.excel
        book4.Readonly.should == true
        book4.should_not == book2 
        book4.close
        book5 = Book.open(@simple_file, :force_excel => book2)
        book5.should be_alive
        book5.should be_a Book
        book5.excel.should == book2.excel
        book5.Readonly.should == true
        book5.should_not == book2 
        book5.close
        book3.close
      end

      it "should open in a given Excel, provide identity transparency, because book can be readonly, such that the old and the new book are readonly" do
        book2 = Book.open(@simple_file, :force_excel => :new)
        book2.excel.should_not == @book.excel
        book3 = Book.open(@simple_file, :force_excel => :new)
        book3.excel.should_not == book2.excel
        book3.excel.should_not == @book.excel
        book2.close
        book3.close
        @book.close
        book4 = Book.open(@simple_file, :force_excel => book2.excel, :read_only => true)
        book4.should be_alive
        book4.should be_a Book
        book4.excel.should == book2.excel
        book4.ReadOnly.should be_true
        book4.should == book2
        book4.close
        book5 = Book.open(@simple_file, :force_excel => book2, :read_only => true)
        book5.should be_alive
        book5.should be_a Book
        book5.excel.should == book2.excel
        book5.ReadOnly.should be_true
        book5.should == book2
        book5.close
        book3.close
      end

      it "should open in a given Excel, provide identity transparency, because book can be readonly, such that the old and the new book are readonly" do
        book2 = Book.open(@simple_file, :force_excel => :new)
        book2.excel.should_not == @book.excel
        book2.close
        @book.close
        book4 = Book.open(@simple_file, :force_excel => book2, :read_only => true)
        book4.should be_alive
        book4.should be_a Book
        book4.excel.should == book2.excel
        book4.ReadOnly.should be_true
        book4.should == book2
        book4.close
      end

      it "should raise an error if no Excel or Book is given" do
        expect{
          Book.open(@simple_file, :force_excel => :b)
          }.to raise_error(ExcelError, "receiver instance is neither an Excel nor a Book")
      end

      it "should do force_excel even if both force_ and default_excel is given" do
        book2 = Book.open(@simple_file, :default_excel => @book.excel, :force_excel => :new)
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should_not == @book.excel 
        book2.should_not == @book
      end
    end

    context "with another :force_excel" do
      it "should do force_excel even if both force_ and default_excel is given" do
        book2 = Book.open(@simple_file, :force_excel => nil)
        book2.should be_alive
        book2.should be_a Book
      end
    end

    context "with :default_excel" do

      before do
        excel = Excel.new(:reuse => false)
        @book = Book.open(@simple_file)
      end

      after do
        @book.close rescue nil
      end

      it "should use the open book" do
        book2 = Book.open(@simple_file, :default_excel => :reuse)
        book2.excel.should == @book.excel
        book2.should be_alive
        book2.should be_a Book
        book2.should == @book
        book2.close
      end

      it "should reopen the book in the excel instance where it was opened before" do
        excel = Excel.new(:reuse => false)
        @book.close
        book2 = Book.open(@simple_file)
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
        book2 = Book.open(@simple_file, :default_excel => :reuse)
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
        book2 = Book.open(@simple_file, :default_excel => :reuse)
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
        book2 = Book.open(@different_file, :default_excel => :reuse)
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should == excel1
        book2.excel.should_not == excel2
        book2.close
      end

      it "should reopen a book in the excel instance where it was opened most recently" do
        book2 = Book.open(@simple_file, :force_excel => :new)
        @book.close
        book2.close
        book3 = Book.open(@simple_file)
        book2.should be_alive
        book2.should be_a Book
        book3.excel.should == book2.excel
        book3.excel.should_not == @book.excel
        book3.should == book2
        book3.should_not == @book
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
        book2 = Book.open(@simple_file)
        book2.excel.should == @book.excel
        book2.should == @book
      end
    end

    context "with :if_unsaved" do

      before do
        @book = Book.open(@simple_file)
        @sheet = @book[0]
        @book.add_sheet(@sheet, :as => 'a_name')
      end

      after do
        @book.close(:if_unsaved => :forget)
        @new_book.close rescue nil
      end

      it "should raise an error, if :if_unsaved is :raise" do
        expect {
          @new_book = Book.open(@simple_file, :if_unsaved => :raise)
        }.to raise_error(ExcelErrorOpen, /workbook is already open but not saved: "workbook.xls"/)
      end

      it "should let the book open, if :if_unsaved is :accept" do
        expect {
          @new_book = Book.open(@simple_file, :if_unsaved => :accept)
          }.to_not raise_error
        @book.should be_alive
        @new_book.should be_alive
        @new_book.should == @book
      end

      it "should open book and close old book, if :if_unsaved is :forget" do
        @new_book = Book.open(@simple_file, :if_unsaved => :forget)
        @book.should_not be_alive
        @new_book.should be_alive
        @new_book.filename.downcase.should == @simple_file.downcase
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
          @new_book = Book.open(@simple_file, :if_unsaved => :alert)
          @new_book.should be_alive
          @new_book.filename.downcase.should == @simple_file.downcase
          @book.should_not be_alive
        end

        it "should not open the new book and not close the unsaved book, if user answers 'no'" do
          # "No" is right to "Yes" (the  default). --> language independent
          # strangely, in the "no" case, the question will sometimes be repeated three times
          #@book.excel.Visible = true
          @key_sender.puts "{right}{enter}"
          @key_sender.puts "{right}{enter}"
          @key_sender.puts "{right}{enter}"
          expect{
            Book.open(@simple_file, :if_unsaved => :alert)
            }.to raise_error(ExcelErrorOpen, "open: user canceled or open error")
          @book.should be_alive
        end
      end

      it "should open the book in a new excel instance, if :if_unsaved is :new_excel" do
        @new_book = Book.open(@simple_file, :if_unsaved => :new_excel)
        @book.should be_alive
        @new_book.should be_alive
        @new_book.filename.should == @book.filename
        @new_book.excel.should_not == @book.excel       
        @new_book.close
      end

      it "should raise an error, if :if_unsaved is default" do
        expect {
          @new_book = Book.open(@simple_file, :if_unsaved => :raise)
        }.to raise_error(ExcelErrorOpen, /workbook is already open but not saved: "workbook.xls"/)
      end

      it "should raise an error, if :if_unsaved is invalid option" do
        expect {
          @new_book = Book.open(@simple_file, :if_unsaved => :invalid_option)
        }.to raise_error(ExcelErrorOpen, ":if_unsaved: invalid option: :invalid_option")
      end
    end

    context "with :if_obstructed" do

      for i in 1..2 do

        context "with and without reopen" do

          before do        
            if i == 1 then 
              book_before = Book.open(@simple_file)
              book_before.close
            end
            @book = Book.open(@simple_file_other_path)
            @sheet_count = @book.workbook.Worksheets.Count
            @sheet = @book[0]
            @book.add_sheet(@sheet, :as => 'a_name')
          end

          after do
            @book.close(:if_unsaved => :forget)
            @new_book.close rescue nil
          end

          it "should raise an error, if :if_obstructed is :raise" do
            expect {
              @new_book = Book.open(@simple_file, :if_obstructed => :raise)
            }.to raise_error(ExcelErrorOpen, /blocked by a book with the same name in a different path: "workbook.xls"/)
          end

          it "should close the other book and open the new book, if :if_obstructed is :forget" do
            @new_book = Book.open(@simple_file, :if_obstructed => :forget)
            @book.should_not be_alive
            @new_book.should be_alive
            @new_book.filename.downcase.should == @simple_file.downcase
          end

          it "should save the old book, close it, and open the new book, if :if_obstructed is :save" do
            @new_book = Book.open(@simple_file, :if_obstructed => :save)
            @book.should_not be_alive
            @new_book.should be_alive
            @new_book.filename.downcase.should == @simple_file.downcase
            old_book = Book.open(@simple_file_other_path, :if_obstructed => :forget)
            old_book.workbook.Worksheets.Count.should ==  @sheet_count + 1
            old_book.close
          end

          it "should raise an error, if the old book is unsaved, and close the old book and open the new book, 
              if :if_obstructed is :close_if_saved" do
            expect{
              @new_book = Book.open(@simple_file, :if_obstructed => :close_if_saved)
            }.to raise_error(ExcelErrorOpen, /workbook with the same name in a different path is unsaved: "workbook.xls"/)
            @book.save
            @new_book = Book.open(@simple_file, :if_obstructed => :close_if_saved)
            @book.should_not be_alive
            @new_book.should be_alive
            @new_book.filename.downcase.should == @simple_file.downcase
            old_book = Book.open(@simple_file_other_path, :if_obstructed => :forget)
            old_book.workbook.Worksheets.Count.should ==  @sheet_count + 1
            old_book.close
          end

          it "should open the book in a new excel instance, if :if_obstructed is :new_excel" do
            @new_book = Book.open(@simple_file, :if_obstructed => :new_excel)
            @book.should be_alive
            @new_book.should be_alive
            @new_book.filename.should_not == @book.filename
            @new_book.excel.should_not == @book.excel
          end

          it "should raise an error, if :if_obstructed is default" do
            expect {
              @new_book = Book.open(@simple_file)              
            }.to raise_error(ExcelErrorOpen, /blocked by a book with the same name in a different path: "workbook.xls"/)
          end         

          it "should raise an error, if :if_obstructed is invalid option" do
            expect {
              @new_book = Book.open(@simple_file, :if_obstructed => :invalid_option)
            }.to raise_error(ExcelErrorOpen, ":if_obstructed: invalid option: :invalid_option")
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
            @new_book = Book.open(@simple_file, :reuse=> true, :if_unsaved => options_value)
            @different_book = Book.new(@different_file, :reuse=> true, :if_unsaved => options_value)
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

      it "should raise an exception" do
        File.delete @simple_save_file rescue nil
        expect {
          Book.open(@simple_save_file, :if_absent => :raise)
        }.to raise_error(ExcelErrorOpen, "file \"#{@simple_save_file}\" not found")
      end

      it "should create a workbook" do
        File.delete @simple_save_file rescue nil
        book = Book.open(@simple_save_file, :if_absent => :create)
        book.should be_a Book
        book.close
        File.exist?(@simple_save_file).should be_true
      end

      it "should raise an exception by default" do
        File.delete @simple_save_file rescue nil
        expect {
          Book.open(@simple_save_file)
        }.to raise_error(ExcelErrorOpen, "file \"#{@simple_save_file}\" not found")
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

    context "with :read_only" do
      
      it "should reopen the book with writable (unsaved changes from readonly will not be saved)" do
        book = Book.open(@simple_file, :read_only => true)
        book.ReadOnly.should be_true
        book.should be_alive
        sheet = book[0]
        old_cell_value = sheet[1,1].value
        sheet[1,1] = sheet[1,1].value == "foo" ? "bar" : "foo"
        book.Saved.should be_false
        new_book = Book.open(@simple_file, :read_only => false, :if_unsaved => :accept)
        new_book.ReadOnly.should be_false 
        new_book.should be_alive
        book.should be_alive   
        new_book.should == book 
        new_sheet = new_book[0]
        new_cell_value = new_sheet[1,1].value
        new_cell_value.should == old_cell_value
      end

      it "should not raise an error when trying to reopen the book as read_only while the writable book had unsaved changes" do
        book = Book.open(@simple_file, :read_only => false)
        book.ReadOnly.should be_false
        book.should be_alive
        sheet = book[0]
        old_cell_value = sheet[1,1].value        
        sheet[1,1] = sheet[1,1].value == "foo" ? "bar" : "foo"
        book.Saved.should be_false
        new_book = Book.open(@simple_file, :read_only => true, :if_unsaved => :accept)
        new_book.ReadOnly.should be_false
        new_book.Saved.should be_false
        new_book.should == book
      end

      it "should reopen the book with writable in the same Excel instance (unsaved changes from readonly will not be saved)" do
        book = Book.open(@simple_file, :read_only => true)
        book.ReadOnly.should be_true
        book.should be_alive
        sheet = book[0]
        old_cell_value = sheet[1,1].value
        sheet[1,1] = sheet[1,1].value == "foo" ? "bar" : "foo"
        book.Saved.should be_false
        new_book = Book.open(@simple_file, :if_unsaved => :accept, :force_excel => book.excel, :read_only => false)
        new_book.ReadOnly.should be_false 
        new_book.should be_alive
        book.should be_alive   
        new_book.should == book 
        new_sheet = new_book[0]
        new_cell_value = new_sheet[1,1].value
        new_cell_value.should == old_cell_value
      end

      it "should reopen the book with readonly (unsaved changes of the writable should be saved)" do
        book = Book.open(@simple_file, :force_excel => :new, :read_only => false)
        book.ReadOnly.should be_false
        book.should be_alive
        sheet = book[0]
        old_cell_value = sheet[1,1].value        
        sheet[1,1] = sheet[1,1].value == "foo" ? "bar" : "foo"
        book.Saved.should be_false
        new_book = Book.open(@simple_file, :force_excel => book.excel, :read_only => true, :if_unsaved => :accept)
        new_book.ReadOnly.should be_false
        new_book.Saved.should be_false
        new_book.should == book
      end

      it "should open the second book in another Excel as writable" do
        book = Book.open(@simple_file, :read_only => true)
        book.ReadOnly.should be_true
        new_book = Book.open(@simple_file, :force_excel => :new, :read_only => false)
        new_book.ReadOnly.should be_false
        new_book.close
        book.close
      end

      it "should be able to save, if :read_only => false" do
        book = Book.open(@simple_file, :read_only => false)
        book.should be_a Book
        expect {
          book.save_as(@simple_save_file, :if_exists => :overwrite)
        }.to_not raise_error
        book.close
      end

      it "should be able to save, if :read_only is default" do
        book = Book.open(@simple_file)
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

      it "should open xlsm file" do
        book = Book.open(@simple_file_xlsm, :visible => true)
        book.close
      end

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
        }.to raise_error(ExcelErrorOpen, "file \"#{path}\" not found")
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
        workbook = @book.workbook
        book1 = Book.new(workbook)
        book1.should be_a Book
        book1.should be_alive
        book1.should == @book
      end
    end
  end
end