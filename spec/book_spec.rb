# -*- coding: utf-8 -*-

require File.join(File.dirname(__FILE__), './spec_helper')


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
    @more_simple_file = @dir + '/more_workbook.xls'
    @connected_file = @dir + '/workbook.xlsx'
  end

  after do
    Excel.close_all
    rm_tmp(@dir)
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
  
  describe "open" do

    context "with connected workbook" do
      it "should open connected workbook" do
        p "filename: #{@connected_file}"
        book = Book.open(@connected_file, :visible => true)
        book.close
      end
    end

    context "standard use cases" do

      it "should read in a seperate excel instance" do
        first_excel = Excel.new
        book = Book.open(@simple_file, :read_only => true, :force_excel => :new)
        book.should be_a Book
        book.should be_alive
        book.ReadOnly.should be_true
        book.Saved.should be_true
        book.excel.should_not == first_excel
        sheet = book[0]
        sheet[0,0].value.should == "simple"
        book.close
      end

      it "should read not bothering about excel instances" do
        first_excel = Excel.new
        book = Book.open(@simple_file, :read_only => true)
        book.should be_a Book
        book.should be_alive
        book.ReadOnly.should be_true
        book.Saved.should be_true
        book.excel.should == first_excel
        sheet = book[0]
        sheet[0,0].value.should == "simple"
        book.close
      end

      it "should open writable" do
        book = Book.open(@simple_file, :if_locked => :take_writable, 
                                      :if_unsaved => :forget, :if_obstructed => :save)
        book.close
      end

      it "should open unobtrusively" do
        book = Book.open(@simple_file, :if_locked => :take_writable, 
                                      :if_unsaved => :accept, :if_obstructed => :new_excel)
        book.close
      end

      it "should open in a given instance" do
        book1 = Book.open(@simple_file)
        book2 = Book.open(@simple_file, :force_excel => book1.excel, :if_locked => :force_writable) 
        book2.close
        book1.close
      end
     
      it "should open writable" do
        book = Book.open(@simple_file, :if_locked => :take_writable, 
                                        :if_unsaved => :save, :if_obstructed => :save)
        book.close
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
        book3.close
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
        }.to raise_error(ExcelErrorOpen, "book is already open but not saved (#{File.basename(@simple_file)})")
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
         @key_sender = IO.popen  'ruby "' + File.join(File.dirname(__FILE__), '/helpers/key_sender.rb') + '" "Microsoft Office Excel" '  , "w"
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
        }.to raise_error(ExcelErrorOpen, "book is already open but not saved (#{File.basename(@simple_file)})")
      end

      it "should raise an error, if :if_unsaved is invalid option" do
        expect {
          @new_book = Book.open(@simple_file, :if_unsaved => :invalid_option)
        }.to raise_error(ExcelErrorOpen, ":if_unsaved: invalid option")
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
            }.to raise_error(ExcelErrorOpen, "blocked by a book with the same name in a different path")
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
            }.to raise_error(ExcelErrorOpen, "book with the same name in a different path is unsaved")
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
            }.to raise_error(ExcelErrorOpen, "blocked by a book with the same name in a different path")
          end         

          it "should raise an error, if :if_obstructed is invalid option" do
            expect {
              @new_book = Book.open(@simple_file, :if_obstructed => :invalid_option)
            }.to raise_error(ExcelErrorOpen, ":if_obstructed: invalid option")
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
        }.to raise_error(ExcelErrorOpen, "file #{@simple_save_file} not found")
      end

      it "should create a workbook" do
        File.delete @simple_save_file rescue nil
        book = Book.open(@simple_save_file, :if_absent => :create)
        book.should be_a Book
        book.close
        File.exist?(@simple_save_file).should be_true
      end

      it "should create a workbook by default" do
        File.delete @simple_save_file rescue nil
        book = Book.open(@simple_save_file)
        book.should be_a Book
        book.close
        File.exist?(@simple_save_file).should be_true
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
        old_cell_value = sheet[0,0].value
        sheet[0,0] = sheet[0,0].value == "simple" ? "complex" : "simple"
        book.Saved.should be_false
        new_book = Book.open(@simple_file, :read_only => false, :if_unsaved => :accept)
        new_book.ReadOnly.should be_false 
        new_book.should be_alive
        book.should be_alive   
        new_book.should == book 
        new_sheet = new_book[0]
        new_cell_value = new_sheet[0,0].value
        new_cell_value.should == old_cell_value
      end

      it "should not raise an error when trying to reopen the book as read_only while the writable book had unsaved changes" do
        book = Book.open(@simple_file, :read_only => false)
        book.ReadOnly.should be_false
        book.should be_alive
        sheet = book[0]
        old_cell_value = sheet[0,0].value        
        sheet[0,0] = sheet[0,0].value == "simple" ? "complex" : "simple"
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
        old_cell_value = sheet[0,0].value
        sheet[0,0] = sheet[0,0].value == "simple" ? "complex" : "simple"
        book.Saved.should be_false
        new_book = Book.open(@simple_file, :if_unsaved => :accept, :force_excel => book.excel, :read_only => false)
        new_book.ReadOnly.should be_false 
        new_book.should be_alive
        book.should be_alive   
        new_book.should == book 
        new_sheet = new_book[0]
        new_cell_value = new_sheet[0,0].value
        new_cell_value.should == old_cell_value
      end

      it "should reopen the book with readonly (unsaved changes of the writable should be saved)" do
        book = Book.open(@simple_file, :force_excel => :new, :read_only => false)
        book.ReadOnly.should be_false
        book.should be_alive
        sheet = book[0]
        old_cell_value = sheet[0,0].value        
        sheet[0,0] = sheet[0,0].value == "simple" ? "complex" : "simple"
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
          Book.open(path, :if_absent => :raise)
        }.to raise_error(ExcelErrorOpen, "file #{path} not found")
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
        book2 = Book.open(@simple_file, :force_excel => @book.book_store.hidden_excel)
        book2.excel.should_not == @book.excel
        book2.excel.Visible.should be_false
        book2.excel.DisplayAlerts.should be_false
        book2.close 
      end
    end
  end

  describe "unobtrusively" do

    def unobtrusively_ok? # :nodoc: #
      Book.unobtrusively(@simple_file) do |book|
        book.should be_a Book
        sheet = book[0]
        sheet[0,0] = sheet[0,0].value == "simple" ? "complex" : "simple"
        book.should be_alive
        book.Saved.should be_false
      end
    end

    context "with no open book" do

      it "should open unobtrusively in a new Excel" do
        expect{ unobtrusively_ok? }.to_not raise_error
      end

      it "should open unobtrusively in the first opened Excel" do
        excel = Excel.new(:reuse => false)
        new_excel = Excel.new(:reuse => false)
        Book.unobtrusively(@simple_file, :if_closed => :reuse) do |book|
          book.should be_a Book
          book.should be_alive
          book.excel.should == excel
          book.excel.should_not == new_excel
        end
      end

      it "should open unobtrusively in a new Excel" do
        excel = Excel.new(:reuse => false)
        new_excel = Excel.new(:reuse => false)
        Book.unobtrusively(@simple_file, :if_closed => :hidden) do |book|
          book.should be_a Book
          book.should be_alive
          book.excel.should_not == excel
          book.excel.should_not == new_excel
        end
      end

      it "should open unobtrusively in a given Excel" do
        excel = Excel.new(:reuse => false)
        new_excel = Excel.new(:reuse => false)
        Book.unobtrusively(@simple_file, :if_closed => new_excel) do |book|
          book.should be_a Book
          book.should be_alive
          book.excel.should_not == excel
        end
      end

    end

    context "with an open book" do

      before do
        @book = Book.open(@simple_file)
      end

      after do
        @book.close(:if_unsaved => :forget)
        @book2.close(:if_unsaved => :forget) rescue nil
      end

      it "should let a saved book saved" do
        @book.Saved.should be_true
        @book.should be_alive
        sheet = @book[0]
        old_cell_value = sheet[0,0].value
        unobtrusively_ok?
        @book.Saved.should be_true
        @book.should be_alive
        sheet = @book[0]
        sheet[0,0].value.should_not == old_cell_value
      end

     it "should let the unsaved book unsaved" do
        sheet = @book[0]
        sheet[0,0] = sheet[0,0].value == "simple" ? "complex" : "simple"
        old_cell_value = sheet[0,0].value
        @book.Saved.should be_false
        unobtrusively_ok?
        @book.should be_alive
        @book.Saved.should be_false
        @book.close(:if_unsaved => :forget)
        @book2 = Book.open(@simple_file)
        sheet2 = @book2[0]
        sheet2[0,0].value.should_not == old_cell_value
      end

      it "should modify unobtrusively the second, writable book" do
        @book2 = Book.open(@simple_file, :force_excel => :new)
        @book.ReadOnly.should be_false
        @book2.ReadOnly.should be_true
        sheet = @book2[0]
        old_cell_value = sheet[0,0].value
        sheet[0,0] = sheet[0,0].value == "simple" ? "complex" : "simple"
        unobtrusively_ok?
        @book2.should be_alive
        @book2.Saved.should be_false
        @book2.close(:if_unsaved => :forget)
        @book.close
        @book = Book.open(@simple_file)
        sheet2 = @book[0]
        sheet2[0,0].value.should_not == old_cell_value
      end    
    end
    
    context "with a closed book" do
      
      before do
        @book = Book.open(@simple_file)
      end

      after do
        @book.close(:if_unsaved => :forget)
      end

      it "should let the closed book closed by default" do
        sheet = @book[0]
        old_cell_value = sheet[0,0].value
        @book.close
        @book.should_not be_alive
        unobtrusively_ok?
        @book.should_not be_alive
        @book = Book.open(@simple_file)
        sheet = @book[0]
        sheet[0,0].value.should_not == old_cell_value
      end

      # The bold reanimation of the @book
      it "should use the excel of the book and keep open the book" do
        excel = Excel.new(:reuse => false)
        sheet = @book[0]
        old_cell_value = sheet[0,0].value
        @book.close
        @book.should_not be_alive
        Book.unobtrusively(@simple_file, :if_closed => :reuse, :keep_open => true) do |book|
          book.should be_a Book
          book.excel.should == @book.excel
          book.excel.should_not == excel
          sheet = book[0]
          cell = sheet[0,0]
          sheet[0,0] = cell.value == "simple" ? "complex" : "simple"
          book.Saved.should be_false
        end
        @book.should be_alive
        @book.close
        new_book = Book.open(@simple_file)
        sheet = new_book[0]
        sheet[0,0].value.should_not == old_cell_value
      end

      # book shall be reanimated even with if_closed => hidden
      it "should use the excel of the book and keep open the book" do
        excel = Excel.new(:reuse => false)
        sheet = @book[0]
        old_cell_value = sheet[0,0].value
        @book.close
        @book.should_not be_alive
        Book.unobtrusively(@simple_file) do |book|
          book.should be_a Book
          book.should be_alive
          book.excel.should_not == @book.excel
          book.excel.should_not == excel
          sheet = book[0]
          cell = sheet[0,0]
          sheet[0,0] = cell.value == "simple" ? "complex" : "simple"
          book.Saved.should be_false
        end
        @book.should_not be_alive
        new_book = Book.open(@simple_file)
        sheet = new_book[0]
        sheet[0,0].value.should_not == old_cell_value
      end

      it "should use another excel if the Excels are closed" do
        excel = Excel.new(:reuse => false)
        sheet = @book[0]
        old_cell_value = sheet[0,0].value
        @book.close
        @book.should_not be_alive
        Excel.close_all
        Book.unobtrusively(@simple_file, :keep_open => true, :if_closed => :reuse) do |book|
          book.should be_a Book
          book.excel.should == @book.excel
          book.excel.should_not == excel
          sheet = book[0]
          cell = sheet[0,0]
          sheet[0,0] = cell.value == "simple" ? "complex" : "simple"
          book.Saved.should be_false
        end
        @book.should be_alive
        @book.close
        new_book = Book.open(@simple_file)
        sheet = new_book[0]
        sheet[0,0].value.should_not == old_cell_value
      end

      it "should use another excel if the Excels are closed" do
        excel = Excel.new(:reuse => false)
        sheet = @book[0]
        old_cell_value = sheet[0,0].value
        @book.close
        @book.should_not be_alive
        Excel.close_all
        Book.unobtrusively(@simple_file, :keep_open => true) do |book|
          book.should be_a Book
          book.excel.should_not == @book.excel
          book.excel.should_not == excel
          sheet = book[0]
          cell = sheet[0,0]
          sheet[0,0] = cell.value == "simple" ? "complex" : "simple"
          book.Saved.should be_false
        end
        @book.should_not be_alive
        new_book = Book.open(@simple_file)
        sheet = new_book[0]
        sheet[0,0].value.should_not == old_cell_value
      end

      it "should use a given Excel" do
        @book.close
        excel = Excel.new(:reuse => false)
        new_excel = Excel.new(:reuse => false)
        Book.unobtrusively(@simple_file, :if_closed => new_excel) do |book|
          book.should be_a Book
          book.excel.should_not == @book.excel
          book.excel.should == new_excel
        end
      end

      it "should use not the given excel if the Excels are closed" do
        excel = Excel.new(:reuse => false)
        sheet = @book[0]
        old_cell_value = sheet[0,0].value
        old_excel = @book.excel
        @book.close
        @book.should_not be_alive
        Excel.close_all
        Book.unobtrusively(@simple_file, :if_closed => excel, :keep_open => true) do |book|
          book.should be_a Book
          book.excel.should_not == old_excel
          book.excel.should == @book.excel
          book.excel.should_not == excel
          sheet = book[0]
          cell = sheet[0,0]
          sheet[0,0] = cell.value == "simple" ? "complex" : "simple"
          book.Saved.should be_false
        end
        @book.should be_alive
        new_book = Book.open(@simple_file)
        sheet = new_book[0]
        sheet[0,0].value.should_not == old_cell_value
      end
    end


    context "with a read_only book" do

      before do
        @book = Book.open(@simple_file, :read_only => true)
      end

      after do
        @book.close
      end

      it "should let the saved book saved" do
        @book.ReadOnly.should be_true
        @book.Saved.should be_true
        sheet = @book[0]
        old_cell_value = sheet[0,0].value
        unobtrusively_ok?
        @book.should be_alive
        @book.Saved.should be_true
        @book.ReadOnly.should be_true
        @book.close
        book2 = Book.open(@simple_file)
        sheet2 = book2[0]
        sheet2[0,0].value.should_not == old_cell_value
      end

      it "should let the unsaved book unsaved" do
        @book.ReadOnly.should be_true
        sheet = @book[0]
        sheet[0,0] = sheet[0,0].value == "simple" ? "complex" : "simple" 
        @book.Saved.should be_false
        @book.should be_alive
        cell_value = sheet[0,0].value
        unobtrusively_ok?
        @book.should be_alive
        @book.Saved.should be_false
        @book.ReadOnly.should be_true
        @book.close
        book2 = Book.open(@simple_file)
        sheet2 = book2[0]
        # modifies unobtrusively the saved version, not the unsaved version
        sheet2[0,0].value.should == cell_value        
      end

      it "should open unobtrusively by default the writable book" do
        book2 = Book.open(@simple_file, :force_excel => :new, :read_only => false)
        @book.ReadOnly.should be_true
        book2.Readonly.should be_false
        sheet = @book[0]
        cell_value = sheet[0,0].value
        Book.unobtrusively(@simple_file) do |book|
          book.should be_a Book
          book.excel.should == book2.excel
          book.excel.should_not == @book.excel
          sheet = book[0]
          sheet[0,0] = sheet[0,0].value == "simple" ? "complex" : "simple"
          book.should be_alive
          book.Saved.should be_false          
        end  
        @book.Saved.should be_true
        @book.ReadOnly.should be_true
        @book.close
        book2.close
        book3 = Book.open(@simple_file)
        new_sheet = book3[0]
        new_sheet[0,0].value.should_not == cell_value
        book3.close
      end

      it "should open unobtrusively by default the book in a new Excel such that the book is writable" do
        book2 = Book.open(@simple_file, :force_excel => :new, :read_only => true)
        @book.ReadOnly.should be_true
        book2.Readonly.should be_true
        sheet = @book[0]
        cell_value = sheet[0,0].value
        Book.unobtrusively(@simple_file) do |book|
          book.should be_a Book
          book.excel.should_not == book2.excel
          book.excel.should_not == @book.excel
          sheet = book[0]
          sheet[0,0] = sheet[0,0].value == "simple" ? "complex" : "simple"
          book.should be_alive
          book.Saved.should be_false          
        end  
        @book.Saved.should be_true
        @book.ReadOnly.should be_true
        @book.close
        book2.close
        book3 = Book.open(@simple_file)
        new_sheet = book3[0]
        new_sheet[0,0].value.should_not == cell_value
        book3.close
      end

      it "should open unobtrusively the book in a new Excel such that the book is writable" do
        book2 = Book.open(@simple_file, :force_excel => :new, :read_only => true)
        @book.ReadOnly.should be_true
        book2.Readonly.should be_true
        sheet = @book[0]
        cell_value = sheet[0,0].value
        Book.unobtrusively(@simple_file, :use_this => false) do |book|
          book.should be_a Book
          book.excel.should_not == book2.excel
          book.excel.should_not == @book.excel
          sheet = book[0]
          sheet[0,0] = sheet[0,0].value == "simple" ? "complex" : "simple"
          book.should be_alive
          book.Saved.should be_false          
        end  
        @book.Saved.should be_true
        @book.ReadOnly.should be_true
        @book.close
        book2.close
        book3 = Book.open(@simple_file)
        new_sheet = book3[0]
        new_sheet[0,0].value.should_not == cell_value
        book3.close
      end

      it "should open unobtrusively the book in a new Excel to open the book writable" do
        excel1 = Excel.new(:reuse => false)
        excel2 = Excel.new(:reuse => false)
        book2 = Book.open(@simple_file, :force_excel => :new, :read_only => true)
        @book.ReadOnly.should be_true
        book2.Readonly.should be_true
        sheet = @book[0]
        cell_value = sheet[0,0].value
        Book.unobtrusively(@simple_file, :use_readonly_excel => false) do |book|
          book.should be_a Book
          book.ReadOnly.should be_false
          book.excel.should_not == book2.excel
          book.excel.should_not == @book.excel
          book.excel.should_not == excel1
          book.excel.should_not == excel2
          sheet = book[0]
          sheet[0,0] = sheet[0,0].value == "simple" ? "complex" : "simple"
          book.should be_alive
          book.Saved.should be_false          
        end  
        @book.Saved.should be_true
        @book.ReadOnly.should be_true
        @book.close
        book2.close
        book3 = Book.open(@simple_file)
        new_sheet = book3[0]
        new_sheet[0,0].value.should_not == cell_value
        book3.close
      end

      it "should open unobtrusively the book in the same Excel to open the book writable" do
        excel1 = Excel.new(:reuse => false)
        excel2 = Excel.new(:reuse => false)
        book2 = Book.open(@simple_file, :force_excel => :new, :read_only => true)
        @book.ReadOnly.should be_true
        book2.Readonly.should be_true
        sheet = @book[0]
        cell_value = sheet[0,0].value
        Book.unobtrusively(@simple_file, :use_readonly_excel => true) do |book|
          book.should be_a Book
          book.excel.should == book2.excel
          book.ReadOnly.should be_false
          sheet = book[0]
          sheet[0,0] = sheet[0,0].value == "simple" ? "complex" : "simple"
          book.should be_alive
          book.Saved.should be_false          
        end  
        book2.Saved.should be_true
        book2.ReadOnly.should be_false
        @book.close
        book2.close
        book3 = Book.open(@simple_file)
        new_sheet = book3[0]
        new_sheet[0,0].value.should_not == cell_value
        book3.close
      end

      it "should open unobtrusively the book in the Excel where it was opened most recently" do
        book2 = Book.open(@simple_file, :force_excel => :new, :read_only => true)
        @book.ReadOnly.should be_true
        book2.Readonly.should be_true
        sheet = @book[0]
        cell_value = sheet[0,0].value
        Book.unobtrusively(@simple_file, :read_only => true) do |book|
          book.should be_a Book
          book.excel.should == book2.excel
          book.excel.should_not == @book.excel
          book.should be_alive
          book.Saved.should be_true         
        end  
        @book.Saved.should be_true
        @book.ReadOnly.should be_true
        @book.close
        book2.close
      end

    end

    context "with a virgin Book class" do
      before do
        class Book
          @@bookstore = nil
        end
      end
      it "should work" do
        expect{ unobtrusively_ok? }.to_not raise_error
      end
    end

    context "with a book never opened before" do
      before do
        class Book
          @@bookstore = nil
        end
        other_book = Book.open(@different_file)
      end
      it "should open the book" do
        expect{ unobtrusively_ok? }.to_not raise_error
      end
    end

    context "with a saved book" do

      before do
        @book1 = Book.open(@simple_file)
      end

      after do
        @book1.close(:if_unsaved => :forget)
      end

      it "should save if the book was modified during unobtrusively" do
        m_time = File.mtime(@book1.stored_filename)
        Book.unobtrusively(@simple_file) do |book|
          @book1.Saved.should be_true
          book.Saved.should be_true  
          sheet = book[0]
          cell = sheet[0,0]
          sheet[0,0] = cell.value == "simple" ? "complex" : "simple"
          @book1.Saved.should be_false
          book.Saved.should be_false
          sleep 1
        end
        @book1.Saved.should be_true
        m_time2 = File.mtime(@book1.stored_filename)
        m_time2.should_not == m_time
      end      

      it "should not save the book if it was not modified during unobtrusively" do
        m_time = File.mtime(@book1.stored_filename)
        Book.unobtrusively(@simple_file) do |book|
          @book1.Saved.should be_true
          book.Saved.should be_true 
          sleep 1
        end
        m_time2 = File.mtime(@book1.stored_filename)
        m_time2.should == m_time
      end            
    end

    context "with block result" do
      before do
        @book1 = Book.open(@simple_file)
      end

      after do
        @book1.close(:if_unsaved => :forget)
      end      

      it "should yield the block result true" do
        result = 
          Book.unobtrusively(@simple_file) do |book| 
            @book1.Saved.should be_true
          end
        result.should == true
      end

      it "should yield the block result nil" do
        result = 
          Book.unobtrusively(@simple_file) do |book| 
          end
        result.should == nil
      end

      it "should yield the block result with an unmodified book" do
        sheet1 = @book1[0]
        cell1 = sheet1[0,0].value
        result = 
          Book.unobtrusively(@simple_file) do |book| 
            sheet = book[0]
            cell = sheet[0,0].value
          end
        result.should == cell1
      end

      it "should yield the block result even if the book gets saved" do
        sheet1 = @book1[0]
        @book1.save
        result = 
          Book.unobtrusively(@simple_file) do |book| 
            sheet = book[0]
            sheet[0,0].value = 22
            @book1.Saved.should be_false
            42
          end
        result.should == 42
        @book1.Saved.should be_true
      end
    end

    context "with visible" do

      after do
        @book1.close
      end    

      it "should let the book invisible" do
        @book1 = Book.open(@simple_file)
        @book1.excel.Visible.should be_false
        Book.unobtrusively(@simple_file) do |book| 
          book.excel.Visible.should be_false
        end
        @book1.excel.Visible.should be_false
        Book.unobtrusively(@simple_file, :visible => false) do |book| 
          book.excel.Visible.should be_false
        end
        @book1.excel.Visible.should be_false
      end
      
      it "should let the book visible" do
        @book1 = Book.open(@simple_file, :visible => true)
        @book1.excel.Visible.should be_true
        Book.unobtrusively(@simple_file) do |book| 
          book.excel.Visible.should be_true
        end
        @book1.excel.Visible.should be_true
        Book.unobtrusively(@simple_file, :visible => true) do |book| 
          book.excel.Visible.should be_true
        end
        @book1.excel.Visible.should be_true
      end

    end

    context "with several Excel instances" do

      before do
        @book1 = Book.open(@simple_file)
        @book2 = Book.open(@simple_file, :force_excel => :new)
        @book1.Readonly.should == false
        @book2.Readonly.should == true
        old_sheet = @book1[0]
        @old_cell_value = old_sheet[0,0].value
        @book1.close
        @book2.close
        @book1.should_not be_alive
        @book2.should_not be_alive
      end

      it "should open unobtrusively the closed book in the most recent Excel where it was open before" do      
        Book.unobtrusively(@simple_file, :if_closed => :reuse) do |book| 
          book.excel.should == @book2.excel
          book.excel.should_not == @book1.excel
          book.ReadOnly.should == false
          sheet = book[0]
          cell = sheet[0,0]
          sheet[0,0] = cell.value == "simple" ? "complex" : "simple"
          book.Saved.should be_false
        end
        new_book = Book.open(@simple_file)
        sheet = new_book[0]
        sheet[0,0].value.should_not == @old_cell_value
      end

      it "should open unobtrusively the closed book in the new hidden Excel" do
        Book.unobtrusively(@simple_file) do |book| 
          book.excel.should_not == @book2.excel
          book.excel.should_not == @book1.excel
          book.ReadOnly.should == false
          sheet = book[0]
          cell = sheet[0,0]
          sheet[0,0] = cell.value == "simple" ? "complex" : "simple"
          book.Saved.should be_false
        end
        new_book = Book.open(@simple_file)
        sheet = new_book[0]
        sheet[0,0].value.should_not == @old_cell_value
      end

      it "should open unobtrusively the closed book in a new Excel if the Excel is not alive anymore" do
        Excel.close_all
        Book.unobtrusively(@simple_file) do |book| 
          book.ReadOnly.should == false
          book.excel.should_not == @book1.excel
          book.excel.should_not == @book2.excel
          sheet = book[0]
          cell = sheet[0,0]
          sheet[0,0] = cell.value == "simple" ? "complex" : "simple"
          book.Saved.should be_false
        end
        new_book = Book.open(@simple_file)
        sheet = new_book[0]
        sheet[0,0].value.should_not == @old_cell_value
      end
    end

    context "with :hidden" do

      before do
        @book1 = Book.open(@simple_file)
        @book1.close
      end
    
      it "should create a new hidden Excel instance" do
        Book.unobtrusively(@simple_file, :if_closed => :hidden) do |book| 
          book.should be_a Book
          book.should be_alive
          book.excel.Visible.should be_false
          book.excel.DisplayAlerts.should be_false
        end
      end

      it "should create a new hidden Excel instance and use this afterwards" do
        hidden_excel = nil
        Book.unobtrusively(@simple_file, :if_closed => :hidden) do |book| 
          book.should be_a Book
          book.should be_alive
          book.excel.Visible.should be_false
          book.excel.DisplayAlerts.should be_false
          hidden_excel = book.excel
        end
        Book.unobtrusively(@different_file, :if_closed => :hidden) do |book| 
          book.should be_a Book
          book.should be_alive
          book.excel.Visible.should be_false
          book.excel.DisplayAlerts.should be_false
          book.excel.should == hidden_excel
        end
      end

      it "should create a new hidden Excel instance if the Excel is closed" do
        Excel.close_all
        Book.unobtrusively(@simple_file, :if_closed => :hidden) do |book| 
          book.should be_a Book
          book.should be_alive
          book.excel.Visible.should be_false
          book.excel.DisplayAlerts.should be_false
          book.excel.should_not == @book1.excel
        end
      end

      it "should exclude hidden Excel when reuse in unobtrusively" do
        hidden_excel = nil
        Book.unobtrusively(@simple_file, :if_closed => :hidden) do |book| 
          book.should be_a Book
          book.should be_alive
          book.excel.Visible.should be_false
          book.excel.DisplayAlerts.should be_false
          book.excel.should_not == @book1.excel
          hidden_excel = book.excel
        end
        Book.unobtrusively(@simple_file, :if_closed => :reuse) do |book| 
          book.should be_a Book
          book.should be_alive
          book.excel.Visible.should be_false
          book.excel.DisplayAlerts.should be_false
          book.excel.should_not == hidden_excel
        end
      end

      it "should exclude hidden Excel when reuse in open" do
        hidden_excel = nil
        Book.unobtrusively(@simple_file, :if_closed => :hidden) do |book| 
          book.should be_a Book
          book.should be_alive
          book.excel.Visible.should be_false
          book.excel.DisplayAlerts.should be_false
          book.excel.should_not == @book1.excel
          hidden_excel = book.excel
        end
        book2 = Book.open(@simple_file, :default_excel => :reuse)
        book2.excel.should_not == hidden_excel
      end

      it "should exclude hidden Excel when reuse in open" do
        book1 = Book.open(@simple_file)
        book1.close
        book2 = Book.open(@simple_file, :default_excel => :reuse)
        book2.excel.should == book1.excel
        book1.should be_alive
        book2.close
      end
    end
  end

  describe "nvalue" do
    context "with standard" do
      before do
        @book1 = Book.open(@more_simple_file)
      end

      after do
        @book1.close
      end   

      it "should return value of a range" do
        @book1.nvalue("new").should == "foo"
        @book1.nvalue("one").should == 1
        @book1.nvalue("firstrow").should == [[1,2]]        
        @book1.nvalue("four").should == [[1,2],[3,4]]
        @book1.nvalue("firstrow").should_not == "12"
      end

      it "should raise an error if name not defined" do
        expect {
          value = @book1.nvalue("foo")
        }.to raise_error(ExcelErrorNValue, "name foo not in more_workbook.xls")
      end

      it "should raise an error if name was defined but contents is calcuated" do
        expect {
          value = @book1.nvalue("named_formula")
        }.to raise_error(ExcelErrorNValue, "range error in more_workbook.xls")
      end
    end
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
        }.to raise_error(ExcelErrorClose, "book is unsaved (#{File.basename(@simple_file)})")
      end

      it "should raise error by default" do
        expect{
          @book.close(:if_unsaved => :raise)
        }.to raise_error(ExcelErrorClose, "book is unsaved (#{File.basename(@simple_file)})")
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

  describe "save" do

    context "with simple save" do
      
      it "should save for a file opened without :read_only" do
        @book = Book.open(@simple_file)
        @book.add_sheet(@sheet, :as => 'a_name')
        @new_sheet_count = @book.workbook.Worksheets.Count
        expect {
          @book.save
        }.to_not raise_error
        @book.workbook.Worksheets.Count.should ==  @new_sheet_count
        @book.close
      end

      it "should raise error with read_only" do
        @book = Book.open(@simple_file, :read_only => true)
        expect {
          @book.save
        }.to raise_error(ExcelErrorSave, "Not opened for writing (opened with :read_only option)")
        @book.close
      end

    end

    context "with open with read only" do
      before do
        @book = Book.open(@simple_file, :read_only => true)
      end

      after do
        @book.close
      end

      it {
        expect {
          @book.save_as(@simple_file)
        }.to raise_error(IOError,
                     "Not opened for writing(open with :read_only option)")
      }
    end

    context "with argument" do
      before do
        Book.open(@simple_file) do |book|
          book.save_as(@simple_save_file, :if_exists => :overwrite)
        end
      end

      it "should save to 'simple_save_file.xlsx'" do
        File.exist?(@simple_save_file).should be_true
      end
    end

    context "with different extensions" do
      before do
        @book = Book.open(@simple_file)
      end

      after do
        @book.close
      end

      possible_extensions = ['xls', 'xlsm', 'xlsx']
      possible_extensions.each do |extensions_value|
        it "should save to 'simple_save_file.#{extensions_value}'" do
          simple_save_file = @dir + '/simple_save_file.' + extensions_value
          File.delete simple_save_file rescue nil
          @book.save_as(simple_save_file, :if_exists => :overwrite)
          File.exist?(simple_save_file).should be_true
          new_book = Book.open(simple_save_file)
          new_book.should be_a Book
          new_book.close
        end
      end
    end

    # options :overwrite, :raise, :excel, no option, invalid option
    possible_displayalerts = [true, false]
    possible_displayalerts.each do |displayalert_value|
      context "with displayalerts=#{displayalert_value}" do
        before do
          @book = Book.open(@simple_file, :displayalerts => displayalert_value)
        end

        after do
          @book.close
        end

        it "should raise an error if the book is open" do
          File.delete @simple_save_file rescue nil
          FileUtils.copy @simple_file, @simple_save_file
          book_save = Book.open(@simple_save_file, :excel => :new)
          expect{
            @book.save_as(@simple_save_file, :if_exists => :overwrite)
            }.to raise_error(ExcelErrorSave, "book is open and used in Excel")
          book_save.close
        end

        it "should save to simple_save_file.xls with :if_exists => :overwrite" do
          File.delete @simple_save_file rescue nil
          File.open(@simple_save_file,"w") do | file |
            file.puts "garbage"
          end
          @book.save_as(@simple_save_file, :if_exists => :overwrite)
          File.exist?(@simple_save_file).should be_true
          new_book = Book.open(@simple_save_file)
          new_book.should be_a Book
          new_book.close
        end
        it "should save to 'simple_save_file.xls' with :if_exists => :raise" do
          dirname, basename = File.split(@simple_save_file)
          File.delete @simple_save_file rescue nil
          File.open(@simple_save_file,"w") do | file |
            file.puts "garbage"
          end
          File.exist?(@simple_save_file).should be_true
          booklength = File.size?(@simple_save_file)
          expect {
            @book.save_as(@simple_save_file, :if_exists => :raise)
            }.to raise_error(ExcelErrorSave, 'book already exists: ' + basename)
          File.exist?(@simple_save_file).should be_true
          File.size?(@simple_save_file).should == booklength
        end

        context "with :if_exists => :alert" do
          before do
            File.delete @simple_save_file rescue nil
            File.open(@simple_save_file,"w") do | file |
              file.puts "garbage"
            end
            @garbage_length = File.size?(@simple_save_file)
            @key_sender = IO.popen  'ruby "' + File.join(File.dirname(__FILE__), '/helpers/key_sender.rb') + '" "Microsoft Excel" '  , "w"
          end

          after do
            @key_sender.close
          end

          it "should save if user answers 'yes'" do
            # "Yes" is to the left of "No", which is the  default. --> language independent
            @key_sender.puts "{left}{enter}" #, :initial_wait => 0.2, :if_target_missing=>"Excel window not found")
            @book.save_as(@simple_save_file, :if_exists => :alert)
            File.exist?(@simple_save_file).should be_true
            File.size?(@simple_save_file).should > @garbage_length
            @book.excel.DisplayAlerts.should == displayalert_value
            new_book = Book.open(@simple_save_file, :excel => :new)
            new_book.should be_a Book
            new_book.close
            @book.excel.DisplayAlerts.should == displayalert_value
          end

          it "should not save if user answers 'no'" do
            # Just give the "Enter" key, because "No" is the default. --> language independent
            # strangely, in the "no" case, the question will sometimes be repeated three times
            @key_sender.puts "{enter}"
            @key_sender.puts "{enter}"
            @key_sender.puts "{enter}"
            #@key_sender.puts "%{n}" #, :initial_wait => 0.2, :if_target_missing=>"Excel window not found")
            expect{
              @book.save_as(@simple_save_file, :if_exists => :alert)
              }.to raise_error(ExcelErrorSave, "not saved or canceled by user")
            File.exist?(@simple_save_file).should be_true
            File.size?(@simple_save_file).should == @garbage_length
            @book.excel.DisplayAlerts.should == displayalert_value
          end

          it "should not save if user answers 'cancel'" do
            # 'Cancel' is right from 'yes'
            # strangely, in the "no" case, the question will sometimes be repeated three times
            @key_sender.puts "{right}{enter}"
            @key_sender.puts "{right}{enter}"
            @key_sender.puts "{right}{enter}"
            #@key_sender.puts "%{n}" #, :initial_wait => 0.2, :if_target_missing=>"Excel window not found")
            expect{
              @book.save_as(@simple_save_file, :if_exists => :alert)
              }.to raise_error(ExcelErrorSave, "not saved or canceled by user")
            File.exist?(@simple_save_file).should be_true
            File.size?(@simple_save_file).should == @garbage_length
            @book.excel.DisplayAlerts.should == displayalert_value
          end

          it "should report save errors and leave DisplayAlerts unchanged" do
            #@key_sender.puts "{left}{enter}" #, :initial_wait => 0.2, :if_target_missing=>"Excel window not found")
            @book.workbook.Close
            expect{
              @book.save_as(@simple_save_file, :if_exists => :alert)
              }.to raise_error(ExcelErrorSaveUnknown)
            File.exist?(@simple_save_file).should be_true
            File.size?(@simple_save_file).should == @garbage_length
            @book.excel.DisplayAlerts.should == displayalert_value
          end

        end

        it "should save to 'simple_save_file.xls' with :if_exists => nil" do
          dirname, basename = File.split(@simple_save_file)
          File.delete @simple_save_file rescue nil
          File.open(@simple_save_file,"w") do | file |
            file.puts "garbage"
          end
          File.exist?(@simple_save_file).should be_true
          booklength = File.size?(@simple_save_file)
          expect {
            @book.save_as(@simple_save_file)
            }.to raise_error(ExcelErrorSave, 'book already exists: ' + basename)
          File.exist?(@simple_save_file).should be_true
          File.size?(@simple_save_file).should == booklength
        end

        it "should save to 'simple_save_file.xls' with :if_exists => :invalid_option" do
          File.delete @simple_save_file rescue nil
          @book.save_as(@simple_save_file)
          expect {
            @book.save_as(@simple_save_file, :if_exists => :invalid_option)
            }.to raise_error(ExcelErrorSave, ':if_exists: invalid option')
        end
      end
    end
  end

  describe "== , alive?, filename, visible, empty_workbook" do

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
        @new_book = Book.new(@simple_file_other_path, :excel => :new)
        @new_book.should_not == @book
      end

      it "should be false with same book names but different excel instances" do
        @new_book = Book.new(@simple_file, :excel => :new)
        @new_book.should_not == @book
      end

      it "should be false with non-Books" do
        @book.should_not == "hallo"
        @book.should_not == 7
        @book.should_not == nil
      end
    end

    context "with visible and displayalerts" do

      before do
        @book = Book.open(@simple_file)
      end

      after do
        @book.close
      end

      it "should make Excel visible" do
        @book.excel.visible = false
        @book.excel.visible.should be_false
        @book.excel.visible = true
        @book.excel.visible.should be_true
      end

      it "should enable DisplayAlerts in Excel" do
        @book.excel.displayalerts = false
        @book.excel.displayalerts.should be_false
        @book.excel.displayalerts = true
        @book.excel.displayalerts.should be_true
      end

    end

  end

  describe "#add_sheet" do
    before do
      @book = Book.open(@simple_file)
      @sheet = @book[0]
    end

    after do
      @book.close(:if_unsaved => :forget)
    end

    context "only first argument" do
      it "should add worksheet" do
        expect { @book.add_sheet @sheet }.to change{ @book.workbook.Worksheets.Count }.from(3).to(4)
      end

      it "should return copyed sheet" do
        sheet = @book.add_sheet @sheet
        copyed_sheet = @book.workbook.Worksheets.Item(@book.workbook.Worksheets.Count)
        sheet.name.should eq copyed_sheet.name
      end
    end

    context "with first argument" do
      context "with second argument is {:as => 'copyed_name'}" do
        it "copyed sheet name should be 'copyed_name'" do
          @book.add_sheet(@sheet, :as => 'copyed_name').name.should eq 'copyed_name'
        end
      end

      context "with second argument is {:before => @sheet}" do
        it "should add the first sheet" do
          @book.add_sheet(@sheet, :before => @sheet).name.should eq @book[0].name
        end
      end

      context "with second argument is {:after => @sheet}" do
        it "should add the first sheet" do
          @book.add_sheet(@sheet, :after => @sheet).name.should eq @book[1].name
        end
      end

      context "with second argument is {:before => @book[2], :after => @sheet}" do
        it "should arguments in the first is given priority" do
          @book.add_sheet(@sheet, :before => @book[2], :after => @sheet).name.should eq @book[2].name
        end
      end

    end

    context "without first argument" do
      context "second argument is {:as => 'new sheet'}" do
        it "should return new sheet" do
          @book.add_sheet(:as => 'new sheet').name.should eq 'new sheet'
        end
      end

      context "second argument is {:before => @sheet}" do
        it "should add the first sheet" do
          @book.add_sheet(:before => @sheet).name.should eq @book[0].name
        end
      end

      context "second argument is {:after => @sheet}" do
        it "should add the second sheet" do
          @book.add_sheet(:after => @sheet).name.should eq @book[1].name
        end
      end
    end

    context "without argument" do
      it "should add empty sheet" do
        expect { @book.add_sheet }.to change{ @book.workbook.Worksheets.Count }.from(3).to(4)
      end

      it "should return copyed sheet" do
        sheet = @book.add_sheet
        copyed_sheet = @book.workbook.Worksheets.Item(@book.workbook.Worksheets.Count)
        sheet.name.should eq copyed_sheet.name
      end
    end

    context "should raise error if the sheet name already exists" do
      it "should raise error with giving a name that already exists" do
        @book.add_sheet(@sheet, :as => 'new_sheet')
        expect{
          @book.add_sheet(@sheet, :as => 'new_sheet')
          }.to raise_error(ExcelErrorSheet, "sheet name already exists")
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
      @book['Sheet1'].should be_kind_of Sheet
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
end
