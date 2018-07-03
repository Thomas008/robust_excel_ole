# -*- coding: utf-8 -*-

require File.join(File.dirname(__FILE__), './spec_helper')

$VERBOSE = nil

include General

module RobustExcelOle

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
    @simple_file_other_path1 = @simple_file_other_path
    @simple_save_file1 = @simple_save_file
  end

  after do
    Excel.kill_all
    #rm_tmp(@dir)
  end

  describe "create file" do
    context "with standard" do
      it "open an existing file" do
        expect {
          @book = Book.open(@simple_file)
        }.to_not raise_error
        @book.should be_a Book
        @book.close
      end
    end
  end

  describe "Book::save" do

    it "should save a file, if it is open" do
      @book = Book.open(@simple_file)
      @book.add_sheet(@sheet, :as => 'a_name')
      @new_sheet_count = @book.ole_workbook.Worksheets.Count
      expect {
        Book.save(@simple_file)
      }.to_not raise_error
      @book.ole_workbook.Worksheets.Count.should ==  @new_sheet_count
      @book.close
    end

    it "should not save a file, if it is not open" do
      @book = Book.open(@simple_file)
      @book.add_sheet(@sheet, :as => 'a_name')
      @new_sheet_count = @book.ole_workbook.Worksheets.Count
      @book.close(:if_unsaved => :forget)
      expect {
        Book.save(@simple_file)
      }.to_not raise_error
    end

  end

  describe "Book::save_as" do
    
    it "should save to 'simple_save_file.xls'" do
      book = Book.open(@simple_file1)
      Book.save_as(@simple_file1, @simple_save_file1, :if_exists => :overwrite)
      File.exist?(@simple_save_file1).should be_true
    end
  end

  describe "Book::close" do

    it "should close the book if it is open" do
      book = Book.open(@simple_file1)
      Book.close(@simple_file1)
      book.should_not be_alive
    end

    it "should not close the book if it is not open" do
      book = Book.open(@simple_file1, :visible => true)
      book.close
      Book.close(@simple_file1)
      book.should_not be_alive
    end

    it "should raise error if the book is unsaved and open" do
      book = Book.open(@simple_file1)
      sheet = book.sheet(1)
      book.add_sheet(sheet, :as => 'a_name')
      expect{
        Book.close(@simple_file1)
      }.to raise_error(WorkbookNotSaved, /workbook is unsaved: "workbook.xls"/)
      expect{
        Book.close(@simple_file, :if_unsaved => :raise)
      }.to raise_error(WorkbookNotSaved, /workbook is unsaved: "workbook.xls"/)
    end

    it "should save and close the book" do
      book = Book.open(@simple_file1)
      sheet_count = book.ole_workbook.Worksheets.Count
      sheet = book.sheet(1)
      book.add_sheet(sheet, :as => 'a_name')
      ole_workbook = book.ole_workbook
      excel = book.excel
      excel.Workbooks.Count.should == 1
      Book.close(@simple_file1, {:if_unsaved => :save})
      excel.Workbooks.Count.should == 0
      book.ole_workbook.should == nil
      book.should_not be_alive
      expect{ole_workbook.Name}.to raise_error(WIN32OLERuntimeError)
      new_book = Book.open(@simple_file1)
      begin
        new_book.ole_workbook.Worksheets.Count.should == sheet_count + 1
      ensure
        new_book.close
      end
    end
  end

  describe "open" do

    context "with various file formats" do

      it "should open linked workbook" do
        book = Book.open(@linked_file, :visible => true)
        book.close
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

      it "should yield an identical Book and set visible and displayalerts values" do
        workbook = @book.ole_workbook
        new_book = Book.new(workbook, :visible => true)
        new_book.excel.displayalerts = true
        new_book.should be_a Book
        new_book.should be_alive
        new_book.should == @book
        new_book.filename.should == @book.filename
        new_book.excel.should == @book.excel
        new_book.should === @book
        new_book.excel.Visible.should be_true
        new_book.excel.DisplayAlerts.should be_true
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

      it "should yield identical Book objects for identical Excel books when reopening" do
        @book.should be_alive
        @book.close
        @book.should_not be_alive
        book2 = Book.open(@simple_file)
        book2.should === @book
        book2.should be_alive
        book2.close
      end

      it "should yield different Book objects when reopening in a new Excel" do
        @book.should be_alive
        old_excel = @book.excel
        @book.close
        @book.should_not be_alive
        book2 = Book.open(@simple_file, :force => {:excel => :new})
        book2.should_not === @book
        book2.should be_alive
        book2.excel.should_not == old_excel
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
        book2 = Book.open(@simple_file, :force => {:excel => :new})
        book2.should be_alive
        book2.should be_a Book
        book2.excel.should_not == @book.excel 
        book2.should_not == @book
        @book.Readonly.should be_false
        book2.Readonly.should be_true
        book2.close
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
        book4.ReadOnly.should be_true
        book4.should == book2
        book4.close
        book5 = Book.open(@simple_file1, :force => {:excel => book2}, :read_only => true)
        book5.should be_alive
        book5.should be_a Book
        book5.excel.should == book2.excel
        book5.ReadOnly.should be_true
        book5.should == book2
        book5.close
        book3.close
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

      it "should open in a given Excel provided as Excel, Book, or WIN32OLE representing an Excel or Workbook" do
        book2 = Book.open(@another_simple_file)
        different_file1 = @different_file
        book3 = Book.open(different_file1, :default => {:excel => book2.excel})
        book3.excel.should === book2.excel
        book3.close
        book4 = Book.open(different_file1, :default => {:excel => book2}) 
        book4.excel.should === book2.excel
        book4.close
        book5 = Book.open(different_file1, :default_excel => book2.ole_workbook)
        book5.excel.should ===  book2.excel
        book5.close
        win32ole_excel1 = WIN32OLE.connect(book2.ole_workbook.Fullname).Application
        book6 = Book.open(different_file1, :default => {:excel => win32ole_excel1})
        book6.excel.should === book2.excel
        book6.close
      end


    end

    context "with :if_unsaved" do

      before do
        @book = Book.open(@simple_file)
        @sheet = @book.sheet(1)
        @book.add_sheet(@sheet, :as => 'a_name')
      end

      after do
        @book.close(:if_unsaved => :forget)
        @new_book.close rescue nil
      end

      it "should raise an error, if :if_unsaved is :raise" do
        expect {
          @new_book = Book.open(@simple_file, :if_unsaved => :raise)
        }.to raise_error(WorkbookNotSaved, /workbook is already open but not saved: "workbook.xls"/)
      end

      it "should let the book open, if :if_unsaved is :accept" do
        expect {
          @new_book = Book.open(@simple_file, :if_unsaved => :accept)
          }.to_not raise_error
        @book.should be_alive
        @new_book.should be_alive
        @new_book.should == @book
      end

      context "with :if_unsaved => :alert or :if_unsaved => :excel" do
        before do
         @key_sender = IO.popen  'ruby "' + File.join(File.dirname(__FILE__), '/helpers/key_sender.rb') + '" "Microsoft Office Excel" '  , "w"
        end

        after do
          @key_sender.close
        end

        it "should open the new book and close the unsaved book, if user answers 'yes'" do
          # "Yes" is the  default. --> language independent
          @key_sender.puts "{enter}"
          @new_book = Book.open(@simple_file1, :if_unsaved => :alert)
          @new_book.should be_alive
          @new_book.filename.downcase.should == @simple_file1.downcase
          @book.should_not be_alive
        end

        it "should not open the new book and not close the unsaved book, if user answers 'no'" do
          # "No" is right to "Yes" (the  default). --> language independent
          # strangely, in the "no" case, the question will sometimes be repeated three times
          #@book.excel.Visible = true
          @key_sender.puts "{right}{enter}"
          @key_sender.puts "{right}{enter}"
          @key_sender.puts "{right}{enter}"
          #expect{
          #  Book.open(@simple_file, :if_unsaved => :alert)
          #  }.to raise_error(ExcelREOError, /user canceled or runtime error/)
          @book.should be_alive
        end

        it "should open the new book and close the unsaved book, if user answers 'yes'" do
          # "Yes" is the  default. --> language independent
          @key_sender.puts "{enter}"
          @new_book = Book.open(@simple_file1, :if_unsaved => :excel)
          @new_book.should be_alive
          @new_book.filename.downcase.should == @simple_file1.downcase
          #@book.should_not be_alive
        end

        it "should not open the new book and not close the unsaved book, if user answers 'no'" do
          # "No" is right to "Yes" (the  default). --> language independent
          # strangely, in the "no" case, the question will sometimes be repeated three times
          #@book.excel.Visible = true
          @key_sender.puts "{right}{enter}"
          @key_sender.puts "{right}{enter}"
          @key_sender.puts "{right}{enter}"
          #expect{
          #  Book.open(@simple_file, :if_unsaved => :excel)
          #  }.to raise_error(ExcelREOError, /user canceled or runtime error/)
          @book.should be_alive
        end

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
            @book = Book.open(@simple_file_other_path1)
            @sheet_count = @book.ole_workbook.Worksheets.Count
            @sheet = @book.sheet(1)
            @book.add_sheet(@sheet, :as => 'a_name')
          end

          after do
            @book.close(:if_unsaved => :forget)
            @new_book.close rescue nil
          end

          it "should save the old book, close it, and open the new book, if :if_obstructed is :save" do
            @new_book = Book.open(@simple_file1, :if_obstructed => :save)
            @book.should_not be_alive
            @new_book.should be_alive
            @new_book.filename.downcase.should == @simple_file1.downcase
            old_book = Book.open(@simple_file_other_path1, :if_obstructed => :forget)
            old_book.ole_workbook.Worksheets.Count.should ==  @sheet_count + 1
            old_book.close
          end

          it "should raise an error, if the old book is unsaved, and close the old book and open the new book, 
              if :if_obstructed is :close_if_saved" do
            expect{
              @new_book = Book.open(@simple_file1, :if_obstructed => :close_if_saved)
            }.to raise_error(WorkbookBlocked, /workbook with the same name in a different path is unsaved/)
            @book.save
            @new_book = Book.open(@simple_file1, :if_obstructed => :close_if_saved)
            @book.should_not be_alive
            @new_book.should be_alive
            @new_book.filename.downcase.should == @simple_file1.downcase
            old_book = Book.open(@simple_file_other_path1, :if_obstructed => :forget)
            old_book.ole_workbook.Worksheets.Count.should ==  @sheet_count + 1
            old_book.close
          end
        end
      end
    end

    context "with non-existing file" do

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
        }.to raise_error(FileNotFound, "file #{General::absolute_path(@simple_save_file).gsub("/","\\").inspect} not found")
      end
    end

    context "with :read_only" do
      
      it "should reopen the book with writable (unsaved changes from readonly will not be saved)" do
        book = Book.open(@simple_file1, :read_only => true)
        book.ReadOnly.should be_true
        book.should be_alive
        sheet = book.sheet(1)
        old_cell_value = sheet[1,1].value
        sheet[1,1] = sheet[1,1].value == "foo" ? "bar" : "foo"
        book.Saved.should be_false
        new_book = Book.open(@simple_file1, :read_only => false, :if_unsaved => :accept)
        new_book.ReadOnly.should be_false 
        new_book.should be_alive
        book.should be_alive   
        new_book.should == book 
        new_sheet = new_book.sheet(1)
        new_cell_value = new_sheet[1,1].value
        new_cell_value.should == old_cell_value
      end

    context "with block" do
      it 'block parameter should be instance of Book' do
        Book.open(@simple_file) do |book|
          book.should be_a Book
        end
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

  describe "visible" do

    it "should preserve :visible if they are not set" do
      excel1 = Excel.create(:visible => true)
      book1 = Book.open(@simple_file)
      book1.excel.Visible.should be_true
      book1.close
    end

    it "should preserve :visible if they are not set" do
      excel1 = Excel.create
      book1 = Book.open(@simple_file, :visible => true)
      book1.excel.Visible.should be_true
    end

    it "should preserve :visible if they are not set" do
      excel1 = Excel.create(:visible => true)
      book1 = Book.open(@different_file, :default => {:excel => :new})
      book1.excel.Visible.should be_false
    end

    it "should preserve :visible if they are not set" do
      excel1 = Excel.create(:visible => true)
      excel2 = Excel.create(:visible => true)
      book1 = Book.open(@different_file, :force => {:excel => excel2})
      book1.excel.Visible.should be_true
      book1.close
    end

    it "should let an open Book open" do
      @book = Book.open(@simple_file, :visible => true)
      Book.unobtrusively(@simple_file) do |book|
        book.should be_a Book
        book.should be_alive
        book.excel.should == @book.excel
        book.excel.Visible.should be_true
      end        
      @book.should be_alive
      @book.should be_a Book
      @book.excel.Visible.should be_true
      @book.close(:if_unsaved => :forget)
      @book2.close(:if_unsaved => :forget) rescue nil
    end
  end


  describe "unobtrusively" do

    def unobtrusively_ok? # :nodoc: #
      Book.unobtrusively(@simple_file) do |book|
        book.should be_a Book
        sheet = book.sheet(1)
        sheet[1,1] = sheet[1,1].value == "foo" ? "bar" : "foo"
        book.should be_alive
        book.Saved.should be_false
      end
    end

    context "with no open book" do

      it "should open unobtrusively in a new Excel" do
        expect{ unobtrusively_ok? }.to_not raise_error
      end
    end

    context "with two running excel instances" do
      before :all do
        Excel.close_all
      end

      before do
        @excel1 = Excel.new(:reuse => false)
        @excel2 = Excel.new(:reuse => false)
      end

      after do
        #Excel.close_all
        begin
          @excel1.close
          @excel2.close 
        rescue ExcelREOError => msg
          # puts "ExcelREOError: #{msg.message}" if msg.message =~ /Excel instance not alive or damaged/
        end
      end

      it "should open unobtrusively in a new Excel" do
        Book.unobtrusively(@simple_file, :if_closed => :current) do |book|
          book.should be_a Book
          book.should be_alive
            book.excel.should == @excel1
            book.excel.should_not == @excel2
          end
      end

      it "should open unobtrusively in a given Excel" do
        Book.unobtrusively(@simple_file, :if_closed => @excel2) do |book|
          book.should be_a Book
          book.should be_alive
            book.excel.should_not == @excel1
            book.excel.should     == @excel2
        end
      end
    end
  
    context "with an open book" do

      before do
        @book = Book.open(@simple_file1)
      end

      after do
        @book.close(:if_unsaved => :forget)
        @book2.close(:if_unsaved => :forget) rescue nil
      end

      it "should let an open Book open if two books have been opened and one has been closed and opened again" do
        book2 = Book.open(@different_file, :force => {:excel => :new})
        @book.close
        book2.close
        @book.reopen
        Book.unobtrusively(@simple_file1) do |book|
          book.should be_a Book
          book.should be_alive
          book.excel.should == @book.excel
        end        
        @book.should be_alive
        @book.should be_a Book
      end

      it "should let a saved book saved" do
        @book.Saved.should be_true
        @book.should be_alive
        sheet = @book.sheet(1)
        old_cell_value = sheet[1,1].value
        unobtrusively_ok?
        @book.Saved.should be_true
        @book.should be_alive
        sheet = @book.sheet(1)
        sheet[1,1].value.should_not == old_cell_value
      end

     it "should let the unsaved book unsaved" do
        sheet = @book.sheet(1)
        sheet[1,1] = sheet[1,1].value == "foo" ? "bar" : "foo"
        old_cell_value = sheet[1,1].value
        @book.Saved.should be_false
        unobtrusively_ok?
        @book.should be_alive
        @book.Saved.should be_false
        @book.close(:if_unsaved => :forget)
        @book2 = Book.open(@simple_file1)
        sheet2 = @book2.sheet(1)
        sheet2[1,1].value.should_not == old_cell_value
      end
    end
    
    context "with a closed book" do
      
      before do
        @book = Book.open(@simple_file1)
      end

      after do
        @book.close(:if_unsaved => :forget)
      end

      it "should let the closed book closed by default" do
        sheet = @book.sheet(1)
        old_cell_value = sheet[1,1].value
        @book.close
        @book.should_not be_alive
        unobtrusively_ok?
        @book.should_not be_alive
        book2 = Book.open(@simple_file1)
        sheet = book2.sheet(1)
        sheet[1,1].Value.should_not == old_cell_value
      end

      # The bold reanimation of the @book
      it "should use the excel of the book and keep open the book" do
        excel = Excel.new(:reuse => false)
        sheet = @book.sheet(1)
        old_cell_value = sheet[1,1].value
        @book.close
        @book.should_not be_alive
        Book.unobtrusively(@simple_file1, :keep_open => true) do |book|
          book.should be_a Book
          book.excel.should == @book.excel
          book.excel.should_not == excel
          sheet = book.sheet(1)
          cell = sheet[1,1]
          sheet[1,1] = cell.value == "foo" ? "bar" : "foo"
          book.Saved.should be_false
        end
        @book.should be_alive
        @book.close
        new_book = Book.open(@simple_file1)
        sheet = new_book.sheet(1)
        sheet[1,1].value.should_not == old_cell_value
      end

      it "should use the excel of the book and keep open the book" do
        excel = Excel.new(:reuse => false)
        sheet = @book.sheet(1)
        old_cell_value = sheet[1,1].value
        @book.close
        @book.should_not be_alive
        Book.unobtrusively(@simple_file1, :if_closed => :current) do |book|
          book.should be_a Book
          book.should be_alive
          book.excel.should == @book.excel
          book.excel.should_not == excel
          sheet = book.sheet(1)
          cell = sheet[1,1]
          sheet[1,1] = cell.value == "foo" ? "bar" : "foo"
          book.Saved.should be_false
        end
        @book.should_not be_alive
        new_book = Book.open(@simple_file1)
        sheet = new_book.sheet(1)
        sheet[1,1].value.should_not == old_cell_value
      end
    end

    
    context "with a read_only book" do

      before do
        @book = Book.open(@simple_file1, :read_only => true)
      end

      after do
        @book.close
      end

      it "should open unobtrusively the book in a new Excel such that the book is writable" do
        book2 = Book.open(@simple_file1, :force => {:excel => :new}, :read_only => true)
        @book.ReadOnly.should be_true
        book2.Readonly.should be_true
        sheet = @book.sheet(1)
        cell_value = sheet[1,1].value
        Book.unobtrusively(@simple_file1, :rw_change_excel => :new, :if_closed => :current, :writable => true) do |book|
          book.should be_a Book
          book.excel.should_not == book2.excel
          book.excel.should_not == @book.excel
          sheet = book.sheet(1)
          sheet[1,1] = sheet[1,1].value == "foo" ? "bar" : "foo"
          book.should be_alive
          book.Saved.should be_false          
        end  
        @book.Saved.should be_true
        @book.ReadOnly.should be_true
        @book.close
        book2.close
        book3 = Book.open(@simple_file1)
        new_sheet = book3.sheet(1)
        new_sheet[1,1].value.should_not == cell_value
        book3.close
      end
    end

    context "with a virgin Book class" do
      before do
        class Book  # :nodoc: #
          @@bookstore = nil
        end
      end
      it "should work" do
        expect{ unobtrusively_ok? }.to_not raise_error
      end
    end

    context "with a book never opened before" do
      before do
        class Book   # :nodoc: #
          @@bookstore = nil
        end
        other_book = Book.open(@different_file)
      end
      it "should open the book" do
        expect{ unobtrusively_ok? }.to_not raise_error
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
    end

    context "with several Excel instances" do

      before do
        @book1 = Book.open(@simple_file1)
        @book2 = Book.open(@simple_file1, :force => {:excel => :new})
        @book1.Readonly.should == false
        @book2.Readonly.should == true
        old_sheet = @book1.sheet(1)
        @old_cell_value = old_sheet[1,1].value
        @book1.close
        @book2.close
        @book1.should_not be_alive
        @book2.should_not be_alive
      end

      it "should open unobtrusively the closed book in the most recent Excel where it was open before" do      
        Book.unobtrusively(@simple_file) do |book| 
          book.excel.should_not == @book2.excel
          book.excel.should == @book1.excel
          book.ReadOnly.should == false
          sheet = book.sheet(1)
          cell = sheet[1,1]
          sheet[1,1] = cell.value == "foo" ? "bar" : "foo"
          book.Saved.should be_false
        end
        new_book = Book.open(@simple_file1)
        sheet = new_book.sheet(1)
        sheet[1,1].value.should_not == @old_cell_value
      end

      it "should open unobtrusively the closed book in the new hidden Excel" do
        Book.unobtrusively(@simple_file, :if_closed => :current) do |book| 
          book.excel.should_not == @book2.excel
          book.excel.should == @book1.excel
          book.ReadOnly.should == false
          sheet = book.sheet(1)
          cell = sheet[1,1]
          sheet[1,1] = cell.value == "foo" ? "bar" : "foo"
          book.Saved.should be_false
        end
        new_book = Book.open(@simple_file1)
        sheet = new_book.sheet(1)
        sheet[1,1].Value.should_not == @old_cell_value
      end
    end
  end

=begin
    context "with :hidden" do

      before do
        @book1 = Book.open(@simple_file1)
        @book1.close
      end
    
      it "should create a new hidden Excel instance and use this afterwards" do
        hidden_excel = nil
        Book.unobtrusively(@simple_file1, :hidden) do |book| 
          book.should be_a Book
          book.should be_alive
          book.excel.Visible.should be_false
          book.excel.DisplayAlerts.should be_false
          hidden_excel = book.excel
        end
        Book.unobtrusively(@different_file, :hidden) do |book| 
          book.should be_a Book
          book.should be_alive
          book.excel.Visible.should be_false
          book.excel.DisplayAlerts.should be_false
          book.excel.should == hidden_excel
        end
      end
    end
  end
=end
  
  describe "for_reading, for_modifying" do

    context "open unobtrusively for reading and modifying" do

      before do
        @book = Book.open(@simple_file1)
        sheet = @book.sheet(1)
        @old_cell_value = sheet[1,1].value
        @book.close
      end

      it "should not change the value" do
        Book.for_reading(@simple_file) do |book|
          book.should be_a Book
          book.should be_alive
          book.Saved.should be_true  
          sheet = book.sheet(1)
          cell = sheet[1,1]
          sheet[1,1] = cell.value == "foo" ? "bar" : "foo"
          book.Saved.should be_false
          book.excel.should == @book.excel
        end
        new_book = Book.open(@simple_file1, :visible => true)
        sheet = new_book.sheet(1)
        sheet[1,1].Value.should == @old_cell_value
      end


      it "should not change the value and use the hidden Excel instance" do
        new_excel = Excel.new(:reuse => false)
        Book.for_reading(@simple_file1, :if_closed => :new) do |book|
          sheet = book.sheet(1)
          cell = sheet[1,1]
          sheet[1,1] = cell.value == "foo" ? "bar" : "foo"
          book.excel.should_not == @book.excel
          book.excel.should_not == new_excel
          book.excel.visible.should be_false
          book.excel.displayalerts.should == :if_visible
        end
        new_book = Book.open(@simple_file1, :visible => true)
        sheet = new_book.sheet(1)
        sheet[1,1].Value.should == @old_cell_value
      end

      it "should change the value" do
        Book.for_modifying(@simple_file1) do |book|
          sheet = book.sheet(1)
          cell = sheet[1,1]
          sheet[1,1] = cell.value == "foo" ? "bar" : "foo"
          book.excel.should == @book.excel
        end
        new_book = Book.open(@simple_file1, :visible => true)
        sheet = new_book.sheet(1)
        sheet[1,1].Value.should_not == @old_cell_value
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

    #it "should evaluate a formula" do
    #  @book1.nameval("named_formula").should == 4      
    #end

    #it "should evaluate a formula via []" do
    #  @book1["named_formula"].should == 4      
    #end

    #it "should return default value if name not defined" do
    #  @book1.nameval("foo", :default => 2).should == 2
    #end

  end

  describe "close" do

    context "with unsaved read_only book" do
      before do
        @book = Book.open(@simple_file, :read_only => true)
        @sheet_count = @book.ole_workbook.Worksheets.Count
        @book.add_sheet(@sheet, :as => 'a_name')
      end

      it "should close the unsaved book without error and without saving" do
        expect{
          @book.close
          }.to_not raise_error
        new_book = Book.open(@simple_file)
        new_book.ole_workbook.Worksheets.Count.should ==  @sheet_count
        new_book.close
      end
    end

    context "with unsaved book" do
      before do
        @book = Book.open(@simple_file)
        @sheet_count = @book.ole_workbook.Worksheets.Count
        @book.add_sheet(@sheet, :as => 'a_name')
        @sheet = @book.sheet(1)
      end

      after do
        @book.close(:if_unsaved => :forget) rescue nil
      end

      it "should raise error by default" do
        expect{
          @book.close(:if_unsaved => :raise)
        }.to raise_error(WorkbookNotSaved, /workbook is unsaved: "workbook.xls"/)
      end

      it "should save the book before close with option :save" do
        ole_workbook = @book.ole_workbook
        excel = @book.excel
        excel.Workbooks.Count.should == 1
        @book.close(:if_unsaved => :save)
        excel.Workbooks.Count.should == 0
        @book.ole_workbook.should == nil
        @book.should_not be_alive
        expect{
          ole_workbook.Name}.to raise_error(WIN32OLERuntimeError)
        new_book = Book.open(@simple_file)
        begin
          new_book.ole_workbook.Worksheets.Count.should == @sheet_count + 1
        ensure
          new_book.close
        end
      end
    end
  end

  describe "save" do

    context "with simple save" do
      
      it "should save for a file opened without :read_only" do
        @book = Book.open(@simple_file)
        @book.add_sheet(@sheet, :as => 'a_name')
        @new_sheet_count = @book.ole_workbook.Worksheets.Count
        expect {
          @book.save
        }.to_not raise_error
        @book.ole_workbook.Worksheets.Count.should ==  @new_sheet_count
        @book.close
      end
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
  end

  describe "alive?, filename, ==, focus, saved" do

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
    end

    context "with focus" do

      before do
        @key_sender = IO.popen  'ruby "' + File.join(File.dirname(__FILE__), '/helpers/key_sender.rb') + '" "Microsoft Office Excel" '  , "w"        
        @book = Book.open(@simple_file, :visible => true)
        @book2 = Book.open(@another_simple_file, :visible => true)
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
        @book2.focus
        @key_sender.puts "{a}{enter}"
        sleep 0.2
        sheet2[3,2].Value.should == "a"
        @book.focus
        @book.Windows(1).Visible.should be_true
        @book.Windows(@book.Name).Visible.should be_true
        @key_sender.puts "{a}{enter}"
        sleep 0.2
        sheet[2,3].Value.should == "a"
        Excel.current.should == @book.excel
      end
    end
  end

  describe "#add_sheet" do
    before do
      @book = Book.open(@simple_file)
      @sheet = @book.sheet(1)
    end

    after do
      @book.close(:if_unsaved => :forget)
    end

    context "only first argument" do
      it "should add worksheet" do
        @book.ole_workbook.Worksheets.Count.should == 3
        @book.add_sheet @sheet
        @book.ole_workbook.Worksheets.Count.should == 4
        #expect { @book.add_sheet @sheet }.to change{ @book.workbook.Worksheets.Count }.from(3).to(4)
      end

      it "should return copyed sheet" do
        sheet = @book.add_sheet @sheet
        copyed_sheet = @book.ole_workbook.Worksheets.Item(@book.ole_workbook.Worksheets.Count)
        sheet.name.should eq copyed_sheet.name
      end
    end

    context "with first argument" do

      context "with second argument is {:before => @book.sheet(3), :after => @sheet}" do
        it "should arguments in the first is given priority" do
          @book.add_sheet(@sheet, :before => @book.sheet(3), :after => @sheet)
          @book.Worksheets.Count.should == 4
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
          @book.add_sheet(:before => @sheet).name.should eq @book.sheet(1).name
        end
      end

      context "second argument is {:after => @sheet}" do
        it "should add the second sheet" do
          @book.add_sheet(:after => @sheet).name.should eq @book.sheet(2).name
        end
      end
    end

    context "without argument" do

      it "should return copyed sheet" do
        sheet = @book.add_sheet
        copyed_sheet = @book.ole_workbook.Worksheets.Item(@book.ole_workbook.Worksheets.Count)
        sheet.name.should eq copyed_sheet.name
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

    context "standard" do

      it 'with sheet name' do
        @book.sheet('Sheet1').should be_kind_of Sheet
      end

      it 'with integer' do
        @book.sheet(1).should be_kind_of Sheet
      end

      it 'with block' do
        @book.each do |sheet|
          sheet.should be_kind_of Sheet
        end
      end

      it 'with each_with_index' do
        @book.each_with_index do |sheet,i|
          sheet.should be_kind_of Sheet
        end
      end
    end

    describe "with retain_saved" do

      before do
        @book = Book.open(@simple_file)
      end

      after do
        @book.close(:if_unsaved => :forget)
      end

      it "should keep the save state 'unsaved'" do
        sheet = @book.sheet(1)
        sheet[1,1] = sheet[1,1].value == "foo" ? "bar" : "foo"
        @book.Saved.should be_false
        @book.retain_saved do
          sheet = @book.sheet(1)
          a = sheet[1,1]
          b = @book.visible
        end
        @book.Saved.should be_false
      end

      it "should keep the save state 'unsaved' even when the workbook was saved before" do
        sheet = @book.sheet(1)
        sheet[1,1] = sheet[1,1].value == "foo" ? "bar" : "foo"
        @book.Saved.should be_false
        @book.retain_saved do
          @book.save
          @book.Saved.should be_true
        end
        @book.Saved.should be_false
      end
    end

=begin    
    context "with test what happens with save-status when setting calculation status" do

      it "should keep the save status" do
        book1 = Book.open(@simple_file, :visible => true)
        book1.Saved.should be_true
        book2 = Book.open(@another_simple_file, :visible => true)
        book1.Saved.should be_true
        book2.Saved.should be_true
        sheet2 = book2.sheet(1)
        sheet2[1,1] = sheet2[1,1].value == "foo" ? "bar" : "foo"
        book1.Saved.should be_true
        book2.Saved.should be_false
        book3 = Book.open(@different_file, :visible => true)
        book1.Saved.should be_true
        book2.Saved.should be_false
        book3.Saved.should be_true
      end
    end
=end

    context 'open with block' do
      it {
        Book.open(@simple_file) do |book|
          book.sheet('Sheet1').should be_a Sheet
        end
      }
    end
  end
end
end
end
