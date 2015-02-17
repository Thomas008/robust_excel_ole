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
    @simple_file = @dir + '/simple.xls'
    @simple_save_file = @dir + '/simple_save.xls'
    @different_file = @dir + '/different_simple.xls'
    @simple_file_other_path = @dir + '/more_data/simple.xls'
  end

  after do
    #Excel.close_all
    rm_tmp(@dir)
  end


  describe "create file" do
    context "with standard" do
      it "simple file with default" do
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
    end
  end

  describe "connect" do

    context "with one excel instance" do

      before do
        @book = Book.open(@simple_file)
      end

      after do
        @book.close rescue nil
        @connected_book.close rescue nil 
      end

      it "should connect to the open book" do
        @connected_book = Book.connect(@simple_file)
        @connected_book.should be_a Book
        @connected_book.should == @book
      end

      it "should connect to a closed book" do
        @connected_book = Book.connect(@simple_file)
        @book.close
        @connected_book = Book.connect(@simple_file)
        @connected_book.should be_a Book
        @connected_book.should == @book
      end

      it "should yield nil to a non-existing book" do
        @connected_book = Book.connect('foo')
        @connected_book.should == nil 
      end

      it "should connect to two different open books in the same excel instance" do
        book2 = Book.open(@different_file)
        @connected_book = Book.connect(@simple_file)
        connected_book2 = Book.connect(@different_file)        
        @connected_book.should == @book
        connected_book2.should == book2
        book2.close
        connected_book2.close
      end

    end

    context "with several excel instances" do

      before do
        @book = Book.open(@simple_file)
      end

      after do
        @book.close(:if_unsaved => :forget)
        @connected_book.close
      end

      it "should connect to two different books open in several excel instances" do
        excel = Excel.new(:reuse => false)
        book2 = Book.open(@different_file, :excel => excel)
        @connected_book = Book.connect(@simple_file)
        connected_book2 = Book.connect(@different_file)        
        @connected_book.should == @book
        connected_book2.should == book2
        book2.close
        connected_book2.close
      end

      it "should connect to the writable, first book" do
        excel = Excel.new(:reuse => false)
        book2 = Book.open(@simple_file, :excel => excel)
        @connected_book = Book.connect(@simple_file)        
        @connected_book.should == @book
        book2.close
      end

    end

    context "with read_only" do

      before do
        @book = Book.open(@simple_file, :read_only => true)
      end

      after do
        @book.close(:if_unsaved => :forget)
        @connected_book.close
      end

      it "should connect to the only writable, second, book" do
        excel = Excel.new(:reuse => false)
        book2 = Book.open(@simple_file, :excel => excel)
        @connected_book = Book.connect(@simple_file)        
        @connected_book.should == book2
        book2.close
      end

      it "should connect to the last read_only book, if only read_only books exist" do
        @book.ReadOnly.should be_true
        excel = Excel.new(:reuse => false)
        book2 = Book.open(@simple_file, :excel => excel, :read_only => true)
        book2.ReadOnly.should be_true
        @connected_book = Book.connect(@simple_file)        
        @connected_book.should == book2
        book2.close
      end

      it "should connect to the unsaved read_only book, if only read_only books exist" do
        @book.ReadOnly.should be_true
        sheet = @book[0]
        sheet[0,0] = sheet[0,0].value == "simple" ? "complex" : "simple"
        @book.Saved. should be_false
        excel = Excel.new(:reuse => false)
        book2 = Book.open(@simple_file, :excel => excel, :read_only => true)
        book2.ReadOnly.should be_true
        book2.Saved. should be_true
        @connected_book = Book.connect(@simple_file)        
        @connected_book.should == @book
        book2.close
      end

      it "should connect to the writable book, if otherwise only read_only and unsaved books exist" do
        @book.ReadOnly.should be_true
        excel2 = Excel.new(:reuse => false)
        book2 = Book.open(@simple_file, :excel => excel2)
        book2.ReadOnly.should be_false
        excel3 = Excel.new(:reuse => false)
        book3 = Book.open(@simple_file, :excel => excel3, :read_only => true)
        book3.ReadOnly.should be_true
        sheet = book3[0]
        sheet[0,0] = sheet[0,0].value == "simple" ? "complex" : "simple"
        book3.Saved.should be_false
        @connected_book = Book.connect(@simple_file)        
        @connected_book.should == book2
        book2.close
        book3.close(:if_unsaved => :forget)
      end

    end

    context "with save_as" do

      before do
        @book = Book.open(@simple_file)
      end

      after do
        @book.close(:if_unsaved => :forget)
        @connected_book.close
      end

      it "should simple save a connected book" do
        @connected_book = Book.connect(@simple_file)
        @connected_book.save
      end

      it "should simple save changes of a connected book" do
        sheet = @book[0]
        sheet[0,0] = sheet[0,0].value == "simple" ? "complex" : "simple"
        @book.Saved.should be_false
        @connected_book = Book.connect(@simple_file)
        @connected_book.save    
        @connected_book.Saved.should be_true  
      end

      it "should save changes of a connected book with another filename" do
        File.delete @simple_save_file rescue nil
        File.open(@simple_save_file,"w") do | file |
          file.puts "garbage"
        end
        sheet = @book[0]
        old_cell = sheet[0,0]
        sheet[0,0] = old_cell.value == "simple" ? "complex" : "simple"
        @book.Saved.should be_false
        @connected_book = Book.connect(@simple_file)
        @connected_book.should == @book
        @connected_book.save_as(@simple_save_file, :if_exists => :overwrite)
        @connected_book.Saved.should be_true  
        File.exist?(@simple_save_file).should be_true
        @connected_book.close
        new_book = Book.connect(@simple_save_file)
        new_book.should_not == nil
        new_book.should be_a Book
        new_book.should_not be_alive
        new_book.should == @book
        new_book2 = Book.connect(@simple_file)
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

    context "with an open book" do

      before do
        @book = Book.open(@simple_file)
      end

      after do
        @book.close(:if_unsaved => :forget)
      end

      it "should let a saved book saved" do
        @book.Saved.should be_true
        @book.should be_alive
        sheet = @book[0]
        @old_cell_value = sheet[0,0].value
        unobtrusively_ok?
        @book.Saved.should be_true
        @book.should be_alive
        sheet = @book[0]
        sheet[0,0].value.should_not == @old_cell_value
      end

      it "should let an unsaved book unsaved" do
        sheet = @book[0]
        sheet[0,0] = sheet[0,0].value == "simple" ? "complex" : "simple" 
        @book.Saved.should be_false
        @book.should be_alive
        @old_cell_value = sheet[0,0].value
        unobtrusively_ok?
        @book.Saved.should be_false
        @book.should be_alive
        sheet = @book[0]
        sheet[0,0].value.should_not == @old_cell_value
      end
    end
    
    context "with a closed book" do
      
      before do
        @book = Book.open(@simple_file)
      end

      after do
        @book.close(:if_unsaved => :forget) 
        Excel.close_all
      end

      it "should let the closed book closed by default" do
        sheet = @book[0]
        @old_cell_value = sheet[0,0].value
        @book.close
        @book.should_not be_alive
        unobtrusively_ok?
        @book.should_not be_alive
        @book = Book.open(@simple_file)
        sheet = @book[0]
        sheet[0,0].value.should_not == @old_cell_value
      end

      it "should keep open the book" do
        @book.close
        @book.should_not be_alive
        Book.unobtrusively(@simple_file, :keep_open => true) do |book|
          book.should be_a Book
          sheet = book[0]
          cell = sheet[0,0]
          sheet[0,0] = cell.value == "simple" ? "complex" : "simple"
          book.Saved.should be_false
        end
        @book.should be_alive
      end
    end

    context "with an unsaved book" do

      before do
        @book = Book.open(@simple_file)
        excel = Excel.new(:reuse => false)
        @book2 = Book.open(@simple_file, :excel => excel)
      end

      after do
        @book.close(:if_unsaved => :forget)
        @book2.close(:if_unsaved => :forget)
        Excel.close_all
      end

      it "should modify unobtrusively the first, unsaved book" do
        sheet = @book[0]
        @old_cell_value = sheet[0,0].value
        sheet2 = @book2[0]
        @old_cell_value2 = sheet2[0,0].value
        sheet[0,0] = sheet[0,0].value == "simple" ? "complex" : "simple"
        unobtrusively_ok?
        @book.should be_alive
        @book.Saved.should be_false
        sheet = @book[0]
        sheet[0,0].value.should_not == @old_cell_value
      end

      it "should modify unobtrusively the second, unsaved book" do
        sheet = @book[0]
        @old_cell_value = sheet[0,0].value
        sheet2 = @book2[0]
        @old_cell_value2 = sheet2[0,0].value
        sheet2[0,0] = sheet2[0,0].value == "simple" ? "complex" : "simple"
        unobtrusively_ok?
        @book2.should be_alive
        @book2.Saved.should be_false
        sheet2 = @book[0]
        sheet2[0,0].value.should_not == @old_cell_value2
      end

    end


  end    

  describe "open" do

    context "with non-existing file" do
      it "should raise an exception" do
        File.delete @simple_save_file rescue nil
        expect {
          Book.open(@simple_save_file)
        }.to raise_error(ExcelErrorOpen, "file #{@simple_save_file} not found")
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

    context "with attr_reader excel" do
      before do
        @new_book = Book.open(@simple_file)
      end
      after do
        @new_book.close
      end
      it "should provide the excel application of the book" do
        excel = @new_book.excel
        excel.class.should == Excel
        excel.should be_a Excel
      end
    end

    context "with :excel" do
      it "should reuse the given excel application of the book" do
        Excel.close_all
        book1 = Book.open(@simple_file)
        excel1 = book1.excel
        book2 = Book.open(@simple_file, :reuse => false)
        excel2 = book2.excel
        excel2.should_not == excel1
        book3 = Book.open(@simple_file)
        excel3 = book3.excel
        book4 = Book.open(@simple_file, :excel => excel2)
        excel4 = book4.excel
        excel3.should == excel1
        excel4.should == excel2
        excel4.class.should == Excel
        excel4.should be_a Excel
        book4.close
        book3.close
        book2.close
        book1.close
      end
    end


    context "with :read_only" do
      it "should be able to save, if :read_only => false" do
        book = Book.open(@simple_file, :read_only => false)
        book.should be_a Book
        expect {
          book.save_as(@simple_save_file, :if_exists => :overwrite)
        }.to_not raise_error
        book.close
      end

      it "should be able to save, if :read_only is set to default value" do
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
          Book.open(path)
        }.to raise_error(ExcelErrorOpen, "file #{path} not found")
      end
    end

    context "with an already opened book" do

      before do
        @book = Book.open(@simple_file)
      end

      after do
        @book.close
      end

      context "with an already saved book" do
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
            it "should belong to the same Excel application" do
              @new_book.excel.should == @book.excel
              @different_book.excel.should == @book.excel
            end
          end
        end
      end

      context "with an unsaved book" do

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
            @book.should_not be_alive
            @new_book.should be_alive
            @new_book.filename.downcase.should == @simple_file.downcase
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
              }.to raise_error(ExcelUserCanceled, "open: canceled by user")
            @book.should be_alive
          end
        end

        it "should open the book in a new excel application, if :if_unsaved is :new_excel" do
          @new_book = Book.open(@simple_file, :if_unsaved => :new_excel)
          @book.should be_alive
          @new_book.should be_alive
          @new_book.filename.should == @book.filename
          @new_book.excel.should_not == @book.excel
          @new_book.close
        end

        it "should raise an error, if :if_unsaved is default" do
          expect {
            @new_book = Book.open(@simple_file)
          }.to raise_error(ExcelErrorOpen, "book is already open but not saved (#{File.basename(@simple_file)})")
        end

        it "should raise an error, if :if_unsaved is invalid option" do
          expect {
            @new_book = Book.open(@simple_file, :if_unsaved => :invalid_option)
          }.to raise_error(ExcelErrorOpen, ":if_unsaved: invalid option")
        end

      end
    end

    context "with a book in a different path" do

      before do        
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

      it "should open the book in a new excel application, if :if_obstructed is :new_excel" do
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

    context "with unsaved book and with :read_only" do
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

      it "should raise error for default" do
        expect{
          @book.close
        }.to raise_error(ExcelErrorClose, "book is unsaved (#{File.basename(@simple_file)})")
      end

      it "should raise error for invalid option" do
        expect{
          @book.close(:if_unsaved => :invalid)
        }.to raise_error(ExcelErrorClose, ":if_unsaved: invalid option")
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
          book_save = Book.open(@simple_save_file, :reuse => false)
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
          (File.size?(@simple_save_file) == booklength).should be_true
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
            new_book = Book.open(@simple_save_file)
            new_book.should be_a Book
            @book.excel.DisplayAlerts.should == displayalert_value
            new_book.close
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
          (File.size?(@simple_save_file) == booklength).should be_true
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
        @new_book = Book.new(@simple_file_other_path, :reuse => false)
        @new_book.should_not == @book
      end

      it "should be false with same book names but different excel apps" do
        @new_book = Book.new(@simple_file, :reuse => false)
        @new_book.should_not == @book
      end

      it "should be false with non-Books" do
        @book.should_not == "hallo"
        @book.should_not == 7
        @book.should_not == nil
      end
    end

    context "with visible" do

      before do
        @book = Book.open(@simple_file)
      end

      after do
        @book.close
      end

      it "should make Excel visible" do
        @book.visible = false
        Excel.current.visible.should be_false
        @book.visible.should be_false
        @book.visible = true
        Excel.current.visible.should be_true
        @book.visible.should be_true
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
