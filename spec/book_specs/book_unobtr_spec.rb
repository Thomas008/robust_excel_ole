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
    rm_tmp(@dir)
  end

  
  describe "unobtrusively" do

    def unobtrusively_ok? # :nodoc: #
      Book.unobtrusively(@simple_file) do |book|
        book.should be_a Book
        sheet = book[0]
        sheet[1,1] = sheet[1,1].value == "foo" ? "bar" : "foo"
        book.should be_alive
        book.Saved.should be_false
      end
    end

    context "with no open book" do

      it "should open unobtrusively if no Excel is open" do
        Excel.close_all
        Book.unobtrusively(@simple_file) do |book|
          book.should be_a Book
        end
      end

      it "should open unobtrusively in a new Excel" do
        expect{ unobtrusively_ok? }.to_not raise_error
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
          rescue RuntimeError => msg
            #puts "RuntimeError: #{msg.message}" if msg.message =~ /Excel instance not alive or damaged/
            nil
          end
        end

      it "should open unobtrusively in the first opened Excel" do
        Book.unobtrusively(@simple_file) do |book|
          book.should be_a Book
          book.should be_alive
            book.excel.should     == @excel1
            book.excel.should_not == @excel2
          end
      end

      it "should open unobtrusively in a new Excel" do
        Book.unobtrusively(@simple_file, :hidden) do |book|
          book.should be_a Book
          book.should be_alive
            book.excel.should_not == @excel1
            book.excel.should_not == @excel2
          end
      end

      it "should open unobtrusively in a given Excel" do
        Book.unobtrusively(@simple_file, @excel2) do |book|
          book.should be_a Book
          book.should be_alive
            book.excel.should_not == @excel1
            book.excel.should     == @excel2
        end
      end
  
      it "should raise an error if the excel instance is not alive" do
        Excel.close_all
        expect{
          Book.unobtrusively(@simple_file, @excel2) do |book|
          end
        }.to raise_error(ExcelErrorOpen, "Excel instance not alive or damaged")
        end
      end

      it "should raise an error if the option is invalid" do
        expect{
          Book.unobtrusively(@simple_file, :invalid_option) do |book|
          end
        }.to raise_error(ExcelError, "receiver instance is neither an Excel nor a Book")
      end

      it "should be visible and displayalerts" do
        Book.unobtrusively(@simple_file, :visible => true, :displayalerts => true) do |book|
          book.should be_a Book
          book.should be_alive
          book.excel.visible.should be_true
          book.excel.displayalerts.should be_true
        end
      end

      it "should be visible" do
        excel = Excel.new(:reuse => false, :displayalerts => true)
        Book.unobtrusively(@simple_file, :visible => true) do |book|
          book.should be_a Book
          book.should be_alive
          book.excel.visible.should be_true
          book.excel.displayalerts.should be_true
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

      #it "should let an open Book open" do
      #  Book.unobtrusively(@simple_file) do |book|
      #    book.should be_a Book
      #    book.should be_alive
      #    book.excel.should == @book.excel
      #  end        
      #  @book.should be_alive
      #  @book.should be_a Book
      #end

      it "should let an open Book open if it has been closed and opened again" do
        @book.close
        @book.reopen
        Book.unobtrusively(@simple_file) do |book|
          book.should be_a Book
          book.should be_alive
          book.excel.should == @book.excel
        end        
        @book.should be_alive
        @book.should be_a Book
      end

      it "should let an open Book open if two books have been opened and one has been closed and opened again" do
        book2 = Book.open(@different_file, :force_excel => :new)
        @book.close
        book2.close
        @book.reopen
        Book.unobtrusively(@simple_file) do |book|
          book.should be_a Book
          book.should be_alive
          book.excel.should == @book.excel
        end        
        @book.should be_alive
        @book.should be_a Book
      end

      it "should open in the Excel of the given Book" do
        #book1 = Book.open(@different_file)
        @book2 = Book.open(@another_simple_file, :force_excel => :new)
        Book.unobtrusively(@different_file, @book2) do |book|
          book.should be_a Book
          book.should be_alive
          book.excel.should_not == @book.excel
          book.excel.should     == @book2.excel
        end
      end

      it "should let a saved book saved" do
        @book.Saved.should be_true
        @book.should be_alive
        sheet = @book[0]
        old_cell_value = sheet[1,1].value
        unobtrusively_ok?
        @book.Saved.should be_true
        @book.should be_alive
        sheet = @book[0]
        sheet[1,1].value.should_not == old_cell_value
      end

     it "should let the unsaved book unsaved" do
        sheet = @book[0]
        sheet[1,1] = sheet[1,1].value == "foo" ? "bar" : "foo"
        old_cell_value = sheet[1,1].value
        @book.Saved.should be_false
        unobtrusively_ok?
        @book.should be_alive
        @book.Saved.should be_false
        @book.close(:if_unsaved => :forget)
        @book2 = Book.open(@simple_file)
        sheet2 = @book2[0]
        sheet2[1,1].value.should_not == old_cell_value
      end

      it "should modify unobtrusively the second, writable book" do
        @book2 = Book.open(@simple_file, :force_excel => :new)
        @book.ReadOnly.should be_false
        @book2.ReadOnly.should be_true
        sheet = @book2[0]
        old_cell_value = sheet[1,1].value
        sheet[1,1] = sheet[1,1].value == "foo" ? "bar" : "foo"
        unobtrusively_ok?
        @book2.should be_alive
        @book2.Saved.should be_false
        @book2.close(:if_unsaved => :forget)
        @book.close
        @book = Book.open(@simple_file)
        sheet2 = @book[0]
        sheet2[1,1].value.should_not == old_cell_value
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
        old_cell_value = sheet[1,1].value
        @book.close
        @book.should_not be_alive
        unobtrusively_ok?
        @book.should_not be_alive
        @book = Book.open(@simple_file)
        sheet = @book[0]
        sheet[1,1].Value.should_not == old_cell_value
      end

      # The bold reanimation of the @book
      it "should use the excel of the book and keep open the book" do
        excel = Excel.new(:reuse => false)
        sheet = @book[0]
        old_cell_value = sheet[1,1].value
        @book.close
        @book.should_not be_alive
        Book.unobtrusively(@simple_file, :keep_open => true) do |book|
          book.should be_a Book
          book.excel.should == @book.excel
          book.excel.should_not == excel
          sheet = book[0]
          cell = sheet[1,1]
          sheet[1,1] = cell.value == "foo" ? "bar" : "foo"
          book.Saved.should be_false
        end
        @book.should be_alive
        @book.close
        new_book = Book.open(@simple_file)
        sheet = new_book[0]
        sheet[1,1].value.should_not == old_cell_value
      end

      # book shall be reanimated even with :hidden
      it "should use the excel of the book and keep open the book" do
        excel = Excel.new(:reuse => false)
        sheet = @book[0]
        old_cell_value = sheet[1,1].value
        @book.close
        @book.should_not be_alive
        Book.unobtrusively(@simple_file, :hidden) do |book|
          book.should be_a Book
          book.should be_alive
          book.excel.should_not == @book.excel
          book.excel.should_not == excel
          sheet = book[0]
          cell = sheet[1,1]
          sheet[1,1] = cell.value == "foo" ? "bar" : "foo"
          book.Saved.should be_false
        end
        @book.should_not be_alive
        new_book = Book.open(@simple_file)
        sheet = new_book[0]
        sheet[1,1].value.should_not == old_cell_value
      end

      it "should use another excel if the Excels are closed" do
        excel = Excel.new(:reuse => false)
        sheet = @book[0]
        old_cell_value = sheet[1,1].value
        @book.close
        @book.should_not be_alive
        Excel.kill_all
        Book.unobtrusively(@simple_file, :keep_open => true) do |book|
          book.should be_a Book
          book.excel.should == @book.excel
          book.excel.should_not == excel
          sheet = book[0]
          cell = sheet[1,1]
          sheet[1,1] = cell.value == "foo" ? "bar" : "foo"
          book.Saved.should be_false
        end
        @book.should be_alive
        @book.close
        new_book = Book.open(@simple_file)
        sheet = new_book[0]
        sheet[1,1].value.should_not == old_cell_value
      end

      it "should use another excel if the Excels are closed" do
        excel = Excel.new(:reuse => false)
        sheet = @book[0]
        old_cell_value = sheet[1,1].value
        @book.close
        @book.should_not be_alive
        Excel.close_all
        Book.unobtrusively(@simple_file, :hidden, :keep_open => true) do |book|
          book.should be_a Book
          book.excel.should_not == @book.excel
          book.excel.should_not == excel
          sheet = book[0]
          cell = sheet[1,1]
          sheet[1,1] = cell.value == "foo" ? "bar" : "foo"
          book.Saved.should be_false
        end
        @book.should_not be_alive
        new_book = Book.open(@simple_file)
        sheet = new_book[0]
        sheet[1,1].value.should_not == old_cell_value
      end      

      it "should modify unobtrusively the copied file" do
        sheet = @book[0]
        old_cell_value = sheet[1,1].value
        File.delete simple_save_file rescue nil
        @book.save_as(@simple_save_file)
        @book.close
        Book.unobtrusively(@simple_save_file) do |book|
          sheet = book[0]
          cell = sheet[1,1]
          sheet[1,1] = cell.Value == "foo" ? "bar" : "foo"
        end
        old_book = Book.open(@simple_file)
        old_sheet = old_book[0]
        old_sheet[1,1].Value.should == old_cell_value
        old_book.close
        new_book = Book.open(@simple_save_file)
        new_sheet = new_book[0]
        new_sheet[1,1].Value.should_not == old_cell_value
        new_book.close
      end
    end

    context "with a visible book" do

      before do
        @book = Book.open(@simple_file, :visible => true)
      end

      after do
        @book.close(:if_unsaved => :forget)
        @book2.close(:if_unsaved => :forget) rescue nil
      end

      it "should let an open Book open" do
        Book.unobtrusively(@simple_file) do |book|
          book.should be_a Book
          book.should be_alive
          book.excel.should == @book.excel
          book.excel.Visible.should be_true
        end        
        @book.should be_alive
        @book.should be_a Book
        @book.excel.Visible.should be_true
      end
      
    end

    context "with various options for an Excel instance in which to open a closed book" do

      before do
        @book = Book.open(@simple_file)
        @book.close
      end

      it "should use a given Excel" do
        new_excel = Excel.new(:reuse => false)
        another_excel = Excel.new(:reuse => false)
        Book.unobtrusively(@simple_file, another_excel) do |book|
          book.excel.should_not == @book.excel
          book.excel.should_not == new_excel
          book.excel.should == another_excel
        end
      end

      it "should use the hidden Excel" do
        new_excel = Excel.new(:reuse => false)
        Book.unobtrusively(@simple_file, :hidden) do |book|
          book.excel.should_not == @book.excel
          book.excel.should_not == new_excel
          book.excel.visible.should be_false
          book.excel.displayalerts.should be_false
          @hidden_excel = book.excel
        end
        Book.unobtrusively(@simple_file, :hidden) do |book|
          book.excel.should_not == @book.excel
          book.excel.should_not == new_excel
          book.excel.visible.should be_false
          book.excel.displayalerts.should be_false
          book.excel.should == @hidden_excel 
        end
      end

      it "should reuse Excel" do
        new_excel = Excel.new(:reuse => false)
        Book.unobtrusively(@simple_file, :reuse) do |book|
          book.excel.should == @book.excel
          book.excel.should_not == new_excel
        end
      end

      it "should reuse Excel by default" do
        new_excel = Excel.new(:reuse => false)
        Book.unobtrusively(@simple_file) do |book|
          book.excel.should == @book.excel
          book.excel.should_not == new_excel
        end
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
        old_cell_value = sheet[1,1].value
        unobtrusively_ok?
        @book.should be_alive
        @book.Saved.should be_true
        @book.ReadOnly.should be_true
        @book.close
        book2 = Book.open(@simple_file)
        sheet2 = book2[0]
        sheet2[1,1].value.should_not == old_cell_value
      end

      it "should let the unsaved book unsaved" do
        @book.ReadOnly.should be_true
        sheet = @book[0]
        sheet[1,1] = sheet[1,1].value == "foo" ? "bar" : "foo" 
        @book.Saved.should be_false
        @book.should be_alive
        Book.unobtrusively(@simple_file) do |book|
          book.should be_a Book
          sheet = book[0]
          sheet[1,1] = sheet[1,1].value == "foo" ? "bar" : "foo"
          @cell_value = sheet[1,1].Value
          book.should be_alive
          book.Saved.should be_false
        end
        @book.should be_alive
        @book.Saved.should be_false
        @book.ReadOnly.should be_true
        @book.close
        book2 = Book.open(@simple_file)
        sheet2 = book2[0]
        # modifies unobtrusively the saved version, not the unsaved version
        sheet2[1,1].value.should == @cell_value        
      end

      it "should open unobtrusively by default the writable book" do
        book2 = Book.open(@simple_file, :force_excel => :new, :read_only => false)
        @book.ReadOnly.should be_true
        book2.Readonly.should be_false
        sheet = @book[0]
        cell_value = sheet[1,1].value
        Book.unobtrusively(@simple_file, :hidden) do |book|
          book.should be_a Book
          book.excel.should == book2.excel
          book.excel.should_not == @book.excel
          sheet = book[0]
          sheet[1,1] = sheet[1,1].value == "foo" ? "bar" : "foo"
          book.should be_alive
          book.Saved.should be_false          
        end  
        @book.Saved.should be_true
        @book.ReadOnly.should be_true
        @book.close
        book2.close
        book3 = Book.open(@simple_file)
        new_sheet = book3[0]
        new_sheet[1,1].value.should_not == cell_value
        book3.close
      end

      it "should open unobtrusively by default the book in a new Excel such that the book is writable" do
        book2 = Book.open(@simple_file, :force_excel => :new, :read_only => true)
        @book.ReadOnly.should be_true
        book2.Readonly.should be_true
        sheet = @book[0]
        cell_value = sheet[1,1].value
        Book.unobtrusively(@simple_file, :hidden) do |book|
          book.should be_a Book
          book.excel.should_not == book2.excel
          book.excel.should_not == @book.excel
          sheet = book[0]
          sheet[1,1] = sheet[1,1].value == "foo" ? "bar" : "foo"
          book.should be_alive
          book.Saved.should be_false          
        end  
        @book.Saved.should be_true
        @book.ReadOnly.should be_true
        @book.close
        book2.close
        book3 = Book.open(@simple_file)
        new_sheet = book3[0]
        new_sheet[1,1].value.should_not == cell_value
        book3.close
      end

      it "should open unobtrusively the book in a new Excel such that the book is writable" do
        book2 = Book.open(@simple_file, :force_excel => :new, :read_only => true)
        @book.ReadOnly.should be_true
        book2.Readonly.should be_true
        sheet = @book[0]
        cell_value = sheet[1,1].value
        Book.unobtrusively(@simple_file, :hidden) do |book|
          book.should be_a Book
          book.excel.should_not == book2.excel
          book.excel.should_not == @book.excel
          sheet = book[0]
          sheet[1,1] = sheet[1,1].value == "foo" ? "bar" : "foo"
          book.should be_alive
          book.Saved.should be_false          
        end  
        @book.Saved.should be_true
        @book.ReadOnly.should be_true
        @book.close
        book2.close
        book3 = Book.open(@simple_file)
        new_sheet = book3[0]
        new_sheet[1,1].value.should_not == cell_value
        book3.close
      end

      it "should open unobtrusively the book in a new Excel to open the book writable" do
        excel1 = Excel.new(:reuse => false)
        excel2 = Excel.new(:reuse => false)
        book2 = Book.open(@simple_file, :force_excel => :new, :read_only => true)
        @book.ReadOnly.should be_true
        book2.Readonly.should be_true
        sheet = @book[0]
        cell_value = sheet[1,1].value
        Book.unobtrusively(@simple_file, :hidden, :readonly_excel => false) do |book|
          book.should be_a Book
          book.ReadOnly.should be_false
          book.excel.should_not == book2.excel
          book.excel.should_not == @book.excel
          book.excel.should_not == excel1
          book.excel.should_not == excel2
          sheet = book[0]
          sheet[1,1] = sheet[1,1].value == "foo" ? "bar" : "foo"
          book.should be_alive
          book.Saved.should be_false          
        end  
        @book.Saved.should be_true
        @book.ReadOnly.should be_true
        @book.close
        book2.close
        book3 = Book.open(@simple_file)
        new_sheet = book3[0]
        new_sheet[1,1].value.should_not == cell_value
        book3.close
      end

      it "should open unobtrusively the book in the same Excel to open the book writable" do
        excel1 = Excel.new(:reuse => false)
        excel2 = Excel.new(:reuse => false)
        book2 = Book.open(@simple_file, :force_excel => :new, :read_only => true)
        @book.ReadOnly.should be_true
        book2.Readonly.should be_true
        sheet = @book[0]
        cell_value = sheet[1,1].value
        Book.unobtrusively(@simple_file, :hidden, :readonly_excel => true) do |book|
          book.should be_a Book
          book.excel.should == book2.excel
          book.ReadOnly.should be_false
          sheet = book[0]
          sheet[1,1] = sheet[1,1].value == "foo" ? "bar" : "foo"
          book.should be_alive
          book.Saved.should be_false          
        end  
        book2.Saved.should be_true
        book2.ReadOnly.should be_false
        @book.close
        book2.close
        book3 = Book.open(@simple_file)
        new_sheet = book3[0]
        new_sheet[1,1].value.should_not == cell_value
        book3.close
      end

      it "should open unobtrusively the book in the Excel where it was opened most recently" do
        book2 = Book.open(@simple_file, :force_excel => :new, :read_only => true)
        @book.ReadOnly.should be_true
        book2.Readonly.should be_true
        sheet = @book[0]
        cell_value = sheet[1,1].value
        Book.unobtrusively(@simple_file, :hidden, :read_only => true) do |book|
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
        Book.unobtrusively(@simple_file, :hidden) do |book|
          @book1.Saved.should be_true
          book.Saved.should be_true  
          sheet = book[0]
          cell = sheet[1,1]
          sheet[1,1] = cell.value == "foo" ? "bar" : "foo"
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
        cell1 = sheet1[1,1].value
        result = 
          Book.unobtrusively(@simple_file) do |book| 
            sheet = book[0]
            cell = sheet[1,1].value
          end
        result.should == cell1
      end

      it "should yield the block result even if the book gets saved" do
        sheet1 = @book1[0]
        @book1.save
        result = 
          Book.unobtrusively(@simple_file) do |book| 
            sheet = book[0]
            sheet[1,1] = 22
            @book1.Saved.should be_false
            42
          end
        result.should == 42
        @book1.Saved.should be_true
      end
    end

    context "with several Excel instances" do

      before do
        @book1 = Book.open(@simple_file)
        @book2 = Book.open(@simple_file, :force_excel => :new)
        @book1.Readonly.should == false
        @book2.Readonly.should == true
        old_sheet = @book1[0]
        @old_cell_value = old_sheet[1,1].value
        @book1.close
        @book2.close
        @book1.should_not be_alive
        @book2.should_not be_alive
      end

      it "should open unobtrusively the closed book in the most recent Excel where it was open before" do      
        Book.unobtrusively(@simple_file) do |book| 
          book.excel.should == @book2.excel
          book.excel.should_not == @book1.excel
          book.ReadOnly.should == false
          sheet = book[0]
          cell = sheet[1,1]
          sheet[1,1] = cell.value == "foo" ? "bar" : "foo"
          book.Saved.should be_false
        end
        new_book = Book.open(@simple_file)
        sheet = new_book[0]
        sheet[1,1].value.should_not == @old_cell_value
      end

      it "should open unobtrusively the closed book in the new hidden Excel" do
        Book.unobtrusively(@simple_file, :hidden) do |book| 
          book.excel.should_not == @book2.excel
          book.excel.should_not == @book1.excel
          book.ReadOnly.should == false
          sheet = book[0]
          cell = sheet[1,1]
          sheet[1,1] = cell.value == "foo" ? "bar" : "foo"
          book.Saved.should be_false
        end
        new_book = Book.open(@simple_file)
        sheet = new_book[0]
        sheet[1,1].Value.should_not == @old_cell_value
      end

      it "should open unobtrusively the closed book in a new Excel if the Excel is not alive anymore" do
        Excel.close_all
        Book.unobtrusively(@simple_file, :hidden) do |book| 
          book.ReadOnly.should == false
          book.excel.should_not == @book1.excel
          book.excel.should_not == @book2.excel
          sheet = book[0]
          cell = sheet[1,1]
          sheet[1,1] = cell.value == "foo" ? "bar" : "foo"
          book.Saved.should be_false
        end
        new_book = Book.open(@simple_file)
        sheet = new_book[0]
        sheet[1,1].Value.should_not == @old_cell_value
      end
    end

    context "with :hidden" do

      before do
        @book1 = Book.open(@simple_file)
        @book1.close
      end
    
      it "should create a new hidden Excel instance" do
        Book.unobtrusively(@simple_file, :hidden) do |book| 
          book.should be_a Book
          book.should be_alive
          book.excel.Visible.should be_false
          book.excel.DisplayAlerts.should be_false
        end
      end

      it "should create a new hidden Excel instance and use this afterwards" do
        hidden_excel = nil
        Book.unobtrusively(@simple_file, :hidden) do |book| 
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

      it "should create a new hidden Excel instance if the Excel is closed" do
        Excel.close_all
        Book.unobtrusively(@simple_file, :hidden) do |book| 
          book.should be_a Book
          book.should be_alive
          book.excel.Visible.should be_false
          book.excel.DisplayAlerts.should be_false
          book.excel.should_not == @book1.excel
        end
      end

      it "should exclude hidden Excel when reuse in unobtrusively" do
        hidden_excel = nil
        Book.unobtrusively(@simple_file, :hidden) do |book| 
          book.should be_a Book
          book.should be_alive
          book.excel.Visible.should be_false
          book.excel.DisplayAlerts.should be_false
          book.excel.should_not == @book1.excel
          hidden_excel = book.excel
        end
        Book.unobtrusively(@simple_file) do |book| 
          book.should be_a Book
          book.should be_alive
          book.excel.Visible.should be_false
          book.excel.DisplayAlerts.should be_false
          book.excel.should_not == hidden_excel
        end
      end

      it "should exclude hidden Excel when reuse in open" do
        hidden_excel = nil
        Book.unobtrusively(@simple_file, :hidden) do |book| 
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

  describe "for_reading, for_modifying" do

    context "open unobtrusively for reading and modifying" do

      before do
        @book = Book.open(@simple_file)
        sheet = @book[0]
        @old_cell_value = sheet[1,1].value
        @book.close
      end

      it "should not change the value" do
        Book.for_reading(@simple_file) do |book|
          book.should be_a Book
          book.should be_alive
          book.Saved.should be_true  
          sheet = book[0]
          cell = sheet[1,1]
          sheet[1,1] = cell.value == "foo" ? "bar" : "foo"
          book.Saved.should be_false
          book.excel.should == @book.excel
        end
        new_book = Book.open(@simple_file, :visible => true)
        sheet = new_book[0]
        sheet[1,1].Value.should == @old_cell_value
      end

      it "should not change the value and use a given Excel" do
        new_excel = Excel.new(:reuse => false)
        another_excel = Excel.new(:reuse => false)
        Book.for_reading(@simple_file, another_excel) do |book|
          sheet = book[0]
          cell = sheet[1,1]
          sheet[1,1] = cell.value == "foo" ? "bar" : "foo"
          book.excel.should == another_excel
        end
        new_book = Book.open(@simple_file, :visible => true)
        sheet = new_book[0]
        sheet[1,1].Value.should == @old_cell_value
      end

      it "should not change the value and use the hidden Excel instance" do
        new_excel = Excel.new(:reuse => false)
        Book.for_reading(@simple_file, :hidden) do |book|
          sheet = book[0]
          cell = sheet[1,1]
          sheet[1,1] = cell.value == "foo" ? "bar" : "foo"
          book.excel.should_not == @book.excel
          book.excel.should_not == new_excel
          book.excel.visible.should be_false
          book.excel.displayalerts.should be_false
        end
        new_book = Book.open(@simple_file, :visible => true)
        sheet = new_book[0]
        sheet[1,1].Value.should == @old_cell_value
      end

      it "should change the value" do
        Book.for_modifying(@simple_file) do |book|
          sheet = book[0]
          cell = sheet[1,1]
          sheet[1,1] = cell.value == "foo" ? "bar" : "foo"
          book.excel.should == @book.excel
        end
        new_book = Book.open(@simple_file, :visible => true)
        sheet = new_book[0]
        sheet[1,1].Value.should_not == @old_cell_value
      end

      it "should change the value and use a given Excel" do
        new_excel = Excel.new(:reuse => false)
        another_excel = Excel.new(:reuse => false)
        Book.for_modifying(@simple_file, another_excel) do |book|
          sheet = book[0]
          cell = sheet[1,1]
          sheet[1,1] = cell.value == "foo" ? "bar" : "foo"
          book.excel.should == another_excel
        end
        new_book = Book.open(@simple_file, :visible => true)
        sheet = new_book[0]
        sheet[1,1].Value.should_not == @old_cell_value
      end

      it "should change the value and use the hidden Excel instance" do
        new_excel = Excel.new(:reuse => false)
        Book.for_modifying(@simple_file, :hidden) do |book|
          sheet = book[0]
          cell = sheet[1,1]
          sheet[1,1] = cell.value == "foo" ? "bar" : "foo"
          book.excel.should_not == @book.excel
          book.excel.should_not == new_excel
          book.excel.visible.should be_false
          book.excel.displayalerts.should be_false
        end
        new_book = Book.open(@simple_file, :visible => true)
        sheet = new_book[0]
        sheet[1,1].Value.should_not == @old_cell_value
      end
    end
  end

  
end
