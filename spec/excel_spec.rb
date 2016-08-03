# -*- coding: utf-8 -*-

require File.join(File.dirname(__FILE__), './spec_helper')

$VERBOSE = nil

include General

module RobustExcelOle

  describe Excel do

    before(:all) do
      Excel.kill_all
    end

    before do
      @dir = create_tmpdir
      @simple_file = @dir + '/workbook.xls'
      @another_simple_file = @dir + '/another_workbook.xls'
      @different_file = @dir + '/different_workbook.xls'
      @invalid_name_file = 'b/workbook.xls'
    end

    after do
      Excel.kill_all
      #rm_tmp(@dir)
    end

    context "Illegal Refrence" do

      before do

        book1 = Book.open(@simple_file)
        book2 = Book.open(@simple_file, :force_excel => :new)
        a = book1.saved 
      end

      it "should not cause warning 'Illegal Reference probably recycled'" do
        Excel.close_all
        book = Book.open(@simple_file)
      end
    end

    context "excel creation" do
      
      def creation_ok? # :nodoc: #
        @excel.alive?.should be_true
        @excel.Visible.should be_false
        @excel.DisplayAlerts.should be_false
        @excel.Name.should == "Microsoft Excel"
      end

      it "should access excel.excel" do
        excel = Excel.new(:reuse => false)
        excel.excel.should == excel
        excel.excel.should be_a Excel
      end

      it "should work with 'new' " do
        @excel = Excel.new
        creation_ok?
      end

      it "should work with 'new' " do
        @excel = Excel.new(:reuse => false)
        creation_ok?
      end

      it "should work with 'reuse' " do
        excel = Excel.create
        @excel = Excel.new(:reuse => true)
        creation_ok?
        @excel.should === excel
        @excel = Excel.current
        creation_ok?
        @excel.should === excel
      end

      it "should work with 'create' " do
        excel = Excel.create
        @excel = Excel.new(:reuse => false)
        creation_ok?
        @excel.should_not == excel
        @excel = Excel.create
        creation_ok?
        @excel.should_not == excel
      end

      context "lifting an Excel instance given as WIN32Ole object" do

        before do
          @book = Book.open(@simple_file)
          @excel = @book.excel          
        end

        it "lifts an Excel instance given as WIN32Ole object" do    
          win32ole_excel = WIN32OLE.connect(@book.ole_workbook.Fullname).Application
          excel = Excel.new(win32ole_excel)
          excel.should be_a Excel
          excel.should be_alive
          excel.should === @excel
        end

        it "lifts an Excel instance given as WIN32Ole object with options" do    
          @excel.Visible = true
          @excel.DisplayAlerts = true
          win32ole_excel = WIN32OLE.connect(@book.ole_workbook.Fullname).Application
          excel = Excel.new(win32ole_excel)
          excel.should be_a Excel
          excel.should be_alive
          excel.should === @excel
          excel.Visible.should be_true
          excel.DisplayAlerts.should be_true
        end

      end
    end

    context "identity transparence" do

      before do
        @excel1 = Excel.create
      end

      it "should create different Excel instances" do
        excel2 = Excel.create
        excel2.should_not == @excel1
        excel2.Hwnd.should_not == @excel1.Hwnd
      end

      it "should reuse the existing Excel instances" do
        excel2 = Excel.current
        excel2.should === @excel1
        excel2.Hwnd.should == @excel1.Hwnd
      end

      it "should reuse existing Excel instance with default options for 'new'" do
        excel2 = Excel.new
        excel2.should === @excel1
        excel2.Hwnd.should == @excel1.Hwnd
      end

      it "should yield the same Excel instances for the same Excel objects" do
        excel2 = @excel1
        excel2.Hwnd.should == @excel1.Hwnd
        excel2.should === @excel1
      end
    end

    context "current" do

      it "should connect to the Excel" do
        excel1 = Excel.create
        excel2 = Excel.current
        excel2.should === excel1
      end

      it "should create a new Excel if there is no Excel to connect with" do
        excel1 = Excel.create
        excel1.close
        excel2 = Excel.current
        excel1.should_not be_alive
        excel2.should be_alive
        Excel.excels_number.should == 1
      end

      it "should make the Excel instance not alive if the Excel that was connected with was closed" do
        excel1 = Excel.create
        excel2 = Excel.current
        excel1.close
        excel1.should_not be_alive
        excel2.should_not be_alive
        sleep 0.2
        Excel.excels_number.should == 0
      end

      it "should reuse the first opened Excel instance if not the first opened Excel instance was closed" do
        excel1 = Excel.create
        excel2 = Excel.create
        excel2.close
        excel3 = Excel.current
        excel3.should === excel1
      end

      it "should reuse the Excel that was not closed" do
        excel1 = Excel.create
        excel2 = Excel.create
        excel1.close
        sleep 0.2
        excel3 = Excel.current
        excel3.should === excel2        
        excel3.Hwnd.should == excel2.Hwnd
      end
    end

    context "excels_number" do
        
      it "should return right number of excel instances" do        
        Excel.excels_number.should == 0
        e1 = Excel.create
        Excel.excels_number.should == 1
        e2 = Excel.create
        Excel.excels_number.should == 2
      end
    end

    context "kill Excel processes hard" do

      before do
        @excel1 = Excel.create
        @excel2 = Excel.create
      end

      it "should kill Excel processes" do
        Excel.kill_all
        @excel1.alive?.should be_false
        @excel2.alive?.should be_false
      end
    end

    context "recreating Excel instances" do

      context "with a single Excel instance" do

        before do
          @book1 = Book.open(@simple_file)
          @excel1 = @book1.excel
        end

        it "should recreate an Excel instance" do
          @excel1.close
          @excel1.should_not be_alive
          @excel1.recreate
          @excel1.should be_a Excel
          @excel1.should be_alive
          @excel1.Visible.should be_false
          @excel1.DisplayAlerts.should be_false
          @book1.should_not be_alive
          @book1.reopen
          @book1.should be_alive
          @excel1.close
          @excel1.should_not be_alive        
        end

        it "should recreate an Excel instance with old visible and displayalerts values" do
          @excel1.visible = true
          @excel1.displayalerts = true
          @excel1.close
          @excel1.should_not be_alive
          @excel1.recreate
          @excel1.should be_a Excel
          @excel1.should be_alive
          @excel1.Visible.should be_true
          @excel1.DisplayAlerts.should be_true
          @book1.reopen
          @book1.should be_alive
          @excel1.close
          @excel1.should_not be_alive
        end

        it "should recreate an Excel instance with new visible and displayalerts values" do
          @excel1.close
          @excel1.should_not be_alive
          @excel1.recreate(:visible => true, :displayalerts => true)
          @excel1.should be_a Excel
          @excel1.should be_alive
          @excel1.Visible.should be_true
          @excel1.DisplayAlerts.should be_true
          @book1.reopen
          @book1.should be_alive
          @excel1.close
          @excel1.should_not be_alive
        end

        it "should recreate an Excel instance and reopen the book" do
          @excel1.close
          @excel1.should_not be_alive
          @excel1.recreate(:reopen_workbooks => true)
          @excel1.should be_a Excel
          @excel1.should be_alive
          @excel1.Visible.should be_false
          @excel1.DisplayAlerts.should be_false
          @book1.should be_alive
          @excel1.close
          @excel1.should_not be_alive
        end
      end

      context "with several Excel instances" do

        before do
          @book1 = Book.open(@simple_file)      
          @book2 = Book.open(@another_simple_file, :force_excel => @book1)
          @book3 = Book.open(@different_file, :force_excel => :new)
          @excel1 = @book1.excel
          @excel3 = @book3.excel
          @excel1.visible = true
          @excel3.displayalerts = true
        end

        it "should recreate several Excel instances" do  
          @excel1.close
          @excel3.close
          @excel1.should_not be_alive
          @excel3.should_not be_alive
          @excel1.recreate(:reopen_workbooks => true, :displayalerts => true)
          @excel1.should be_alive
          @excel1.should be_a Excel
          @excel1.Visible.should be_true
          @excel1.DisplayAlerts.should be_true
          @book1.should be_alive
          @book2.should be_alive
          @excel3.recreate(:visible => true)
          @excel3.should be_alive
          @excel3.should be_a Excel
          @excel3.Visible.should be_true
          @excel3.DisplayAlerts.should be_true
          @book3.reopen
          @book3.should be_alive
          @book3.excel.should == @excel3
          @excel1.close
          @excel1.should_not be_alive
          @excel3.close
          @excel3.should_not be_alive
        end
      end    
    end

    # Error
    context "close excel instances" do
      def direct_excel_creation_helper  # :nodoc: #
        expect { WIN32OLE.connect("Excel.Application") }.to raise_error
        sleep 0.1
        ole_excel1 = WIN32OLE.new("Excel.Application")
        ole_excel1.Workbooks.Add
        ole_excel2 = WIN32OLE.new("Excel.Application")
        ole_excel2.Workbooks.Add
        expect { WIN32OLE.connect("Excel.Application") }.to_not raise_error
      end

      it "simple file with default" do
        Excel.close_all
        direct_excel_creation_helper
        sleep 3
        puts "excels_number: #{Excel.excels_number}"
        sleep 1
        Excel.close_all
        sleep 0.1
        expect { WIN32OLE.connect("Excel.Application") }.to raise_error
      end
    end

    describe "close_all" do

      context "with saved workbooks" do

        it "should do with no Excel instances" do
          expect{
            Excel.close_all
          }.to_not raise_error
        end

        it "should close one Excel instance" do
          excel1 = Excel.create
          Excel.close_all
          excel1.should_not be_alive
          Excel.excels_number.should == 0
        end

        it "should close two Excel instances" do
          excel1 = Excel.create
          excel2 = Excel.create
          Excel.close_all
          excel1.should_not be_alive
          excel2.should_not be_alive
          Excel.excels_number.should == 0
        end
      end

      context "with unsaved_workbooks" do
        
        before do
          book1 = Book.open(@simple_file, :read_only => true)
          book2 = Book.open(@simple_file, :force_excel => :new)
          book3 = Book.open(@another_simple_file, :force_excel => book2.excel)
          book4 = Book.open(@different_file, :force_excel => :new)
          @excel1 = book1.excel
          @excel2 = book2.excel
          @excel4 = book4.excel
          sheet2 = book2.sheet(1)
          @old_cell_value2 = sheet2[1,1].value
          sheet2[1,1] = sheet2[1,1].value == "foo" ? "bar" : "foo"
          sheet3 = book3.sheet(1)
          @old_cell_value3 = sheet3[1,1].value
          sheet3[1,1] = sheet3[1,1].value == "foo" ? "bar" : "foo"
        end

        it "should close the first Excel without unsaved workbooks and then raise an error" do
          expect{
            Excel.close_all(:if_unsaved => :raise)
          }.to raise_error(ExcelErrorClose, "Excel contains unsaved workbooks")
          @excel1.should_not be_alive
          @excel2.should be_alive
          @excel4.should be_alive
        end

        it "should close the first Excel without unsaved workbooks and then raise an error per default" do
          expect{
            Excel.close_all(:if_unsaved => :raise)
          }.to raise_error(ExcelErrorClose, "Excel contains unsaved workbooks")
          @excel1.should_not be_alive
          @excel2.should be_alive
          @excel4.should be_alive
        end

        it "should raise an error for invalid option" do
          expect {
            Excel.close_all(:if_unsaved => :invalid_option)
          }.to raise_error(ExcelErrorClose, ":if_unsaved: invalid option: :invalid_option") 
        end

      end
    end

=begin
      context "unsaved_workbooks" do

        # Error
        it "should save the unsaved workbook" do
          book1 = Book.open(@simple_file, :visible => true)
          excel1 = book1.excel
          book1.sheet(1)[1,1] = "bar"
          book1.Saved.should be_false
          Excel.close_all(:if_unsaved => :save)
        end

        # Error
        it "should forget the unsaved workbook" do
          book1 = Book.open(@simple_file, :visible => true)
          excel1 = book1.excel
          excel1.displayalerts = false
          sheet1 = book1.sheet(1)
          old_cell_value1 = sheet1[1,1].value
          sheet1[1,1] = sheet1[1,1].value == "foo" ? "bar" : "foo"
          book1.Saved.should be_false
          Excel.close_all(:if_unsaved => :forget)
        end

        # Error
        it "should close the Excel instances with saving the unsaved workbooks" do
          Excel.close_all(:if_unsaved => :save)
          @excel1.should_not be_alive
          @excel2.should_not be_alive
          @excel4.should_not be_alive
          new_book2 = Book.open(@simple_file)
          new_sheet2 = new_book2.sheet(1)
          new_sheet2[1,1].value.should_not == @old_cell_value2
          new_book2.close   
          new_book3 = Book.open(@another_simple_file)
          new_sheet3 = new_book3.sheet(1)
          new_sheet3[1,1].value.should_not == @old_cell_value3
          new_book3.close
        end

        # Error: timeout
        it "should close the Excel instances without saving the unsaved workbooks" do          
          @excel1.displayalerts = @excel2.displayalerts = @excel4.displayalerts = false
          Excel.close_all(:if_unsaved => :forget)
          @excel1.should_not be_alive
          @excel2.should_not be_alive
          @excel4.should_not be_alive
          new_book2 = Book.open(@simple_file)
          new_sheet2 = new_book2.sheet(1)
          new_sheet2[1,1].value.should == @old_cell_value2
          new_book2.close   
          new_book3 = Book.open(@another_simple_file)
          new_sheet3 = new_book3.sheet(1)
          new_sheet3[1,1].value.should == @old_cell_value3
          new_book3.close   
        end

        
      end
=end

=begin
      context "with :if_unsaved => :alert" do
        before do
          @key_sender = IO.popen  'ruby "' + File.join(File.dirname(__FILE__), '/helpers/key_sender.rb') + '" "Microsoft Excel" '  , "w"
          book2 = Book.open(@simple_file, :force_excel => :new)
          #@excel1 = book1.excel
          @excel2 = book2.excel
          sheet2 = book2.sheet(1)
          @old_cell_value2 = sheet2[1,1].value
          sheet2[1,1] = sheet2[1,1].value == "foo" ? "bar" : "foo"
        end

        after do
          @key_sender.close
        end

        it "should save if user answers 'yes'" do
          @key_sender.puts "{enter}"
          Excel.close_all(:if_unsaved => :alert)
          @excel2.should_not be_alive
          new_book2 = Book.open(@simple_file)
          new_sheet2 = new_book2.sheet(1)
          new_sheet2[1,1].value.should_not == @old_cell_value2
          new_book2.close   
        end

        it "should not save if user answers 'no'" do            
          @key_sender.puts "{right}{enter}"
          @key_sender.puts "{right}{enter}"
          Excel.close_all(:if_unsaved => :alert)
          @excel2.should_not be_alive
          new_book2 = Book.open(@simple_file)
          new_sheet2 = new_book2.sheet(1)
          new_sheet2[1,1].value.should == @old_cell_value2
          new_book2.close   
        end

        #it "should not save if user answers 'cancel'" do
        #  @key_sender.puts "{left}{enter}"
        #  @key_sender.puts "{left}{enter}"
        #  @key_sender.puts "{left}{enter}"
        #  expect{
        #    Excel.close_all(:if_unsaved => :alert)
        #  }.to raise_error(ExcelUserCanceled, "close: canceled by user")
        # end

        it "should save if user answers 'yes'" do
          @excel1.displayalerts = @excel2.displayalerts = @excel4.displayalerts = true
          @key_sender.puts "{enter}"          
          Excel.close_all(:if_unsaved => :forget)
          @excel2.should_not be_alive
          new_book2 = Book.open(@simple_file)
          new_sheet2 = new_book2.sheet(1)
          new_sheet2[1,1].value.should_not == @old_cell_value2
          new_book2.close   
        end

        it "should not save if user answers 'no'" do  
          @excel1.displayalerts = @excel2.displayalerts = @excel4.displayalerts = true        
          @key_sender.puts "{right}{enter}"
          @key_sender.puts "{right}{enter}"
          Excel.close_all(:if_unsaved => :forget)
          @excel2.should_not be_alive
          new_book2 = Book.open(@simple_file)
          new_sheet2 = new_book2.sheet(1)
          new_sheet2[1,1].value.should == @old_cell_value2
          new_book2.close   
        end
      end
    end
=end    
    
    describe "close" do

      context "with a saved workbook" do

        before do
          @excel = Excel.create
          @book = Book.open(@simple_file)
        end

        it "should close the Excel" do
          @excel.should be_alive
          @book.should be_alive
          @excel.close
          @excel.should_not be_alive
          @book.should_not be_alive
        end

        it "should close the Excel hard" do
          @excel.should be_alive
          @book.should be_alive
          @excel.close(:hard => true)
          @excel.should_not be_alive
          @book.should_not be_alive
        end

        it "should close the Excel without destroying the others" do
          @excel.should be_alive
          excel2 = Excel.create
          @excel.close
          @excel.should_not be_alive
          excel2.should be_alive
        end

        it "should close the Excel hard without destroying the others" do
          @excel.should be_alive
          excel2 = Excel.create
          @excel.close(:hard => true)
          @excel.should_not be_alive
          excel2.should be_alive
        end
      end

      context "with unsaved workbooks" do

        before do
          @excel = Excel.create
          @book = Book.open(@simple_file)
          sheet = @book.sheet(1)
          @old_cell_value = sheet[1,1].value
          sheet[1,1] = sheet[1,1].value == "foo" ? "bar" : "foo"
          @book2 = Book.open(@another_simple_file)
          sheet2 = @book2.sheet(1)
          @old_cell_value2 = sheet2[1,1].value
          sheet2[1,1] = sheet2[1,1].value == "foo" ? "bar" : "foo"
        end

        it "should raise an error" do
          @excel.should be_alive
          @book.should be_alive
          @book.saved.should be_false
          @book2.should be_alive
          @book2.saved.should be_false
          expect{
            @excel.close(:if_unsaved => :raise)
          }.to raise_error(ExcelErrorClose, "Excel contains unsaved workbooks")
        end        

        it "should close the Excel without saving the workbook" do
          @excel.displayalerts = false
          @excel.should be_alive
          @excel.displayalerts = false
          @excel.close(:if_unsaved => :forget)
          @excel.should_not be_alive
          new_book = Book.open(@simple_file)
          new_sheet = new_book.sheet(1)
          new_sheet[1,1].value.should == @old_cell_value
          new_book.close          
          new_book2 = Book.open(@another_simple_file)
          new_sheet2 = new_book2.sheet(1)
          new_sheet2[1,1].value.should == @old_cell_value2
          new_book2.close 
        end

        it "should close the Excel with saving the workbook" do
          @excel.should be_alive
          @excel.close(:if_unsaved => :save)
          @excel.should_not be_alive
          new_book = Book.open(@simple_file)
          new_sheet = new_book.sheet(1)
          new_sheet[1,1].value.should_not == @old_cell_value
          new_book.close          
          new_book2 = Book.open(@another_simple_file)
          new_sheet2 = new_book2.sheet(1)
          new_sheet2[1,1].value.should_not == @old_cell_value2
          new_book2.close   
        end

        it "should close the Excel with saving the workbook" do
          @excel.should be_alive
          @excel.close(:if_unsaved => :keep_open)
          @excel.should be_alive
          @excel.close(:if_unsaved => :forget)
        end

        it "should raise an error for invalid option" do
          expect {
            @excel.close(:if_unsaved => :invalid_option)
          }.to raise_error(ExcelErrorClose, ":if_unsaved: invalid option: :invalid_option") 
        end

        it "should raise an error by default" do
          @excel.should be_alive
          expect{
            @excel.close
          }.to raise_error(ExcelErrorClose, "Excel contains unsaved workbooks")
        end
  
        it "should close the Excel without saving the workbook hard" do
          @excel.should be_alive
          @book.should be_alive
          @book.saved.should be_false
          @excel.close(:if_unsaved => :forget, :hard => true)
          @excel.should_not be_alive
          @book.should_not be_alive
          new_book = Book.open(@simple_file)
          new_sheet = new_book.sheet(1)
          new_sheet[1,1].value.should == @old_cell_value
          new_book.close    
          new_book.excel.close(:hard => true)
          procs = WIN32OLE.connect("winmgmts:\\\\.")
          processes = procs.InstancesOf("win32_process")     
          result = []
          processes.each do |p|
            result << p if p.name == "EXCEL.EXE"
          end
          result.should be_empty
        end

      end

      context "with :if_unsaved => :alert" do

        before do
          @key_sender = IO.popen  'ruby "' + File.join(File.dirname(__FILE__), '/helpers/key_sender.rb') + '" "Microsoft Excel" '  , "w"
          @excel = Excel.create
          @book = Book.open(@simple_file)
          sheet = @book.sheet(1)
          @old_cell_value = sheet[1,1].value
          sheet[1,1] = sheet[1,1].value == "foo" ? "bar" : "foo"
        end

        after do
          @key_sender.close
        end

        it "should save if user answers 'yes'" do
          # "Yes" is to the left of "No", which is the  default. --> language independent
          @excel.should be_alive
          @key_sender.puts "{enter}" 
          @excel.close(:if_unsaved => :alert)
          @excel.should_not be_alive
          new_book = Book.open(@simple_file)
          new_sheet = new_book.sheet(1)
          new_sheet[1,1].value.should_not == @old_cell_value
          new_book.close   
        end

        it "should not save if user answers 'no'" do            
          @excel.should be_alive
          @book.should be_alive
          @book.saved.should be_false
          @key_sender.puts "{right}{enter}"
          @excel.close(:if_unsaved => :alert)
          @excel.should_not be_alive
          @book.should_not be_alive
          new_book = Book.open(@simple_file)
          new_sheet = new_book.sheet(1)
          new_sheet[1,1].value.should == @old_cell_value
          new_book.close     
        end

        it "should not save if user answers 'cancel'" do
          # strangely, in the "cancel" case, the question will sometimes be repeated twice            
          @excel.should be_alive
          @book.should be_alive
          @book.saved.should be_false
          @key_sender.puts "{left}{enter}"
          @key_sender.puts "{left}{enter}"
          expect{
            @excel.close(:if_unsaved => :alert)
            }.to raise_error(ExcelUserCanceled, "close: canceled by user")
        end

        it "should save if user answers 'yes'" do
          # "Yes" is to the left of "No", which is the  default. --> language independent
          @excel.should be_alive
          @excel.displayalerts = true
          @key_sender.puts "{enter}" 
          @excel.close(:if_unsaved => :forget)
          @excel.should_not be_alive
          new_book = Book.open(@simple_file)
          new_sheet = new_book.sheet(1)
          new_sheet[1,1].value.should_not == @old_cell_value
          new_book.close   
        end

        it "should not save if user answers 'no'" do            
          @excel.should be_alive
          @excel.displayalerts = true
          @book.should be_alive
          @book.saved.should be_false
          @key_sender.puts "{right}{enter}"
          @excel.close(:if_unsaved => :forget)
          @excel.should_not be_alive
          @book.should_not be_alive
          new_book = Book.open(@simple_file)
          new_sheet = new_book.sheet(1)
          new_sheet[1,1].value.should == @old_cell_value
          new_book.close     
        end

        it "should not save if user answers 'cancel'" do
          # strangely, in the "cancel" case, the question will sometimes be repeated twice            
          @excel.should be_alive
          @excel.displayalerts = true
          @book.should be_alive
          @book.saved.should be_false
          @key_sender.puts "{left}{enter}"
          @key_sender.puts "{left}{enter}"
          @excel.close(:if_unsaved => :forget)
        end

      end
    
    end

    describe "alive" do

      it "should yield alive" do
        excel = Excel.create
        excel.alive?.should be_true
      end

      it "should yield not alive" do
        excel = Excel.create
        excel.close
        excel.alive?.should be_false
      end

    end

    describe "==" do
      before do
        @excel1 = Excel.create
      end

      it "should be true with two identical excel instances" do
        excel2 = Excel.current
        excel2.should == @excel1
      end

      it "should be false with two different excel instances" do
        excel2 = Excel.create
        excel2.should_not == @excel1
      end

      it "should be false with non-Excel objects" do
        @excel1.should_not == "hallo"
        @excel1.should_not == 7
        @excel1.should_not == nil
      end

      # Error
      it "should be false with dead Excel objects" do
        excel2 = Excel.current
        sleep 3
        Excel.close_all
        sleep 2
        excel2.should_not == @excel1
      end

    end

    context "workbooks_visible" do

      it "should not raise an error for an empty Excel instance" do
        excel = Excel.create
        expect{
          excel.workbooks_visible = true
        }.to_not raise_error
      end

      it "should make visible a workbook" do
        book1 = Book.open(@simple_file)
        book1.excel.workbooks_visible = true
        book1.excel.Visible.should be_true
        book1.Windows(book1.Name).Visible.should be_true
        book1.visible.should be_true
      end

      it "should make visible and invisible two workbooks" do
        book1 = Book.open(@simple_file)
        book2 = Book.open(@different_file)
        excel = book1.excel
        excel.workbooks_visible = true
        excel.Visible.should be_true
        book1.Windows(book1.Name).Visible.should be_true
        book1.visible.should be_true
        book2.Windows(book2.Name).Visible.should be_true
        book2.visible.should be_true
        excel.workbooks_visible = false
        excel.Visible.should be_true
        book1.Windows(book1.Name).Visible.should be_false
        book1.visible.should be_false
        book2.Windows(book2.Name).Visible.should be_false
        book2.visible.should be_false
      end

      it "should make visible all workbooks" do
        book1 = Book.open(@simple_file, :visible => true)
        book2 = Book.open(@different_file)
        excel = book1.excel
        excel.workbooks_visible = true
        excel.Visible.should be_true
        book1.Windows(book1.Name).Visible.should be_true
        book1.visible.should be_true
        book2.Windows(book2.Name).Visible.should be_true
        book2.visible.should be_true
      end

    end

    context "with Visible and DisplayAlerts" do

      it "should set Excel visible and invisible with current" do
        excel1 = Excel.create
        excel2 = Excel.current(:visible => true)
        excel2.Visible.should be_true
        excel2.visible.should be_true
        excel2.DisplayAlerts.should be_true
        excel2.displayalerts.should == :if_visible
      end

      it "should set Excel visible and invisible with current" do
        excel1 = Excel.new(:reuse => false, :visible => true)
        excel1.Visible.should be_true
        excel1.visible.should be_true
        excel1.DisplayAlerts.should be_true
        excel1.displayalerts.should == :if_visible
        excel1.visible = false
        excel1.Visible.should be_false
        excel1.visible.should be_false
        excel1.DisplayAlerts.should be_false
        excel1.displayalerts.should == :if_visible
        excel2 = Excel.current(:visible => true)
        excel2.Visible.should be_true
        excel2.visible.should be_true
        excel2.displayalerts.should == :if_visible
        excel2.DisplayAlerts.should be_true
      end

      it "should set Excel visible and invisible" do
        excel = Excel.new(:reuse => false, :visible => true)
        excel.Visible.should be_true
        excel.visible.should be_true
        excel.DisplayAlerts.should be_true
        excel.displayalerts.should == :if_visible
        excel6 = Excel.current
        excel6.should === excel
        excel6.Visible.should be_true
        excel.visible = false
        excel.Visible.should be_false
        excel.visible.should be_false
        excel.DisplayAlerts.should be_false
        excel.displayalerts.should == :if_visible
        excel7 = Excel.current
        excel7.should === excel
        excel7.Visible.should be_false
        excel7.DisplayAlerts.should be_false
        excel1 = Excel.create(:visible => true)
        excel1.should_not == excel
        excel1.Visible.should be_true
        excel1.visible.should be_true
        excel1.DisplayAlerts.should be_true
        excel1.displayalerts.should == :if_visible
        excel2 = Excel.create(:visible => false)
        excel2.Visible.should be_false
        excel2.visible.should be_false
        excel2.DisplayAlerts.should be_false
        excel2.displayalerts.should == :if_visible
        excel3 = Excel.current
        excel3.should === excel
        excel3.Visible.should be_false
        excel3.visible.should be_false
        excel3.DisplayAlerts.should be_false
        excel3.displayalerts.should == :if_visible
        excel4 = Excel.current(:visible => true)
        excel4.should === excel
        excel4.Visible.should be_true
        excel4.visible.should be_true
        #Error: 
        #excel4.DisplayAlerts.should be_true
        excel4.displayalerts.should == :if_visible
        excel5 = Excel.current(:visible => false)
        excel5.should === excel
        excel5.Visible.should be_false
        excel5.visible.should be_false
        excel5.DisplayAlerts.should be_false
        excel5.displayalerts.should == :if_visible
      end

      it "should enable or disable Excel DispayAlerts" do        
        excel = Excel.new(:reuse => false, :displayalerts => true)
        excel.DisplayAlerts.should be_true
        excel.displayalerts.should be_true
        excel.Visible.should be_false
        excel.visible.should be_false
        excel6 = Excel.current
        excel6.should === excel
        excel6.DisplayAlerts.should be_true
        excel6.displayalerts.should be_true
        excel6.Visible.should be_false
        excel6.visible.should be_false
        excel.displayalerts = false
        excel.DisplayAlerts.should be_false
        excel.displayalerts.should be_false
        excel.Visible.should be_false
        excel.visible.should be_false
        excel7 = Excel.current
        excel7.should === excel
        excel7.DisplayAlerts.should be_false
        excel7.displayalerts.should be_false
        excel7.Visible.should be_false
        excel7.visible.should be_false
        excel1 = Excel.create(:displayalerts => true)
        excel1.should_not == excel
        excel1.DisplayAlerts.should be_true
        excel1.displayalerts.should be_true
        excel1.Visible.should be_false
        excel1.visible.should be_false
        excel2 = Excel.create(:displayalerts => false)
        excel2.DisplayAlerts.should be_false
        excel2.displayalerts.should be_false
        excel2.Visible.should be_false
        excel2.visible.should be_false
        excel3 = Excel.current
        excel3.should === excel
        excel3.DisplayAlerts.should be_false
        excel3.displayalerts.should be_false
        excel3.Visible.should be_false
        excel3.visible.should be_false
        excel4 = Excel.current(:displayalerts => true)
        excel4.should === excel
        excel4.DisplayAlerts.should be_true
        excel4.displayalerts.should be_true
        excel4.Visible.should be_false
        excel4.visible.should be_false
        excel5 = Excel.current(:displayalerts => false)
        excel5.should === excel
        excel5.DisplayAlerts.should be_false
        excel5.displayalerts.should be_false
        excel5.Visible.should be_false
        excel5.visible.should be_false
      end

      it "should set Excel visible and displayalerts" do        
        excel = Excel.new(:reuse => false, :visible => true, :displayalerts => true)
        excel.DisplayAlerts.should be_true
        excel.displayalerts.should be_true
        excel.Visible.should be_true
        excel.visible.should be_true
        excel6 = Excel.current
        excel6.should === excel
        excel6.DisplayAlerts.should be_true
        excel6.displayalerts.should be_true
        excel6.Visible.should be_true
        excel6.visible.should be_true
        excel.displayalerts = false
        excel.DisplayAlerts.should be_false
        excel.displayalerts.should be_false
        excel.Visible.should be_true
        excel.visible.should be_true
        excel7 = Excel.current
        excel7.should === excel
        excel7.DisplayAlerts.should be_false
        excel7.displayalerts.should be_false
        excel7.Visible.should be_true
        excel7.visible.should be_true        
        excel2 = Excel.new(:reuse => false, :visible => true, :displayalerts => true)
        excel2.visible = false
        excel2.DisplayAlerts.should be_true
        excel2.displayalerts.should be_true
        excel2.Visible.should be_false
        excel2.visible.should be_false
        excel3 = Excel.new(:reuse => false, :visible => true, :displayalerts => false)
        excel3.Visible.should be_true
        excel3.DisplayAlerts.should be_false
        excel3 = Excel.new(:reuse => false, :visible => false, :displayalerts => true)
        excel3.Visible.should be_false
        excel3.DisplayAlerts.should be_true
        excel3 = Excel.new(:reuse => false, :visible => false, :displayalerts => false)
        excel3.Visible.should be_false
        excel3.DisplayAlerts.should be_false
        excel4 = Excel.create(:visible => true, :displayalerts => true)
        excel4.DisplayAlerts.should be_true
        excel4.displayalerts.should be_true
        excel4.Visible.should be_true
        excel4.visible.should be_true
        excel5 = Excel.current(:visible => true, :displayalerts => false)
        excel5.should === excel
        excel5.DisplayAlerts.should be_false
        excel5.displayalerts.should be_false
        excel5.Visible.should be_true
        excel5.visible.should be_true
        excel6 = Excel.current(:visible => false, :displayalerts => true)
        excel6.should === excel
        excel6.DisplayAlerts.should be_true
        excel6.displayalerts.should be_true
        excel6.Visible.should be_false
        excel6.visible.should be_false
      end

      it "should work with displayalerts == if_visible" do
        excel = Excel.new(:reuse => false, :visible => true, :displayalerts => :if_visible)
        excel.Visible.should be_true
        excel.DisplayAlerts.should be_true
        excel2 = Excel.new(:reuse => false, :visible => false, :displayalerts => :if_visible)
        excel2.Visible.should be_false
        excel2.DisplayAlerts.should be_false
        excel3 = Excel.new(:reuse => false, :displayalerts => :if_visible)
        excel3.Visible.should be_false
        excel3.DisplayAlerts.should be_false
        excel3.visible = true
        excel3.Visible.should be_true
        excel3.DisplayAlerts.should be_true
        excel3.visible = false
        excel3.Visible.should be_false
        excel3.DisplayAlerts.should be_false
      end

      it "should keep visible and displayalerts values when reusing Excel" do
        excel = Excel.new(:visible => true)
        excel.Visible.should be_true
        excel.DisplayAlerts.should be_true
        excel2 = Excel.new(:displayalerts => false)
        excel2.should == excel
        excel.Visible.should be_true
        excel.DisplayAlerts.should be_false        
      end

      it "should keep displayalerts and visible values when reusing Excel" do
        excel = Excel.new(:displayalerts => true)
        excel.Visible.should be_false
        excel.DisplayAlerts.should be_true
        excel2 = Excel.new(:visible => true)
        excel2.should == excel
        excel.Visible.should be_true
        excel.DisplayAlerts.should be_true        
      end

    end

    context "with resetting displayalerts values" do
      before do
        @excel1 = Excel.new(:displayalerts => true)
        @excel2 = Excel.new(:displayalerts => false, :reuse => false)
        @excel3 = Excel.new(:displayalerts => false, :visible => true, :reuse => false)
      end

      it "should turn off displayalerts" do
        @excel1.DisplayAlerts.should be_true
        begin
          @excel1.with_displayalerts false do
            @excel1.DisplayAlerts.should be_false
            raise TestError, "any_error"
          end
        rescue TestError
          @excel1.DisplayAlerts.should be_true
        end
      end
    
      it "should turn on displayalerts" do
        @excel2.DisplayAlerts.should be_false
        begin
          @excel1.with_displayalerts true do
            @excel1.DisplayAlerts.should be_true
            raise TestError, "any_error"
          end
        rescue TestError
          @excel2.DisplayAlerts.should be_false
        end
      end

      it "should set displayalerts to :if_visible" do
        @excel1.DisplayAlerts.should be_true
        begin
          @excel1.with_displayalerts :if_visible do
            @excel1.DisplayAlerts.should be_false
            @excel1.Visible.should be_false
            raise TestError, "any_error"
          end
        rescue TestError
          @excel1.DisplayAlerts.should be_true
        end
      end

      it "should set displayalerts to :if_visible" do
        @excel3.DisplayAlerts.should be_false
        begin
          @excel3.with_displayalerts :if_visible do
            @excel3.DisplayAlerts.should be_true
            @excel3.Visible.should be_true
            raise TestError, "any_error"
          end
        rescue TestError
          @excel3.DisplayAlerts.should be_false
        end
      end

    end

    context "with calculation" do

      before do
        @excel1 = Excel.new
      end

      it "should not set calculation mode when no workbook is opened" do
        old_calculation_mode = @excel1.Calculation
        @excel1.with_calculation(:automatic) do
          @excel1.Calculation.should == old_calculation_mode
        end
        @excel1.with_calculation(:manual) do
          @excel1.Calculation.should == old_calculation_mode
        end
        @excel1.set_calculation(:automatic)
        @excel1.Calculation.should == old_calculation_mode
        @excel1.set_calculation(:manual)
        @excel1.Calculation.should == old_calculation_mode
      end

      it "should set calculation mode to manual and reset to the previous value" do
        b = Book.open(@simple_file)
        calculation_mode_old = @excel1.Calculation
        calculation_before_save_old = @excel1.CalculateBeforeSave
        @excel1.with_calculation(:manual) do
          @excel1.Calculation.should == -4135
          @excel1.CalculateBeforeSave.should be_false
        end
        @excel1.Calculation.should == calculation_mode_old
        @excel1.CalculateBeforeSave.should == calculation_before_save_old
      end

      it "should set calculation mode to manual" do
        b = Book.open(@simple_file)
        calculation_mode_old = @excel1.Calculation
        calculation_before_save_old = @excel1.CalculateBeforeSave
        @excel1.set_calculation(:manual)
        @excel1.Calculation.should == -4135
        @excel1.CalculateBeforeSave.should be_false
      end

      it "should set calculation mode automatic" do
        b = Book.open(@simple_file)
        @excel1.with_calculation(:automatic) do
          @excel1.Calculation.should == -4105
          @excel1.CalculateBeforeSave.should be_true
        end
      end

      it "should set calculation mode to automatic as default" do
        b = Book.open(@simple_file)
        @excel1.with_calculation do
          @excel1.Calculation.should == -4105
          @excel1.CalculateBeforeSave.should be_true
        end
      end

    end

    context "method delegation for capitalized methods" do
      before do
        @excel1 = Excel.new
      end

      it "should raise WIN32OLERuntimeError" do
        expect{ @excel1.NonexistingMethod }.to raise_error(VBAMethodMissingError, /unknown VBA property or method :NonexistingMethod/)
      end

      it "should raise NoMethodError for uncapitalized methods" do
        expect{ @excel1.nonexisting_method }.to raise_error(NoMethodError)
      end

      it "should report that Excel is not alive" do
        @excel1.close
        expect{ @excel1.Nonexisting_method }.to raise_error(ExcelError, "method missing: Excel not alive")
      end

    end

    context "with hwnd and hwnd2excel" do
      
      before do
        @excel1 = Excel.new
        @excel2 = Excel.new(:reuse => false)
      end

      it "should yield the correct hwnd" do
        @excel1.Hwnd.should == @excel1.hwnd
        @excel2.Hwnd.should == @excel2.hwnd
      end

      it "should provide the same excel instances" do
        @excel1.should_not == @excel2
        excel3 = Excel.hwnd2excel(@excel1.hwnd)
        excel4 = Excel.hwnd2excel(@excel2.hwnd)
        @excel1.should == excel3
        @excel2.should == excel4
        excel3.should_not == excel4 
      end
    end

    describe "unsaved_workbooks" do

      context "with standard" do
        
        before do
          @book = Book.open(@simple_file)
          @book3 = Book.open(@different_file, :read_only => true)
          sheet = @book.sheet(1)
          sheet[1,1] = sheet[1,1].value == "foo" ? "bar" : "foo"
          sheet3 = @book3.sheet(1)
          sheet3[1,1] = sheet3[1,1].value == "foo" ? "bar" : "foo"
          @book2 = Book.open(@another_simple_file, :force_excel => :new)
          sheet2 = @book2.sheet(1)
          sheet2[1,1] = sheet2[1,1].value == "foo" ? "bar" : "foo"
        end

        it "should list unsaved workbooks" do
          @book.Saved.should be_false
          @book2.Saved.should be_false
          @book3.Saved.should be_false
          excel = @book.excel
          # unsaved_workbooks yields different WIN32OLE objects than book.workbook
          uw_names = []
          excel.unsaved_workbooks.each {|uw| uw_names << uw.Name}
          uw_names.should == [@book.ole_workbook.Name]
        end

        it "should list all unsaved workbooks" do
          result = []
          Excel.unsaved_workbooks_all.each do |unsaved_workbooks| 
            uw_names = []
            unsaved_workbooks.each {|uw| uw_names << uw.Name}
            result << uw_names
          end
          result.include?([@book.ole_workbook.Name]).should be_true
          result.include?([@book2.ole_workbook.Name]).should be_true
        end

      end
    end

    describe "generate workbook" do

      context "with standard" do
        
        before do
          @excel1 = Excel.create
          @file_name = @dir + '/bar.xls'
        end

        it "should generate a workbook" do
          workbook = @excel1.generate_workbook(@file_name)
          workbook.should be_a WIN32OLE
          workbook.Name.should == File.basename(@file_name)
          workbook.FullName.should == General::absolute_path(@file_name)
          workbook.Saved.should be_true
          workbook.ReadOnly.should be_false
          workbook.Sheets.Count.should == 3
          workbooks = @excel1.Workbooks
          workbooks.Count.should == 1
        end

        it "should generate the same workbook twice" do
          workbook = @excel1.generate_workbook(@file_name)
          workbook.should be_a WIN32OLE
          workbook.Name.should == File.basename(@file_name)
          workbook.FullName.should == General::absolute_path(@file_name)
          workbook.Saved.should be_true
          workbook.ReadOnly.should be_false
          workbook.Sheets.Count.should == 3
          workbooks = @excel1.Workbooks
          workbooks.Count.should == 1
          workbook2 = @excel1.generate_workbook(@file_name)
          workbook2.should be_a WIN32OLE
          workbooks = @excel1.Workbooks
          workbooks.Count.should == 2
        end

        it "should generate a workbook if one is already existing" do
          book = Book.open(@simple_file)
          workbook = @excel1.generate_workbook(@file_name)
          workbook.should be_a WIN32OLE
          workbook.Name.should == File.basename(@file_name)
          workbook.FullName.should == General::absolute_path(@file_name)
          workbook.Saved.should be_true
          workbook.ReadOnly.should be_false
          workbook.Sheets.Count.should == 3
          workbooks = @excel1.Workbooks
          workbooks.Count.should == 2
        end

        it "should raise error when book cannot be saved" do
          expect{
            workbook = @excel1.generate_workbook(@invalid_name_file)
          # not always Unknown ???? ToDo #*#
          #}.to raise_error(ExcelErrorSaveUnknown)
          }.to raise_error(ExcelErrorSave)
        end

      end
    end

    describe "nameval, set_nameval" do

      before do
        @book1 = Book.open(@dir + '/another_workbook.xls')
        @excel1 = @book1.excel
      end

      after do
        @book1.close(:if_unsaved => :forget)
      end   

      it "should return value of a defined name" do
        @excel1.nameval("firstcell").should == "foo"
        @excel1["firstcell"].should == "foo"
      end        

      it "should return default value if name not defined and default value is given" do
        @excel1.nameval("foo", :default => 2).should == 2
      end

      it "should evaluate a formula" do
        @excel1.nameval("named_formula").should == 4
        @excel1["named_formula"].should == 4
      end

      it "should raise an error if name not defined" do
        expect {
          @excel1.nameval("foo")
        }.to raise_error(ExcelError, /cannot find name "foo"/)
        expect {
        @excel1["foo"]
        }.to raise_error(ExcelError, /cannot find name "foo"/)
        expect {
          excel2 = Excel.create
          excel2.nameval("one")
        }.to raise_error(ExcelError, /cannot find name "one"/)
        expect {
          excel3 = Excel.create
          excel3["one"]
        }.to raise_error(ExcelError, /cannot find name "one"/)
      end

      it "should set a range to a value" do
        @excel1.nameval("firstcell").should == "foo"
        @excel1.set_nameval("firstcell","bar")
        @excel1.nameval("firstcell").should == "bar"
        @excel1["firstcell"] = "foo"
        @excel1.nameval("firstcell").should == "foo"
      end

      it "should raise an error if name cannot be evaluated" do
        expect{
          @excel1.set_nameval("foo", 1)
          }.to raise_error(ExcelError, /cannot find name "foo"/)
        expect{
          @excel1["foo"] = 1
          }.to raise_error(ExcelError, /cannot find name "foo"/)
      end
    end

    describe "rangeval, set_rangeval" do
      
      before do
        @book1 = Book.open(@dir + '/another_workbook.xls')
        @excel1 = @book1.excel
      end

      after do
        @book1.close(:if_unsaved => :forget)
      end   

      it "should return value of a locally defined name" do
        @excel1.rangeval("firstcell").should == "foo"          
      end        

      it "should return value of a defined name" do
        @excel1.rangeval("new").should == "foo"         
        @excel1.rangeval("one").should == 1.0    
        @excel1.rangeval("four").should == [[1,2],[3,4]]
        @excel1.rangeval("firstrow").should == [[1,2]]
      end    

      it "should return default value if name not defined and default value is given" do
        @excel1.rangeval("foo", :default => 2).should == 2
      end

      it "should raise an error if name not defined for the sheet" do
        expect {
          @excel1.rangeval("foo")
          }.to raise_error(ExcelError, /cannot find name "foo"/)
        expect {
          @excel1.rangeval("named_formula")
          }.to raise_error(ExcelError, /cannot find name "named_formula"/)
        expect {
          excel2 = Excel.create
          excel2.rangeval("one")
        }.to raise_error(ExcelError, /cannot find name "one"/)
      end
    
      it "should set a range to a value" do
        @excel1.rangeval("firstcell").should == "foo"
        @excel1.set_rangeval("firstcell","bar")
        @excel1.rangeval("firstcell").should == "bar"
      end

      it "should raise an error if name cannot be evaluated" do
        expect{
          @excel1.set_nameval("foo", 1)
        }.to raise_error(ExcelError, /cannot find name "foo"/)
      end

    end

  end
end

class TestError < RuntimeError  # :nodoc: #
end
