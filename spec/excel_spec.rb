# -*- coding: utf-8 -*-

require File.join(File.dirname(__FILE__), './spec_helper')

$VERBOSE = nil

module RobustExcelOle

  describe Excel do

    before(:all) do
      Excel.close_all(:hard => true)
    end

    before do
      @dir = create_tmpdir
      #print "tmpdir: "; p @dir
      @simple_file = @dir + '/workbook.xls'
      @another_simple_file = @dir + '/another_workbook.xls'
      @different_file = @dir + '/different_workbook.xls'
      @invalid_name_file = 'b/workbook.xls'
    end

    after do
      Excel.close_all(:hard => true)
      rm_tmp(@dir)
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

      it "should work with 'create' " do
        @excel = Excel.create
        creation_ok?
      end
    end

    context "with identity transparence" do

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

    context "with Excel processes" do
        
      before do
        @excel1 = Excel.create
        p "hwnd: #{@excel1.hwnd}"
      end

      it "should show processes" do        
        Excel.get_excel_processes       
      end

      it "should kill Excel processes" do
        Excel.kill_excel_processes
      end

    end

    context "with reanimating Excel instances" do

      it "should reanimate a single Excel instance" do
        excel1 = Excel.create
        excel1.close
        excel1 = Excel.new(:reuse => false)
        excel1.should be_a Excel
        excel1.should be_alive
      end

      it "should reanimate an Excel instance and keep identity transparence" do        
        excel1 = Excel.create
        excel2 = excel1
        excel3 = Excel.create       
        #Excel.print_hwnd2excel
        excel2.should === excel1
        excel2.Hwnd.should == excel1.Hwnd        
        excel3.should_not == excel1
        excel3.Hwnd.should_not == excel1.Hwnd
        excel1.close
        excel1.should_not be_alive
        excel2.should_not be_alive
        excel3.should be_alive
        #Excel.print_hwnd2excel
        #excel1.reanmiate
        #excel1 = Excel.new(:reuse => false)
        excel1 = Excel.create
        excel1.should be_alive
        
        # necessary?
        #excel2.should be_alive
        #excel2.should === @excel1
        #excel2.Hwnd.should == @excel1.Hwnd
        #excel3.should_not == @excel1
        #excel3.Hwnd.should_not == @excel1.Hwnd
      end

      it "should reanimate an Excel instance and keep visible and displayalerts" do        
        excel1 = Excel.new(:reuse => false, :visible => true, :displayalerts => true)
        excel2 = excel1
        excel3 = Excel.create       
        #Excel.print_hwnd2excel
        excel2.should === excel1
        excel2.Hwnd.should == excel1.Hwnd        
        excel3.should_not == excel1
        excel3.Hwnd.should_not == excel1.Hwnd
        excel1.close
        excel1.should_not be_alive
        excel2.should_not be_alive
        excel3.should be_alive
        #Excel.print_hwnd2excel
        #excel1.reanimate
        excel1 = Excel.create        
        excel1.should be_alive
        excel1.Visible.should be_true
        excel1.DisplayAlerts.should be_true
        
        # necessary?
        #excel2.should be_alive
        #excel2.should === @excel1
        #excel2.Hwnd.should == @excel1.Hwnd
        #excel3.should_not == @excel1
        #excel3.Hwnd.should_not == @excel1.Hwnd
      end
    end    

    context "close excel instances" do
      def direct_excel_creation_helper  # :nodoc: #
        expect { WIN32OLE.connect("Excel.Application") }.to raise_error
        sleep 0.1
        excel1 = WIN32OLE.new("Excel.Application")
        excel1.Workbooks.Add
        excel2 = WIN32OLE.new("Excel.Application")
        excel2.Workbooks.Add
        expect { WIN32OLE.connect("Excel.Application") }.to_not raise_error
      end

      it "simple file with default" do
        Excel.close_all
        direct_excel_creation_helper
        Excel.close_all
        sleep 0.1
        expect { WIN32OLE.connect("Excel.Application") }.to raise_error
      end
    end

=begin
      # testing private methods
      context "close_excel" do

        before do
          @book = Book.open(@simple_file, :visible => true)
          @excel = @book.excel
          @book2 = Book.open(@simple_file, :force_excel => :new, :visible => true)
          @excel2 = @book2.excel
        end

        it "should close one Excel" do
          @excel.should be_alive
          @excel2.should be_alive
          @book.should be_alive
          @book2.should be_alive
          @excel.close_excel(:hard => false)
          @excel.should_not be_alive
          @book.should_not be_alive
          @excel2.should be_alive
          @book2.should be_alive
        end
      end

=end

    describe "close_all" do

      context "with saved workbooks" do

        before do
          book = Book.open(@simple_file, :visible => true)
          book2 = Book.open(@simple_file, :force_excel => :new, :visible => true)
          @excel = book.excel
          @excel2 = book2.excel
        end

        it "should close the Excel instances" do
          @excel.should be_alive
          @excel2.should be_alive
          Excel.close_all
          @excel.should_not be_alive
          @excel2.should_not be_alive
        end
      end
    end

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
      end

      context "with an unsaved workbook" do

        before do
          @excel = Excel.create
          @book = Book.open(@simple_file)
          sheet = @book[0]
          @old_cell_value = sheet[1,1].value
          sheet[1,1] = sheet[1,1].value == "foo" ? "bar" : "foo"
        end

        it "should raise an error" do
          @excel.should be_alive
          @book.should be_alive
          @book.saved.should be_false
          expect{
            @excel.close(:if_unsaved => :raise)
          }.to raise_error(ExcelErrorClose, "Excel contains unsaved workbooks")
          @excel.should be_alive
          @book.should be_alive
        end

        it "should close the Excel without saving the workbook" do
          @excel.should be_alive
          @book.should be_alive
          @book.saved.should be_false
          @excel.close(:if_unsaved => :forget)
          @excel.should_not be_alive
          @book.should_not be_alive
          new_book = Book.open(@simple_file)
          new_sheet = new_book[0]
          new_sheet[1,1].value.should == @old_cell_value
          new_book.close          
        end

        it "should close the Excel with saving the workbook" do
          @excel.should be_alive
          @book.should be_alive
          @book.saved.should be_false
          @excel.close(:if_unsaved => :save)
          @excel.should_not be_alive
          @book.should_not be_alive
          new_book = Book.open(@simple_file)
          new_sheet = new_book[0]
          new_sheet[1,1].value.should_not == @old_cell_value
          new_book.close          
        end

        it "should raise an error for invalid option" do
          expect {
            @excel.close(:if_unsaved => :invalid_option)
          }.to raise_error(ExcelErrorClose, ":if_unsaved: invalid option: invalid_option") 
        end

        it "should raise an error by default" do
          @excel.should be_alive
          @book.should be_alive
          @book.saved.should be_false
          expect{
            @excel.close
          }.to raise_error(ExcelErrorClose, "Excel contains unsaved workbooks")
          @excel.should be_alive
          @book.should be_alive
        end
      end
    end

#        context "with :if_unsaved => :alert" do
#          before do
#            @key_sender = IO.popen  'ruby "' + File.join(File.dirname(__FILE__), '/helpers/key_sender.rb') + '" "Microsoft Excel" '  , "w"
#          end
#
#          after do
#            @key_sender.close
#          end
#
#         possible_answers = [:yes, :no, :cancel]
#          possible_answers.each_with_index do |answer, position|
#            it "should" + (answer == :yes ? "" : " not") + " the unsaved book and" + (answer == :cancel ? " not" : "") + " close it" + "if user answers '#{answer}'" do
#            # "Yes" is the  default. "No" is right of "Yes", "Cancel" is right of "No" --> language independent
#            @key_sender.puts  "{right}" * position + "{enter}"
#          end
#        end
#      end


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

      it "should be false with dead Excel objects" do
        excel2 = Excel.current
        Excel.close_all
        excel2.should_not == @excel1
      end

    end

    context "with Visible and DisplayAlerts" do

      it "should create Excel visible" do
        excel = Excel.new(:visible => true)
        excel.Visible.should be_true
        excel.visible.should be_true
        excel.DisplayAlerts.should be_false
        excel.displayalerts.should be_false
        excel.visible = false
        excel.Visible.should be_false
        excel.visible.should be_false
      end

      it "should create Excel with DispayAlerts enabled" do        
        excel = Excel.new(:displayalerts => true)
        excel.DisplayAlerts.should be_true
        excel.displayalerts.should be_true
        excel.Visible.should be_false
        excel.visible.should be_false
        excel.displayalerts = false
        excel.DisplayAlerts.should be_false
        excel.displayalerts.should be_false
      end

      it "should keep visible and displayalerts values when reusing Excel" do
        excel = Excel.new(:visible => true)
        excel.visible.should be_true
        excel.displayalerts.should be_false
        excel2 = Excel.new(:displayalerts => true)
        excel2.should == excel
        excel.visible.should be_true
        excel.displayalerts.should be_true        
      end

      it "should keep displayalerts and visible values when reusing Excel" do
        excel = Excel.new(:displayalerts => true)
        excel.visible.should be_false
        excel.displayalerts.should be_true
        excel2 = Excel.new(:visible => true)
        excel2.should == excel
        excel.visible.should be_true
        excel.displayalerts.should be_true        
      end

    end

    context "with displayalerts" do
      before do
        @excel1 = Excel.new(:displayalerts => true)
        @excel2 = Excel.new(:displayalerts => false, :reuse => false)
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

    end

    context "method delegation for capitalized methods" do
      before do
        @excel1 = Excel.new
      end

      it "should raise WIN32OLERuntimeError" do
        expect{ @excel1.NonexistingMethod }.to raise_error(VBAMethodMissingError)
      end

      it "should raise NoMethodError for uncapitalized methods" do
        expect{ @excel1.nonexisting_method }.to raise_error(NoMethodError)
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
          @book2 = Book.open(@another_simple_file)
          @book3 = Book.open(@different_file, :read_only => true)
          sheet = @book[0]
          sheet[1,1] = sheet[1,1].value == "foo" ? "bar" : "foo"
          sheet3 = @book3[0]
          sheet3[1,1] = sheet3[1,1].value == "foo" ? "bar" : "foo"
        end

        it "should list unsaved workbooks" do
          @book.Saved.should be_false
          @book2.Save
          @book2.Saved.should be_true
          @book3.Saved.should be_false
          excel = @book.excel
          # unsaved_workbooks yields different WIN32OLE objects than book.workbook
          uw_names = []
          excel.unsaved_workbooks.each {|uw| uw_names << uw.Name}
          uw_names.should == [@book.workbook.Name]
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
          workbook.FullName.should == RobustExcelOle::absolute_path(@file_name)
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
          workbook.FullName.should == RobustExcelOle::absolute_path(@file_name)
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
          workbook.FullName.should == RobustExcelOle::absolute_path(@file_name)
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
  end
end

class TestError < RuntimeError  # :nodoc: #
end
