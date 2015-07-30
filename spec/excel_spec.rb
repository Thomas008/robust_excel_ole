# -*- coding: utf-8 -*-

require File.join(File.dirname(__FILE__), './spec_helper')

$VERBOSE = nil

module RobustExcelOle

  describe Excel do

    before(:all) do
      Excel.close_all
    end

    before do
      @dir = create_tmpdir
      @simple_file = @dir + '/workbook.xls'
      @another_simple_file = @dir + '/another_workbook.xls'
      @invalid_name_file = 'b/workbook.xls'
    end

    after do
      Excel.close_all
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

    context "with existing excel" do

      before do
        @excel1 = Excel.create
      end

      it "should create different excel" do
        excel2 = Excel.create
        excel2.Hwnd.should_not == @excel1.Hwnd
      end

      it "should reuse existing excel" do
        excel2 = Excel.current
        excel2.Hwnd.should == @excel1.Hwnd
      end

      it "should reuse existing excel with default options for 'new'" do
        excel2 = Excel.new
        excel2.should be_a Excel
        excel2.Hwnd.should == @excel1.Hwnd
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
          @excel = Excel.create
          @book = Book.open(@simple_file)
          @book2 = Book.open(@another_simple_file)
          sheet = @book[0]
          sheet[1,1] = sheet[1,1].value == "foo" ? "bar" : "foo"
        end

        it "should list unsaved workbooks" do
          @book.Saved.should be_false
          @book2.Saved.should be_true
          @excel.unsaved_workbooks.should == [@book]
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
          }.to raise_error(ExcelErrorSaveUnknown)
        end

      end
    end
  end

  describe "RobustExcelOle" do
    context "#absolute_path" do
      it "should work" do
        RobustExcelOle::absolute_path("C:/abc").should == "C:\\abc"
        RobustExcelOle::absolute_path("C:\\abc").should == "C:\\abc"
        RobustExcelOle::absolute_path("C:abc").should == Dir.pwd.gsub("/","\\") + "\\abc"
        RobustExcelOle::absolute_path("C:abc").should == File.expand_path("abc").gsub("/","\\")
      end

      it "should return right absolute path name" do
        @filename = 'C:/Dokumente und Einstellungen/Zauberthomas/Eigene Dateien/robust_excel_ole/spec/book_spec.rb'
        RobustExcelOle::absolute_path(@filename).gsub("\\","/").should == @filename
      end
    end
  end
end

class TestError < RuntimeError  # :nodoc: #
end
