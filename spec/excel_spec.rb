# -*- coding: utf-8 -*-

require File.join(File.dirname(__FILE__), './spec_helper')

$VERBOSE = nil

module RobustExcelOle

  describe Excel do

    before(:all) do
      Excel.close_all
    end

    context "excel creation" do
      
      def creation_ok? # :nodoc: #
        @excel.alive?.should be_true
        @excel.Visible.should be_false
        @excel.DisplayAlerts.should be_false
        @excel.Name.should == "Microsoft Excel"
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
        Excel.close_all
        @excel1 = Excel.create
      end

      it "should create different excel" do
        excel2 = Excel.create
        #puts "@excel1 #{@excel1.Hwnd}"
        #puts "excel2  #{excel2.Hwnd}"
        excel2.Hwnd.should_not == @excel1.Hwnd
      end

      it "should reuse existing excel" do
        excel2 = Excel.current
        #puts "@excel1 #{@excel1.Hwnd}"
        #puts "excel2  #{excel2.Hwnd}"
        excel2.Hwnd.should == @excel1.Hwnd
      end

      it "should reuse existing excel with default options for 'new'" do
        excel2 = Excel.new
        #puts "@excel1 #{@excel1.Hwnd}"
        #puts "excel2  #{excel2.Hwnd}"
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
        Excel.close_all
        @excel1 = Excel.create
      end

      it "should be true with two identical excel applications" do
        excel2 = Excel.current
        excel2.should == @excel1
      end

      it "should be false with two different excel applications" do
        excel2 = Excel.create
        excel2.should_not == @excel1
      end

      it "should be false with non-Excel objects" do
        @excel1.should_not == "hallo"
        @excel1.should_not == 7
        @excel1.should_not == nil
      end

    end

    context "with Visible and DisplayAlerts" do

      before do
        Excel.close_all
      end

      it "should be visible" do
        excel = Excel.new(:visible => true)
        excel.Visible.should be_true
        excel.visible.should be_true
        excel.DisplayAlerts.should be_false
        excel.displayalerts.should be_false
        excel.visible = false
        excel.Visible.should be_false
        excel.visible.should be_false
      end

      it "should displayalerts" do        
        excel = Excel.new(:displayalerts => true)
        excel.DisplayAlerts.should be_true
        excel.displayalerts.should be_true
        excel.Visible.should be_false
        excel.visible.should be_false
        excel.displayalerts = false
        excel.DisplayAlerts.should be_false
        excel.displayalerts.should be_false
      end

      it "should visible and displayalerts" do
        excel = Excel.new(:visible => true)
        excel.visible.should be_true
        excel.displayalerts.should be_false
        excel2 = Excel.new(:reuse => true, :visible => true, :displayalerts => true)
        excel2.visible.should be_true
        excel2.displayalerts.should be_true        
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

class TestError < RuntimeError
end
