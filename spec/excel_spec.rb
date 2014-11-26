# -*- coding: utf-8 -*-

require File.join(File.dirname(__FILE__), './spec_helper')

$VERBOSE = nil

module RobustExcelOle

  describe Excel do

    before (:all) do
      Excel.close_all
    end

    context "app creation" do
      
      def creation_ok? # :nodoc: #
        @app.alive?.should == true
        @app.Visible.should == false
        @app.DisplayAlerts.should == false
        @app.Name.should == "Microsoft Excel"
      end

      it "should work with 'new' " do
        @app = Excel.new
        creation_ok?
      end

      it "should work with 'new' " do
        @app = Excel.new(:reuse => false)
        creation_ok?
      end

      it "should work with 'create' " do
        @app = Excel.create
        creation_ok?
      end

    end

    context "with existing app" do

      before do
        Excel.close_all
        @app1 = Excel.create
      end

      it "should create different app" do
        app2 = Excel.create
        #puts "@app1 #{@app1.Hwnd}"
        #puts "app2  #{app2.Hwnd}"
        app2.Hwnd.should_not == @app1.Hwnd
      end

      it "should reuse existing app" do
        app2 = Excel.current
        #puts "@app1 #{@app1.Hwnd}"
        #puts "app2  #{app2.Hwnd}"
        app2.Hwnd.should == @app1.Hwnd
      end

      it "should reuse existing app with default options for 'new'" do
        app2 = Excel.new
        #puts "@app1 #{@app1.Hwnd}"
        #puts "app2  #{app2.Hwnd}"
        app2.Hwnd.should == @app1.Hwnd
      end

    end

    context "close excel instances" do
      def direct_excel_creation_helper  # :nodoc: #
        expect { WIN32OLE.connect("Excel.Application") }.to raise_error
        sleep 0.1
        exl1 = WIN32OLE.new("Excel.Application")
        exl1.Workbooks.Add
        exl2 = WIN32OLE.new("Excel.Application")
        exl2.Workbooks.Add
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
        @app1 = Excel.create
      end

      it "should be true with two identical excel applications" do
        app2 = Excel.current
        app2.should == @app1
      end

      it "should be false with two different excel applications" do
        app2 = Excel.create
        app2.should_not == @app1
      end

      it "should be false with non-Excel objects" do
        @app1.should_not == "hallo"
        @app1.should_not == 7
        @app1.should_not == nil
      end

    end


    context "with :excel" do

      it "should reuse in given excel app" do
        app1 = Excel.new(:reuse => false)
        app2 = Excel.new(:reuse => false)
        app3 = Excel.new(:excel => app1)
        app4 = Excel.new(:excel => app2)
        app3.should == app1
        app4.should == app2
      end

    end

    context "with Visible and DisplayAlerts" do

      before do
        Excel.close_all
      end

      it "should be visible" do
        app = Excel.new(:visible => true)
        app.Visible.should == true
        app.DisplayAlerts.should == false
      end

      it "should displayalerts" do        
        app = Excel.new(:displayalerts => true)
        app.DisplayAlerts.should == true
        app.Visible.should == false
      end

      it "should visible and displayalerts" do
        app = Excel.new(:visible => true)
        app.Visible.should == true
        app.DisplayAlerts.should == false
        app2 = Excel.new(:displayalerts => true)
        app2.Visible.should == true
        app2.DisplayAlerts.should == true
      end

    end


    context "with displayalerts" do
      before do
        @app1 = Excel.new(:displayalerts => true)
      end

      it "should turn off displayalerts" do
        @app1.DisplayAlerts.should == true
        begin
          @app1.with_displayalerts false do
            @app1.DisplayAlerts.should == false
            raise TestError, "any_error"
          end
        rescue TestError
          @app1.DisplayAlerts.should == true
        end
      end
    
    end

    context "method delegation for capitalized methods" do
      before do
        @app1 = Excel.new
      end

      it "should raise WIN32OLERuntimeError" do
        expect{ @app1.NonexistingMethod }.to raise_error(VBAMethodMissingError)
      end

      it "should raise NoMethodError for uncapitalized methods" do
        expect{ @app1.nonexisting_method }.to raise_error(NoMethodError)
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

      it "should return right absoute path name" do
        @filename = 'C:/Dokumente und Einstellungen/Zauberthomas/Eigene Dateien/robust_excel_ole/spec/book_spec.rb'
        RobustExcelOle::absolute_path(@filename).gsub("\\","/").should == @filename
      end
    end

  end

end

class TestError < RuntimeError
end
