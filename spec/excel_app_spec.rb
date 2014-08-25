# -*- coding: utf-8 -*-

require File.join(File.dirname(__FILE__), './spec_helper')

$VERBOSE = nil

module RobustExcelOle

  describe ExcelApp do
    before do
      @dir = create_tmpdir
      @simple_file = @dir + '/simple.xls'
    end

    after do
      rm_tmp(@dir)
    end

    save_path = "C:" + "/" + "simple_save.xls"

    context "app creation" do
      after do
        ExcelApp.close_all
      end

      def creation_ok?
        @app.alive?.should == true
        @app.Visible.should == false
        @app.DisplayAlerts.should == false
        @app.Name.should == "Microsoft Excel"
      end

      it "should work with 'new' " do
        @app = ExcelApp.new
        creation_ok?
      end

      it "should work with 'new' " do
        @app = ExcelApp.new(:reuse => false)
        creation_ok?
      end

      it "should work with 'create' " do
        @app = ExcelApp.create
        creation_ok?
      end

    end



    context "with existing app" do

      before do
        ExcelApp.close_all
        @app1 = ExcelApp.create
      end

      after do
        ExcelApp.close_all
      end

      it "should create different app" do
        app2 = ExcelApp.create
        #puts "@app1 #{@app1.Hwnd}"
        #puts "app2  #{app2.Hwnd}"
        app2.Hwnd.should_not == @app1.Hwnd
      end

      it "should reuse existing app" do
        app2 = ExcelApp.reuse_if_possible
        #puts "@app1 #{@app1.Hwnd}"
        #puts "app2  #{app2.Hwnd}"
        app2.Hwnd.should == @app1.Hwnd
      end

      it "should reuse existing app with default options for 'new'" do
        app2 = ExcelApp.new
        #puts "@app1 #{@app1.Hwnd}"
        #puts "app2  #{app2.Hwnd}"
        app2.Hwnd.should == @app1.Hwnd
      end

    end

    context "close excel instances" do
      it "simple file with default" do
        RobustExcelOle::ExcelApp.close_all
        expect { WIN32OLE.connect("Excel.Application") }.to raise_error
        sleep 0.1      
        exl1 = WIN32OLE.new("Excel.Application")
        exl1.Workbooks.Add 
        exl2 = WIN32OLE.new("Excel.Application")
        exl2.Workbooks.Add 
        expect { WIN32OLE.connect("Excel.Application") }.to_not raise_error
        RobustExcelOle::ExcelApp.close_all
        sleep 0.1
        expect { WIN32OLE.connect("Excel.Application") }.to raise_error
      end
    end

    describe "==" do
      before do
        ExcelApp.close_all
        @app1 = ExcelApp.create
      end

      after do
        ExcelApp.close_all
      end

      it "should be true with two identical excel applications" do
        app2 = ExcelApp.reuse_if_possible
        app2.should == @app1
      end

      it "should be false with two different excel applications" do
        app2 = ExcelApp.create
        app2.should_not == @app1
      end

      it "should be false with non-ExcelApp objects" do
        @app1.should_not == "hallo"
        @app1.should_not == 7
        @app1.should_not == nil
      end

    end

    context "with Visible and DisplayAlerts" do

      before do
        ExcelApp.close_all
      end

      after (:each) do
        ExcelApp.close_all
      end

      it "should be visible" do
        app = ExcelApp.new(:visible => true)
        app.Visible.should == true
        app.DisplayAlerts.should == false
      end

      it "should displayalerts" do        
        app = ExcelApp.new(:displayalerts => true)
        app.DisplayAlerts.should == true
        app.Visible.should == false
      end

      it "should displayalerts" do
        app = ExcelApp.new(:visible => true)
        app.Visible.should == true
        app.DisplayAlerts.should == false
        app2 = ExcelApp.new(:displayalerts => true)
        app2.Visible.should == true
        app2.DisplayAlerts.should == true
        sleep 1        
      end

    end

  end
end
