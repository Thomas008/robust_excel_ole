# -*- coding: utf-8 -*-

require File.join(File.dirname(__FILE__), './spec_helper')

$VERBOSE = nil

include General

module RobustExcelOle

  describe Excel do

    before(:all) do
      Excel.kill_all
      sleep 0.2
    end

    before do
      @dir = create_tmpdir
      @simple_file = @dir + '/workbook.xls'
      @another_simple_file = @dir + '/another_workbook.xls'
      @different_file = @dir + '/different_workbook.xls'
      @invalid_name_file = 'b/workbook.xls'
      @simple_file1 = @simple_file
      @different_file1 = @different_file
    end

    after do
      Excel.kill_all
      rm_tmp(@dir)
    end

    context "with connect and preserving options" do

      before do
        @ole_excel = WIN32OLE.new('Excel.Application')        
      end

      it "should preserve visible false" do
        @ole_excel.Visible.should be false
        excel = Excel.current
        excel.Hwnd.should == @ole_excel.Hwnd
        excel.Visible.should be false
      end

      it "should preserve visible true" do
        @ole_excel.Visible = true
        @ole_excel.Visible.should be true
        excel = Excel.current
        excel.Hwnd.should == @ole_excel.Hwnd
        excel.Visible.should be true
      end

      it "should set visible true" do
        @ole_excel.Visible.should be false
        excel = Excel.current(:visible => true)
        excel.Visible.should be true
      end

      it "should set visible false" do
        excel = Excel.current(:visible => false)
        excel.Visible.should be false
      end

      it "should set visible false, even if it was true before" do
        @ole_excel.Visible = true
        @ole_excel.Visible.should be true
        excel = Excel.current(:visible => false)
        excel.Visible.should be false
      end

      it "should preserve displayalerts true" do
        @ole_excel.DisplayAlerts.should be true
        excel = Excel.current
        excel.Hwnd.should == @ole_excel.Hwnd
        excel.DisplayAlerts.should be true
      end

      it "should preserve displayalerts false" do
        @ole_excel.DisplayAlerts = false
        @ole_excel.DisplayAlerts.should be false
        excel = Excel.current
        excel.DisplayAlerts.should be false
      end

      it "should set displayalerts true" do
        @ole_excel.DisplayAlerts.should be true
        excel = Excel.current(:displayalerts => true)
        excel.DisplayAlerts.should be true
      end

      it "should set displayalerts true, even if false before" do
        @ole_excel.DisplayAlerts = false
        @ole_excel.DisplayAlerts.should be false
        excel = Excel.current(:displayalerts => true)
        excel.DisplayAlerts.should be true
      end

      it "should set displayalerts false" do
        @ole_excel.DisplayAlerts.should be true
        excel = Excel.current(:displayalerts => false)
        excel.DisplayAlerts.should be false
      end

    end

    context "with already open Excel instances and an open unsaved workbook" do

      before do
        @ole_excel1 = WIN32OLE.new('Excel.Application')
        @ole_excel2 = WIN32OLE.new('Excel.Application')
        #ole_workbook1 = @ole_excel1.Workbooks.Open(@another_simple_file, { 'ReadOnly' => false })
        ole_workbook1 = @ole_excel1.Workbooks.Open(@another_simple_file, nil, false)
        ole_workbook1.Names.Item("firstcell").RefersToRange.Value = "foo"
      end

      it "should create a new Excel instance" do
        excel1 = Excel.new(:reuse => false)
        excel1.ole_excel.Hwnd.should_not == @ole_excel1.Hwnd
        excel1.ole_excel.Hwnd.should_not == @ole_excel2.Hwnd
      end

      it "should connect to the already opened Excel instance" do
        excel1 = Excel.new(:reuse => true)
        excel1.ole_excel.Hwnd.should == @ole_excel1.Hwnd
      end

      it "should connect to the already running Excel instance" do
        excel1 = Excel.create
        excel1.close
        sleep 0.2
        excel2 = Excel.current
        excel1.should_not be_alive
        excel2.should be_alive
        excel2.ole_excel.Hwnd.should == @ole_excel1.Hwnd
        Excel.excels_number.should == 2
      end

      it "should make the Excel instance not alive if the Excel that was connected with was closed" do
        excel1 = Excel.create
        excel2 = Excel.current
        excel1.close
        sleep 0.2
        excel1.should_not be_alive
        excel2.should be_alive
        excel2.ole_excel.Hwnd.should == @ole_excel1.Hwnd
      end

      it "should reuse the first opened Excel instance if not the first opened Excel instance was closed" do
        excel1 = Excel.create
        excel2 = Excel.create
        excel2.close
        sleep 0.2
        excel3 = Excel.current
        excel3.ole_excel.Hwnd.should == @ole_excel1.Hwnd
      end

      it "should reuse the Excel that was not closed" do
        excel1 = Excel.create
        excel2 = Excel.create
        excel1.close
        sleep 0.2
        excel3 = Excel.current
        excel3.ole_excel.Hwnd.should == @ole_excel1.Hwnd        
      end

      it "should kill hard all Excel instances" do
        excel1 = Excel.create
        Excel.kill_all
        excel1.should_not be_alive
        expect{
          @ole_excel1.Name
          }.to raise_error
        expect{
          @ole_excel2.Name
          }.to raise_error  
      end

      it "should close all Excel instances" do
        excel1 = Excel.create
        result = Excel.close_all(:if_unsaved => :forget)
        sleep 1
        expect{
          @ole_excel1.Name
          }.to raise_error
        expect{
          @ole_excel2.Name
          }.to raise_error  
        result.should == [2,0]
      end

      it "should recreate an Excel instance" do
        excel1 = Excel.create
        excel1.close
        excel1.should_not be_alive
        excel1.recreate
        excel1.should be_a Excel
        excel1.should be_alive
        excel1.ole_excel.Hwnd.should_not == @ole_excel1.Hwnd
        excel1.ole_excel.Hwnd.should_not == @ole_excel2.Hwnd
        Excel.excels_number.should == 3
      end

    end

    context "Illegal Refrence" do


      before do
        book1 = Workbook.open(@simple_file1)
        book2 = Workbook.open(@simple_file1, :force_excel => :new)
        sleep 1
        a = book1.saved 
      end

      it "should not cause warning 'Illegal Reference probably recycled'" do
        Excel.close_all
        book = Workbook.open(@simple_file)
      end
    end

    context "excel creation" do
      
      def creation_ok? # :nodoc:
        @excel.alive?
        @excel.Screenupdating.should be true
        @excel.Visible.should be false
        @excel.DisplayAlerts.should be false
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

        it "lifts an Excel instance given as WIN32OLE object and has same Hwnd" do
          ole_excel = WIN32OLE.new("Excel.Application")  
          reo_excel = Excel.new(ole_excel)
          reo_excel.ole_excel.Hwnd.should == ole_excel.Hwnd  
          reo_excel.Visible.should == false
          reo_excel.properties[:displayalerts].should == :if_visible
        end

        it "lifts an Excel instance given as WIN32OLE object and set options" do
          app = WIN32OLE.new('Excel.Application')
          ole_excel = WIN32OLE.connect("Excel.Application")  
          reo_excel = Excel.new(ole_excel, {:displayalerts => true, :visible => true})
          ole_excel.Visible.should == true
          ole_excel.DisplayAlerts.should == true
        end


        it "lifts an Excel instance given as WIN32Ole object" do    
          @book = Workbook.open(@simple_file)
          @excel = @book.excel          
          win32ole_excel = WIN32OLE.connect(@book.ole_workbook.Fullname).Application
          excel = Excel.new(win32ole_excel)
          excel.should be_a Excel
          excel.should be_alive
          excel.should === @excel
        end

        it "lifts an Excel instance given as WIN32Ole object with options" do    
          @book = Workbook.open(@simple_file)
          @excel = @book.excel
          @excel.Visible = true
          @excel.DisplayAlerts = true
          win32ole_excel = WIN32OLE.connect(@book.ole_workbook.Fullname).Application
          excel = Excel.new(win32ole_excel)
          excel.should be_a Excel
          excel.should be_alive
          excel.should === @excel
          excel.Visible.should be true
          excel.DisplayAlerts.should be true
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
        sleep 0.2
        excel2 = Excel.current
        excel1.should_not be_alive
        excel2.should be_alive
        Excel.excels_number.should == 1
      end

      it "should make the Excel instance not alive if the Excel that was connected with was closed" do
        excel1 = Excel.create
        excel2 = Excel.current
        excel1.close
        sleep 0.2
        excel1.should_not be_alive
        excel2.should_not be_alive
        sleep 0.2
        Excel.excels_number.should == 0
      end

      it "should reuse the first opened Excel instance if not the first opened Excel instance was closed" do
        excel1 = Excel.create
        excel2 = Excel.create
        excel2.close
        sleep 0.2
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
        Excel.kill_all
        sleep 0.2
        n1 = Excel.excels_number
        e1 = Excel.create
        Excel.excels_number.should == n1 + 1
        e2 = Excel.create
        Excel.excels_number.should == n1 + 2
      end
    end

    context "kill Excel processes hard" do

      before do
        @excel1 = Excel.create
        @excel2 = Excel.create
      end

      it "should kill Excel processes" do
        Excel.kill_all
        @excel1.alive?.should be false
        @excel2.alive?.should be false
      end
    end

    context "recreating Excel instances" do

      context "with a single Excel instance" do

        before do
          @book1 = Workbook.open(@simple_file)
          @excel1 = @book1.excel
        end

        it "should recreate an Excel instance" do
          @excel1.close
          @excel1.should_not be_alive
          @excel1.recreate
          @excel1.should be_a Excel
          @excel1.should be_alive
          @excel1.Visible.should be false
          @excel1.DisplayAlerts.should be false
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
          @excel1.Visible.should be true
          @excel1.DisplayAlerts.should be true
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
          @excel1.Visible.should be true
          @excel1.DisplayAlerts.should be true
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
          @excel1.Visible.should be false
          @excel1.DisplayAlerts.should be false
          @book1.should be_alive
          @excel1.close
          @excel1.should_not be_alive
        end
      end

      context "with several Excel instances" do

        before do
          @book1 = Workbook.open(@simple_file)      
          @book2 = Workbook.open(@another_simple_file, :force_excel => @book1)
          @book3 = Workbook.open(@different_file, :force_excel => :new)
          @excel1 = @book1.excel
          @excel3 = @book3.excel
          @excel1.visible = true
          @excel3.displayalerts = true
        end

        it "should recreate several Excel instances" do  
          @excel1.close(:if_unsaved => :forget)
          @excel3.close
          sleep 0.2
          @excel1.should_not be_alive
          @excel3.should_not be_alive
          @excel1.recreate(:reopen_workbooks => true, :displayalerts => true)
          @excel1.should be_alive
          @excel1.should be_a Excel
          @excel1.Visible.should be true
          @excel1.DisplayAlerts.should be true
          @book1.should be_alive
          @book2.should be_alive
          @book1.visible == true
          @book2.visible == true
          @excel3.recreate(:visible => true)
          @excel3.should be_alive
          @excel3.should be_a Excel
          @excel3.Visible.should be true
          @excel3.DisplayAlerts.should be true
          @book3.reopen
          @book3.should be_alive
          @book3.excel.should == @excel3
          @excel1.close(:if_unsaved => :forget)
          sleep 2
          @excel1.should_not be_alive
          @excel3.close
          sleep 2
          @excel3.should_not be_alive
        end
      end    
    end

    context "close excel instances" do
      # @private
      def direct_excel_creation_helper 
        expect { WIN32OLE.connect("Excel.Application") }.to raise_error
        sleep 0.1
        ole_excel1 = WIN32OLE.new("Excel.Application")
        ole_excel1.Workbooks.Add
        ole_excel2 = WIN32OLE.new("Excel.Application")
        ole_excel2.Workbooks.Add
        expect { WIN32OLE.connect("Excel.Application") }.to_not raise_error
      end

      it "simple file with default" do
        Excel.kill_all
        direct_excel_creation_helper
        sleep 4
        Excel.kill_all
        sleep 4
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
          result = Excel.close_all
          sleep 0.2
          excel1.should_not be_alive
          result.should == [1,0]
        end

        it "should close two Excel instances" do
          excel1 = Excel.create
          excel2 = Excel.create
          result = Excel.close_all
          sleep 0.2
          excel1.should_not be_alive
          excel2.should_not be_alive
          result.should == [2,0]
        end
      end

      context "with unsaved workbooks" do

        context "with one Excel instance" do

          before do
            book1 = Workbook.open(@simple_file1, :visible => true)
            @excel1 = book1.excel
            sheet1 = book1.sheet(1)
            @old_cell_value1 = sheet1[1,1].Value
            sheet1[1,1] = sheet1[1,1].Value == "foo" ? "bar" : "foo"
            book1.Saved.should be false
          end

          it "should save the unsaved workbook" do
            result = Excel.close_all(:if_unsaved => :save)
            sleep 0.5
            @excel1.should_not be_alive
            new_book1 = Workbook.open(@simple_file1)
            new_sheet1 = new_book1.sheet(1)
            new_sheet1[1,1].Value.should_not == @old_cell_value1
            new_book1.close
            result.should == [1,0]
          end

          it "should forget the unsaved workbook" do
            result = Excel.close_all(:if_unsaved => :forget)
            sleep 0.5
            @excel1.should_not be_alive
            new_book1 = Workbook.open(@simple_file1)
            new_sheet1 = new_book1.sheet(1)
            new_sheet1[1,1].Value.should == @old_cell_value1
            new_book1.close
            result.should == [1,0]
          end
        end        

        context "with two Excel instances" do
          
          before do          
            book1 = Workbook.open(@simple_file1, :force_excel => :new)
            book2 = Workbook.open(@different_file, :force_excel => :new)          
            @excel1 = book1.excel
            @excel2 = book2.excel
            sheet2 = book2.sheet(1)
            @old_cell_value2 = sheet2[1,1].Value
            sheet2[1,1] = sheet2[1,1].Value == "foo" ? "bar" : "foo"
          end

          it "should close the first Excel without unsaved workbooks and then raise an error" do
            expect{
              Excel.close_all(:if_unsaved => :raise)
            }.to raise_error(UnsavedWorkbooks, "Excel contains unsaved workbooks" +
              "\nHint: Use option :if_unsaved with values :forget and :save to close the 
           Excel instance without or with saving the unsaved workbooks before, respectively")
            sleep 0.2
            @excel1.should_not be_alive
            @excel2.should be_alive
            result = Excel.close_all(:if_unsaved => :forget)
            sleep 0.2
            @excel2.should_not be_alive
            result.should == [1,0]
          end

          it "should close the first Excel without unsaved workbooks and then raise an error" do
            expect{
              Excel.close_all
            }.to raise_error(UnsavedWorkbooks, "Excel contains unsaved workbooks" +
              "\nHint: Use option :if_unsaved with values :forget and :save to close the 
           Excel instance without or with saving the unsaved workbooks before, respectively")
            sleep 0.2
            @excel1.should_not be_alive
            @excel2.should be_alive
            result = Excel.close_all(:if_unsaved => :forget)
            sleep 0.2
            @excel2.should_not be_alive
            result.should == [1,0]
          end

          it "should close the Excel instances with saving the unsaved workbooks" do
            result = Excel.close_all(:if_unsaved => :save)
            sleep 0.2
            @excel1.should_not be_alive
            @excel2.should_not be_alive
            new_book2 = Workbook.open(@different_file1)
            new_sheet2 = new_book2.sheet(1)
            new_sheet2[1,1].Value.should_not == @old_cell_value2
            new_book2.close
            result.should == [2,0]
          end

          it "should close the Excel instances without saving the unsaved workbooks" do          
            result = Excel.close_all(:if_unsaved => :forget)
            sleep 0.2
            @excel1.should_not be_alive
            @excel2.should_not be_alive
            new_book2 = Workbook.open(@different_file1)
            new_sheet2 = new_book2.sheet(1)
            new_sheet2[1,1].Value.should == @old_cell_value2
            new_book2.close
            result.should == [2,0]
          end       

          it "should raise an error for invalid option" do
            expect {
              Excel.close_all(:if_unsaved => :invalid_option)
            }.to raise_error(OptionInvalid, ":if_unsaved: invalid option: :invalid_option" +
              "\nHint: Valid values are :raise, :forget, :save and :alert") 
          end
        end

        context "with three Excel instances" do

         before do          
            @book1 = Workbook.open(@simple_file1, :force_excel => :new)
            @book2 = Workbook.open(@another_simple_file, :force_excel => :new) 
            @book3 = Workbook.open(@different_file, :force_excel => :new)
            old_cell_value1 = @book2.sheet(1)[1,1].Value                 
            @book2.sheet(1)[1,1] = old_cell_value1 == "foo" ? "bar" : "foo"
          end

          it "should close the 1st and 3rd Excel instances that have saved workbooks" do  
            expect{
              Excel.close_all(:if_unsaved => :raise)
            }.to raise_error(UnsavedWorkbooks, "Excel contains unsaved workbooks" +
              "\nHint: Use option :if_unsaved with values :forget and :save to close the 
           Excel instance without or with saving the unsaved workbooks before, respectively")
            sleep 0.2
            @book1.excel.should_not be_alive
            @book2.excel.should be_alive
            @book3.excel.should_not be_alive
            result = Excel.close_all(:if_unsaved => :forget)
            @book2.excel.should_not be_alive
            result.should == [1,0]
          end
        end

        context "with unknown Excel instances" do

         before do          
            @ole_xl = WIN32OLE.new('Excel.Application')
            @book1 = Workbook.open(@simple_file1, :force_excel => :new)
            @book2 = Workbook.open(@another_simple_file, :force_excel => :new) 
            @book3 = Workbook.open(@different_file, :force_excel => :new)
            old_cell_value1 = @book2.sheet(1)[1,1].Value                 
            @book2.sheet(1)[1,1] = old_cell_value1 == "foo" ? "bar" : "foo"
          end

          it "should close three Excel instances that have saved workbooks" do  
            expect{
              Excel.close_all(:if_unsaved => :raise)
            }.to raise_error(UnsavedWorkbooks, "Excel contains unsaved workbooks" +
              "\nHint: Use option :if_unsaved with values :forget and :save to close the 
           Excel instance without or with saving the unsaved workbooks before, respectively")
            sleep 0.2
            expect{
              @ole_xl.Name
            }.to raise_error #(WIN32OLERuntimeError)
            @book1.excel.should_not be_alive
            @book2.excel.should be_alive
            @book3.excel.should_not be_alive
            result = Excel.close_all(:if_unsaved => :forget)
            @book2.excel.should_not be_alive
            result.should == [1,0]
          end

          it "should close all four Excel instances" do  
            result = Excel.close_all(:if_unsaved => :forget)
            sleep 0.2
            expect{
              @ole_xl.Name
            }.to raise_error(RuntimeError, /failed to get/)
            @book1.excel.should_not be_alive
            @book2.excel.should_not be_alive
            @book3.excel.should_not be_alive
            result.should == [4,0]
          end
        end

      end
    end

    describe "close" do

      context "with saved workbooks" do

        before do
          @excel = Excel.create
          @book = Workbook.open(@simple_file)
          @excel.should be_alive
        end

        it "should close the Excel" do
          @book.should be_alive
          @excel.close
          sleep 0.2
          @excel.should_not be_alive
          @book.should_not be_alive
        end

        it "should close the Excel without destroying the others" do
          excel2 = Excel.create
          @excel.close
          sleep 0.2
          @excel.should_not be_alive
          excel2.should be_alive
        end
      end

      context "with unsaved workbooks" do

        before do
          @excel = Excel.create
          @book = Workbook.open(@simple_file)
          sheet = @book.sheet(1)
          @old_cell_value = sheet[1,1].Value
          sheet[1,1] = sheet[1,1].Value == "foo" ? "bar" : "foo"
          @book2 = Workbook.open(@another_simple_file)
          sheet2 = @book2.sheet(1)
          @old_cell_value2 = sheet2[1,1].Value
          sheet2[1,1] = sheet2[1,1].Value == "foo" ? "bar" : "foo"
          @excel.should be_alive
          @book.should be_alive
          @book.saved.should be false
          @book2.should be_alive
          @book2.saved.should be false
        end

        it "should raise an error" do
          expect{
            @excel.close(:if_unsaved => :raise)
          }.to raise_error(UnsavedWorkbooks, "Excel contains unsaved workbooks" +
            "\nHint: Use option :if_unsaved with values :forget and :save to close the 
           Excel instance without or with saving the unsaved workbooks before, respectively")
        end        

        it "should raise an error per default" do
          expect{
            @excel.close(:if_unsaved => :raise)
          }.to raise_error(UnsavedWorkbooks, "Excel contains unsaved workbooks" +
            "\nHint: Use option :if_unsaved with values :forget and :save to close the 
           Excel instance without or with saving the unsaved workbooks before, respectively")
        end        

        it "should close the Excel without saving the workbook" do
          result = @excel.close(:if_unsaved => :forget)
          sleep 0.2
          @excel.should_not be_alive
          result.should == 1
          new_book = Workbook.open(@simple_file)
          new_sheet = new_book.sheet(1)
          new_sheet[1,1].Value.should == @old_cell_value
          new_book.close          
        end
        
        it "should close the Excel without saving the workbook even with displayalerts true" do
          @excel.displayalerts = false
          @excel.should be_alive
          @excel.displayalerts = true
          result = @excel.close(:if_unsaved => :forget)
          sleep 0.2
          result.should == 1
          @excel.should_not be_alive
          new_book = Workbook.open(@simple_file)
          new_sheet = new_book.sheet(1)
          new_sheet[1,1].Value.should == @old_cell_value
          new_book.close          
        end

        it "should close the Excel with saving the workbook" do
          @excel.should be_alive
          result = @excel.close(:if_unsaved => :save)
          sleep 0.2
          result.should == 1
          @excel.should_not be_alive
          new_book = Workbook.open(@simple_file)
          new_sheet = new_book.sheet(1)
          new_sheet[1,1].Value.should_not == @old_cell_value
          new_book.close          
        end

        it "should raise an error for invalid option" do
          expect {
            @excel.close(:if_unsaved => :invalid_option)
          }.to raise_error(OptionInvalid, ":if_unsaved: invalid option: :invalid_option" +
            "\nHint: Valid values are :raise, :forget, :save and :alert") 
        end
      end

      context "with :if_unsaved => :alert" do

        before do
          @key_sender = IO.popen  'ruby "' + File.join(File.dirname(__FILE__), '/helpers/key_sender.rb') + '" "Microsoft Excel" '  , "w"
          @excel = Excel.create(:visible => true)
          @book = Workbook.open(@simple_file, :visible => true)
          sheet = @book.sheet(1)
          @old_cell_value = sheet[1,1].Value
          sheet[1,1] = sheet[1,1].Value == "foo" ? "bar" : "foo"
        end

        after do
          @key_sender.close
        end

        it "should save if user answers 'yes'" do
          # "Yes" is to the left of "No", which is the  default. --> language independent
          @excel.should be_alive
          @key_sender.puts "{enter}" 
          @key_sender.puts "{enter}" 
          @key_sender.puts "{enter}" 
          result = @excel.close(:if_unsaved => :alert)
          @excel.should_not be_alive
          result.should == 1
          new_book = Workbook.open(@simple_file)
          new_sheet = new_book.sheet(1)
          new_sheet[1,1].Value.should_not == @old_cell_value
          new_book.close   
        end

        it "should not save if user answers 'no'" do            
          @excel.should be_alive
          @book.should be_alive
          @book.saved.should be false
          @key_sender.puts "{right}{enter}"
          result = @excel.close(:if_unsaved => :alert)
          @excel.should_not be_alive
          result.should == 1
          @book.should_not be_alive
          new_book = Workbook.open(@simple_file)
          new_sheet = new_book.sheet(1)
          new_sheet[1,1].Value.should == @old_cell_value
          new_book.close     
        end

        it "should not save if user answers 'cancel'" do
          # strangely, in the "cancel" case, the question will sometimes be repeated twice            
          @excel.should be_alive
          @book.should be_alive
          @book.saved.should be false
          @key_sender.puts "{left}{enter}"
          @key_sender.puts "{left}{enter}"
          @key_sender.puts "{left}{enter}"
          expect{
            @excel.close(:if_unsaved => :alert)
            }.to raise_error(ExcelREOError, "user canceled or runtime error")
        end
      end
    end

    describe "close_workbooks" do

      context "with standard" do
        
        before do
          @book = Workbook.open(@simple_file)
          sheet = @book.sheet(1)
          @old_cell_value = sheet[1,1].Value
          sheet[1,1] = sheet[1,1].Value == "foo" ? "bar" : "foo"          
          @book3 = Workbook.open(@different_file, :read_only => true)
          sheet3 = @book3.sheet(1)
          sheet3[1,1] = sheet3[1,1].Value == "foo" ? "bar" : "foo"
          @excel = @book.excel
          @book2 = Workbook.open(@another_simple_file, :force_excel => :new)
        end

        it "should be ok if there are no unsaved workbooks" do
          expect{
            @book2.excel.close_workbooks
          }.to_not raise_error
        end

        it "should raise error" do
          expect{
            @excel.close_workbooks(:if_unsaved => :raise)
          }.to raise_error(UnsavedWorkbooks, "Excel contains unsaved workbooks" +
          "\nHint: Use option :if_unsaved with values :forget and :save to close the 
           Excel instance without or with saving the unsaved workbooks before, respectively" )

        end

        it "should raise error per default" do
          expect{
            @excel.close_workbooks
          }.to raise_error(UnsavedWorkbooks, "Excel contains unsaved workbooks" +
            "\nHint: Use option :if_unsaved with values :forget and :save to close the 
           Excel instance without or with saving the unsaved workbooks before, respectively")
        end

        it "should close the workbook with forgetting the workbook" do
          @excel.close_workbooks(:if_unsaved => :forget)
          sleep 0.2
          @excel.should be_alive
          @excel.Workbooks.Count.should == 0          
          new_book = Workbook.open(@simple_file)
          new_sheet = new_book.sheet(1)
          new_sheet[1,1].Value.should == @old_cell_value
          new_book.close          
        end

        it "should close the workbook with saving the workbook" do
          @excel.close_workbooks(:if_unsaved => :save)
          sleep 0.2
          @excel.should be_alive
          @excel.Workbooks.Count.should == 0          
          new_book = Workbook.open(@simple_file)
          new_sheet = new_book.sheet(1)
          new_sheet[1,1].Value.should_not == @old_cell_value
          new_book.close          
        end

        it "should raise an error for invalid option" do
          expect {
            @excel.close_workbooks(:if_unsaved => :invalid_option)
          }.to raise_error(OptionInvalid, ":if_unsaved: invalid option: :invalid_option" +
            "\nHint: Valid values are :raise, :forget, :save and :alert") 
        end
      end
    end

=begin
    describe "retain_saved_workbooks" do

      before do
        @book1 = Workbook.open(@simple_file)
        @book2 = Workbook.open(@another_simple_file)
        @book3 = Workbook.open(@different_file)
        sheet2 = @book2.sheet(1)
        sheet2[1,1] = sheet2[1,1].Value == "foo" ? "bar" : "foo"
        @book2.Saved.should be false
        @excel = Excel.current
      end

      it "should retain saved workbooks" do
        @excel.retain_saved_workbooks do
          sheet1 = @book1.sheet(1)
          sheet1[1,1] = sheet1[1,1].Value == "foo" ? "bar" : "foo"
          @book1.Saved.should be false
          sheet3 = @book3.sheet(1)
          sheet3[1,1] = sheet3[1,1].Value == "foo" ? "bar" : "foo"
          @book3.Saved.should be false
        end
        @book1.Saved.should be true
        @book2.Saved.should be false
        @book3.Saved.should be true
      end
    end

=end

    describe "unsaved_workbooks" do

      context "with standard" do
        
        before do
          @book = Workbook.open(@simple_file)
          sheet = @book.sheet(1)
          sheet[1,1] = sheet[1,1].Value == "foo" ? "bar" : "foo"          
          @book3 = Workbook.open(@different_file, :read_only => true)
          sheet3 = @book3.sheet(1)
          sheet3[1,1] = sheet3[1,1].Value == "foo" ? "bar" : "foo"
          @book.Saved.should be false
          @book3.Saved.should be false
        end

        it "should list unsaved workbooks" do          
          excel = @book.excel
          # unsaved_workbooks yields different WIN32OLE objects than book.workbook
          uw_names = []
          excel.unsaved_workbooks.each {|uw| uw_names << uw.Name}
          uw_names.should == [@book.ole_workbook.Name]
        end

        it "should yield true, that there are unsaved workbooks" do
          Excel.contains_unsaved_workbooks?.should be true
        end
      end
    end

    describe "workbooks, each, each_with_index" do

      before do
        @excel = Excel.create
        @book1 = Workbook.open(@simple_file)
        @book2 = Workbook.open(@different_file)
      end

      it "should list workbooks" do
        workbooks = @excel.workbooks
        workbooks.should == [@book1,@book2]
      end

      it "should each_workbook" do
        i = 0
        @excel.each_workbook do |workbook|
          workbook.should be_alive
          workbook.should be_a Workbook
          workbook.filename.should == @simple_file if i == 0
          workbook.filename.should == @different_file if i == 1
          i += 1
        end
      end

      it "should each_workbook_with_index" do
        @excel.each_workbook_with_index do |workbook,i|
          workbook.should be_alive
          workbook.should be_a Workbook
          workbook.filename.should == @simple_file if i == 0
          workbook.filename.should == @different_file if i == 1
        end
      end

      it "should each_workbook with options" do
        i = 0
        @excel.each_workbook(:visible => true) do |workbook|
          workbook.should be_alive
          workbook.should be_a Workbook
          workbook.visible.should be true
          workbook.filename.should == @simple_file if i == 0
          workbook.filename.should == @different_file if i == 1
          i += 1
        end
      end

      it "should set options" do
        @excel.each_workbook(:visible => true)
        [1,2].each do |i|
          ole_workbook = @excel.Workbooks.Item(i) 
          ole_workbook.Windows(ole_workbook.Name).Visible.should be true
        end
      end


    end

    describe "unsaved_known_workbooks" do

      it "should return empty list" do
        Excel.unsaved_known_workbooks.should be_empty
      end

      it "should return empty list for first Excel instance" do
        book = Workbook.open(@simple_file)
        Excel.unsaved_known_workbooks.should == [[]]
        book.close
      end
       
      it "should return one unsaved book" do
        book = Workbook.open(@simple_file)
        sheet = book.sheet(1)
        sheet[1,1] = sheet[1,1].Value == "foo" ? "bar" : "foo"
        # Excel.unsaved_known_workbooks.should == [[book.ole_workbook]]
        unsaved_known_wbs = Excel.unsaved_known_workbooks
        unsaved_known_wbs.size.should == 1
        unsaved_known_wbs.each do |ole_wb_list|
          ole_wb_list.size.should == 1
          ole_wb_list.each do |ole_workbook|
            ole_workbook.Fullname.tr('\\','/').should == @simple_file  
          end
        end
        book2 = Workbook.open(@another_simple_file)
        # Excel.unsaved_known_workbooks.should == [[book.ole_workbook]]
        unsaved_known_wbs = Excel.unsaved_known_workbooks
        unsaved_known_wbs.size.should == 1
        unsaved_known_wbs.each do |ole_wb_list|
          ole_wb_list.size.should == 1
          ole_wb_list.each do |ole_workbook|
            ole_workbook.Fullname.tr('\\','/').should == @simple_file  
          end
        end        
        book2.close
        book.close(:if_unsaved => :forget)
      end

      it "should return two unsaved books" do
        book = Workbook.open(@simple_file)
        sheet = book.sheet(1)
        sheet[1,1] = sheet[1,1].Value == "foo" ? "bar" : "foo"
        book2 = Workbook.open(@another_simple_file)
        sheet2 = book2.sheet(1)
        sheet2[1,1] = sheet2[1,1].Value == "foo" ? "bar" : "foo"
        #Excel.unsaved_known_workbooks.should == [[book.ole_workbook, book2.ole_workbook]]
        unsaved_known_wbs = Excel.unsaved_known_workbooks
        unsaved_known_wbs.size.should == 1
        unsaved_known_wbs.each do |ole_wb_list|
          ole_wb_list.size.should == 2
          ole_workbook1, ole_workbook2 = ole_wb_list
          ole_workbook1.Fullname.tr('\\','/').should == @simple_file  
          ole_workbook2.Fullname.tr('\\','/').should == @another_simple_file  
        end        
        book2.close(:if_unsaved => :forget)
        book.close(:if_unsaved => :forget)
      end

      it "should return two unsaved books" do
        book = Workbook.open(@simple_file)
        book2 = Workbook.open(@another_simple_file, :force_excel => :new)
        open_books = [book, book2]
        begin 
          open_books.each do |wb|
            sheet = wb.sheet(1)
            sheet[1,1] = (sheet[1,1].Value == "foo") ? "bar" : "foo"
          end 
        #Excel.unsaved_known_workbooks.should == [[book.ole_workbook], [book2.ole_workbook]]
          unsaved_known_wbs = Excel.unsaved_known_workbooks
          unsaved_known_wbs.size.should == 2
          unsaved_known_wbs.flatten.map{|ole_wb| ole_wb.Fullname.tr('\\','/') }.sort.should == open_books.map{|b| b.filename}.sort
        ensure
          open_books.each {|wb| wb.close(:if_unsaved => :forget)}
        end
      end

    end

    describe "alive" do

      it "should yield alive" do
        excel = Excel.create
        excel.alive?.should be true
      end

      it "should yield not alive" do
        excel = Excel.create
        excel.close
        excel.alive?.should be false
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
        sleep 3
        Excel.close_all
        sleep 2
        excel2.should_not == @excel1
      end

    end

    describe "focus" do

      it "should focus" do
        excel = Excel.create
        excel.focus
        excel.Visible.should be true
      end

    end

    context "with Visible and DisplayAlerts, focus" do

      it "should bring Excel in focus" do
        excel1 = Excel.create
        excel2 = Excel.create
        excel1.focus
        excel1.Visible.should be true
        excel1.properties[:visible].should be true
      end

      it "should set default values" do
        excel1 = Excel.new
        excel1.Visible.should be false
        excel1.DisplayAlerts.should be false
        excel1.properties[:visible].should be false
        excel1.properties[:displayalerts].should == :if_visible
      end

      it "should set visible true" do
        excel1 = Excel.new(:visible => true)
        excel1.Visible.should be true
        excel1.DisplayAlerts.should be true
        excel1.properties[:visible].should be true
        excel1.properties[:displayalerts].should == :if_visible
      end

      it "should set visible false" do
        excel1 = Excel.new(:visible => false)
        excel1.Visible.should be false
        excel1.DisplayAlerts.should be false
        excel1.properties[:visible].should be false
        excel1.properties[:displayalerts].should == :if_visible
      end

      it "should set displayalerts true" do
        excel1 = Excel.new(:displayalerts => true)
        excel1.Visible.should be false
        excel1.DisplayAlerts.should be true
        excel1.properties[:visible].should be false
        excel1.properties[:displayalerts].should be true
      end

      it "should set displayalerts false" do
        excel1 = Excel.new(:displayalerts => false)
        excel1.Visible.should be false
        excel1.DisplayAlerts.should be false
        excel1.properties[:visible].should be false
        excel1.properties[:displayalerts].should be false
      end

      it "should use values of the current Excel when reusing" do
        excel1 = Excel.create
        excel1.Visible.should be false
        excel1.DisplayAlerts.should be false
        excel1.properties[:visible].should be false
        excel1.properties[:displayalerts].should == :if_visible
        excel1.Visible = true
        excel1.DisplayAlerts = true
        excel1.Visible.should be true
        excel1.DisplayAlerts.should be true
        excel2 = Excel.new(:reuse => true)
        excel2.Visible.should be true
        excel2.DisplayAlerts.should be true
      end

      it "should take visible and displayalerts from Visible and DisplayAlerts of the connected Excel" do
        excel1 = Excel.create
        excel2 = Excel.current
        excel2.Visible.should be false
        excel2.properties[:visible].should be false
        excel2.DisplayAlerts.should be false
        excel2.properties[:displayalerts].should == :if_visible
      end

      it "should take Visible and DisplayAlerts from the connected Excel" do
        excel1 = Excel.create
        excel2 = Excel.current(:visible => true)
        excel2.Visible.should be true
        excel2.properties[:visible].should be true
        excel2.DisplayAlerts.should be true
        excel2.properties[:displayalerts].should == :if_visible
      end

      it "should set Excel visible and invisible with current" do
        excel1 = Excel.new(:reuse => false, :visible => true)
        excel1.Visible.should be true
        excel1.properties[:visible].should be true
        excel1.DisplayAlerts.should be true
        excel1.properties[:displayalerts].should == :if_visible
        excel1.visible = false
        excel1.Visible.should be false
        excel1.properties[:visible].should be false
        excel1.DisplayAlerts.should be false
        excel1.properties[:displayalerts].should == :if_visible
        excel2 = Excel.current(:visible => true)
        excel2.Visible.should be true
        excel2.properties[:visible].should be true
        excel2.properties[:displayalerts].should == :if_visible
        excel2.DisplayAlerts.should be true
      end

      it "should set Excel visible and invisible" do
        excel = Excel.new(:reuse => false, :visible => true)
        excel.Visible.should be true
        excel.properties[:visible].should be true
        excel.DisplayAlerts.should be true
        excel.properties[:displayalerts].should == :if_visible
        excel.visible = false
        excel.Visible.should be false
        excel.properties[:visible].should be false
        excel.DisplayAlerts.should be false
        excel.properties[:displayalerts].should == :if_visible
        excel7 = Excel.current
        excel7.should === excel
        excel7.Visible.should be false
        excel7.DisplayAlerts.should be false
        excel1 = Excel.create(:visible => true)
        excel1.should_not == excel
        excel1.Visible.should be true
        excel1.properties[:visible].should be true
        excel1.DisplayAlerts.should be true
        excel1.properties[:displayalerts].should == :if_visible
        excel2 = Excel.create(:visible => false)
        excel2.Visible.should be false
        excel2.properties[:visible].should be false
        excel2.DisplayAlerts.should be false
        excel2.properties[:displayalerts].should == :if_visible
        excel3 = Excel.current
        excel3.should === excel
        excel3.Visible.should be false
        excel3.properties[:visible].should be false
        excel3.DisplayAlerts.should be false
        excel3.properties[:displayalerts].should == :if_visible
        excel4 = Excel.current(:visible => true)
        excel4.should === excel
        excel4.Visible.should be true
        excel4.properties[:visible].should be true
        excel4.DisplayAlerts.should be true
        excel4.properties[:displayalerts].should == :if_visible
        excel5 = Excel.current(:visible => false)
        excel5.should === excel
        excel5.Visible.should be false
        excel5.properties[:visible].should be false
        excel5.DisplayAlerts.should be false
        excel5.properties[:displayalerts].should == :if_visible
      end

      it "should enable or disable Excel DispayAlerts" do        
        excel = Excel.new(:reuse => false, :displayalerts => true)
        excel.DisplayAlerts.should be true
        excel.properties[:displayalerts].should be true
        excel.Visible.should be false
        excel.properties[:visible].should be false
        excel6 = Excel.current
        excel6.should === excel
        excel6.DisplayAlerts.should be true
        excel6.properties[:displayalerts].should be true
        excel6.Visible.should be false
        excel6.properties[:visible].should be false
        excel.displayalerts = false
        excel.DisplayAlerts.should be false
        excel.properties[:displayalerts].should be false
        excel.Visible.should be false
        excel.properties[:visible].should be false
        excel7 = Excel.current
        excel7.should === excel
        excel7.DisplayAlerts.should be false
        excel7.properties[:displayalerts].should be false
        excel7.Visible.should be false
        excel7.properties[:visible].should be false
        excel1 = Excel.create(:displayalerts => true)
        excel1.should_not == excel
        excel1.DisplayAlerts.should be true
        excel1.properties[:displayalerts].should be true
        excel1.Visible.should be false
        excel1.properties[:visible].should be false
        excel2 = Excel.create(:displayalerts => false)
        excel2.DisplayAlerts.should be false
        excel2.properties[:displayalerts].should be false
        excel2.Visible.should be false
        excel2.properties[:visible].should be false
        excel3 = Excel.current
        excel3.should === excel
        excel3.DisplayAlerts.should be false
        excel3.properties[:displayalerts].should be false
        excel3.Visible.should be false
        excel3.properties[:visible].should be false
        excel4 = Excel.current(:displayalerts => true)
        excel4.should === excel
        excel4.DisplayAlerts.should be true
        excel4.properties[:displayalerts].should be true
        excel4.Visible.should be false
        excel4.properties[:visible].should be false
        excel5 = Excel.current(:displayalerts => false)
        excel5.should === excel
        excel5.DisplayAlerts.should be false
        excel5.properties[:displayalerts].should be false
        excel5.Visible.should be false
        excel5.properties[:visible].should be false
      end

      it "should set Excel visible and displayalerts" do        
        excel = Excel.new(:reuse => false, :visible => true, :displayalerts => true)
        excel.DisplayAlerts.should be true
        excel.properties[:displayalerts].should be true
        excel.Visible.should be true
        excel.properties[:visible].should be true
        excel6 = Excel.current
        excel6.should === excel
        excel6.DisplayAlerts.should be true
        excel6.properties[:displayalerts].should be true
        excel6.Visible.should be true
        excel6.properties[:visible].should be true
        excel.displayalerts = false
        excel.DisplayAlerts.should be false
        excel.properties[:displayalerts].should be false
        excel.Visible.should be true
        excel.properties[:visible].should be true
        excel7 = Excel.current
        excel7.should === excel
        excel7.DisplayAlerts.should be false
        excel7.properties[:displayalerts].should be false
        excel7.Visible.should be true
        excel7.properties[:visible].should be true        
        excel2 = Excel.new(:reuse => false, :visible => true, :displayalerts => true)
        excel2.visible = false
        excel2.DisplayAlerts.should be true
        excel2.properties[:displayalerts].should be true
        excel2.Visible.should be false
        excel2.properties[:visible].should be false
        excel3 = Excel.new(:reuse => false, :visible => true, :displayalerts => false)
        excel3.Visible.should be true
        excel3.DisplayAlerts.should be false
        excel3 = Excel.new(:reuse => false, :visible => false, :displayalerts => true)
        excel3.Visible.should be false
        excel3.DisplayAlerts.should be true
        excel3 = Excel.new(:reuse => false, :visible => false, :displayalerts => false)
        excel3.Visible.should be false
        excel3.DisplayAlerts.should be false
        excel4 = Excel.create(:visible => true, :displayalerts => true)
        excel4.DisplayAlerts.should be true
        excel4.properties[:displayalerts].should be true
        excel4.Visible.should be true
        excel4.properties[:visible].should be true
        excel5 = Excel.current(:visible => true, :displayalerts => false)
        excel5.should === excel
        excel5.DisplayAlerts.should be false
        excel5.properties[:displayalerts].should be false
        excel5.Visible.should be true
        excel5.properties[:visible].should be true
        excel6 = Excel.current(:visible => false, :displayalerts => true)
        excel6.should === excel
        excel6.DisplayAlerts.should be true
        excel6.properties[:displayalerts].should be true
        excel6.Visible.should be false
        excel6.properties[:visible].should be false
      end

      it "should work with displayalerts == if_visible" do
        excel = Excel.new(:reuse => false, :visible => true, :displayalerts => :if_visible)
        excel.Visible.should be true
        excel.DisplayAlerts.should be true
        excel2 = Excel.new(:reuse => false, :visible => false, :displayalerts => :if_visible)
        excel2.Visible.should be false
        excel2.DisplayAlerts.should be false
        excel3 = Excel.new(:reuse => false, :displayalerts => :if_visible)
        excel3.Visible.should be false
        excel3.DisplayAlerts.should be false
        excel3.visible = true
        excel3.Visible.should be true
        excel3.DisplayAlerts.should be true
        excel3.visible = false
        excel3.Visible.should be false
        excel3.DisplayAlerts.should be false
      end

      it "should keep visible and displayalerts values when reusing Excel" do
        excel = Excel.new(:visible => true)
        excel.Visible.should be true
        excel.DisplayAlerts.should be true
        excel2 = Excel.new(:displayalerts => false)
        excel2.should == excel
        excel.Visible.should be true
        excel.DisplayAlerts.should be false        
      end

      it "should keep displayalerts and visible values when reusing Excel" do
        excel = Excel.new(:displayalerts => true)
        excel.Visible.should be false
        excel.DisplayAlerts.should be true
        excel2 = Excel.new(:visible => true)
        excel2.should == excel
        excel.Visible.should be true
        excel.DisplayAlerts.should be true        
      end

    end

    context "with resetting displayalerts values" do
      before do
        @excel1 = Excel.new(:displayalerts => true)
        @excel2 = Excel.new(:displayalerts => false, :reuse => false)
        @excel3 = Excel.new(:displayalerts => false, :visible => true, :reuse => false)
      end

      it "should turn off displayalerts" do
        @excel1.DisplayAlerts.should be true
        begin
          @excel1.with_displayalerts false do
            @excel1.DisplayAlerts.should be false
            raise TestError, "any_error"
          end
        rescue TestError
          @excel1.DisplayAlerts.should be true
        end
      end
    
      it "should turn on displayalerts" do
        @excel2.DisplayAlerts.should be false
        begin
          @excel1.with_displayalerts true do
            @excel1.DisplayAlerts.should be true
            raise TestError, "any_error"
          end
        rescue TestError
          @excel2.DisplayAlerts.should be false
        end
      end

      it "should set displayalerts to :if_visible" do
        @excel1.DisplayAlerts.should be true
        begin
          @excel1.with_displayalerts :if_visible do
            @excel1.DisplayAlerts.should be false
            @excel1.Visible.should be false
            raise TestError, "any_error"
          end
        rescue TestError
          @excel1.DisplayAlerts.should be true
        end
      end

      it "should set displayalerts to :if_visible" do
        @excel3.DisplayAlerts.should be false
        begin
          @excel3.with_displayalerts :if_visible do
            @excel3.DisplayAlerts.should be true
            @excel3.Visible.should be true
            raise TestError, "any_error"
          end
        rescue TestError
          @excel3.DisplayAlerts.should be false
        end
      end

    end

    context "with screen updating" do

      it "should set screen updating" do
        excel1 = Excel.new
        excel1.ScreenUpdating.should be true
        excel2 = Excel.create(:screenupdating => false)
        excel2.ScreenUpdating.should be false
        excel3 = Excel.new
        excel3.ScreenUpdating.should be true
        excel4 = Excel.new(:screenupdating => false)
        excel4.ScreenUpdating.should be false
      end

    end

    context "with calculation" do

      it "should create and reuse Excel with calculation mode" do
        excel1 = Excel.create(:calculation => :manual)
        excel1.properties[:calculation].should == :manual
        excel2 = Excel.create(:calculation => :automatic)
        excel2.properties[:calculation].should == :automatic
        excel3 = Excel.current
        excel3.properties[:calculation].should == :manual
        excel4 = Excel.current(:calculation => :automatic)
        excel4.properties[:calculation].should == :automatic
        excel5 = Excel.new(:reuse => false)
        excel5.properties[:calculation].should == nil
        excel6 = Excel.new(:reuse => false, :calculation => :manual)
        excel6.properties[:calculation].should == :manual
      end

=begin
      it "should do with_calculation mode without workbooks" do
        @excel1 = Excel.new
        old_calculation_mode = @excel1.Calculation
        old_calculatebeforesave = @excel1.CalculateBeforeSave
        @excel1.with_calculation(:automatic) do
          @excel1.Calculation.should == old_calculation_mode
          @excel1.CalculateBeforeSave.should == old_calculatebeforesave
        end
        @excel1.with_calculation(:manual) do
          @excel1.Calculation.should == old_calculation_mode
          @excel1.CalculateBeforeSave.should == old_calculatebeforesave
        end
      end
=end

      it "should set calculation mode without workbooks" do
        @excel1 = Excel.new
        old_calculation_mode = @excel1.Calculation
        old_calculatebeforesave = @excel1.CalculateBeforeSave
        @excel1.calculation = :automatic
        @excel1.properties[:calculation].should == :automatic
        @excel1.Calculation.should == old_calculation_mode 
        @excel1.CalculateBeforeSave.should == old_calculatebeforesave
        @excel1.calculation = :manual
        @excel1.properties[:calculation].should == :manual
        @excel1.Calculation.should == old_calculation_mode
        @excel1.CalculateBeforeSave.should == old_calculatebeforesave
      end

      context "with no visible workbook" do
        
        it "should set calculation mode (change from automatic to manual)" do
          excel1 = Excel.create(:calculation => :automatic)
          book1 = Workbook.open(@simple_file1, :visible => false)
          expect( book1.Windows(1).Visible ).to be true # false
          expect {       excel1.calculation = :manual 
            }.to change{ excel1.properties[:calculation] 
          }.from(        :automatic
          ).to(          :manual )
        end

        it "should set calculation mode (change from manual to automatic)" do
          excel1 = Excel.create(:calculation => :manual)
          book1 = Workbook.open(@simple_file1, :visible => false)
          expect( book1.Windows(1).Visible ).to be true # false
          expect {       excel1.calculation = :automatic 
            }.to change{ excel1.properties[:calculation] 
          }.from(        :manual
          ).to(          :automatic )
        end
      end

      it "should do with_calculation with workbook" do
        @excel1 = Excel.new
        book = Workbook.open(@simple_file, :visible => true)
        old_calculation_mode = @excel1.Calculation
        @excel1.with_calculation(:manual) do
          @excel1.properties[:calculation].should == :manual
          @excel1.Calculation.should == XlCalculationManual
          @excel1.CalculateBeforeSave.should be false
          book.Saved.should be true
        end
        @excel1.Calculation.should == old_calculation_mode
        @excel1.CalculateBeforeSave.should be false
        @excel1.with_calculation(:automatic) do
          @excel1.properties[:calculation].should == :automatic
          @excel1.Calculation.should == XlCalculationAutomatic
          @excel1.CalculateBeforeSave.should be false
          book.Saved.should be false
        end
        @excel1.Calculation.should == old_calculation_mode
        @excel1.CalculateBeforeSave.should be false
      end

      it "should set calculation mode to manual with workbook" do
        @excel1 = Excel.new
        book = Workbook.open(@simple_file, :visible => true)
        book.Saved.should be true
        book.Windows(book.Name).Visible = true
        @excel1.calculation = :manual
        @excel1.properties[:calculation].should == :manual
        @excel1.Calculation.should == XlCalculationManual
        @excel1.CalculateBeforeSave.should be false
        book.Saved.should be true
      end

      it "should set calculation mode to automatic with workbook" do
        @excel1 = Excel.new
        book = Workbook.open(@simple_file, :visible => true)
        book.Saved.should be true
        @excel1.calculation = :automatic
        @excel1.properties[:calculation].should == :automatic
        @excel1.Calculation.should == XlCalculationAutomatic
        @excel1.CalculateBeforeSave.should be false
        book.Saved.should be true
      end

      it "should set calculation mode to manual with unsaved workbook" do
        @excel1 = Excel.new
        book = Workbook.open(@simple_file, :visible => true)
        book.sheet(1)[1,1] = "foo"
        book.Saved.should be false
        book.Windows(book.Name).Visible = true
        @excel1.calculation = :manual
        @excel1.properties[:calculation].should == :manual
        @excel1.Calculation.should == XlCalculationManual
        @excel1.CalculateBeforeSave.should be false
        book.Saved.should be false
      end

      it "should set calculation mode to automatic with unsaved workbook" do
        @excel1 = Excel.new
        book = Workbook.open(@simple_file, :visible => true)
        book.sheet(1)[1,1] = "foo"
        book.Saved.should be false
        @excel1.calculation = :automatic
        @excel1.properties[:calculation].should == :automatic
        @excel1.Calculation.should == XlCalculationAutomatic
        @excel1.CalculateBeforeSave.should be false
        book.Saved.should be false
      end

      it "should set Calculation without workbooks" do
        @excel1 = Excel.new
        expect{
          @excel1.Calculation = XlCalculationManual
        }.to raise_error # (WIN32OLERuntimeError)
      end

      it "should do Calculation to manual with workbook" do
        @excel1 = Excel.new
        b = Workbook.open(@simple_file, :visible => true)
        @excel1.Calculation = XlCalculationManual
        @excel1.properties[:calculation].should == :manual
        @excel1.Calculation.should == XlCalculationManual
      end

      it "should do Calculation to automatic with workbook" do
        @excel1 = Excel.new
        b = Workbook.open(@simple_file, :visible => true)
        @excel1.Calculation = XlCalculationAutomatic
        @excel1.properties[:calculation].should == :automatic
        @excel1.Calculation.should == XlCalculationAutomatic
      end

    end

    context "method delegation for capitalized methods" do
      before do
        @excel1 = Excel.new
      end

      it "should raise WIN32OLERuntimeError" do
        expect{ @excel1.NonexistingMethod }.to raise_error
        #(VBAMethodMissingError, /unknown VBA property or method :NonexistingMethod/)
      end

      it "should raise NoMethodError for uncapitalized methods" do
        expect{ @excel1.nonexisting_method }.to raise_error(NoMethodError)
      end

      it "should report that Excel is not alive" do
        @excel1.close
        expect{ @excel1.Nonexisting_method }.to raise_error(ObjectNotAlive, "method missing: Excel not alive")
      end

    end

    describe "for_this_instance" do

      before do
        @excel = Excel.new(:reuse => false)
        book = Workbook.open(@simple_file)
        book.excel.calculation = :manual
        book.close(:if_unsaved => :save)
      end

      it "should set options in the Excel instance" do
        @excel.for_this_instance(:displayalerts => true, :visible => true, :screenupdating => true, :calculation => :manual)
        @excel.DisplayAlerts.should be true
        @excel.Visible.should be true
        @excel.ScreenUpdating.should be true
        book = Workbook.open(@simple_file)
        book.excel.calculation = :manual
        book.save
        @excel.Calculation.should == XlCalculationManual
        book.close
      end

    end

    context "for_all_workbooks" do

      it "should not raise an error for an empty Excel instance" do
        excel = Excel.create
        expect{
          excel.for_all_workbooks(:visible => true, :read_only => true, :check_compatibility => true)
        }.to_not raise_error
      end

      it "should set options to true for a workbook" do
        book1 = Workbook.open(@simple_file1)
        excel1 = book1.excel
        excel1.for_all_workbooks(:visible => true, :read_only => true, :check_compatibility => true)
        excel1.Visible.should be true
        ole_workbook1 = book1.ole_workbook
        ole_workbook1.Windows(ole_workbook1.Name).Visible.should be true
        ole_workbook1.ReadOnly.should be true
        ole_workbook1.CheckCompatibility.should be true
        excel1.for_all_workbooks(:visible => false, :read_only => false, :check_compatibility => false)
        excel1.Visible.should be true
        ole_workbook1 = book1.ole_workbook
        ole_workbook1.Windows(ole_workbook1.Name).Visible.should be false
        ole_workbook1.ReadOnly.should be false
        ole_workbook1.CheckCompatibility.should be false
      end

      it "should set options for two workbooks" do
        book1 = Workbook.open(@simple_file1)
        book2 = Workbook.open(@different_file1)
        excel = book1.excel
        excel.for_all_workbooks(:visible => true, :check_compatibility => true, :read_only => true)
        excel.Visible.should be true
        ole_workbook1 = book1.ole_workbook
        ole_workbook2 = book2.ole_workbook
        ole_workbook1.Windows(ole_workbook1.Name).Visible.should be true
        ole_workbook2.Windows(ole_workbook2.Name).Visible.should be true
        ole_workbook1.CheckCompatibility.should be true
        ole_workbook2.CheckCompatibility.should be true
        ole_workbook1.ReadOnly.should be true
        ole_workbook2.ReadOnly.should be true    
        excel.for_all_workbooks(:visible => false, :check_compatibility => false, :read_only => false)
        excel.Visible.should be true
        ole_workbook1 = book1.ole_workbook
        ole_workbook2 = book2.ole_workbook
        ole_workbook1.Windows(ole_workbook1.Name).Visible.should be false
        ole_workbook2.Windows(ole_workbook2.Name).Visible.should be false
        ole_workbook1.CheckCompatibility.should be false
        ole_workbook2.CheckCompatibility.should be false
        ole_workbook1.ReadOnly.should be false
        ole_workbook2.ReadOnly.should be false    
      end

    end

    describe "known_excel_instances" do

      it "should return empty list" do
        Excel.known_excel_instances.should be_empty
      end

      it "should return list of one Excel process" do
        excel = Excel.new
        Excel.known_excel_instances.should == [excel]
        excel.close
      end

      it "should return list of two Excel processes" do
        excel1 = Excel.create
        excel2 = Excel.create
        Excel.known_excel_instances.should == [excel1,excel2]
      end

      it "should return list of two Excel processes" do
        excel1 = Excel.new
        excel2 = Excel.current
        excel3 = Excel.create
        Excel.known_excel_instances.should == [excel1,excel3]
      end

    end

    context "with hwnd and hwnd2excel" do
      
      before do
        Excel.kill_all
        @excel1 = Excel.new(:visible => true)
        @excel2 = Excel.new(:reuse => false, :visible => false)
      end

      it "should yield the correct hwnd" do
        @excel1.Hwnd.should == @excel1.hwnd
        @excel2.Hwnd.should == @excel2.hwnd
      end

   #   it "should provide the same excel instances" do
   #     @excel1.should_not == @excel2
   #     excel3 = Excel.hwnd2excel(@excel1.hwnd)
   #     excel4 = Excel.hwnd2excel(@excel2.hwnd)
   #     @excel1.should == excel3
   #     @excel2.should == excel4
   #     excel3.should_not == excel4
   #   end

=begin
      # does not work yet
      it "should not say 'probably recycled'" do
        e1_hwnd = @excel1.hwnd
        @excel1.close_workbooks
        weak_xl = WeakRef.new(@excel1.ole_excel)
        @excel1.Quit
        @excel1 = nil
        GC.start
        sleep 2
        process_id = Win32API.new("user32", "GetWindowThreadProcessId", ["I","P"], "I")
        pid_puffer = " " * 32
        process_id.call(e1_hwnd, pid_puffer)
        pid = pid_puffer.unpack("L")[0]
        begin
          Process.kill("KILL", pid) 
        rescue 
          trace "kill_error: #{$!}"
        end
        if weak_xl.weakref_alive? then
           #if WIN32OLE.ole_reference_count(weak_xlapp) > 0
          begin
            #weak_xl.ole_free
          rescue
            trace "weakref_probl_olefree"
          end
        end
        excel5 = Excel.new(:reuse => false)
        e1_again = Excel.hwnd2excel(e1_hwnd)
        e1_again.Hwnd.should == e1_hwnd
        e1_again.should == nil 
      end
=end
    end
    
    describe "generate workbook" do

      context "with standard" do
        
        before do
          @excel1 = Excel.create
          @file_name = @dir + '/bar.xls'
        end

        it "should generate a workbook" do
          workbook = @excel1.generate_workbook(@file_name)
          workbook.should be_a Workbook
          workbook.Name.should == File.basename(@file_name)
          workbook.FullName.should == General::absolute_path(@file_name)
          workbook.Saved.should be true
          workbook.ReadOnly.should be false
          workbook.Sheets.Count.should == 3
          workbooks = @excel1.Workbooks
          workbooks.Count.should == 1
        end

       
        it "should generate a workbook if one is already existing" do
          book = Workbook.open(@simple_file)
          workbook = @excel1.generate_workbook(@file_name)
          workbook.should be_a Workbook
          workbook.Name.should == File.basename(@file_name)
          workbook.FullName.should == General::absolute_path(@file_name)
          workbook.Saved.should be true
          workbook.ReadOnly.should be false
          workbook.Sheets.Count.should == 3
          workbooks = @excel1.Workbooks
          workbooks.Count.should == 2
        end

        it "should raise error if filename is with wrong path" do
          expect{
            workbook = @excel1.generate_workbook(@invalid_name_file)
          }.to raise_error(FileNotFound, /could not save workbook with filename/)
        end

        it "should raise error if filename is nil" do
          expect{
            workbook = @excel1.generate_workbook(@nil)
          }.to raise_error(FileNameNotGiven, "filename is nil")
        end

      end
    end

    context "setting the name of a range" do

       before do
        @book1 = Workbook.open(@dir + '/another_workbook.xls', :read_only => true, :visible => true)
        @book1.excel.displayalerts = false
        @excel1 = @book1.excel
      end

      after do
        @book1.close
      end   

      it "should name an unnamed range with a giving address" do
        @excel1.add_name("foo",[1,2])
        @excel1.Names.Item("foo").Name.should == "foo"
        @excel1.Names.Item("foo").Value.should == "=Sheet1!$B$1:$B$1"
      end

      it "should rename an already named range with a giving address" do
        @excel1.add_name("foo",[1,1])
        @excel1.Names.Item("foo").Name.should == "foo"
        @excel1.Names.Item("foo").Value.should == "=Sheet1!$A$1:$A$1"
      end

      it "should raise an error" do
        expect{
          @excel1.add_name("foo", [-2, 1])
        }.to raise_error(RangeNotEvaluatable, /cannot add name "foo" to range/)
      end

      it "should rename a range" do
        @excel1.add_name("foo",[1,1])
        @excel1.rename_range("foo","bar")
        @excel1.namevalue_glob("bar").should == "foo"
      end

      it "should delete a name of a range" do
        @excel1.add_name("foo",[1,1])
        @excel1.delete_name("foo")
        expect{
          @excel1.namevalue_glob("foo")
        }.to raise_error(NameNotFound, /name "foo"/)
      end

      it "should add a name of a rectangular range" do
        @excel1.add_name("foo",[1..3,1..4])
        @excel1["foo"].should == [["foo", "workbook", "sheet1", nil], ["foo", 1.0, 2.0, 4.0], ["matz", 3.0, 4.0, 4.0]] 
      end

      it "should accept the old interface" do
        @excel1.add_name("foo",1..3,1..4)
        @excel1["foo"].should == [["foo", "workbook", "sheet1", nil], ["foo", 1.0, 2.0, 4.0], ["matz", 3.0, 4.0, 4.0]] 
      end

      it "should add a name of an infinite row range" do
        @excel1.add_name("foo",[1..3, nil])
        @excel1.Names.Item("foo").Value.should == "=Sheet1!$1:$3"
      end

      it "should add a name of an infinite column range" do
        @excel1.add_name("foo",[nil, "A".."C"])
        @excel1.Names.Item("foo").Value.should == "=Sheet1!$A:$C"
      end

      it "should add a name of an infinite row range" do
        @excel1.add_name("foo",[nil, 1..3])
        @excel1.Names.Item("foo").Value.should == "=Sheet1!$A:$C"
      end

      it "should add a name of an infinite column range" do
        @excel1.add_name("foo",["A:C"])
        @excel1.Names.Item("foo").Value.should == "=Sheet1!$A:$C"
      end

      it "should add a name of an infinite column range" do
        @excel1.add_name("foo",["1:2"])
        @excel1.Names.Item("foo").Value.should == "=Sheet1!$1:$2"
      end

    end

    describe "namevalue_glob, set_namevalue_glob" do

      before do
        @book1 = Workbook.open(@dir + '/another_workbook.xls')
        @book1.Windows(@book1.Name).Visible = true
        @excel1 = @book1.excel
      end

      after do
        @book1.close(:if_unsaved => :forget)
      end   

      it "should return value of a defined name" do
        @excel1.namevalue_glob("firstcell").should == "foo"
        @excel1["firstcell"].should == "foo"
      end        

      #it "should evaluate a formula" do
      #  @excel1.namevalue_glob("named_formula").should == 4
      #  @excel1["named_formula"].should == 4
      #end

      it "should raise an error if name not defined" do
        expect {
          @excel1.namevalue_glob("foo")
        }.to raise_error(NameNotFound, /name "foo"/)
        expect {
        @excel1["foo"]
        }.to raise_error(NameNotFound, /name "foo"/)
        expect {
          excel2 = Excel.create
          excel2.namevalue_glob("one")
        }.to raise_error(NameNotFound, /name "one"/)
        expect {
          excel3 = Excel.create(:visible => true)
          excel3["one"]
        }.to raise_error(NameNotFound, /name "one"/)
      end

      it "should set a range to a value" do
        @excel1.namevalue_glob("firstcell").should == "foo"
        @excel1.set_namevalue_glob("firstcell","bar")
        @excel1.namevalue_glob("firstcell").should == "bar"
        @excel1["firstcell"] = "foo"
        @excel1.namevalue_glob("firstcell").should == "foo"
      end

      it "should raise an error if name cannot be evaluated" do
        expect{
          @excel1.set_namevalue_glob("foo", 1)
          }.to raise_error(RangeNotEvaluatable, /cannot assign value to range named "foo"/)
        expect{
          @excel1["foo"] = 1
          }.to raise_error(RangeNotEvaluatable, /cannot assign value to range named "foo"/)
      end

      it "should color the cell (deprecated)" do
        @excel1.set_namevalue_glob("firstcell", "foo")
        @book1.Names.Item("firstcell").RefersToRange.Interior.ColorIndex.should == -4142 
        @excel1.set_namevalue_glob("firstcell", "foo", :color => 4)
        @book1.Names.Item("firstcell").RefersToRange.Interior.ColorIndex.should == 4
        @excel1["firstcell"].should == "foo"
        @excel1["firstcell"] = "foo"
        @excel1.Names.Item("firstcell").RefersToRange.Interior.ColorIndex.should == 42
        @book1.save
      end

      it "should color the cell" do
        @excel1.set_namevalue_glob("firstcell", "foo")
        @book1.Names.Item("firstcell").RefersToRange.Interior.ColorIndex.should == -4142 
        @excel1.set_namevalue_glob("firstcell", "foo", :color => 4)
        @book1.Names.Item("firstcell").RefersToRange.Interior.ColorIndex.should == 4
        @excel1["firstcell"].should == "foo"
        @excel1["firstcell"] = "foo"
        @excel1.Names.Item("firstcell").RefersToRange.Interior.ColorIndex.should == 42
        @book1.save
      end


    end

    describe "namevalue, set_namevalue" do
      
      before do
        @book1 = Workbook.open(@another_simple_file)
        @excel1 = @book1.excel
        # for some reason the workbook must be visible
        @book1.visible = true
      end

      after do
        @book1.close(:if_unsaved => :forget)
      end   

      it "should return value of a locally defined name" do
        @excel1.namevalue("firstcell").should == "foo"          
      end        

      it "should return value of a defined name" do
        @excel1.namevalue("new").should == "foo"         
        @excel1.namevalue("one").should == 1.0    
        @excel1.namevalue("four").should == [[1,2],[3,4]]
        @excel1.namevalue("firstrow").should == [[1,2]]
      end    

      it "should return default value if name not defined and default value is given" do
        @excel1.namevalue("foo", :default => 2).should == 2
      end

      it "should raise an error if name not defined for the sheet" do
        expect {
          @excel1.namevalue("foo")
          }.to raise_error(NameNotFound, /name "foo" not in/)
        expect {
          @excel1.namevalue("named_formula")
          }.to raise_error(NameNotFound, /name "named_formula" not in/)
        expect {
          excel2 = Excel.create
          excel2.namevalue("one")
        }.to raise_error(NameNotFound, /name "one" not in/)
      end
    
      it "should set a range to a value" do
        @excel1.namevalue("firstcell").should == "foo"
        @excel1.set_namevalue("firstcell","bar")
        @excel1.namevalue("firstcell").should == "bar"
      end

      it "should raise an error if name cannot be evaluated" do
        expect{
          @excel1.set_namevalue_glob("foo", 1)
        }.to raise_error(RangeNotEvaluatable, /cannot assign value to range named "foo" in/)
      end

      it "should color the cell (depracated)" do
        @excel1.set_namevalue("firstcell", "foo")
        @book1.Names.Item("firstcell").RefersToRange.Interior.ColorIndex.should == -4142 
        @excel1.set_namevalue("firstcell", "foo", :color => 4)
        @book1.Names.Item("firstcell").RefersToRange.Interior.ColorIndex.should == 4
      end

      it "should color the cell" do
        @excel1.set_namevalue("firstcell", "foo")
        @book1.Names.Item("firstcell").RefersToRange.Interior.ColorIndex.should == -4142 
        @excel1.set_namevalue("firstcell", "foo", :color => 4)
        @book1.Names.Item("firstcell").RefersToRange.Interior.ColorIndex.should == 4
      end


    end

  end
end

# @private
class TestError < RuntimeError 
end
