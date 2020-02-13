# -*- coding: utf-8 -*-

require File.join(File.dirname(__FILE__), './../spec_helper')


$VERBOSE = nil

include RobustExcelOle
include General

describe Workbook do

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
    @simple_file_xlsm = @dir + '/workbook.xlsm'
    @simple_file_xlsx = @dir + '/workbook.xlsx'
    @simple_file1 = @simple_file
  end

  after do
    Excel.kill_all
    rm_tmp(@dir)
  end

  
  describe "unobtrusively" do

    # @private
    def unobtrusively_ok?
      Workbook.unobtrusively(@simple_file) do |book|
        book.should be_a Workbook
        sheet = book.sheet(1)
        sheet[1,1] = sheet[1,1].Value == "foo" ? "bar" : "foo"
        book.should be_alive
        book.Saved.should be false
      end
    end

    it "should do simple unobtrusively" do
      expect{unobtrusively_ok?}.to_not raise_error
    end

    describe "block transparency" do

      it "should return correct value of the block" do
        (Workbook.unobtrusively(@simple_file1) do |book| 
          22
         end).should == 22
      end

      it "should return value of the last block" do
        (Workbook.unobtrusively(@simple_file1) do |book|
          Workbook.unobtrusively(@different_file) do |book2|
            11
          end
          22
        end).should == 22
      end

      it "should return correct value in several blocks" do
        (Workbook.unobtrusively(@simple_file1) do |book|
          Workbook.unobtrusively(@different_file) do |book2|
            22
          end
        end).should == 22
      end

    end

    describe "unknown workbooks" do

      context "with one invisible saved writable workbook" do

        before do
          @ole_e1 = WIN32OLE.new('Excel.Application')
          ws = ole_e1.Workbooks
          @abs_filename = General.absolute_path(@simple_file1)
          @ole_wb = ws.Open(abs_filename)
          @old_value = @ole_wb.Worksheets.Item(1).Cells.Item(1,1).Value
        end

        it "should connect" do
          Workbooks.unobtrusively(@simple_file1) do |book|
            book.excel.Workbook.Count.should == 1
            Excel.excels_number.should == 1
            book.FullName.should == General.absolute_path(@simple_file1)
            book.saved.should be true
            book.visible.should be false
            book.writable.should be true
          end
          ole_wb = WIN32OLE.connect(@abs_filename)
          ole_wb.Saved.should be true
          @ole_e1.Visible.should be false
          ole_wb.ReadOnly.should be false
        end

        it "should set visible => true and remain invisiblity" do
          Workbooks.unobtrusively(@simple_file1, :visible => true) do |book|
            book.saved.should be true
            book.visible.should be true
            book.writable.should be true
          end
          ole_wb = WIN32OLE.connect(@abs_filename)
          ole_wb.Saved.should be true
          @ole_e1.Visible.should be false
          ole_wb.ReadOnly.should be false
        end

        it "should set read_only => true and remain writability" do
          Workbooks.unobtrusively(@simple_file1, :read_only => true) do |book|
            book.saved.should be true
            book.visible.should be false
            book.writable.should be false
          end
          ole_wb = WIN32OLE.connect(@abs_filename)
          ole_wb.Saved.should be true
          @ole_e1.Visible.should be false
          ole_wb.ReadOnly.should be false
        end

        it "should set visible => true, read_only => true" do
          Workbooks.unobtrusively(@simple_file1, :visible => true, :read_only => true) do |book|
            book.saved.should be true
            book.visible.should be true
            book.writable.should be false
          end
          ole_wb = WIN32OLE.connect(@abs_filename)
          ole_wb.Saved.should be true
          @ole_e1.Visible.should be false
          ole_wb.ReadOnly.should be false
        end

        it "should modify and remain saved-status" do
          Workbooks.unobtrusively(@simple_file1) do |book|
            book.saved.should be true
            book.visible.should be false
            book.writable.should be true
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo"
            @new_value = book.sheet(1)[1,1].value
            book.Saved.should be false
          end
          ole_wb = WIN32OLE.connect(@abs_filename)
          ole_wb.Saved.should be true
          @ole_e1.Visible.should be false
          ole_wb.ReadOnly.should be false
          ole_wb.Close
          book2 = Workbook.open(@simple_file1)
          book2.sheet(1)[1,1].value.should_not == @old_value
          book2.sheet(1)[1,1].value.should == @new_value
        end

        it "should modify and remain saved-status and not save the new value when writable => false" do
          Workbooks.unobtrusively(@simple_file1, :writable => false) do |book|
            book.saved.should be true
            book.visible.should be false
            book.writable.should be true
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo"
            @new_value = book.sheet(1)[1,1].value
            book.Saved.should be false
          end
          ole_wb = WIN32OLE.connect(@abs_filename)
          ole_wb.Saved.should be true
          @ole_e1.Visible.should be false
          ole_wb.ReadOnly.should be false
          ole_wb.Close
          book2 = Workbook.open(@simple_file1)
          book2.sheet(1)[1,1].value.should == @old_value
          book2.sheet(1)[1,1].value.should_not == @new_value
        end

      end

      context "with one visible saved writable workbook" do

        before do
          @ole_e1 = WIN32OLE.new('Excel.Application')
          ws = ole_e1.Workbooks
          @abs_filename = General.absolute_path(@simple_file1)
          @ole_wb = ws.Open(abs_filename)
          @ole_e1.Visible = true
          @ole_wb.Windows(@ole_wb.Name).Visible = true
        end

        it "should remain visibility" do
          Workbooks.unobtrusively(@simple_file1) do |book|
            book.saved.should be true
            book.visible.should be true
            book.writable.should be true
          end
          ole_wb = WIN32OLE.connect(@abs_filename)
          ole_wb.Saved.should be true
          @ole_e1.Visible.should be true
          ole_wb.ReadOnly.should be false
        end

        it "should set visible => false and remain visibility" do
          Workbooks.unobtrusively(@simple_file1, :visible => false) do |book|
            book.saved.should be true
            book.visible.should be false
            book.writable.should be true
          end
          ole_wb = WIN32OLE.connect(@abs_filename)
          ole_wb.Saved.should be true
          @ole_e1.Visible.should be true
          ole_wb.ReadOnly.should be false
        end

      end

      context "with one unsaved writable workbook" do

        before do
          @ole_e1 = WIN32OLE.new('Excel.Application')
          ws = ole_e1.Workbooks
          @abs_filename = General.absolute_path(@simple_file1)
          @ole_wb = ws.Open(abs_filename)
          @ole_e1.Visible = true
          @ole_wb.Windows(@ole_wb.Name).Visible = true
          @old_value = @ole_wb.Worksheets.Item(1).Cells.Item(1,1).Value
          @ole_wb.Worksheets.Item(name).Cells.Item(1,1).Value = @old_value == "foo" ? "bar" : "foo"
          @new_value = @ole_wb.Worksheets.Item(1).Cells.Item(1,1).Value
          @ole_wb.Saved.should be false
        end

        it "should connect and remain unsaved" do
          Workbooks.unobtrusively(@simple_file1) do |book|
            book.saved.should be false
            book.visible.should be true
            book.writable.should be true
          end
          ole_wb = WIN32OLE.connect(@abs_filename)
          ole_wb.Saved.should be false
          @ole_e1.Visible.should be true
          ole_wb.ReadOnly.should be false
        end

        it "should remain writable" do
          Workbooks.unobtrusively(@simple_file1, :read_only => true) do |book|
            book.saved.should be false
            book.visible.should be true
            book.writable.should be false
          end
          ole_wb = WIN32OLE.connect(@abs_filename)
          ole_wb.Saved.should be false
          @ole_e1.Visible.should be true
          ole_wb.ReadOnly.should be false
        end

        it "should remain unsaved when modifying" do
          Workbooks.unobtrusively(@simple_file1) do |book|
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo"
            @new_value = book.sheet(1)[1,1].Value
            book.saved.should be false
            book.visible.should be true
            book.writable.should be true
          end
          ole_wb = WIN32OLE.connect(@abs_filename)
          ole_wb.Saved.should be false
          @ole_e1.Visible.should be true
          ole_wb.ReadOnly.should be false
          ole_wb.Close
          book2 = Workbook.open(@simple_file1)
          book2.sheet(1)[1,1].value.should_not == @old_value
          book2.sheet(1)[1,1].value.should == @new_value
        end

        it "should not write with :writable => false" do
          Workbooks.unobtrusively(@simple_file1, :writable => false) do |book|
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo"
            @new_value = book.sheet(1)[1,1].Value
            book.saved.should be false
            book.visible.should be true
            book.writable.should be true
          end
          ole_wb = WIN32OLE.connect(@abs_filename)
          ole_wb.Saved.should be false
          @ole_e1.Visible.should be true
          ole_wb.ReadOnly.should be false
          ole_wb.Close
          book2 = Workbook.open(@simple_file1)
          book2.sheet(1)[1,1].value.should == @old_value
          book2.sheet(1)[1,1].value.should_not == @new_value
        end

      end

      context "with one read-only workbook" do

        before do
          @ole_e1 = WIN32OLE.new('Excel.Application')
          ws = ole_e1.Workbooks
          @abs_filename = General.absolute_path(@simple_file1)
          @ole_wb = ws.Open(abs_filename, RobustExcelOle::XlUpdateLinksNever, true)
          @ole_e1.Visible = true
          @ole_wb.Windows(@ole_wb.Name).Visible = true
          @ole_wb.ReadOnly.should be true
        end

        it "should connect and remain read-only" do
          Workbooks.unobtrusively(@simple_file1) do |book|
            book.saved.should be true
            book.visible.should be true
            book.writable.should be false
          end
          ole_wb = WIN32OLE.connect(@abs_filename)
          ole_wb.Saved.should be true
          @ole_e1.Visible.should be true
          ole_wb.ReadOnly.should be true
        end

        it "should remain read-only" do
          Workbooks.unobtrusively(@simple_file1, :read_only => false) do |book|
            book.saved.should be false
            book.visible.should be true
            book.writable.should be true
          end
          ole_wb = WIN32OLE.connect(@abs_filename)
          ole_wb.Saved.should be true
          @ole_e1.Visible.should be true
          ole_wb.ReadOnly.should be true
        end

        it "should remain read-only when modifying" do
          Workbooks.unobtrusively(@simple_file1) do |book|
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo"
            @new_value = book.sheet(1)[1,1].Value
            book.saved.should be false
            book.visible.should be true
            book.writable.should be true
          end
          ole_wb = WIN32OLE.connect(@abs_filename)
          ole_wb.Saved.should be true
          @ole_e1.Visible.should be true
          ole_wb.ReadOnly.should be true
          ole_wb.Close
          book2 = Workbook.open(@simple_file1)
          book2.sheet(1)[1,1].value.should_not == @old_value
          book2.sheet(1)[1,1].value.should == @new_value
        end

        it "should remain read-only when modifying" do
          Workbooks.unobtrusively(@simple_file1, :read_only => false, :writable => true) do |book|
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo"
            @new_value = book.sheet(1)[1,1].Value
            book.saved.should be false
            book.visible.should be true
            book.writable.should be true
          end
          ole_wb = WIN32OLE.connect(@abs_filename)
          ole_wb.Saved.should be true
          @ole_e1.Visible.should be true
          ole_wb.ReadOnly.should be true
          ole_wb.Close
          book2 = Workbook.open(@simple_file1)
          book2.sheet(1)[1,1].value.should_not == @old_value
          book2.sheet(1)[1,1].value.should == @new_value
        end

        it "should remain read-only when modifying and not save changes, when :writable => false" do
          Workbooks.unobtrusively(@simple_file1, :writable => false) do |book|
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo"
            @new_value = book.sheet(1)[1,1].Value
            book.saved.should be false
            book.visible.should be true
            book.writable.should be true
          end
          ole_wb = WIN32OLE.connect(@abs_filename)
          ole_wb.Saved.should be true
          @ole_e1.Visible.should be true
          ole_wb.ReadOnly.should be true
          ole_wb.Close
          book2 = Workbook.open(@simple_file1)
          book2.sheet(1)[1,1].value.should == @old_value
          book2.sheet(1)[1,1].value.should_not == @new_value
        end

        it "should remain read-only when modifying and not save changes, when :writable => false" do
          Workbooks.unobtrusively(@simple_file1, :read_only => false, :writable => false) do |book|
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo"
            @new_value = book.sheet(1)[1,1].Value
            book.saved.should be false
            book.visible.should be true
            book.writable.should be true
          end
          ole_wb = WIN32OLE.connect(@abs_filename)
          ole_wb.Saved.should be true
          @ole_e1.Visible.should be true
          ole_wb.ReadOnly.should be true
          ole_wb.Close
          book2 = Workbook.open(@simple_file1)
          book2.sheet(1)[1,1].value.should == @old_value
          book2.sheet(1)[1,1].value.should_not == @new_value
        end

      end

    end

    describe "excels number" do

      it "should open one excel instance and workbook should be closed" do
        Workbook.unobtrusively(@simple_file1){ |book| nil }
        Excel.excels_number.should == 1
      end

    end

    describe "closed workbook" do

      it "should close the workbook by default" do
        Workbook.unobtrusively(@simple_file1){ |book| nil}
        Excel.current.Workbooks.Count.should == 0  
      end

    end

    describe "writability" do

      context "with no book" do

        it "should open read-write" do
          Workbook.unobtrusively(@simple_file1) do |book|            
            book.ReadOnly.should be false
            @old_value = book.sheet(1)[1,1].Value
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo"
            book.Saved.should be false
          end          
          Excel.current.Workbooks.Count.should == 0
          b1 = Workbook.open(@simple_file1)
          b1.sheet(1)[1,1].Value.should_not == @old_value
        end

        it "should open as read-write" do
          Workbook.unobtrusively(@simple_file, :read_only => false) do |book|
            book.ReadOnly.should be false
            @old_value = book.sheet(1)[1,1].Value
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo"
            book.Saved.should be false
          end
          Excel.current.Workbooks.Count.should == 0
          b1 = Workbook.open(@simple_file1)
          b1.sheet(1)[1,1].Value.should_not == @old_value
        end

        it "should open as read-write" do
          Workbook.unobtrusively(@simple_file, :read_only => false, :writable => false) do |book|
            book.ReadOnly.should be false
            @old_value = book.sheet(1)[1,1].Value
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo"
            book.Saved.should be false
          end
          Excel.current.Workbooks.Count.should == 0
          b1 = Workbook.open(@simple_file1)
          b1.sheet(1)[1,1].Value.should_not == @old_value
        end

        it "should open as read-write" do
          Workbook.unobtrusively(@simple_file, :writable => true) do |book|
            book.ReadOnly.should be false
            @old_value = book.sheet(1)[1,1].Value
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo"
            book.Saved.should be false
          end
          Excel.current.Workbooks.Count.should == 0
          b1 = Workbook.open(@simple_file1)
          b1.sheet(1)[1,1].Value.should_not == @old_value
        end

        it "should open as read-write" do
          Workbook.unobtrusively(@simple_file, :writable => true, :read_only => false) do |book|
            book.ReadOnly.should be false
            @old_value = book.sheet(1)[1,1].Value
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo"
            book.Saved.should be false
          end
          Excel.current.Workbooks.Count.should == 0
          b1 = Workbook.open(@simple_file1)
          b1.sheet(1)[1,1].Value.should_not == @old_value
        end

        it "should open read-only" do
          Workbook.unobtrusively(@simple_file, :read_only => true) do |book|            
            book.ReadOnly.should be true
            @old_value = book.sheet(1)[1,1].Value
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo"
            book.Saved.should be false
          end
          Excel.current.Workbooks.Count.should == 0
          b1 = Workbook.open(@simple_file1)
          b1.sheet(1)[1,1].Value.should == @old_value
        end

        it "should open read-only" do
          Workbook.unobtrusively(@simple_file, :read_only => true, :writable => false) do |book|            
            book.ReadOnly.should be true
            @old_value = book.sheet(1)[1,1].Value
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo"
            book.Saved.should be false
          end
          Excel.current.Workbooks.Count.should == 0
          b1 = Workbook.open(@simple_file1)
          b1.sheet(1)[1,1].Value.should == @old_value
        end

        it "should open not writable" do
          Workbook.unobtrusively(@simple_file, :writable => false) do |book|
            book.ReadOnly.should be true
            @old_value = book.sheet(1)[1,1].Value
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo"
            book.Saved.should be false            
          end
          Excel.current.Workbooks.Count.should == 0          
          b1 = Workbook.open(@simple_file1)
          b1.sheet(1)[1,1].Value.should == @old_value
        end

        it "should raise error if both options are true" do          
          expect{
            Workbook.unobtrusively(@simple_file, :writable => true, :read_only => true) {|book|}
          }.to raise_error(OptionInvalid, "contradicting options")
        end
      end
  
      context "with open writable book" do

        before do
           @book = Workbook.open(@simple_file1)
           @old_value = @book.sheet(1)[1,1].Value
        end

        it "should open as read-write by default" do
          Workbook.unobtrusively(@simple_file1) do |book|
            book.Readonly.should be false
            book.should == @book
            book.filename.should == @book.filename
            book.excel.should == @book.excel
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo"
          end
          @book.close
          book2 = Workbook.open(@simple_file1)
          book2.sheet(1)[1,1].Value.should_not == @old_value
        end

        it "should open as read-write" do
          Workbook.unobtrusively(@simple_file1, :read_only => false) do |book|
            book.Readonly.should be false
            book.should == @book
            book.filename.should == @book.filename
            book.excel.should == @book.excel
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo"
          end
          @book.close
          book2 = Workbook.open(@simple_file1)
          book2.sheet(1)[1,1].Value.should_not == @old_value
        end

        it "should open as read-write" do
          Workbook.unobtrusively(@simple_file1, :read_only => false, :writable => false) do |book|
            book.Readonly.should be false
            book.should == @book
            book.filename.should == @book.filename
            book.excel.should == @book.excel
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo"
          end
          @book.close
          book2 = Workbook.open(@simple_file1)
          book2.sheet(1)[1,1].Value.should_not == @old_value
        end

        it "should open as read-write" do
          Workbook.unobtrusively(@simple_file1, :writable => true) do |book|
            book.Readonly.should be false
            book.should == @book
            book.filename.should == @book.filename
            book.excel.should == @book.excel
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo"
          end
          @book.close
          book2 = Workbook.open(@simple_file1)
          book2.sheet(1)[1,1].Value.should_not == @old_value
        end

        it "should open as read-write" do          
          Workbook.unobtrusively(@simple_file1, :writable => true, :read_only => false) do |book|
            book.Readonly.should be false
            book.should == @book
            book.filename.should == @book.filename
            book.excel.should == @book.excel
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo"
          end
          @book.close
          book2 = Workbook.open(@simple_file1)
          book2.sheet(1)[1,1].Value.should_not == @old_value
        end

        it "should force to read-only" do
          Workbook.unobtrusively(@simple_file1, :read_only => true) do |book|            
            book.ReadOnly.should be true
            book.should == @book
            book.filename.should == @book.filename
            book.excel.should == @book.excel
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo" 
          end
          @book.close
          book2 = Workbook.open(@simple_file1)
          book2.sheet(1)[1,1].Value.should == @old_value
        end

        it "should force to read-only" do
          Workbook.unobtrusively(@simple_file1, :read_only => true, :writable => false) do |book|            
            book.ReadOnly.should be true
            book.should == @book
            book.filename.should == @book.filename
            book.excel.should == @book.excel
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo" 
          end
          @book.close
          book2 = Workbook.open(@simple_file1)
          book2.sheet(1)[1,1].Value.should == @old_value
        end

        it "should open not writable" do
          Workbook.unobtrusively(@simple_file1, :writable => false) do |book|
            book.ReadOnly.should be false
            book.should == @book
            book.filename.should == @book.filename
            book.excel.should == @book.excel
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo"
          end
          @book.close
          book2 = Workbook.open(@simple_file1)
          book2.sheet(1)[1,1].Value.should == @old_value
        end
      end

      context "with open read-only book" do

        before do
          @book = Workbook.open(@simple_file1, :read_only => true)
          @old_value = @book.sheet(1)[1,1].Value 
        end

        it "should not write" do
          Workbook.unobtrusively(@simple_file1, :writable => false) do |book|
            book.Readonly.should be true
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo"
            book.sheet(1)[1,1].Value.should_not == @old_value
          end
          @book.ReadOnly.should be true
          @book.close
          Workbook.unobtrusively(@simple_file1, :writable => false) do |book|
            book.sheet(1)[1,1].Value.should == @old_value
          end
          #@book.close
          #book2 = Workbook.open(@simple_file1)
          #book2.sheet(1)[1,1].Value.should == @old_value
        end

        it "should not change the read_only mode" do
          Workbook.unobtrusively(@simple_file1) do |book|
            book.Readonly.should be true
            book.should == @book
            book.filename.should == @book.filename
            book.excel.should == @book.excel
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo"
          end
          @book.ReadOnly.should be true
          @book.close
          book2 = Workbook.open(@simple_file1)
          book2.sheet(1)[1,1].Value.should == @old_value
        end

        it "should open as read-write" do
          Workbook.unobtrusively(@simple_file1, :read_only => false) do |book|
            book.Readonly.should be true
            book.should == @book
            book.filename.should == @book.filename
            book.excel.should == @book.excel
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo"
          end
          @book.close
          book2 = Workbook.open(@simple_file1)
          book2.sheet(1)[1,1].Value.should == @old_value
        end

        it "should open as read-write" do
          Workbook.unobtrusively(@simple_file1, :read_only => false, :writable => false) do |book|
            book.Readonly.should be true
            book.should == @book
            book.filename.should == @book.filename
            book.excel.should == @book.excel
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo"
          end
          @book.close
          book2 = Workbook.open(@simple_file1)
          book2.sheet(1)[1,1].Value.should == @old_value
        end

        it "should force to read-write" do
          Workbook.unobtrusively(@simple_file1, :writable => true) do |book|
            book.ReadOnly.should be false
            book.should == @book
            book.filename.should == @book.filename
            book.excel.should == @book.excel
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo"
          end
          @book.close
          book2 = Workbook.open(@simple_file1)
          book2.sheet(1)[1,1].Value.should_not == @old_value
        end

        it "should force to read-write" do
          Workbook.unobtrusively(@simple_file1, :writable => true, :read_only => false) do |book|
            book.Readonly.should be false
            book.should == @book
            book.filename.should == @book.filename
            book.excel.should == @book.excel
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo"
          end
          @book.close
          book2 = Workbook.open(@simple_file1)
          book2.sheet(1)[1,1].Value.should_not == @old_value
        end

=begin
        it "should force to read-write" do
          e1 = Excel.create
          Workbook.unobtrusively(@simple_file1, :writable => true, :rw_change_excel => e1) do |book|
            book.Readonly.should be false
            book.filename.should == @book.filename
            book.excel.should == e1
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo"
          end
          @book.close
          book2 = Workbook.open(@simple_file1)
          book2.sheet(1)[1,1].Value.should_not == @old_value
        end

        it "should force to read-write" do
          Workbook.unobtrusively(@simple_file1, :writable => true, :rw_change_excel => :current) do |book|
            book.Readonly.should be false
            book.should == @book
            book.filename.should == @book.filename
            book.excel.should == @book.excel
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo"
          end
          @book.close
          book2 = Workbook.open(@simple_file1)
          book2.sheet(1)[1,1].Value.should_not == @old_value
        end

        it "should force to read-write" do
          Workbook.unobtrusively(@simple_file1, :writable => true, :rw_change_excel => :new) do |book|
            book.Readonly.should be false
            book.filename.should == @book.filename
            book.excel.should_not == @book.excel
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo"
          end
          @book.close
          book2 = Workbook.open(@simple_file1)
          book2.sheet(1)[1,1].Value.should_not == @old_value
        end
=end

        it "should force to read-write" do
          Workbook.unobtrusively(@simple_file1, :writable => true, :read_only => false) do |book|
            book.Readonly.should be false
            book.should == @book
            book.filename.should == @book.filename
            book.excel.should == @book.excel
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo"
          end
          @book.close
          book2 = Workbook.open(@simple_file1)
          book2.sheet(1)[1,1].Value.should_not == @old_value
        end

        it "should force to read-only" do
          Workbook.unobtrusively(@simple_file1, :read_only => true) do |book|            
            book.ReadOnly.should be true
            book.should == @book
            book.filename.should == @book.filename
            book.excel.should == @book.excel
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo" 
          end
          @book.close
          book2 = Workbook.open(@simple_file1)
          book2.sheet(1)[1,1].Value.should == @old_value
        end

        it "should force to read-only" do
          Workbook.unobtrusively(@simple_file1, :read_only => true, :writable => false) do |book|            
            book.ReadOnly.should be true
            book.should == @book
            book.filename.should == @book.filename
            book.excel.should == @book.excel
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo" 
          end
          @book.close
          book2 = Workbook.open(@simple_file1)
          book2.sheet(1)[1,1].Value.should == @old_value
        end

        it "should open not writable" do
          Workbook.unobtrusively(@simple_file1, :writable => false) do |book|
            book.ReadOnly.should be true
            book.should == @book
            book.filename.should == @book.filename
            book.excel.should == @book.excel
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo"
          end
          @book.close
          book2 = Workbook.open(@simple_file1)
          book2.sheet(1)[1,1].Value.should == @old_value
        end
      end

      context "with open unsaved writable book" do

        before do
           @book = Workbook.open(@simple_file1)           
           @book.sheet(1)[1,1] = @book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo"           
           @old_value = @book.sheet(1)[1,1].Value
        end

        it "should open as read-write by default" do
          Workbook.unobtrusively(@simple_file1) do |book|
            book.Readonly.should be false
            book.should == @book
            book.filename.should == @book.filename
            book.excel.should == @book.excel
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo"
          end
          @book.Saved.should be false
          @book.sheet(1)[1,1].Value.should_not == @old_value
        end

        it "should open as read-write" do
          Workbook.unobtrusively(@simple_file1, :read_only => false) do |book|
            book.Readonly.should be false
            book.should == @book
            book.filename.should == @book.filename
            book.excel.should == @book.excel
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo"
          end
          @book.Saved.should be false
          @book.sheet(1)[1,1].Value.should_not == @old_value
        end

        it "should open as read-write" do
          Workbook.unobtrusively(@simple_file1, :read_only => false, :writable => false) do |book|
            book.Readonly.should be false
            book.should == @book
            book.filename.should == @book.filename
            book.excel.should == @book.excel
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo"
          end
          @book.Saved.should be false
          @book.sheet(1)[1,1].Value.should_not == @old_value
        end

        it "should force to read-write" do
          Workbook.unobtrusively(@simple_file1, :writable => true) do |book|
            book.Readonly.should be false
            book.should == @book
            book.filename.should == @book.filename
            book.excel.should == @book.excel
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo"
          end
          @book.Saved.should be false
          @book.sheet(1)[1,1].Value.should_not == @old_value
        end

        it "should force to read-write" do          
          Workbook.unobtrusively(@simple_file1, :writable => true, :read_only => false) do |book|
            book.Readonly.should be false
            book.should == @book
            book.filename.should == @book.filename
            book.excel.should == @book.excel
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo"
          end
          @book.Saved.should be false
          @book.sheet(1)[1,1].Value.should_not == @old_value
        end

        it "should force to read-only (not implemented)" do
          expect{
            Workbook.unobtrusively(@simple_file1, :read_only => true) 
          }.to raise_error(NotImplementedREOError)
        end

        it "should force to read-only (not implemented)" do
          expect{
            Workbook.unobtrusively(@simple_file1, :read_only => true, :writable => false) 
          }.to raise_error(NotImplementedREOError)
        end

        it "should open not writable" do
          Workbook.unobtrusively(@simple_file1, :writable => false) do |book|
            book.ReadOnly.should be false
            book.should == @book
            book.filename.should == @book.filename
            book.excel.should == @book.excel
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo"
          end
          @book.Saved.should be false
          @book.sheet(1)[1,1].Value.should_not == @old_value
          @book.close(:if_unsaved => :forget)
          book2 = Workbook.open(@simple_file1)
          book2.sheet(1)[1,1].Value.should_not == @old_value
        end
      end

      context "with open unsaved read-only book" do

        before do
          @book = Workbook.open(@simple_file1, :read_only => true)          
          @book.sheet(1)[1,1] = @book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo"
          @old_value = @book.sheet(1)[1,1].Value
        end

        it "should open as read-only by default" do
          Workbook.unobtrusively(@simple_file1) do |book|
            book.Readonly.should be true
            book.should == @book
            book.filename.should == @book.filename
            book.excel.should == @book.excel
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo"
          end
          @book.Saved.should be false
          @book.ReadOnly.should be true
          @book.sheet(1)[1,1].Value.should_not == @old_value
          @book.close
          book2 = Workbook.open(@simple_file1)
          book2.sheet(1)[1,1].Value.should_not == @old_value
        end

        it "should open as read-only" do
          Workbook.unobtrusively(@simple_file1, :read_only => false) do |book|
            book.Readonly.should be true
            book.should == @book
            book.filename.should == @book.filename
            book.excel.should == @book.excel
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo"
          end
          @book.Saved.should be false
          @book.ReadOnly.should be true
          @book.sheet(1)[1,1].Value.should_not == @old_value
          @book.close
          book2 = Workbook.open(@simple_file1)
          book2.sheet(1)[1,1].Value.should_not == @old_value
        end

        it "should open as read-only" do
          Workbook.unobtrusively(@simple_file1, :read_only => false, :writable => false) do |book|
            book.Readonly.should be true
            book.should == @book
            book.filename.should == @book.filename
            book.excel.should == @book.excel
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo"
          end
          @book.Saved.should be false
          @book.ReadOnly.should be true
          @book.sheet(1)[1,1].Value.should_not == @old_value
          @book.close
          book2 = Workbook.open(@simple_file1)
          book2.sheet(1)[1,1].Value.should_not == @old_value
        end
    
        it "should raise an error" do
          expect{
            Workbook.unobtrusively(@simple_file1, :writable => true)
            }.to raise_error(NotImplementedREOError, "unsaved read-only workbook shall be written")
        end

        it "should raise an error" do
           expect{
            Workbook.unobtrusively(@simple_file1, :writable => true)
            }.to raise_error(NotImplementedREOError, "unsaved read-only workbook shall be written")
        end

        it "should force to read-only" do
          Workbook.unobtrusively(@simple_file1, :read_only => true) do |book|            
            book.ReadOnly.should be true
            book.should == @book
            book.filename.should == @book.filename
            book.excel.should == @book.excel
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo" 
          end
          @book.Saved.should be false
          @book.ReadOnly.should be true
          @book.sheet(1)[1,1].Value.should_not == @old_value
          @book.close
          book2 = Workbook.open(@simple_file1)
          book2.sheet(1)[1,1].Value.should_not == @old_value
        end

        it "should force to read-only" do
          Workbook.unobtrusively(@simple_file1, :read_only => true, :writable => false) do |book|            
            book.ReadOnly.should be true
            book.should == @book
            book.filename.should == @book.filename
            book.excel.should == @book.excel
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo" 
          end
          @book.Saved.should be false
          @book.ReadOnly.should be true
          @book.sheet(1)[1,1].Value.should_not == @old_value
          @book.close
          book2 = Workbook.open(@simple_file1)
          book2.sheet(1)[1,1].Value.should_not == @old_value
        end

        it "should open not writable" do
          Workbook.unobtrusively(@simple_file1, :writable => false) do |book|
            book.ReadOnly.should be true
            book.should == @book
            book.filename.should == @book.filename
            book.excel.should == @book.excel
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo"
          end
          @book.Saved.should be false
          @book.ReadOnly.should be true
          @book.sheet(1)[1,1].Value.should_not == @old_value
          @book.close
          book2 = Workbook.open(@simple_file1)
          book2.sheet(1)[1,1].Value.should_not == @old_value
        end
      end

    end

    describe "intertweened blocks" do

      context "with no book" do
      
        it "should not close the book in the outer block" do
          Workbook.unobtrusively(@simple_file1) do |book|
            Workbook.unobtrusively(@simple_file1) do |book2|
              book2.should == book
            end
            book.should be_alive
            Workbook.unobtrusively(@simple_file1) do |book3|
              book3.should == book
            end
            book.should be_alive
          end
        end

        it "should not close the book in the outer block with writable false" do
          Workbook.unobtrusively(@simple_file1, :writable => false) do |book|
            Workbook.unobtrusively(@simple_file1, :writable => false) do |book2|
              book2.should == book
            end
            book.should be_alive
          end
        end

        it "should write in the outer and inner block" do
          Workbook.unobtrusively(@simple_file1) do |book|
            @old_value = book.sheet(1)[1,1].Value
            book.ReadOnly.should be false
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo" 
            book.Saved.should be false
            book.sheet(1)[1,1].Value.should_not == @old_value
            Workbook.unobtrusively(@simple_file1) do |book2|
              book2.should == book
              book2.ReadOnly.should be false
              book2.Saved.should be false
              book2.sheet(1)[1,1].Value.should_not == @old_value
              book2.sheet(1)[1,1] = book2.sheet(1)[1,1].Value == "foo" ? "bar" : "foo" 
              book2.sheet(1)[1,1].Value.should == @old_value
            end
            book.should be_alive
            book.Saved.should be false
            book.sheet(1)[1,1].Value.should == @old_value            
          end
          book = Workbook.open(@simple_file1)
          book.sheet(1)[1,1].Value.should == @old_value 
        end

        it "should write in the outer and not in the inner block" do
          expect{
          Workbook.unobtrusively(@simple_file1) do |book|
            @old_value = book.sheet(1)[1,1].Value
            book.ReadOnly.should be false
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo" 
            book.Saved.should be false
            book.sheet(1)[1,1].Value.should_not == @old_value
            Workbook.unobtrusively(@simple_file1, :read_only => true) do |book2|
              book2.should == book
              book2.ReadOnly.should be true
              #book2.Saved.should be false
              book2.sheet(1)[1,1].Value.should_not == @old_value
              book2.sheet(1)[1,1] = book2.sheet(1)[1,1].Value == "foo" ? "bar" : "foo" 
              book2.sheet(1)[1,1].Value.should == @old_value
            end
            book.should be_alive
            book.Saved.should be false
            book.sheet(1)[1,1].Value.should_not == @old_value            
          end
          book = Workbook.open(@simple_file1)
          book.sheet(1)[1,1].Value.should_not == @old_value
          }.to raise_error(NotImplementedREOError)
        end

        it "should write in the outer and not in the inner block" do
          Workbook.unobtrusively(@simple_file1) do |book|
            @old_value = book.sheet(1)[1,1].Value
            book.ReadOnly.should be false
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo" 
            book.Saved.should be false
            book.sheet(1)[1,1].Value.should_not == @old_value
            Workbook.unobtrusively(@simple_file1, :writable => false) do |book2|
              book2.should == book
              book2.ReadOnly.should be false
              book2.Saved.should be false
              book2.sheet(1)[1,1].Value.should_not == @old_value
              book2.sheet(1)[1,1] = book2.sheet(1)[1,1].Value == "foo" ? "bar" : "foo" 
              book2.sheet(1)[1,1].Value.should == @old_value
            end
            book.should be_alive
            book.Saved.should be false
            book.sheet(1)[1,1].Value.should == @old_value            
          end
          book = Workbook.open(@simple_file1)
          book.sheet(1)[1,1].Value.should == @old_value
        end

        it "should be read-only in the outer and write in the inner block" do
          Workbook.unobtrusively(@simple_file1, :read_only => true) do |book|
            @old_value = book.sheet(1)[1,1].Value
            book.ReadOnly.should be true
            book.sheet(1)[1,1] = book.sheet(1)[1,1].Value == "foo" ? "bar" : "foo" 
            book.Saved.should be false
            book.sheet(1)[1,1].Value.should_not == @old_value
            Workbook.unobtrusively(@simple_file1) do |book2|
              book2.should == book
              book2.ReadOnly.should be true
              book2.Saved.should be false
              book2.sheet(1)[1,1].Value.should_not == @old_value
              book2.sheet(1)[1,1] = book2.sheet(1)[1,1].Value == "foo" ? "bar" : "foo" 
              book2.sheet(1)[1,1].Value.should == @old_value
            end
            book.should be_alive
            book.Saved.should be false
            book.ReadOnly.should be true
            book.sheet(1)[1,1].Value.should == @old_value            
          end
          book = Workbook.open(@simple_file1)
          book.sheet(1)[1,1].Value.should == @old_value
        end

      end

    end

    describe "unchanging" do

      context "with openess" do

        it "should remain closed" do
          Workbook.unobtrusively(@simple_file) do |book|
          end
          Excel.current.Workbooks.Count.should == 0
        end

        it "should remain open" do
          book1 = Workbook.open(@simple_file1)
          Workbook.unobtrusively(@simple_file1) do |book|
            book.should be_a Workbook
            book.should be_alive
          end
          book1.should be_alive
        end

      end

      context "with writability" do

        it "should remain read_only" do
          book1 = Workbook.open(@simple_file1, :read_only => true)
          Workbook.unobtrusively(@simple_file1) do |book|
          end
          book1.ReadOnly.should be true
        end

        it "should remain writable" do
          book1 = Workbook.open(@simple_file1, :read_only => false)
          Workbook.unobtrusively(@simple_file1) do |book|
          end
          book1.ReadOnly.should be false
        end

        it "should write and remain read_only and open the workbook in another Excel" do
          book1 = Workbook.open(@simple_file1, :read_only => true)
          old_value = book1.sheet(1)[1,1].Value
          Workbook.unobtrusively(@simple_file1) do |book|
            sheet = book.sheet(1)
            sheet[1,1] = sheet[1,1].Value == "foo" ? "bar" : "foo"
            book.excel.should == book1.excel
          end
          book1.ReadOnly.should be true
          book1.close
          book2 = Workbook.open(@simple_file1)
          book2.sheet(1)[1,1].Value.should == old_value
        end

        it "should write and remain read_only and open the workbook in the same Excel" do
          book1 = Workbook.open(@simple_file1, :read_only => true)
          old_value = book1.sheet(1)[1,1].Value
          Workbook.unobtrusively(@simple_file1, :writable => true) do |book|
            book.ReadOnly.should be false
            sheet = book.sheet(1)
            sheet[1,1] = sheet[1,1].Value == "foo" ? "bar" : "foo"
            book.excel.should == book1.excel
          end
          book1.ReadOnly.should be true
          book1.close
          book2 = Workbook.open(@simple_file1)
          book2.sheet(1)[1,1].Value.should_not == old_value
        end

      end

      context "with visibility" do

        it "should remain invisible" do
          book1 = Workbook.open(@simple_file1, :visible => false)
          Workbook.unobtrusively(@simple_file1) do |book|
          end
          book1.excel.Visible.should be false
        end

        it "should remain visible" do
          book1 = Workbook.open(@simple_file1, :visible => true)
          Workbook.unobtrusively(@simple_file1) do |book|
          end
          book1.excel.Visible.should be true
          book1.Windows(book1.Name).Visible.should be true
        end

      end

      context "with check-compatibility" do

        it "should remain check-compatibility false" do
          book1 = Workbook.open(@simple_file1, :check_compatibility => false)
          Workbook.unobtrusively(@simple_file1) do |book|
          end
          book1.CheckCompatibility.should be false
        end

        it "should remain check-compatibility true" do
          book1 = Workbook.open(@simple_file1, :check_compatibility => true)
          Workbook.unobtrusively(@simple_file1) do |book|
          end
          book1.CheckCompatibility.should be true
        end

      end

      context "with calculation" do

        it "should remain the calculation mode" do
          book1 = Workbook.open(@simple_file1)
          old_calculation = book1.excel.calculation
          Workbook.unobtrusively(@simple_file1) do |book|
          end
          book1.excel.calculation.should == old_calculation
        end
      
        it "should remain calculation manual" do
          book1 = Workbook.open(@simple_file1, :calculation => :manual)
          Workbook.unobtrusively(@simple_file1) do |book|
          end
          book1.excel.Calculation.should == XlCalculationManual
        end

        it "should remain calculation automatic" do
          book1 = Workbook.open(@simple_file1, :calculation => :automatic)
          Workbook.unobtrusively(@simple_file1) do |book|
          end
          book1.excel.Calculation.should == XlCalculationAutomatic
        end
       
      end
    end

    context "with no open book" do

      it "should open unobtrusively if no Excel is open" do
        Excel.kill_all
        Workbook.unobtrusively(@simple_file) do |book|
          book.should be_a Workbook
          book.excel.Visible.should be false
          book.CheckCompatibility.should be false
          book.ReadOnly.should be false
        end
      end

      it "should open unobtrusively in a new Excel" do
        expect{ unobtrusively_ok? }.to_not raise_error
      end
    end

    context "with two running excel instances" do
      
      before :all do
        Excel.kill_all
      end

      before do
        @excel1 = Excel.new(:reuse => false)
        @excel2 = Excel.new(:reuse => false)
      end

      after do
        Excel.kill_all
      end

      it "should open unobtrusively in the first opened Excel" do
        Workbook.unobtrusively(@simple_file) do |book|
          book.excel.should     == @excel1
          book.excel.should_not == @excel2
        end
      end

      it "should open unobtrusively in a new Excel" do
        Workbook.unobtrusively(@simple_file, :if_closed => :new) do |book|
          book.should be_a Workbook
          book.should be_alive
          book.excel.should_not == @excel1
          book.excel.should_not == @excel2
        end
      end

      it "should open unobtrusively in a given Excel" do
        Workbook.unobtrusively(@simple_file, :if_closed => @excel2) do |book|
          book.should be_a Workbook
          book.should be_alive
          book.excel.should_not == @excel1
          book.excel.should     == @excel2
        end
      end

      it "should open in another Excel instance if the given Excel instance is not alive" do
        @excel1.close
        sleep 2
        expect{
          Workbook.unobtrusively(@simple_file, :if_closed => @excel2) do |book|
            book.should be_alive
            book.excel.should == @excel2
          end
        }.to_not raise_error
      end

      it "should raise an error if the option is invalid" do
        expect{
          Workbook.unobtrusively(@simple_file, :if_closed => :invalid_option) do |book|
          end
        }.to raise_error(TypeREOError, "given object is neither an Excel, a Workbook, nor a Win32ole")
      end

    end

    context "with an open book" do

      before do
        @book = Workbook.open(@simple_file1)
      end

      after do
        @book.close(:if_unsaved => :forget)
        @book2.close(:if_unsaved => :forget) rescue nil
      end

      it "should let an open Workbook open if it has been closed and opened again" do
        @book.close
        @book.reopen
        Workbook.unobtrusively(@simple_file1) do |book|
          book.should be_a Workbook
          book.should be_alive
          book.excel.should == @book.excel
        end        
        @book.should be_alive
        @book.should be_a Workbook
      end

      it "should let an open Workbook open if two books have been opened and one has been closed and opened again" do
        book2 = Workbook.open(@different_file, :force_excel => :new)
        @book.close
        book2.close
        @book.reopen
        Workbook.unobtrusively(@simple_file1) do |book|
          book.should be_a Workbook
          book.should be_alive
          book.excel.should == @book.excel
        end        
        @book.should be_alive
        @book.should be_a Workbook
      end

      it "should open in the Excel of the given Workbook" do
        @book2 = Workbook.open(@another_simple_file, :force_excel => :new)
        Workbook.unobtrusively(@different_file, :if_closed => @book2) do |book|
          book.should be_a Workbook
          book.should be_alive
          book.excel.should_not == @book.excel
          book.excel.should     == @book2.excel
        end
      end

      it "should let a saved book saved" do
        @book.Saved.should be true
        @book.should be_alive
        sheet = @book.sheet(1)
        old_cell_value = sheet[1,1].Value
        unobtrusively_ok?
        @book.Saved.should be true
        @book.should be_alive
        sheet = @book.sheet(1)
        sheet[1,1].Value.should_not == old_cell_value
      end

     it "should let the unsaved book unsaved" do
        sheet = @book.sheet(1)
        sheet[1,1] = sheet[1,1].Value == "foo" ? "bar" : "foo"
        old_cell_value = sheet[1,1].Value
        @book.Saved.should be false
        unobtrusively_ok?
        @book.should be_alive
        @book.Saved.should be false
        @book.close(:if_unsaved => :forget)
        @book2 = Workbook.open(@simple_file1)
        sheet2 = @book2.sheet(1)
        sheet2[1,1].Value.should_not == old_cell_value
      end

      it "should modify unobtrusively the second, writable book" do
        @book2 = Workbook.open(@simple_file1, :force_excel => :new)
        @book.ReadOnly.should be false
        @book2.ReadOnly.should be true
        sheet = @book2.sheet(1)
        old_cell_value = sheet[1,1].Value
        sheet[1,1] = sheet[1,1].Value == "foo" ? "bar" : "foo"
        unobtrusively_ok?
        @book2.should be_alive
        @book2.Saved.should be false
        @book2.close(:if_unsaved => :forget)
        @book.close
        new_book = Workbook.open(@simple_file1)
        sheet2 = new_book.sheet(1)
        sheet2[1,1].Value.should_not == old_cell_value
      end    
    end


    context "with a closed book" do
      
      before do
        @book = Workbook.open(@simple_file)
      end

      after do
        @book.close(:if_unsaved => :forget)
      end

      it "should let the closed book closed by default" do
        sheet = @book.sheet(1)
        old_cell_value = sheet[1,1].Value
        @book.close
        @book.should_not be_alive
        unobtrusively_ok?
        @book.should_not be_alive
        new_book = Workbook.open(@simple_file1)
        sheet = new_book.sheet(1)
        sheet[1,1].Value.should_not == old_cell_value
      end

      # The bold reanimation of the @book
      it "should use the excel of the book and keep open the book" do
        excel = Excel.new(:reuse => false)
        sheet = @book.sheet(1)
        old_cell_value = sheet[1,1].Value
        @book.close
        @book.should_not be_alive
        Workbook.unobtrusively(@simple_file, :keep_open => true) do |book|
          book.should be_a Workbook
          book.excel.should == @book.excel
          book.excel.should_not == excel
          sheet = book.sheet(1)
          cell = sheet[1,1]
          sheet[1,1] = cell.Value == "foo" ? "bar" : "foo"
          book.Saved.should be false
        end
        @book.should be_alive
        @book.close
        new_book = Workbook.open(@simple_file1)
        sheet = new_book.sheet(1)
        sheet[1,1].Value.should_not == old_cell_value
      end

      # book shall be reanimated
      it "should use the excel of the book and keep open the book" do
        excel = Excel.new(:reuse => false)
        sheet = @book.sheet(1)
        old_cell_value = sheet[1,1].Value
        @book.close
        @book.should_not be_alive
        Workbook.unobtrusively(@simple_file, :if_closed => :new) do |book|
          book.should be_a Workbook
          book.should be_alive
          book.excel.should_not == @book.excel
          book.excel.should_not == excel
          sheet = book.sheet(1)
          cell = sheet[1,1]
          sheet[1,1] = cell.Value == "foo" ? "bar" : "foo"
          book.Saved.should be false
        end
        @book.should_not be_alive
        new_book = Workbook.open(@simple_file1)
        sheet = new_book.sheet(1)
        sheet[1,1].Value.should_not == old_cell_value
      end

      it "should use another excel if the Excels are closed" do
        sheet = @book.sheet(1)
        old_cell_value = sheet[1,1].Value
        @book.close
        @book.should_not be_alive
        Workbook.unobtrusively(@simple_file, :keep_open => true) do |book|
          book.should be_a Workbook
          sheet = book.sheet(1)
          cell = sheet[1,1]
          sheet[1,1] = cell.Value == "foo" ? "bar" : "foo"
          book.Saved.should be false
        end
        @book.should be_alive
        @book.close
        new_book = Workbook.open(@simple_file1)
        sheet = new_book.sheet(1)
        sheet[1,1].Value.should_not == old_cell_value
      end

      it "should use another excel if the Excels are closed" do
        excel = Excel.new(:reuse => false)
        sheet = @book.sheet(1)
        old_cell_value = sheet[1,1].Value
        @book.close
        @book.should_not be_alive
        Excel.kill_all
        Workbook.unobtrusively(@simple_file, :if_closed => :new, :keep_open => true) do |book|
          book.should be_a Workbook
          book.excel.should_not == @book.excel
          book.excel.should_not == excel
          sheet = book.sheet(1)
          cell = sheet[1,1]
          sheet[1,1] = cell.Value == "foo" ? "bar" : "foo"
          book.Saved.should be false
        end
        @book.should_not be_alive
        new_book = Workbook.open(@simple_file1)
        sheet = new_book.sheet(1)
        sheet[1,1].Value.should_not == old_cell_value
      end      

      it "should modify unobtrusively the copied file" do
        sheet = @book.sheet(1)
        old_cell_value = sheet[1,1].Value
        File.delete simple_save_file rescue nil
        @book.save_as(@simple_save_file)
        @book.close
        Workbook.unobtrusively(@simple_save_file) do |book|
          sheet = book.sheet(1)
          cell = sheet[1,1]
          sheet[1,1] = cell.Value == "foo" ? "bar" : "foo"
        end
        old_book = Workbook.open(@simple_file1)
        old_sheet = old_book.sheet(1)
        old_sheet[1,1].Value.should == old_cell_value
        old_book.close
        new_book = Workbook.open(@simple_save_file)
        new_sheet = new_book.sheet(1)
        new_sheet[1,1].Value.should_not == old_cell_value
        new_book.close
      end
    end

    context "with a visible book" do

      before do
        @book = Workbook.open(@simple_file1, :visible => true)
      end

      after do
        @book.close(:if_unsaved => :forget)
        @book2.close(:if_unsaved => :forget) rescue nil
      end

      it "should let an open Workbook open" do
        Workbook.unobtrusively(@simple_file1) do |book|
          book.should be_a Workbook
          book.should be_alive
          book.excel.should == @book.excel
          book.excel.Visible.should be true
        end        
        @book.should be_alive
        @book.should be_a Workbook
        @book.excel.Visible.should be true
      end
      
    end

    context "with various options for an Excel instance in which to open a closed book" do

      before do
        @book = Workbook.open(@simple_file1)
        @book.close
      end

      it "should use a given Excel" do
        new_excel = Excel.new(:reuse => false)
        another_excel = Excel.new(:reuse => false)
        Workbook.unobtrusively(@simple_file1, :if_closed => another_excel) do |book|
          book.excel.should_not == @book.excel
          book.excel.should_not == new_excel
          book.excel.should == another_excel
        end
      end

      it "should use another Excel" do
        new_excel = Excel.new(:reuse => false)
        Workbook.unobtrusively(@simple_file1, :if_closed => :new) do |book|
          book.excel.should_not == @book.excel
          book.excel.should_not == new_excel
          book.excel.visible.should be false
          book.excel.displayalerts.should == :if_visible
          @another_excel = book.excel
        end
        Workbook.unobtrusively(@simple_file1, :if_closed => :current) do |book|
          book.excel.should_not == @book.excel
          book.excel.should_not == new_excel
          book.excel.should == @another_excel
          book.excel.visible.should be false
          book.excel.displayalerts.should == :if_visible
        end
      end

      it "should reuse Excel" do
        new_excel = Excel.new(:reuse => false)
        Workbook.unobtrusively(@simple_file1, :if_closed => :current) do |book|
          book.excel.should == @book.excel
          book.excel.should_not == new_excel
        end
      end

      it "should reuse Excel by default" do
        new_excel = Excel.new(:reuse => false)
        Workbook.unobtrusively(@simple_file1) do |book|
          book.excel.should == @book.excel
          book.excel.should_not == new_excel
        end
      end

    end

    context "with a read_only book" do

      before do
        @book = Workbook.open(@simple_file1, :read_only => true)
      end

      after do
        @book.close
      end

      it "should let the saved book saved" do
        @book.ReadOnly.should be true
        @book.Saved.should be true
        sheet = @book.sheet(1)
        old_cell_value = sheet[1,1].Value
        unobtrusively_ok?
        @book.should be_alive
        @book.Saved.should be true
        @book.ReadOnly.should be true
        @book.close
        book2 = Workbook.open(@simple_file1)
        sheet2 = book2.sheet(1)
        sheet2[1,1].Value.should == old_cell_value
      end

      it "should let the unsaved book unsaved" do
        @book.ReadOnly.should be true
        sheet = @book.sheet(1)
        sheet[1,1] = sheet[1,1].Value == "foo" ? "bar" : "foo" 
        @book.Saved.should be false
        @book.should be_alive
        Workbook.unobtrusively(@simple_file1) do |book|
          book.should be_a Workbook
          sheet = book.sheet(1)
          sheet[1,1] = sheet[1,1].Value == "foo" ? "bar" : "foo"
          @cell_value = sheet[1,1].Value
          book.should be_alive
          book.Saved.should be false
        end
        @book.should be_alive
        @book.Saved.should be false
        @book.ReadOnly.should be true
        @book.close
        book2 = Workbook.open(@simple_file1)
        sheet2 = book2.sheet(1)
        # modifies unobtrusively the saved version, not the unsaved version
        sheet2[1,1].Value.should == @cell_value        
      end

      it "should open unobtrusively by default the writable book" do
        book2 = Workbook.open(@simple_file1, :force_excel => :new, :read_only => false)
        @book.ReadOnly.should be true
        book2.Readonly.should be false
        sheet = @book.sheet(1)
        cell_value = sheet[1,1].Value
        Workbook.unobtrusively(@simple_file1, :if_closed => :new) do |book|
          book.should be_a Workbook
          book.excel.should == book2.excel
          book.excel.should_not == @book.excel
          sheet = book.sheet(1)
          sheet[1,1] = sheet[1,1].Value == "foo" ? "bar" : "foo"
          book.should be_alive
          book.Saved.should be false          
        end  
        @book.Saved.should be true
        @book.ReadOnly.should be true
        @book.close
        book2.close
        book3 = Workbook.open(@simple_file1)
        new_sheet = book3.sheet(1)
        new_sheet[1,1].Value.should_not == cell_value
        book3.close
      end

=begin
      it "should open unobtrusively the book in a new Excel to open the book writable" do
        excel1 = Excel.new(:reuse => false)
        excel2 = Excel.new(:reuse => false)
        book2 = Workbook.open(@simple_file1, :force_excel => :new, :read_only => true)
        @book.ReadOnly.should be true
        book2.Readonly.should be true
        sheet = @book.sheet(1)
        cell_value = sheet[1,1].Value
        Workbook.unobtrusively(@simple_file1, :writable => true, :if_closed => :new, :rw_change_excel => :new) do |book|
          book.should be_a Workbook
          book.ReadOnly.should be false
          book.excel.should_not == book2.excel
          book.excel.should_not == @book.excel
          book.excel.should_not == excel1
          book.excel.should_not == excel2
          sheet = book.sheet(1)
          sheet[1,1] = sheet[1,1].Value == "foo" ? "bar" : "foo"
          book.should be_alive
          book.Saved.should be false          
        end  
        @book.Saved.should be true
        @book.ReadOnly.should be true
        @book.close
        book2.close
        book3 = Workbook.open(@simple_file1)
        new_sheet = book3.sheet(1)
        new_sheet[1,1].Value.should_not == cell_value
        book3.close
      end

      it "should open unobtrusively the book in the same Excel to open the book writable" do
        excel1 = Excel.new(:reuse => false)
        excel2 = Excel.new(:reuse => false)
        book2 = Workbook.open(@simple_file1, :force_excel => :new, :read_only => true)
        @book.ReadOnly.should be true
        book2.Readonly.should be true
        sheet = @book.sheet(1)
        cell_value = sheet[1,1].Value
        Workbook.unobtrusively(@simple_file1, :writable => true, :if_closed => :new, :rw_change_excel => :current) do |book|
          book.should be_a Workbook
          book.excel.should == book2.excel
          book.ReadOnly.should be false
          sheet = book.sheet(1)
          sheet[1,1] = sheet[1,1].Value == "foo" ? "bar" : "foo"
          book.should be_alive
          book.Saved.should be false          
        end  
        book2.Saved.should be true
        book2.ReadOnly.should be true
        @book.close
        book2.close
        book3 = Workbook.open(@simple_file1)
        new_sheet = book3.sheet(1)
        new_sheet[1,1].Value.should_not == cell_value
        book3.close
      end
=end

      it "should open unobtrusively the book in the Excel where it was opened most recently" do
        book2 = Workbook.open(@simple_file1, :force_excel => :new, :read_only => true)
        @book.ReadOnly.should be true
        book2.Readonly.should be true
        sheet = @book.sheet(1)
        cell_value = sheet[1,1].Value
        Workbook.unobtrusively(@simple_file1, :if_closed => :new, :read_only => true) do |book|
          book.should be_a Workbook
          book.excel.should == book2.excel
          book.excel.should_not == @book.excel
          book.should be_alive
          book.Saved.should be true         
        end  
        @book.Saved.should be true
        @book.ReadOnly.should be true
        @book.close
        book2.close
      end

    end

    context "with a virgin Workbook class" do
      before do
        class Workbook
          @@bookstore = nil
        end
      end
      it "should work" do
        expect{ unobtrusively_ok? }.to_not raise_error
      end
    end

    context "with a book never opened before" do
      before do
        class Workbook
          @@bookstore = nil
        end
        other_book = Workbook.open(@different_file)
      end
      it "should open the book" do
        expect{ unobtrusively_ok? }.to_not raise_error
      end
    end

    context "with a saved book" do

      before do
        @book1 = Workbook.open(@simple_file1)
      end

      after do
        @book1.close(:if_unsaved => :forget)
      end

      it "should save if the book was modified during unobtrusively" do
        m_time = File.mtime(@book1.stored_filename)
        Workbook.unobtrusively(@simple_file1, :if_closed => :new) do |book|
          @book1.Saved.should be true
          book.Saved.should be true  
          sheet = book.sheet(1)
          cell = sheet[1,1]
          sheet[1,1] = cell.Value == "foo" ? "bar" : "foo"
          @book1.Saved.should be false
          book.Saved.should be false
          sleep 1
        end
        @book1.Saved.should be true
        m_time2 = File.mtime(@book1.stored_filename)
        m_time2.should_not == m_time
      end      

      it "should not save the book if it was not modified during unobtrusively" do
        m_time = File.mtime(@book1.stored_filename)
        Workbook.unobtrusively(@simple_file1) do |book|
          @book1.Saved.should be true
          book.Saved.should be true 
          sleep 1
        end
        m_time2 = File.mtime(@book1.stored_filename)
        m_time2.should == m_time
      end            
    end

    context "with block result" do
      before do
        @book1 = Workbook.open(@simple_file1)
      end

      after do
        @book1.close(:if_unsaved => :forget)
      end      

      it "should yield the block result true" do
        result = 
          Workbook.unobtrusively(@simple_file1) do |book| 
            @book1.Saved.should be true
          end
        result.should == true
      end

      it "should yield the block result nil" do
        result = 
          Workbook.unobtrusively(@simple_file1) do |book| 
          end
        result.should == nil
      end

      it "should yield the block result with an unmodified book" do
        sheet1 = @book1.sheet(1)
        cell1 = sheet1[1,1].Value
        result = 
          Workbook.unobtrusively(@simple_file1) do |book| 
            sheet = book.sheet(1)
            cell = sheet[1,1].Value
          end
        result.should == cell1
      end

      it "should yield the block result even if the book gets saved" do
        sheet1 = @book1.sheet(1)
        @book1.save
        result = 
          Workbook.unobtrusively(@simple_file1) do |book| 
            sheet = book.sheet(1)
            sheet[1,1] = 22
            @book1.Saved.should be false
            42
          end
        result.should == 42
        @book1.Saved.should be true
      end
    end

    context "with several Excel instances" do

      before do
        @book1 = Workbook.open(@simple_file1)
        @book2 = Workbook.open(@simple_file1, :force_excel => :new)
        @book1.Readonly.should == false
        @book2.Readonly.should == true
        old_sheet = @book1.sheet(1)
        @old_cell_value = old_sheet[1,1].Value
        @book1.close
        @book2.close
        @book1.should_not be_alive
        @book2.should_not be_alive
      end

      it "should open unobtrusively the closed book in the most recent Excel where it was open before" do      
        Workbook.unobtrusively(@simple_file1) do |book| 
          book.excel.should == @book2.excel
          book.excel.should_not == @book1.excel
          book.ReadOnly.should == false
          sheet = book.sheet(1)
          cell = sheet[1,1]
          sheet[1,1] = cell.Value == "foo" ? "bar" : "foo"
          book.Saved.should be false
        end
        new_book = Workbook.open(@simple_file1)
        sheet = new_book.sheet(1)
        sheet[1,1].Value.should_not == @old_cell_value
      end

      it "should open unobtrusively the closed book in the new Excel" do
        Workbook.unobtrusively(@simple_file1, :if_closed => :new) do |book| 
          book.excel.should_not == @book2.excel
          book.excel.should_not == @book1.excel
          book.ReadOnly.should == false
          sheet = book.sheet(1)
          cell = sheet[1,1]
          sheet[1,1] = cell.Value == "foo" ? "bar" : "foo"
          book.Saved.should be false
        end
        new_book = Workbook.open(@simple_file1)
        sheet = new_book.sheet(1)
        sheet[1,1].Value.should_not == @old_cell_value
      end

      it "should open unobtrusively the closed book in a new Excel if the Excel is not alive anymore" do
        Excel.kill_all
        Workbook.unobtrusively(@simple_file1, :if_closed => :new) do |book| 
          book.ReadOnly.should == false
          book.excel.should_not == @book1.excel
          book.excel.should_not == @book2.excel
          sheet = book.sheet(1)
          cell = sheet[1,1]
          sheet[1,1] = cell.Value == "foo" ? "bar" : "foo"
          book.Saved.should be false
        end
        new_book = Workbook.open(@simple_file1)
        sheet = new_book.sheet(1)
        sheet[1,1].Value.should_not == @old_cell_value
      end
    end
  end

  describe "promoting" do

    context "with standard" do

      before do
        @book = Workbook.open(@simple_file1)
        @ole_workbook = @book.ole_workbook 
      end

      after do
        @book.close
      end

      it "should promote" do
        Workbook.unobtrusively(@ole_workbook) do |book|
          book.should === @book
          book.equal?(@book).should be true
        end
      end
    
    end

  end

  describe "for_reading, for_modifying" do

    context "open unobtrusively for reading and modifying" do

      before do
        @book = Workbook.open(@simple_file1)
        sheet = @book.sheet(1)
        @old_cell_value = sheet[1,1].Value
        @book.close
      end

      it "should not change the value" do
        Workbook.for_reading(@simple_file1) do |book|
          book.should be_a Workbook
          book.should be_alive
          book.Saved.should be true  
          sheet = book.sheet(1)
          cell = sheet[1,1]
          sheet[1,1] = cell.Value == "foo" ? "bar" : "foo"
          book.Saved.should be false
          book.excel.should == @book.excel
        end
        new_book = Workbook.open(@simple_file1, :visible => true)
        sheet = new_book.sheet(1)
        sheet[1,1].Value.should == @old_cell_value
      end

      it "should not change the value and use a given Excel" do
        new_excel = Excel.new(:reuse => false)
        another_excel = Excel.new(:reuse => false)
        Workbook.for_reading(@simple_file1, :if_closed => another_excel) do |book|
          sheet = book.sheet(1)
          cell = sheet[1,1]
          sheet[1,1] = cell.Value == "foo" ? "bar" : "foo"
          book.excel.should == another_excel
        end
        new_book = Workbook.open(@simple_file1, :visible => true)
        sheet = new_book.sheet(1)
        sheet[1,1].Value.should == @old_cell_value
      end

      it "should not change the value and use the new Excel instance" do
        new_excel = Excel.new(:reuse => false)
        Workbook.for_reading(@simple_file1, :if_closed => :new) do |book|
          sheet = book.sheet(1)
          cell = sheet[1,1]
          sheet[1,1] = cell.Value == "foo" ? "bar" : "foo"
          book.excel.should_not == @book.excel
          book.excel.should_not == new_excel
          book.excel.visible.should be false
          book.excel.displayalerts.should == :if_visible
        end
        new_book = Workbook.open(@simple_file1, :visible => true)
        sheet = new_book.sheet(1)
        sheet[1,1].Value.should == @old_cell_value
      end

      it "should change the value" do
        Workbook.for_modifying(@simple_file1) do |book|
          sheet = book.sheet(1)
          cell = sheet[1,1]
          sheet[1,1] = cell.Value == "foo" ? "bar" : "foo"
          book.excel.should == @book.excel
        end
        new_book = Workbook.open(@simple_file, :visible => true)
        sheet = new_book.sheet(1)
        sheet[1,1].Value.should_not == @old_cell_value
      end

      it "should change the value and use a given Excel" do
        new_excel = Excel.new(:reuse => false)
        another_excel = Excel.new(:reuse => false)
        Workbook.for_modifying(@simple_file1, :if_closed => another_excel) do |book|
          sheet = book.sheet(1)
          cell = sheet[1,1]
          sheet[1,1] = cell.Value == "foo" ? "bar" : "foo"
          book.excel.should == another_excel
        end
        new_book = Workbook.open(@simple_file1, :visible => true)
        sheet = new_book.sheet(1)
        sheet[1,1].Value.should_not == @old_cell_value
      end

      it "should change the value and use the new Excel instance" do
        new_excel = Excel.new(:reuse => false)
        Workbook.for_modifying(@simple_file1, :if_closed => :new) do |book|
          sheet = book.sheet(1)
          cell = sheet[1,1]
          sheet[1,1] = cell.Value == "foo" ? "bar" : "foo"
          book.excel.should_not == @book.excel
          book.excel.should_not == new_excel
          book.excel.visible.should be false
          book.excel.displayalerts.should == :if_visible
        end
        new_book = Workbook.open(@simple_file1, :visible => true)
        sheet = new_book.sheet(1)
        sheet[1,1].Value.should_not == @old_cell_value
      end
    end
  end
  
end
