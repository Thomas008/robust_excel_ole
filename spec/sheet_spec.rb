# -*- coding: utf-8 -*-

require File.join(File.dirname(__FILE__), './spec_helper')

$VERBOSE = nil

include RobustExcelOle
include General

describe Sheet do
 
  before(:all) do
    excel = Excel.new(:reuse => true)
    open_books = excel == nil ? 0 : excel.Workbooks.Count
    puts "*** open books *** : #{open_books}" if open_books > 0
    Excel.close_all
  end 

  before do
    @dir = create_tmpdir
    @simple_file = @dir + '/workbook.xls'
    @protected_file = @dir + '/protected_sheet.xls'
    @blank_file = @dir + '/book_with_blank.xls'
    @merge_file = @dir + '/merge_cells.xls'
    @book = Book.open(@simple_file)
    @sheet = @book.sheet(1)
  end

  after do
    @book.close(:if_unsaved => :forget)
    Excel.kill_all
    rm_tmp(@dir)
  end

  describe ".initialize" do
    context "when open sheet protected(with password is 'protect')" do
      before do
        @key_sender = IO.popen  'ruby "' + File.join(File.dirname(__FILE__), '/helpers/key_sender.rb') + '" "Microsoft Office Excel" '  , "w"
        @book_protect = Book.open(@protected_file, :visible => true, :read_only => true, :force_excel => :new)
        @key_sender.puts "{p}{r}{o}{t}{e}{c}{t}{enter}"
        @protected_sheet = @book_protect['protect']
      end

      after do
        @book_protect.close
        @key_sender.close
      end

      it "should be a protected sheet" do
        @protected_sheet.ProtectContents.should be_true
      end

      it "protected sheet can't be write" do
        expect { @protected_sheet[1,1] = 'write' }.to raise_error
      end
    end

  end

  shared_context "sheet 'open book with blank'" do
    before do
      @book_with_blank = Book.open(@blank_file, :read_only => true)
      @sheet_with_blank = @book_with_blank.sheet(1)
    end

    after do
      @book_with_blank.close
    end
  end

  describe "access sheet name" do
    describe "#name" do
      it 'get sheet1 name' do
        @sheet.name.should eq 'Sheet1'
      end
    end

    describe "#name=" do
      
      it 'change sheet1 name to foo' do
        @sheet.name = 'foo'
        @sheet.name.should eq 'foo'
      end

      it "should raise error when adding the same name" do
        @sheet.name = 'foo'
        @sheet.name.should eq 'foo'
        new_sheet = @book.add_sheet @sheet
        expect{
          new_sheet.name = 'foo'
        }.to raise_error(ExcelErrorSheet, /sheet name "foo" already exists/)
      end
    end
  end

  describe 'access cell' do

    describe "#[]" do      

      context "access [1,1]" do

        it { @sheet[1, 1].should be_kind_of Cell }
        it { @sheet[1, 1].value.should eq 'foo' }
      end

      context "access [1, 1], [1, 2], [3, 1]" do
        it "should get every values" do
          @sheet[1, 1].value.should eq 'foo'
          @sheet[1, 2].value.should eq 'workbook'
          @sheet[3, 1].value.should eq 'matz'
        end
      end

      context "supplying nil as parameter" do
        it "should access [1,1]" do
          @sheet[1, nil].value.should eq 'foo'
          @sheet[nil, 1].value.should eq 'foo'
        end
      end

    end

    it "change a cell to 'bar'" do
      @sheet[1, 1] = 'bar'
      @sheet[1, 1].value.should eq 'bar'
    end

    it "should change a cell to nil" do
      @sheet[1, 1] = nil
      @sheet[1, 1].value.should eq nil
    end

    describe '#each' do
      it "should sort line in order of column" do
        @sheet.each_with_index do |cell, i|
          case i
          when 0
            cell.value.should eq 'foo'
          when 1
            cell.value.should eq 'workbook'
          when 2
            cell.value.should eq 'sheet1'
          when 3
            cell.value.should eq 'foo'
          when 4
            cell.value.should be_nil
          when 5
            cell.value.should eq 'foobaaa'
          end
        end
      end

      context "read sheet with blank" do
        include_context "sheet 'open book with blank'"

        it 'should get from ["A1"]' do
          @sheet_with_blank.each_with_index do |cell, i|
            case i
            when 5
              cell.value.should be_nil
            when 6
              cell.value.should eq 'simple'
            when 7
              cell.value.should be_nil
            when 8
              cell.value.should eq 'workbook'
            when 9
              cell.value.should eq 'sheet1'
            end
          end
        end
      end

    end

    describe "#each_row" do
      it "items should RobustExcelOle::Range" do
        @sheet.each_row do |rows|
          rows.should be_kind_of RobustExcelOle::Range
        end
      end

      context "with argument 1" do
        it 'should read from second row' do
          @sheet.each_row(1) do |rows|
            case rows.row
            when 2
              rows.values.should eq ['foo', nil, 'foobaaa']
            when 3
              rows.values.should eq ['matz', 'is', 'nice']
            end
          end
        end
      end

      context "read sheet with blank" do
        include_context "sheet 'open book with blank'"

        it 'should get from ["A1"]' do
          @sheet_with_blank.each_row do |rows|
            case rows.row - 1
            when 0
              rows.values.should eq [nil, nil, nil, nil, nil]
            when 1
              rows.values.should eq [nil, 'simple', nil, 'workbook', 'sheet1']
            when 2
              rows.values.should eq [nil, 'foo', nil, nil, 'foobaaa']
            when 3
              rows.values.should eq [nil, nil, nil, nil, nil]
            when 4
              rows.values.should eq [nil, 'matz', nil, 'is', 'nice']
            end
          end
        end
      end

    end

    describe "#each_row_with_index" do
      it "should read with index" do
        @sheet.each_row_with_index do |rows, idx|
          case idx
          when 0
            rows.values.should eq ['foo', 'workbook', 'sheet1']
          when 1
            rows.values.should eq ['foo', nil, 'foobaaa']
          when 2
            rows.values.should eq ['matz', 'is', 'nice']
          end
        end
      end

      context "with argument 1" do
        it "should read from second row, index is started 0" do
          @sheet.each_row_with_index(1) do |rows, idx|
            case idx
            when 0
              rows.values.should eq ['foo', nil, 'foobaaa']
            when 1
              rows.values.should eq ['matz', 'is', 'nice']
            end
          end
        end
      end

    end

    describe "#each_column" do
      it "items should RobustExcelOle::Range" do
        @sheet.each_column do |columns|
          columns.should be_kind_of RobustExcelOle::Range
        end
      end

      context "with argument 1" do
        it "should read from second column" do
          @sheet.each_column(1) do |columns|
            case columns.column
            when 2
              columns.values.should eq ['workbook', nil, 'is']
            when 3
              columns.values.should eq ['sheet1', 'foobaaa', 'nice']
            end
          end
        end
      end

      context "read sheet with blank" do
        include_context "sheet 'open book with blank'"

        it 'should get from ["A1"]' do
          @sheet_with_blank.each_column do |columns|
            case columns.column- 1
            when 0
              columns.values.should eq [nil, nil, nil, nil, nil]
            when 1
              columns.values.should eq [nil, 'simple', 'foo', nil, 'matz']
            when 2
              columns.values.should eq [nil, nil, nil, nil, nil]
            when 3
              columns.values.should eq [nil, 'workbook', nil, nil, 'is']
            when 4
              columns.values.should eq [nil, 'sheet1', 'foobaaa', nil, 'nice']
            end
          end
        end
      end

      context "read sheet which last cell is merged" do
        before do
          @book_merge_cells = Book.open(@merge_file)
          @sheet_merge_cell = @book_merge_cells.sheet(1)
        end

        after do
          @book_merge_cells.close
        end

        it "should get from ['A1'] to ['C2']" do
          columns_values = []
          @sheet_merge_cell.each_column do |columns|
            columns_values << columns.values
          end
          columns_values.should eq [
                                [nil, 'first merged', nil, 'merged'],
                                [nil, 'first merged', 'first', 'merged'],
                                [nil, 'first merged', 'second', 'merged'],
                                [nil, nil, 'third', 'merged']
                           ]
        end
      end
    end

    describe "#each_column_with_index" do
      it "should read with index" do
        @sheet.each_column_with_index do |columns, idx|
          case idx
          when 0
            columns.values.should eq ['foo', 'foo', 'matz']
          when 1
            columns.values.should eq ['workbook', nil, 'is']
          when 2
            columns.values.should eq ['sheet1', 'foobaaa', 'nice']
          end
        end
      end

      context "with argument 1" do
        it "should read from second column, index is started 0" do
          @sheet.each_column_with_index(1) do |column_range, idx|
            case idx
            when 0
              column_range.values.should eq ['workbook', nil, 'is']
            when 1
              column_range.values.should eq ['sheet1', 'foobaaa', 'nice']
            end
          end
        end
      end
    end

    describe "#row_range" do
      context "with second argument" do
        before do
          @row_range = @sheet.row_range(1, 2..3)
        end

        it { @row_range.should be_kind_of RobustExcelOle::Range }

        it "should get range cells of second argument" do
          @row_range.values.should eq ['workbook', 'sheet1']
        end
      end

      context "without second argument" do
        before do
          @row_range = @sheet.row_range(3)
        end

        it "should get all cells" do
          @row_range.values.should eq ['matz', 'is', 'nice']
        end
      end

    end

    describe "#col_range" do
      context "with second argument" do
        before do
          @col_range = @sheet.col_range(1, 2..3)
        end

        it { @col_range.should be_kind_of RobustExcelOle::Range }

        it "should get range cells of second argument" do
          @col_range.values.should eq ['foo', 'matz']
        end
      end

      context "without second argument" do
        before do
          @col_range = @sheet.col_range(2)
        end

        it "should get all cells" do
          @col_range.values.should eq ['workbook', nil, 'is']
        end
      end
    end

    describe "rangeval, []" do

      context "getting the value of a range" do
      
        before do
          @book1 = Book.open(@dir + '/another_workbook.xls')
          @sheet1 = @book1.sheet(1)
        end

        after do
          @book1.close
        end   

        it "should return value of a defined name" do
          @sheet1.rangeval("firstcell").should == "foo"
          @sheet1["firstcell"].should == "foo"
        end        

        it "should return default value if name not defined and default value is given" do
          @sheet1.rangeval("foo", :default => 2).should == 2
        end

        it "should evaluate a formula" do
          @sheet1.rangeval("named_formula").should == 4
          @sheet1["named_formula"].should == 4
        end

        it "should raise an error if name not defined" do
          expect {
            @sheet1.rangeval("foo")
          }.to raise_error(SheetError, /cannot evaluate name "foo" in sheet/)
        end
        expect {
            @sheet1["foo"]
          }.to raise_error(SheetError, /cannot evaluate name "foo" in sheet/)
        end
      end
    end

    describe "set_rangeval" do

      context "setting the value of a range" do
      
        before do
          @book1 = Book.open(@dir + '/another_workbook.xls', :read_only => true)
          @sheet1 = @book1.sheet(1)
        end

        after do
          @book1.close(:if_unsaved => :forget)
        end   

        it "should set a range to a value" do
          @sheet1.rangeval("firstcell").should == "foo"
          @sheet1[1,1].Value.should == "foo"
          @sheet1.set_rangeval("firstcell","bar")
          @sheet1.rangeval("firstcell").should == "bar"
          @sheet1[1,1].Value.should == "bar"
          @sheet1["firstcell"] = "foo"
          @sheet1.rangeval("firstcell").should == "foo"
          @sheet1[1,1].Value.should == "foo"
        end

        it "should raise an error if name cannot be evaluated" do
          expect{
            @sheet1.set_nameval("foo", 1)
          }.to raise_error(SheetError, /cannot evaluate name "foo" in sheet/)
          expect{
            @sheet1["foo"].should == 1
          }.to raise_error(SheetError, /cannot evaluate name "foo" in sheet/)
        end
      end
    end


    describe "nameval" do

      context "getting the value of a name" do
      
        before do
          @book1 = Book.open(@dir + '/another_workbook.xls')
          @sheet1 = @book1.sheet(1)
        end

        after do
          @book1.close
        end   

        it "should return value of a defined name" do
          @sheet1.nameval("firstcell").should == "foo"
        end        

        it "should return default value if name not defined and default value is given" do
          @sheet1.nameval("foo", :default => 2).should == 2
        end

        it "should evaluate a formula" do
          @sheet1.nameval("named_formula").should == 4
        end

        it "should raise an error if name not defined" do
          expect {
            @sheet1.nameval("foo")
          }.to raise_error(SheetError, /cannot evaluate name "foo" in sheet/)
        end
      end
    end

    describe "set_nameval" do

      context "setting the value of a name" do
      
        before do
          @book1 = Book.open(@dir + '/another_workbook.xls', :read_only => true)
          @sheet1 = @book1.sheet(1)
        end

        after do
          @book1.close(:if_unsaved => :forget)
        end   

        it "should set a range to a value" do
          @sheet1.nameval("firstcell").should == "foo"
          @sheet1[1,1].Value.should == "foo"
          @sheet1.set_nameval("firstcell","bar")
          @sheet1.nameval("firstcell").should == "bar"
          @sheet1[1,1].Value.should == "bar"
        end

        it "should raise an error if name cannot be evaluated" do
          expect{
            @sheet1.set_nameval("foo", 1)
            }.to raise_error(SheetError, /cannot evaluate name "foo" in sheet/)
        end
      end
    end


          #@sheet1["firstcell"] = "bar"
          #@sheet1.nameval("firstcell").should == "bar"
          #@sheet1[1,1].Value.should == "bar"

    describe "set_name" do

      context "setting the name of a range" do

         before do
          @book1 = Book.open(@dir + '/another_workbook.xls', :read_only => true, :visible => true)
          @sheet1 = @book1.sheet(1)
        end

        after do
          @book1.close
        end   

        it "should name an unnamed range with a giving address" do
          expect{
            @sheet1[1,2].Name.Name
          }.to raise_error          
          @sheet1.set_name("foo",1,2)
          @sheet1[1,2].Name.Name.should == "Sheet1!foo"
        end

        it "should rename an already named range with a giving address" do
          @sheet1[1,1].Name.Name.should == "Sheet1!firstcell"
          @sheet1.set_name("foo",1,1)
          @sheet1[1,1].Name.Name.should == "Sheet1!foo"
        end

        it "should raise an error" do
          expect{
            @sheet1.set_name("foo",-2,1)
          }.to raise_error(SheetError, /cannot add name "foo" to cell with row -2 and column 1/)
        end
      end
    end

    describe "send methods to worksheet" do

      it "should send methods to worksheet" do
        @sheet.Cells(1,1).Value.should eq 'foo'
      end

      it "should raise an error for unknown methods or properties" do
        expect{
          @sheet.Foo
        }.to raise_error(VBAMethodMissingError, /unknown VBA property or method :Foo/)
      end

  #    it "should report that worksheet is not alive" do
  #      @book.close
  #      expect{ @sheet.Nonexisting_method }.to raise_error(ExcelError, "method missing: worksheet nil")
  #    end
    end

    describe "#method_missing" do
      it "can access COM method" do
        @sheet.Cells(1,1).Value.should eq 'foo'
      end

      context "unknown method" do
        it { expect { @sheet.hogehogefoo }.to raise_error }
      end
    end

  end
end
