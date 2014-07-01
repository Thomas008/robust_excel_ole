# -*- coding: utf-8 -*-

require File.join(File.dirname(__FILE__), './spec_helper')

$VERBOSE = nil

describe WrapExcel::Book do
  before do
    @dir = create_tmpdir
    @simple_file = @dir + '/simple.xls'
  end

  after do
    rm_tmp(@dir)
  end

  describe ".open" do
    context "exist file" do
      it "simple file with default" do
        expect {
          book = WrapExcel::Book.open(@simple_file)
          book.close
        }.to_not raise_error
      end

      it "simple file with writable" do
        expect {
          book = WrapExcel::Book.open(@simple_file, :read_only => false)
          book.close
        }.to_not raise_error
      end

      it "simple file with visible = true" do
        expect {
          book = WrapExcel::Book.open(@simple_file, :visible => true)
          book.close
        }.to_not raise_error
      end

      context "with block" do
        it 'block parameter should be instance of WrapExcel::Book' do
          WrapExcel::Book.open(@simple_file) do |book|
            book.should be_is_a WrapExcel::Book
          end
        end
      end
    end

    describe "WIN32OLE#GetAbsolutePathName" do
      it "'~' should be HOME directory" do
        path = '~/Abrakadabra.xlsx'
        expected_path = Regexp.new(File.expand_path(path).gsub(/\//, "."))
        expect {
          WrapExcel::Book.open(path)
        }.to raise_error(WIN32OLERuntimeError, expected_path)
      end
    end

    it 'should not output deprecation warning' do
      capture(:stderr) {
        book = WrapExcel::Book.open(@simple_file)
        book.close
      }.should eq ""
    end

  end

  describe ".new" do
=begin
    it 'should output deprecation warning' do
      capture(:stderr) {
        book = WrapExcel::Book.new(@simple_file)
        book.close
      }.should match /DEPRECATION WARNING: WrapExcel::Book.new and WrapExcel::Book.open will be split. If you open existing file, please use WrapExcel::Book.open.\(call from #{File.expand_path(__FILE__)}:#{__LINE__ - 2}.+\)\n/
    end
=end
  end

  describe 'access sheet' do
    before do
      @book = WrapExcel::Book.open(@simple_file)
    end

    after do
      @book.close
    end

    it 'with sheet name' do
      @book['Sheet1'].should be_kind_of WrapExcel::Sheet
    end

    it 'with integer' do
      @book[0].should be_kind_of WrapExcel::Sheet
    end

    it 'with block' do
      @book.each do |sheet|
        sheet.should be_kind_of WrapExcel::Sheet
      end
    end

    context 'open with block' do
      it {
        WrapExcel::Book.open(@simple_file) do |book|
          book['Sheet1'].should be_is_a WrapExcel::Sheet
        end
      }
    end
  end

  describe "#add_sheet" do
    before do
      @book = WrapExcel::Book.open(@simple_file)
      @sheet = @book[0]
    end

    after do
      @book.close
    end

    context "only first argument" do
      it "should add worksheet" do
        expect { @book.add_sheet @sheet }.to change{ @book.book.Worksheets.Count }.from(3).to(4)
      end

      it "should return copyed sheet" do
        sheet = @book.add_sheet @sheet
        copyed_sheet = @book.book.Worksheets.Item(@book.book.Worksheets.Count)
        sheet.name.should eq copyed_sheet.name
      end
    end

    context "with first argument" do
      context "with second argument is {:as => 'copyed_name'}" do
        it "copyed sheet name should be 'copyed_name'" do
          @book.add_sheet(@sheet, :as => 'copyed_name').name.should eq 'copyed_name'
        end
      end

      context "with second argument is {:before => @sheet}" do
        it "should add the first sheet" do
          @book.add_sheet(@sheet, :before => @sheet).name.should eq @book[0].name
        end
      end

      context "with second argument is {:after => @sheet}" do
        it "should add the first sheet" do
          @book.add_sheet(@sheet, :after => @sheet).name.should eq @book[1].name
        end
      end

      context "with second argument is {:before => @book[2], :after => @sheet}" do
        it "should arguments in the first is given priority" do
          @book.add_sheet(@sheet, :before => @book[2], :after => @sheet).name.should eq @book[2].name
        end
      end

    end

    context "without first argument" do
      context "second argument is {:as => 'new sheet'}" do
        it "should return new sheet" do
          @book.add_sheet(:as => 'new sheet').name.should eq 'new sheet'
        end
      end

      context "second argument is {:before => @sheet}" do
        it "should add the first sheet" do
          @book.add_sheet(:before => @sheet).name.should eq @book[0].name
        end
      end

      context "second argument is {:after => @sheet}" do
        it "should add the second sheet" do
          @book.add_sheet(:after => @sheet).name.should eq @book[1].name
        end
      end

    end

    context "without argument" do
      it "should add empty sheet" do
        expect { @book.add_sheet }.to change{ @book.book.Worksheets.Count }.from(3).to(4)
      end

      it "shoule return copyed sheet" do
        sheet = @book.add_sheet
        copyed_sheet = @book.book.Worksheets.Item(@book.book.Worksheets.Count)
        sheet.name.should eq copyed_sheet.name
      end
    end
  end

  describe "#save" do
    context "when open with read only" do
      before do
        @book = WrapExcel::Book.open(@simple_file)
      end

      it {
        expect {
          @book.save
        }.to raise_error(IOError,
                     "Not opened for writing(open with :read_only option)")
      }
    end

    context "with argument" do
      before do
        WrapExcel::Book.open(@simple_file, :read_only => false) do |book|
          book.save("#{@dir}/simple_save.xlsx")
        end
      end

      it "should save to 'simple_save.xlsx'" do
        File.exist?(@dir + "/simple_save.xlsx").should be_true
      end
    end

    context "with file name" do
      before do
        @book = WrapExcel::Book.open(@simple_file, :read_only => false) 
      end
      
      after do
        @book.close
      end

      it "should save to 'simple_save.xlsm'" do
        save_path = "C:" + "/" + "simple_save.xlsm"
        p save_path
        File.delete save_path rescue nil
        @book.save(save_path)
        File.exist?(save_path).should be_true
        book_neu = WrapExcel::Book.open(save_path, :read_only => true) 
        book_neu.should be_a WrapExcel::Book
        book_neu.close
      end

      it "should save to 'simple_save.xlsx'" do
        save_path = "C:" + "/" + "simple_save.xlsx"
        p save_path
        File.delete save_path rescue nil
        @book.save(save_path)
        File.exist?(save_path).should be_true
        book_neu = WrapExcel::Book.open(save_path, :read_only => true) 
        book_neu.should be_a WrapExcel::Book
        book_neu.close
      end

      it "should save to 'simple_save.xls'" do
        save_path = "C:" + "/" + "simple_save.xls"
        p save_path
        File.delete save_path rescue nil
        @book.save(save_path)
        File.exist?(save_path).should be_true
        book_neu = WrapExcel::Book.open(save_path, :read_only => true) 
        book_neu.should be_a WrapExcel::Book
        book_neu.close
      end
    end

    context "save with options"
      before do
        @book = WrapExcel::Book.open(@simple_file, :read_only => false) 
      end
      
      after do
        @book.close
      end

      it "should save to 'simple_save.xlsm' with overwrite" do
        save_path = "C:" + "/" + "simple_save.xlsm"
        p save_path
        File.delete save_path rescue nil
        @book.save(save_path)
        @book.save(save_path, :if_exists => :overwrite)
        File.exist?(save_path).should be_true
        book_neu = WrapExcel::Book.open(save_path, :read_only => true) 
        book_neu.should be_a WrapExcel::Book
        book_neu.close
      end
      it "should save to 'simple_save.xlsm' with excel" do
        save_path = "C:" + "/" + "simple_save.xlsm"
        p save_path
        File.delete save_path rescue nil
        @book.save(save_path)
        @book.save(save_path, :if_exists => :excel )
        File.exist?(save_path).should be_true
        book_neu = WrapExcel::Book.open(save_path, :read_only => true) 
        book_neu.should be_a WrapExcel::Book
        book_neu.close
      end
      it "should save to 'simple_save.xlsm' with raise" do
        save_path = "C:" + "/" + "simple_save.xlsm"
        p save_path
        File.delete save_path rescue nil
        @book.save(save_path)
        @book.save(save_path, :if_exists => :raise) rescue nil
        File.exist?(save_path).should be_true
        book_neu = WrapExcel::Book.open(save_path, :read_only => true) 
        book_neu.should be_a WrapExcel::Book
        book_neu.close
      end
  end
end
