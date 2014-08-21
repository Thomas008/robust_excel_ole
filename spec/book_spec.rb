# -*- coding: utf-8 -*-

require File.join(File.dirname(__FILE__), './spec_helper')



$VERBOSE = nil

describe RobustExcelOle::Book do
  before do
    @dir = create_tmpdir
    @simple_file = @dir + '/simple.xls'
  end

  after do
    rm_tmp(@dir)
  end

  save_path = "C:" + "/" + "simple_save.xls"

  context "class methods" do
    context "create file" do
      it "simple file with default" do
        expect {
          book = RobustExcelOle::Book.new(@simple_file)
          book.close
        }.to_not raise_error
      end
    end
  end

  describe "open" do

    after do
      RobustExcelOle::ExcelApp.close_all
    end

    context "with non-existing file" do
      it "should raise an exception" do
        File.delete save_path rescue nil
        expect {
          RobustExcelOle::Book.open(save_path)
        }.to raise_error(ExcelErrorOpen, "file #{save_path} not found")
      end
    end

    context "with standard options" do
      before do
        @book = RobustExcelOle::Book.open(@simple_file)
      end

      after do
        @book.close
      end

      it "should say that it lives" do
        @book.alive?.should be_true
      end      
    end

    context "with excel_app" do
      before do
        @new_book = RobustExcelOle::Book.open(@simple_file)
      end
      after do
        @new_book.close
      end
      it "should provide the excel application of the book" do
        excel_app = @new_book.excel_app
        excel_app.class.should == RobustExcelOle::ExcelApp
        excel_app.should be_a RobustExcelOle::ExcelApp
      end
    end


    context "with options writable, visible" do
      it "simple file with writable" do
        expect {
          book = RobustExcelOle::Book.open(@simple_file, :read_only => false)
          book.close
        }.to_not raise_error
      end

      it "simple file with visible = true" do
        expect {
          book = RobustExcelOle::Book.open(@simple_file, :visible => true)
          book.close
        }.to_not raise_error
      end
    end

    context "with block" do
      it 'block parameter should be instance of RobustExcelOle::Book' do
        RobustExcelOle::Book.open(@simple_file) do |book|
          book.should be_is_a RobustExcelOle::Book
        end
      end
    end

    context "with WIN32OLE#GetAbsolutePathName" do
      it "'~' should be HOME directory" do
        path = '~/Abrakadabra.xlsx'
        expected_path = Regexp.new(File.expand_path(path).gsub(/\//, "."))
        expect {
          RobustExcelOle::Book.open(path)
        }.to raise_error(ExcelErrorOpen, "file #{path} not found")
      end
    end

    context "with an already opened and saved book" do

      before do
        @book = RobustExcelOle::Book.open(@simple_file, :read_only => false)
      end

      after do
        @book.close
      end

      possible_options = [:read_only, :raise, :accept, :forget, nil]
      possible_options.each do |options_value|        
        context "with :if_unsaved => #{options_value}" do
          before do
            @new_book = RobustExcelOle::Book.open(@simple_file, :reuse=> true, :if_unsaved => options_value)
          end
          after do
            @new_book.close
          end
          it "should open without problems " do
              @new_book.should be_a RobustExcelOle::Book
          end
          it "should belong to the same Excel application" do
            @new_book.excel_app.should == @book.excel_app
          end
        end
      end
    end

    context "with an already opened book that is not saved" do

      before do
        @book = RobustExcelOle::Book.open(@simple_file, :read_only => false)
        # mappe ändern
        @sheet = @book[0]
        @book.add_sheet(@sheet, :as => 'copyed_name')
      end

      after do
        @book.close
      end

      it "if_unsaved is :raise" do
        expect {
          new_book = RobustExcelOle::Book.open(@simple_file, :if_unsaved => :raise)
          new_book.close
           }.to raise_error(ExcelErrorOpen, "book is already open but not saved (#{File.basename(@simple_file)})")
        #new_book sollte kein Buch sein
        #new_book.should.not be_is_a RobustExcelOle::Book
        #oder: expect{new_book.close}.to raise_error
      end
    end
  end

  describe 'access sheet' do
    before do
      @book = RobustExcelOle::Book.open(@simple_file)
    end

    after do
      @book.close
    end

    it 'with sheet name' do
      @book['Sheet1'].should be_kind_of RobustExcelOle::Sheet
    end

    it 'with integer' do
      @book[0].should be_kind_of RobustExcelOle::Sheet
    end

    it 'with block' do
      @book.each do |sheet|
        sheet.should be_kind_of RobustExcelOle::Sheet
      end
    end

    context 'open with block' do
      it {
        RobustExcelOle::Book.open(@simple_file) do |book|
          book['Sheet1'].should be_is_a RobustExcelOle::Sheet
        end
      }
    end
  end

  describe "#add_sheet" do
    before do
      @book = RobustExcelOle::Book.open(@simple_file)
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

      it "should return copyed sheet" do
        sheet = @book.add_sheet
        copyed_sheet = @book.book.Worksheets.Item(@book.book.Worksheets.Count)
        sheet.name.should eq copyed_sheet.name
      end
    end
  end

  describe "save" do
    context "when open with read only" do
      before do
        @book = RobustExcelOle::Book.open(@simple_file)
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
        RobustExcelOle::Book.open(@simple_file, :read_only => false) do |book|
          book.save("#{@dir}/simple_save.xlsx", :if_exists => :overwrite)
        end
      end

      it "should save to 'simple_save.xlsx'" do
        File.exist?(@dir + "/simple_save.xlsx").should be_true
      end
    end

    context "with different extensions" do
      before do
        @book = RobustExcelOle::Book.open(@simple_file, :read_only => false)
      end

      after do
        @book.close
      end

      possible_extensions = ["xls", "xlsm", "xlsx"]
      possible_extensions.each do |extensions_value|
        it "should save to 'simple_save.#{extensions_value}'" do
          save_path = "C:" + "/" + "simple_save." + extensions_value
          File.delete save_path rescue nil
          @book.save(save_path, :if_exists => :overwrite)
          File.exist?(save_path).should be_true
          new_book = RobustExcelOle::Book.open(save_path, :read_only => true)
          new_book.should be_a RobustExcelOle::Book
          new_book.close
        end
      end
    end

    # options :overwrite, :raise, no option, invalid option
    possible_displayalerts = [true, false]
    possible_displayalerts.each do |displayalert_value|
      context "with options displayalerts=#{displayalert_value}" do
        before do
          @book = RobustExcelOle::Book.open(@simple_file, :read_only => false, :displayalerts => displayalert_value)
        end

        after do
          @book.close
        end

        it "should save to 'simple_save.xlsm' with :overwrite" do
          File.delete save_path rescue nil
          File.open(save_path,"w") do | file |
            file.puts "garbage"
          end
          @book.save(save_path, :if_exists => :overwrite)
          File.exist?(save_path).should be_true
          new_book = RobustExcelOle::Book.open(save_path, :read_only => true)
          new_book.should be_a RobustExcelOle::Book
          new_book.close
        end

        it "should save to 'simple_save.xlsm' with :raise" do
          dirname, basename = File.split(save_path)
          File.delete save_path rescue nil
          File.open(save_path,"w") do | file |
            file.puts "garbage"
          end
          File.exist?(save_path).should be_true
          booklength = File.size?(save_path)
          expect {
            @book.save(save_path, :if_exists => :raise)
            }.to raise_error(ExcelErrorSave, 'book already exists: ' + basename)
          File.exist?(save_path).should be_true
          (File.size?(save_path) == booklength).should be_true
        end

        it "should save to 'simple_save.xlsm' with no option" do
          dirname, basename = File.split(save_path)
          File.delete save_path rescue nil
          File.open(save_path,"w") do | file |
            file.puts "garbage"
          end
          File.exist?(save_path).should be_true
          booklength = File.size?(save_path)
          expect {
            @book.save(save_path)
            }.to raise_error(ExcelErrorSave, 'book already exists: ' + basename)
          File.exist?(save_path).should be_true
          (File.size?(save_path) == booklength).should be_true
        end

        it "should save to 'simple_save.xlsm' with invalid_option" do
          File.delete save_path rescue nil
          @book.save(save_path)
          expect {
            @book.save(save_path, :if_exists => :invalid_option)
            }.to raise_error(ExcelErrorSave, 'invalid option (invalid_option)')
        end
      end
    end

    # option :excel
    possible_displayalerts = [false,true]
    possible_displayalerts.each do |displayalert_value|
      context "save with option excel displayalerts=#{displayalert_value}" do
        before do
          @book = RobustExcelOle::Book.open(@simple_file, :read_only => false, :displayalerts => displayalert_value, :visible => false)
        end

        after do
          @book.close
        end

        while not "should save to 'simple_save.xlsm' with excel" do
          File.delete save_path rescue nil
          File.open(save_path,"w") do | file |
            file.puts "garbage"
          end
          @book.save(save_path, :if_exists => :excel)
          File.exist?(save_path).should be_true
          new_book = RobustExcelOle::Book.open(save_path, :read_only => true)
          new_book.should be_a RobustExcelOle::Book
          new_book.close
        end

        context "with if_exists => :excel" do
          before do
            File.delete save_path rescue nil
            File.open(save_path,"w") do | file |
              file.puts "garbage"
            end
            @garbage_length = File.size?(save_path)
            path = File.join(File.dirname(__FILE__), '/helpers/key_sender.rb')
            p "path:#{path}"
            filename = 'C:/key_sender.rb'
            @key_sender = IO.popen  'ruby ' + filename + '  "Microsoft Excel" '  , "w"
            # findet directory nicht - zu lang?
            #@key_sender = IO.popen  'ruby ' + File.join(File.dirname(__FILE__), '/helpers/key_sender.rb') + '  "Microsoft Excel" '  , "w"
          end

          after do
            @key_sender.close
          end

          it "should save if user answers 'yes'" do
            # "Yes" is to the left of "No", which is the  default. --> language independent
            @key_sender.puts "{left}{enter}" #, :initial_wait => 0.2, :if_target_missing=>"Excel window not found")
            @book.save(save_path, :if_exists => :excel)
            File.exist?(save_path).should be_true
            File.size?(save_path).should > @garbage_length
            new_book = RobustExcelOle::Book.open(save_path, :read_only => true)
            new_book.should be_a RobustExcelOle::Book
            new_book.close
          end

# Abfrage
          it "should not save if user answers 'no'" do
            # Just give the "Enter" key, because "No" is the default. --> language independent
            # strangely, in the "no" case, the question will sometimes be repeated three times
            @key_sender.puts "{enter}"
            @key_sender.puts "{enter}"
            @key_sender.puts "{enter}"
            #@key_sender.puts "%{n}" #, :initial_wait => 0.2, :if_target_missing=>"Excel window not found")
            @book.save(save_path, :if_exists => :excel)
            File.exist?(save_path).should be_true
            File.size?(save_path).should == @garbage_length
          end

        end
      end
    end
  end
end
