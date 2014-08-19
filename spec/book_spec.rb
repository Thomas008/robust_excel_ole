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

    context "close excel instances" do
      it "simple file with default" do
        RobustExcelOle::Book.close_all_excel_apps
        expect { WIN32OLE.connect("Excel.Application") }.to raise_error
        #exl_con = WIN32OLE.connect("Excel.Application") rescue nil
        #exl_con.Visible = true
        sleep 1        
        #exl_con.Quit
        #sleep 2

        #exl_con.should be_nil
        #expect { WIN32OLE.connect("Excel.Application") }.to raise_error
        exl1 = WIN32OLE.new("Excel.Application")
        #exl1.Workbooks.Add 
        exl2 = WIN32OLE.new("Excel.Application")
        exl2.Workbooks.Add 
        expect { WIN32OLE.connect("Excel.Application") }.to_not raise_error
        RobustExcelOle::Book.close_all_excel_apps
        sleep 0.3
        expect { WIN32OLE.connect("Excel.Application") }.to raise_error
      end
    end
  end

  describe "open" do

    after do
      RobustExcelOle::Book.close_all_excel_apps
    end

    context "if file does not exist" do
      it "should raise an exception" do
        File.delete save_path rescue nil
        expect {
          RobustExcelOle::Book.open(save_path)
        }.to raise_error(ExcelErrorOpen, "file #{save_path} not found")
      end
    end

    context "if file exists" do
      before do
        @book = RobustExcelOle::Book.open(@simple_file)
      end

      after do
        @book.close
      end

      it "already open" do
        book_neu = RobustExcelOle::Book.open(@simple_file)
        book_neu.close
      end

      it "should say that it lives" do
        @book.alive?.should be_true
      end      
    end

    context "a book is already open and saved" do

      before do
        @book = RobustExcelOle::Book.open(@simple_file, :read_only => false)
      end

      after do
        @book.close
      end

      possible_options = [:read_only, :raise, :accept, :forget, nil]
      possible_options.each do |options_value|        
        it "if_not_saved is #{options_value}" do
          p "option: #{options_value}"
          expect{
            book_neu = RobustExcelOle::Book.open(@simple_file, :if_book_not_saved => options_value)
            #book_neu = RobustExcelOle::Book.open(save_path, :if_book_not_saved => options_value)
            # sollte nicht ein neues Buch öffnen. sollte also KEIN Buch sein!
            book_neu.should be_a RobustExcelOle::Book
            book_neu.close
          }.to_not raise_error
        end
      end
    end

    
    context "a book is already open and not saved" do

      before do
        @book = RobustExcelOle::Book.open(@simple_file, :read_only => false)
        # mappe ändern
        @sheet = @book[0]
        @book.add_sheet(@sheet, :as => 'copyed_name')
      end

      after do
        @book.close
      end

      it "if_not_saved is :raise" do
        expect {
          book_neu = RobustExcelOle::Book.open(@simple_file, :if_not_saved => :raise)
          book_neu.close
           }.to raise_error(ExcelErrorOpen, "book is already open but not saved (#{File.basename(@simple_file)})")
        #book_neu sollte kein Buch sein
        #book_neu.should.not be_is_a RobustExcelOle::Book
        #oder: expect{book_neu.close}.to raise_error
        
      end

    end
  end


  describe ".open" do
    context "exist file" do
      it "simple file with default" do
        expect {
          book = RobustExcelOle::Book.open(@simple_file)
          book.close
        }.to_not raise_error
      end

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

      context "with block" do
        it 'block parameter should be instance of RobustExcelOle::Book' do
          RobustExcelOle::Book.open(@simple_file) do |book|
            book.should be_is_a RobustExcelOle::Book
          end
        end
      end
    end

    describe "WIN32OLE#GetAbsolutePathName" do
      it "'~' should be HOME directory" do
        path = '~/Abrakadabra.xlsx'
        expected_path = Regexp.new(File.expand_path(path).gsub(/\//, "."))
        expect {
          RobustExcelOle::Book.open(path)
        }.to raise_error(ExcelErrorOpen, "file #{path} not found")
      end
    end

    it 'should not output deprecation warning' do
      capture(:stderr) {
        book = RobustExcelOle::Book.open(@simple_file)
        book.close
      }.should eq ""
    end

  end

  describe ".new" do
=begin
    it 'should output deprecation warning' do
      capture(:stderr) {
        book = RobustExcelOle::Book.new(@simple_file)
        book.close
      }.should match /DEPRECATION WARNING: RobustExcelOle::Book.new and RobustExcelOle::Book.open will be split. If you open existing file, please use RobustExcelOle::Book.open.\(call from #{File.expand_path(__FILE__)}:#{__LINE__ - 2}.+\)\n/
    end
=end
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

# errors
  describe "#add_sheet" do
    before do
      @book = RobustExcelOle::Book.open(@simple_file)
      @sheet = @book[0]
    end

    after do
      @book.close
    end
    
    #error
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
          book_neu = RobustExcelOle::Book.open(save_path, :read_only => true)
          book_neu.should be_a RobustExcelOle::Book
          book_neu.close
        end
      end
    end

    # options :overwrite, :raise, :excel, no option, invalid option
    possible_displayalerts = [true, false]
    possible_displayalerts.each do |displayalert_value|
      context "save with options displayalerts=#{displayalert_value}" do
        before do
          @book = RobustExcelOle::Book.open(@simple_file, :read_only => false, :displayalerts => displayalert_value)
        end

        after do
          @book.close
        end

        it "should save to 'simple_save.xlsm' with overwrite" do
          File.delete save_path rescue nil
          File.open(save_path,"w") do | file |
            file.puts "garbage"
          end
          @book.save(save_path, :if_exists => :overwrite)
          File.exist?(save_path).should be_true
          book_neu = RobustExcelOle::Book.open(save_path, :read_only => true)
          book_neu.should be_a RobustExcelOle::Book
          book_neu.close
        end

        it "should save to 'simple_save.xlsm' with raise" do
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

        it "should save to 'simple_save.xlsm' with excel" do
          File.delete save_path rescue nil
          File.open(save_path,"w") do | file |
            file.puts "garbage"
          end
          expect {
            @book.save(save_path, :if_exists => :excel)
            }.to_not raise_error
          File.exist?(save_path).should be_true
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
          book_neu = RobustExcelOle::Book.open(save_path, :read_only => true)
          book_neu.should be_a RobustExcelOle::Book
          book_neu.close
        end

        context "with if_exists => :excel" do
          before do
            File.delete save_path rescue nil
            File.open(save_path,"w") do | file |
              file.puts "garbage"
            end
            @garbage_length = File.size?(save_path)
            @key_sender = IO.popen  'ruby ' + File.join(File.dirname(__FILE__), '/helpers/key_sender.rb') + '  "Microsoft Excel" '  , "w"
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
            book_neu = RobustExcelOle::Book.open(save_path, :read_only => true)
            book_neu.should be_a RobustExcelOle::Book
            book_neu.close
          end

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
