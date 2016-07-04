# -*- coding: utf-8 -*-

require File.join(File.dirname(__FILE__), './../spec_helper')


$VERBOSE = nil

include RobustExcelOle
include General

describe Book do

  before(:all) do
    excel = Excel.new(:reuse => true)
    open_books = excel == nil ? 0 : excel.Workbooks.Count
    puts "*** open books *** : #{open_books}" if open_books > 0
    Excel.close_all
  end

  before do
    @dir = create_tmpdir
    @simple_file = @dir + '/workbook.xls'
    @simple_save_file = @dir + '/workbook_save.xls'
    @different_file = @dir + '/different_workbook.xls'
    @simple_file_other_path = @dir + '/more_data/workbook.xls'
    @another_simple_file = @dir + '/another_workbook.xls'
    @linked_file = @dir + '/workbook_linked.xlsm'
    @simple_file_xlsm = @dir + '/workbook.xls'
    @simple_file_xlsx = @dir + '/workbook.xlsx'
    @simple_file1 = @simple_file
    @simple_file_other_path1 = @simple_file_other_path
    @simple_save_file1 = @simple_save_file
  end

  after do
    Excel.kill_all
    rm_tmp(@dir)
  end

  describe "save" do

    context "with simple save" do
      
      it "should save for a file opened without :read_only" do
        @book = Book.open(@simple_file)
        @book.add_sheet(@sheet, :as => 'a_name')
        @new_sheet_count = @book.ole_workbook.Worksheets.Count
        expect {
          @book.save
        }.to_not raise_error
        @book.ole_workbook.Worksheets.Count.should ==  @new_sheet_count
        @book.close
      end

      it "should raise error with read_only" do
        @book = Book.open(@simple_file, :read_only => true)
        expect {
          @book.save
        }.to raise_error(ExcelErrorSave, "Not opened for writing (opened with :read_only option)")
        @book.close
      end

    end

    context "with open with read only" do
      before do
        @book = Book.open(@simple_file, :read_only => true)
      end

      after do
        @book.close
      end

      it {
        expect {
          @book.save_as(@simple_file)
        }.to raise_error(ExcelErrorSave,
                     "Not opened for writing (opened with :read_only option)")
      }
    end

    context "with argument" do
      before do
        Book.open(@simple_file) do |book|
          book.save_as(@simple_save_file, :if_exists => :overwrite)
        end
      end

      it "should save to 'simple_save_file.xlsx'" do
        File.exist?(@simple_save_file).should be_true
      end
    end

    context "with different extensions" do
      before do
        @book = Book.open(@simple_file)
      end

      after do
        @book.close
      end

      possible_extensions = ['xls', 'xlsm', 'xlsx']
      possible_extensions.each do |extensions_value|
        it "should save to 'simple_save_file.#{extensions_value}'" do
          simple_save_file = @dir + '/simple_save_file.' + extensions_value
          File.delete simple_save_file rescue nil
          @book.save_as(simple_save_file, :if_exists => :overwrite)
          File.exist?(simple_save_file).should be_true
          new_book = Book.open(simple_save_file)
          new_book.should be_a Book
          new_book.close
        end
      end
    end

    context "with blocked by another file" do

      before do
        @book = Book.open(@simple_file1)
        @book2 = Book.open(@another_simple_file)
      end

      after do
        @book.close(:if_unsaved => :forget)
        @book2.close
      end

      it "should raise an error with :obstructed => :raise" do
        File.delete @simple_file_other_path1 rescue nil
        File.open(@simple_file_other_path1,"w") do | file |
          file.puts "garbage"
        end
        File.exist?(@simple_file_other_path1).should be_true
        expect{
          @book2.save_as(@simple_file_other_path1, :if_exists => :overwrite, :if_obstructed => :raise)
        }.to raise_error(ExcelErrorSave, /blocked by another workbook/)
      end

      it "should close the blocking workbook without saving, and save the current workbook with :if_obstructed => :forget" do
        File.delete @simple_file_other_path1 rescue nil
        File.open(@simple_file_other_path1,"w") do | file |
          file.puts "garbage"
        end
        @book2.save_as(@simple_file_other_path1, :if_exists => :overwrite, :if_obstructed => :forget)
        @book.should_not be_alive
        File.exist?(@simple_file_other_path1).should be_true
        new_book = Book.open(@simple_file_other_path1)
        new_book.should be_a Book
        new_book.close
      end

      it "should close the blocking workbook without saving even if it is unsaved with :if_obstructed => :forget" do
        File.delete @simple_file_other_path1 rescue nil
        File.open(@simple_file_other_path1,"w") do | file |
          file.puts "garbage"
        end
        sheet = @book.sheet(1)
        cell_value = sheet[1,1].value
        sheet[1,1] = sheet[1,1].value == "foo" ? "bar" : "foo"
        @book.Saved.should be_false
        @book2.save_as(@simple_file_other_path1, :if_exists => :overwrite, :if_obstructed => :forget)
        @book.should_not be_alive
        @book2.should be_alive
        File.exist?(@simple_file_other_path1).should be_true
        new_book = Book.open(@simple_file_other_path1)
        new_book.should be_a Book
        new_book.close
        old_book = Book.open(@simple_file1)
        old_sheet = old_book.sheet(1)
        old_sheet[1,1].value.should == cell_value
        old_book.close
      end

      it "should save and close the blocking workbook, and save the current workbook with :if_obstructed => :save" do
        File.delete @simple_file_other_path1 rescue nil
        File.open(@simple_file_other_path1,"w") do | file |
          file.puts "garbage"
        end
        sheet = @book.sheet(1)
        cell_value = sheet[1,1].value
        sheet[1,1] = sheet[1,1].value == "foo" ? "bar" : "foo"
        @book.Saved.should be_false
        @book2.save_as(@simple_file_other_path1, :if_exists => :overwrite, :if_obstructed => :save)
        @book.should_not be_alive
        @book2.should be_alive
        File.exist?(@simple_file_other_path1).should be_true
        new_book = Book.open(@simple_file_other_path1)
        new_book.should be_a Book
        new_book.close
        old_book = Book.open(@simple_file1)
        old_sheet = old_book.sheet(1)
        old_sheet[1,1].value.should_not == cell_value
        old_book.close
      end

      it "should close the blocking workbook if it was saved, and save the current workbook with :if_obstructed => :close_if_saved" do
        File.delete @simple_file_other_path1 rescue nil
        File.open(@simple_file_other_path1,"w") do | file |
          file.puts "garbage"
        end
        @book.Saved.should be_true
        @book2.save_as(@simple_file_other_path1, :if_exists => :overwrite, :if_obstructed => :close_if_saved)
        @book.should_not be_alive
        @book2.should be_alive
        File.exist?(@simple_file_other_path1).should be_true
        new_book = Book.open(@simple_file_other_path1)
        new_book.should be_a Book
        new_book.close
      end

      it "should raise an error if the blocking workbook was unsaved with :if_obstructed => :close_if_saved" do
        sheet = @book.sheet(1)
        cell_value = sheet[1,1].value
        sheet[1,1] = sheet[1,1].value == "foo" ? "bar" : "foo"
        @book.Saved.should be_false      
        expect{
          @book2.save_as(@simple_file_other_path1, :if_exists => :overwrite, :if_obstructed => :close_if_saved)
        }.to raise_error(ExcelErrorSave, /blocking workbook is unsaved: "workbook.xls"/)
      end

      it "should raise an error with an invalid option" do
        File.delete @simple_file_other_path1 rescue nil
        File.open(@simple_file_other_path1,"w") do | file |
          file.puts "garbage"
        end
        File.exist?(@simple_file_other_path1).should be_true
        expect{
          @book2.save_as(@simple_file_other_path1, :if_exists => :overwrite, :if_obstructed => :invalid)
        }.to raise_error(ExcelErrorSave, ":if_obstructed: invalid option: :invalid")
      end

      it "should raise an error by default" do
        File.delete @simple_file_other_path1 rescue nil
        File.open(@simple_file_other_path1,"w") do | file |
          file.puts "garbage"
        end
        File.exist?(@simple_file_other_path1).should be_true
        expect{
          @book2.save_as(@simple_file_other_path1, :if_exists => :overwrite)
        }.to raise_error(ExcelErrorSave, /blocked by another workbook/)
      end

      it "should raise an error if the file does not exist and an workbook with the same name and other path exists" do
        File.delete @simple_file_other_path1 rescue nil
        File.exist?(@simple_file_other_path1).should be_false
        expect{
          @book2.save_as(@simple_file_other_path1, :if_exists => :overwrite, :if_obstructed => :raise)
          }.to raise_error(ExcelErrorSave, /blocked by another workbook/)
      end

      it "should raise an error if the file exists and an workbook with the same name and other path exists" do
        File.delete @simple_file_other_path1 rescue nil
        File.open(@simple_file_other_path1,"w") do | file |
          file.puts "garbage"
        end
        File.exist?(@simple_file_other_path1).should be_true
        expect{
          @book.save_as(@simple_file_other_path1, :if_exists => :raise, :if_obstructed => :raise)
        }.to raise_error(ExcelErrorSave, /file already exists: "workbook.xls"/)
      end

    end


    # options :overwrite, :raise, :excel, no option, invalid option
    possible_displayalerts = [true, false]
    possible_displayalerts.each do |displayalert_value|
      context "with displayalerts=#{displayalert_value}" do
        before do
          @book = Book.open(@simple_file, :displayalerts => displayalert_value)
        end

        after do
          @book.close
        end

        it "should raise an error if the book is open" do
          File.delete @simple_save_file1 rescue nil
          FileUtils.copy @simple_file, @simple_save_file1
          book_save = Book.open(@simple_save_file1, :excel => :new)
          expect{
            @book.save_as(@simple_save_file1, :if_exists => :overwrite)
            }.to raise_error(ExcelErrorSave, "workbook is open and used in Excel")
          book_save.close
        end        

        it "should save to simple_save_file.xls with :if_exists => :overwrite" do
          File.delete @simple_save_file1 rescue nil
          File.open(@simple_save_file1,"w") do | file |
            file.puts "garbage"
          end
          @book.save_as(@simple_save_file1, :if_exists => :overwrite)
          File.exist?(@simple_save_file1).should be_true
          new_book = Book.open(@simple_save_file1)
          new_book.should be_a Book
          new_book.close
        end

        it "should simple save if file name is equal to the old one with :if_exists => :overwrite" do
          sheet = @book.sheet(1)
          sheet[1,1] = sheet[1,1].value == "foo" ? "bar" : "foo"
          new_value = sheet[1,1].value
          @book.save_as(@simple_file1, :if_exists => :overwrite)
          new_book = Book.open(@simple_file1)
          new_sheet = new_book.sheet(1)
          new_sheet[1,1].value.should == new_value
          new_book.close
        end

        it "should save to 'simple_save_file.xls' with :if_exists => :raise" do
          dirname, basename = File.split(@simple_save_file)
          File.delete @simple_save_file1 rescue nil
          File.open(@simple_save_file1,"w") do | file |
            file.puts "garbage"
          end
          File.exist?(@simple_save_file1).should be_true
          booklength = File.size?(@simple_save_file1)
          expect {
            @book.save_as(@simple_save_file1, :if_exists => :raise)
            }.to raise_error(ExcelErrorSave, /file already exists: "workbook_save.xls"/)
          File.exist?(@simple_save_file1).should be_true
          File.size?(@simple_save_file1).should == booklength
        end

        context "with :if_exists => :alert" do
          before do
            File.delete @simple_save_file1 rescue nil
            File.open(@simple_save_file1,"w") do | file |
              file.puts "garbage"
            end
            @garbage_length = File.size?(@simple_save_file1)
            @key_sender = IO.popen  'ruby "' + File.join(File.dirname(__FILE__), '../helpers/key_sender.rb') + '" "Microsoft Excel" '  , "w"
          end

          after do
            @key_sender.close
          end

          it "should save if user answers 'yes'" do
            # "Yes" is to the left of "No", which is the  default. --> language independent
            @key_sender.puts "{left}{enter}" #, :initial_wait => 0.2, :if_target_missing=>"Excel window not found")
            @book.save_as(@simple_save_file1, :if_exists => :alert)
            File.exist?(@simple_save_file1).should be_true
            File.size?(@simple_save_file1).should > @garbage_length
            @book.excel.DisplayAlerts.should == displayalert_value
            new_book = Book.open(@simple_save_file1, :excel => :new)
            new_book.should be_a Book
            new_book.close
            @book.excel.DisplayAlerts.should == displayalert_value
          end

          it "should not save if user answers 'no'" do
            # Just give the "Enter" key, because "No" is the default. --> language independent
            # strangely, in the "no" case, the question will sometimes be repeated three times
            @key_sender.puts "{enter}"
            @key_sender.puts "{enter}"
            @key_sender.puts "{enter}"
            #@key_sender.puts "%{n}" #, :initial_wait => 0.2, :if_target_missing=>"Excel window not found")
            expect{
              @book.save_as(@simple_save_file1, :if_exists => :alert)
              }.to raise_error(ExcelErrorSave, "not saved or canceled by user")
            File.exist?(@simple_save_file1).should be_true
            File.size?(@simple_save_file1).should == @garbage_length
            @book.excel.DisplayAlerts.should == displayalert_value
          end

          it "should not save if user answers 'cancel'" do
            # 'Cancel' is right from 'yes'
            # strangely, in the "no" case, the question will sometimes be repeated three times
            @key_sender.puts "{right}{enter}"
            @key_sender.puts "{right}{enter}"
            @key_sender.puts "{right}{enter}"
            #@key_sender.puts "%{n}" #, :initial_wait => 0.2, :if_target_missing=>"Excel window not found")
            expect{
              @book.save_as(@simple_save_file1, :if_exists => :alert)
              }.to raise_error(ExcelErrorSave, "not saved or canceled by user")
            File.exist?(@simple_save_file1).should be_true
            File.size?(@simple_save_file1).should == @garbage_length
            @book.excel.DisplayAlerts.should == displayalert_value
          end

          it "should report save errors and leave DisplayAlerts unchanged" do
            #@key_sender.puts "{left}{enter}" #, :initial_wait => 0.2, :if_target_missing=>"Excel window not found")
            @book.ole_workbook.Close
            expect{
              @book.save_as(@simple_save_file1, :if_exists => :alert)
              }.to raise_error(ExcelErrorSave, "Workbook is not alive")
            File.exist?(@simple_save_file1).should be_true
            File.size?(@simple_save_file1).should == @garbage_length
            @book.excel.DisplayAlerts.should == displayalert_value
          end

        end

        context "with :if_exists => :excel" do
          before do
            File.delete @simple_save_file1 rescue nil
            File.open(@simple_save_file1,"w") do | file |
              file.puts "garbage"
            end
            @garbage_length = File.size?(@simple_save_file)
            @key_sender = IO.popen  'ruby "' + File.join(File.dirname(__FILE__), '../helpers/key_sender.rb') + '" "Microsoft Excel" '  , "w"
          end

          after do
            @key_sender.close
          end

          it "should save if user answers 'yes'" do
            # "Yes" is to the left of "No", which is the  default. --> language independent
            @key_sender.puts "{left}{enter}" #, :initial_wait => 0.2, :if_target_missing=>"Excel window not found")
            @book.save_as(@simple_save_file1, :if_exists => :excel)
            File.exist?(@simple_save_file1).should be_true
            File.size?(@simple_save_file1).should > @garbage_length
            @book.excel.DisplayAlerts.should == displayalert_value
            new_book = Book.open(@simple_save_file1, :excel => :new)
            new_book.should be_a Book
            new_book.close
            @book.excel.DisplayAlerts.should == displayalert_value
          end

          it "should not save if user answers 'no'" do
            # Just give the "Enter" key, because "No" is the default. --> language independent
            # strangely, in the "no" case, the question will sometimes be repeated three times
            @key_sender.puts "{enter}"
            @key_sender.puts "{enter}"
            @key_sender.puts "{enter}"
            #@key_sender.puts "%{n}" #, :initial_wait => 0.2, :if_target_missing=>"Excel window not found")
            expect{
              @book.save_as(@simple_save_file1, :if_exists => :excel)
              }.to raise_error(ExcelErrorSave, "not saved or canceled by user")
            File.exist?(@simple_save_file1).should be_true
            File.size?(@simple_save_file1).should == @garbage_length
            @book.excel.DisplayAlerts.should == displayalert_value
          end

          it "should not save if user answers 'cancel'" do
            # 'Cancel' is right from 'yes'
            # strangely, in the "no" case, the question will sometimes be repeated three times
            @key_sender.puts "{right}{enter}"
            @key_sender.puts "{right}{enter}"
            @key_sender.puts "{right}{enter}"
            #@key_sender.puts "%{n}" #, :initial_wait => 0.2, :if_target_missing=>"Excel window not found")
            expect{
              @book.save_as(@simple_save_file1, :if_exists => :excel)
              }.to raise_error(ExcelErrorSave, "not saved or canceled by user")
            File.exist?(@simple_save_file1).should be_true
            File.size?(@simple_save_file1).should == @garbage_length
            @book.excel.DisplayAlerts.should == displayalert_value
          end

          it "should report save errors and leave DisplayAlerts unchanged" do
            #@key_sender.puts "{left}{enter}" #, :initial_wait => 0.2, :if_target_missing=>"Excel window not found")
            @book.ole_workbook.Close
            expect{
              @book.save_as(@simple_save_file1, :if_exists => :excel)
              }.to raise_error(ExcelErrorSave, "Workbook is not alive")
            File.exist?(@simple_save_file1).should be_true
            File.size?(@simple_save_file1).should == @garbage_length
            @book.excel.DisplayAlerts.should == displayalert_value
          end

        end

        it "should save to 'simple_save_file.xls' with :if_exists => nil" do
          dirname, basename = File.split(@simple_save_file1)
          File.delete @simple_save_file1 rescue nil
          File.open(@simple_save_file1,"w") do | file |
            file.puts "garbage"
          end
          File.exist?(@simple_save_file1).should be_true
          booklength = File.size?(@simple_save_file1)
          expect {
            @book.save_as(@simple_save_file1)
            }.to raise_error(ExcelErrorSave, /file already exists: "workbook_save.xls"/)
          File.exist?(@simple_save_file1).should be_true
          File.size?(@simple_save_file1).should == booklength
        end

        it "should save to 'simple_save_file.xls' with :if_exists => :invalid" do
          File.delete @simple_save_file1 rescue nil
          @book.save_as(@simple_save_file1)
          expect {
            @book.save_as(@simple_save_file1, :if_exists => :invalid)
            }.to raise_error(ExcelErrorSave, ':if_exists: invalid option: :invalid')
        end
      end
    end
  end
end
  