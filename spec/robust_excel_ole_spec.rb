# -*- coding: utf-8 -*-

require File.join(File.dirname(__FILE__), './spec_helper')

$VERBOSE = nil

include RobustExcelOle

describe RobustExcelOle do

  before(:all) do
    excel = Excel.new(:reuse => true)
    open_books = excel == nil ? 0 : excel.Workbooks.Count
    puts "*** open books *** : #{open_books}" if open_books > 0
    Excel.close_all
  end

  before do
    @dir = create_tmpdir
    @simple_file = @dir + '/workbook.xls'
  end

  after do
    Excel.kill_all
    rm_tmp(@dir)
  end

  describe "t" do

    it "should put some" do
      a = 4
      t "some text #{a}"
    end

    it "should put another text" do
      a = 5
      t "another text #{a}"
    end

  end

  describe "#absolute_path" do
    
    context "with standard" do
      
      before do
        @previous_dir = Dir.getwd
      end
      
      after do
        Dir.chdir @previous_dir
      end

      it "should return the right absolute paths" do
        absolute_path("C:/abc").should == "C:\\abc"
        absolute_path("C:\\abc").should == "C:\\abc"
        Dir.chdir "C:/windows"
        absolute_path("C:abc").downcase.should == Dir.pwd.gsub("/","\\").downcase + "\\abc"
        absolute_path("C:abc").upcase.should   == File.expand_path("abc").gsub("/","\\").upcase
      end

      it "should return right absolute path name" do
        filename = 'C:/Dokumente und Einstellungen/Zauberthomas/Eigene Dateien/robust_excel_ole/spec/book_spec.rb'
        absolute_path(filename).gsub("\\","/").should == filename
      end
    end
  end

  describe "canonize" do

    context "with standard" do
      
      it "should reduce slash at the end" do
        normalize("hallo/").should == "hallo"
        normalize("/this/is/the/Path/").should == "/this/is/the/Path"
      end

      it "should save capital letters" do
        normalize("HALLO/").should == "HALLO"
        normalize("/This/IS/tHe/patH/").should == "/This/IS/tHe/patH"
      end

      it "should reduce multiple shlashes" do
        normalize("/this/is//the/path").should == "/this/is/the/path"
        normalize("///this/////////is//the/path/////").should == "/this/is/the/path"
      end

      it "should reduce dots in the paths" do
        canonize("/this/is/./the/path").should == "/this/is/the/path"
        canonize("this/.is/./the/pa.th/").should == "this/.is/the/pa.th"
        canonize("this//.///.//.is/the/pa.th/").should == "this/.is/the/pa.th"
      end

      it "should change to the upper directory with two dots" do
        canonize("/this/is/../the/path").should == "/this/the/path"
        canonize("this../.i.s/.../..the/..../pa.th/").should == "this../.i.s/.../..the/..../pa.th"
      end

      it "should downcase" do
        canonize("/This/IS/tHe/path").should == "/this/is/the/path"
        canonize("///THIS/.///./////iS//the/../PatH/////").should == "/this/is/path"
      end

      it "should raise an error for no strings" do
        expect{
          canonize(1)
        }.to raise_error(ExcelError, "No string given to canonize, but 1")
      end

    end
  end

end
