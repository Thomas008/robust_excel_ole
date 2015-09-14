# -*- coding: utf-8 -*-

require File.join(File.dirname(__FILE__), './../spec_helper')


module My # :nodoc: #
  class Excel < RobustExcelOle::Excel # :nodoc: #
  end

  class Book < RobustExcelOle::Book   # :nodoc: #
  end
end

describe "subclassed Book" do

  before(:all) do
    excel = RobustExcelOle::Excel.new(:reuse => true)
    open_books = excel == nil ? 0 : excel.Workbooks.Count
    puts "*** open books *** : #{open_books}" if open_books > 0
    RobustExcelOle::Excel.close_all
  end

  before do
    @dir = create_tmpdir
    @simple_file = @dir + '/workbook.xls'
  end

  after do
    RobustExcelOle::Excel.close_all
    rm_tmp(@dir)
  end

  
  describe "open" do

    it "should use the subclassed Excel" do
      #REO::Book.open(@simple_file) do |book|
      My::Book.open(@simple_file) do |book|
        book.should be_a RobustExcelOle::Book
        book.should be_a My::Book
        book.excel.should be_a RobustExcelOle::Excel
        book.excel.class.should == My::Excel
        book.excel.should be_a My::Excel
      end
    end

  end

  
end
