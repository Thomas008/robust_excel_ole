# -*- coding: utf-8 -*-

require File.join(File.dirname(__FILE__), './../spec_helper')

# @private
module My 
  class Excel < RobustExcelOle::Excel 
  end

  class Workbook < RobustExcelOle::Workbook   
  end

end

describe "subclassed Workbook" do

  before(:all) do
    excel = RobustExcelOle::Excel.new(:reuse => true)
    open_books = excel == nil ? 0 : excel.Workbooks.Count
    puts "*** open books *** : #{open_books}" if open_books > 0
    RobustExcelOle::Excel.kill_all
  end

  before do
    @dir = create_tmpdir
    @simple_file = @dir + '/workbook.xls'
  end

  after do
    RobustExcelOle::Excel.kill_all
    rm_tmp(@dir)
  end

  
  describe "open" do

    it "should use the subclassed Excel" do
      My::Workbook.open(@simple_file) do |book|
        book.should be_a RobustExcelOle::Workbook
        book.should be_a My::Workbook
        book.excel.should be_a RobustExcelOle::Excel
        book.excel.class.should == My::Excel
        book.excel.should be_a My::Excel
      end
    end

  end

  
end
