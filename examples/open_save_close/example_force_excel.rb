# example_force_excel.rb: 
# opening books in new or given Excel instances using :force_excel

require_relative '../../lib/robust_excel_ole'
require_relative '../../spec/helpers/create_temporary_dir'
require "fileutils"

include RobustExcelOle

Excel.kill_all
begin
  dir = create_tmpdir
  simple_file = dir + 'workbook.xls'
  book1 = Workbook.open(simple_file, :visible => true)  # open a workbook in a new Excel instance since no Excel is open
  sleep 2
  book2 = Workbook.open(simple_file)             # open a new workbook in the same Excel instance                            
  p "book1 == book2" if book2 == book1      # the workbooks are identical
  sleep 2       
  book3 = Workbook.open(simple_file, :force => {:excel => :new, :visible => true}) # open another workbook in a new Excel instance,   
  p "book3 != book1" if (not (book3 == book1))   # the workbooks are not identical 
  sleep 2   
  new_excel = Excel.new(:reuse => false)        # create a third Excel instance
  book4 = Workbook.open(simple_file, :force => {:excel => new_excel, :visible => true})  # open another workbook in the new Excel instance
  p "book4 != book3 && book4 != book1" if (not (book4 == book3) && (not (book4 == book1)))
  sleep 2                                     # (Excel chooses the first Excel application)        
  book4.close                                 # close the workbooks
  book3.close                                             
  book2.close
  book1.close                                         
ensure
  Excel.kill_all                            # close all workbooks, quit Excel instances
  rm_tmp(dir)
end
