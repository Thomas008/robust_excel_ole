# example_force_excel.rb: 
# opening books in new or given Excel instances using :force_excel

require File.expand_path('../../lib/robust_excel_ole', File.dirname(__FILE__))
require File.join(File.dirname(File.expand_path(__FILE__)), '../../spec/helpers/create_temporary_dir')
require "fileutils"

include RobustExcelOle

Excel.close_all
begin
  dir = create_tmpdir
  simple_file = dir + 'workbook.xls'
  book1 = Book.open(simple_file)            # open a book in a new Excel instance since no Excel is open
  book1.excel.visible = true                # make current Excel visible
  sleep 2
  book2 = Book.open(simple_file)             # open a new book in the same Excel instance                            
  p "book1 == book2" if book2 == book1      # the books are identical
  sleep 2       
  book3 = Book.open(simple_file, :force_excel => :new, :visible => true) # open another book in a new Excel instance,   
  p "book3 != book1" if (not (book3 == book1))   # the books are not identical 
  sleep 2   
  new_excel = Excel.new(:reuse => false)        # create a third Excel instance
  book4 = Book.open(simple_file, :force_excel => new_excel, :visible => true)  # open another book in the new Excel instance
  p "book4 != book3 && book4 != book1" if (not (book4 == book3) && (not (book4 == book1)))
  sleep 2                                     # (Excel chooses the first Excel application)        
  book4.close                                 # close the books
  book3.close                                             
  book2.close
  book1.close                                         
ensure
  Excel.close_all                            # close all workbooks, quit Excel instances
  rm_tmp(dir)
end
