# example_default_excel.rb: 
# reopening books using :default_excel

require File.expand_path('../../lib/robust_excel_ole', File.dirname(__FILE__))
require File.join(File.dirname(File.expand_path(__FILE__)), '../../spec/helpers/create_temporary_dir')
require "fileutils"

include RobustExcelOle

Excel.close_all
begin
  dir = create_tmpdir
  file_name1 = dir + 'workbook.xls'
  file_name2 = dir + 'different_workbook.xls'
  file_name3 = dir + 'book_with_blank.xls'
  file_name4 = dir + 'merge_cells.xls'
  book1 = Book.open(file_name1)            # open a book in a new Excel instance since no Excel is open
  book1.excel.visible = true               # make current Excel visible
  sleep 2
  book1.close                               # close the book
  sleep 2
  book2 = Book.open(file_name1)            # reopen the book                            
  p "book1 == book2" if book2 == book1      # the books are identical
  sleep 2   
  new_excel = Excel.new(:reuse => false)    # create a new Excel    
  book3 = Book.open(file_name2, :default => {:excel => :current}, :visible => true) # open another book
  if book3.excel == book2.excel then     # since this book cannot be reopened, the option :default => {:excel} applies:
    p "book3 opened in the first Excel"  # according to :default => {:excel => :current} the book is opened
  end                                    # in the Excel instance the was created first
  sleep 2                                          
  new_excel = Excel.new(:reuse => false)         
  book4 = Book.open(file_name3, :default_excel => new_excel, :visible => true)  # open another book
  if book4.excel == new_excel then       # since this book cannot be reopened, the option :default_excel applies: 
    p "book4 opened in the second Excel" # according to :default_excel => new_excel the book is opened
  end                                    # in the given Excel, namely the second Excel instance new_excel 
  sleep 2   
  book5 = Book.open(file_name4, :default_excel => :new, :visible => true)  # open another book
  if ((not book5.excel == book1.excel) && (not book5.excel == new_excel)) then  # since this book cannot be reopened, 
    p "book5 opened in a third Excel" # the option :default_excel applies. according to :default_excel => :new 
  end                                 # the book is opened in a new Excel
  sleep 2  
  book5.close
  book4.close
  book3.close
  book2.close
ensure
  Excel.close_all                            # close all workbooks, quit Excel instances
  rm_tmp(dir)
end
