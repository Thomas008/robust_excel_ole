# example_if_obstructed_save.rb:
# open with :if_obstructed: :save

require File.join(File.dirname(__FILE__), '../../lib/robust_excel_ole')

include RobustExcelOle

ExcelApp.close_all
begin
  dir = 'C:/'
  file_name = dir + 'simple.xls'
  other_dir = 'C:/more_data/'
  other_file_name = other_dir + 'simple.xls'
  book = Book.open(file_name, :visible => true)  # open a book, make Excel visible
  sleep 1
  sheet = book[0]
  first_cell = sheet[0,0].value                                   # access a sheet
  sheet[0,0] = first_cell == "simple" ? "complex" : "simple"      # change a cell
  sleep 1
  new_book = Book.open(other_file_name, :if_obstructed => :save)  # open a book with the same file name in a different path
  sleep 1                                                         #save the old book, close it, before
  old_book = Book.open(file_name, :if_obstructed => :forget ,:visible => true) # open the old book    
  sleep 1
  old_sheet = old_book[0]
  old_first_cell = old_sheet[0,0].value
  puts "the old book was saved" unless old_first_cell == first_cell 
  new_book.close                                 # close the books                      
  old_book.close
ensure
  ExcelApp.close_all                         # close all workbooks, quit Excel application
end
