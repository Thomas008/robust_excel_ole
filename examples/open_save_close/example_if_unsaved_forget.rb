# example_ifunsaved_forget.rb:
# open with :if_unsaved => :forget, :new_app, close with :if_unsaved => :save 

require File.join(File.dirname(__FILE__), '../../lib/robust_excel_ole')

include RobustExcelOle

Excel.close_all
begin
  dir = 'C:/'
  file_name = dir + 'simple.xls' 
  book = Book.open(file_name)                      # open a book
  Excel.current.Visible = true        # make Excel visible 
  sleep 1
  sheet = book[0]                                            # access a sheet
  first_cell = sheet[0,0].value
  sheet[0,0] = first_cell == "simple" ? "complex" : "simple" # change a cell
  sleep 1
  new_book = Book.open(file_name, :if_unsaved => :forget) # open another book with the same file name 
                                                          # and close the unsaved book without saving it
  sheet_new_book = new_book[0]
  if (not book.alive?) && new_book.alive? && sheet_new_book[0,0].value == first_cell then # check whether the unsaved book 
    puts "open with :if_unsaved => :forget : the unsaved book is closed and not saved."     # is closed and was not saved
  end
  sleep 1
  sheet_new_book[0,0] = sheet_new_book[0,0].value == "simple" ? "complex" : "simple" # change a cell
  # open another book in a new Excel application, and make Excel visible, leaving the unsaved book open
  another_book = Book.open(file_name, :if_unsaved => :new_app, :visible => true)  
  sleep 3                                                                  # leaving the unsaved book open  
  new_book.close(:if_unsaved => :forget )                                
  another_book.close
ensure
	  Excel.close_all                                    # close all workbooks, quit Excel application
end