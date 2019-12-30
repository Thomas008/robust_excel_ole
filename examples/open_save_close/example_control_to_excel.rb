# example_give_control_to_excel.rb: 
# open, close, save  with giving control to Excel 

require_relative '../../lib/robust_excel_ole'
require_relative '../../spec/helpers/create_temporary_dir'
require "fileutils"

include RobustExcelOle

Excel.kill_all
begin
  dir = create_tmpdir
  file_name = dir + 'workbook.xls' 
  book = Workbook.open(file_name, :visible => true)          # open a book
  sleep 1
  sheet = book.sheet(1)                                                        # access a worksheet
  sheet[1,1] = sheet[1,1].Value == "simple" ? "complex" : "simple"       # change a cell
  sleep 1
  begin
    new_book = Workbook.open(file_name, :visible => true, :if_unsaved => :alert) # open another workbook with the same file name 
  rescue WorkbookREOError => msg                          # if the user chooses not open the workbook,
  	puts "#{msg.message}"                                  #   an exeptions is raised
  end
  puts "new book has opened" if new_book
  Excel.current.visible = true
  begin
  	book.close(:if_unsaved => :alert)                      # close the unsaved workbook. 
  rescue WorkbookREOError => msg                          # user is asked whether the unsaved workbook shall be saved
  	puts "#{msg.message}"                                  # if the user chooses to cancel, then an expeption is raised
  end
  if new_book then
  	begin
  	  new_book.save_as(file_name, :if_exists => :alert)    # save the new workbook, if it was opened
  	rescue WorkbookREOError => msg                           # user is asked, whether the existing file shall be overwritten
  	  puts "save_as: #{msg.message}"                       # if the user chooses "no" or "cancel", an exception is raised
  	end 

  	new_book.close                                         # close the new workbook, if the user chose to open it
  end
ensure                                                              
  Excel.kill_all                                       # close ALL workbooks, quit Excel application
  rm_tmp(dir)
end
