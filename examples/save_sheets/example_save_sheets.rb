# example 11: save the sheets of a book as separate books

require File.join(File.dirname(__FILE__), '../lib/robust_excel_ole')

require "fileutils"

include RobustExcelOle

ExcelApp.close_all
begin
  dir = 'C:/'
  file_name = dir + 'book_with_blank.xls'
  book = Book.open(file_name)                   # open a book
  ExcelApp.reuse.Visible = true     # make Excel visible 
  # 1. leere books speichern, um speichern zu können
  # 2. die sheets als books speichern:
  #    dazu ein book mit dem sheet öffnen
  #    dazu die anderen sheets schließen
  i = 0
  book.each do |sheet|
    i = i + 1
    puts "#{i}. sheet:"
    file_name_sheet = file_name + "_sheet#{i}.xls"
    puts "file_name_sheet: #{file_name_sheet}"
	  # generate empty book
    excel = 
	  book_sheet = ExcelApp.Workbooks.Add
    sleep 1
  end
 ensure                                                              
  ExcelApp.close_all                              # close workbooks, quit Excel application
 end

