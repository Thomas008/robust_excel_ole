# example_saving.rb: 
# save the sheets of a book as separate workbooks

require 'rubygems'
#require 'robust_excel_ole'
require File.join(File.dirname(__FILE__), '../../lib/robust_excel_ole')
require "fileutils"

include RobustExcelOle

begin
  Excel.close_all
  
  dir = "C:/data"
  workbook_name = 'workbook.xls'
  base_name = workbook_name[0,workbook_name.rindex('.')]
  suffix = workbook_name[workbook_name.rindex('.')+1,workbook_name.length]
  file_name = dir + "/" + workbook_name

  Book.unobtrusively(file_name) do |book_orig| 
    book_orig.each do |sheet_orig|
      file_sheet_name = dir + "/" + base_name + "_" + sheet_orig.name + "." + suffix
      Excel.current.generate_workbook(file_sheet_name)
    end
  end  
  Book.unobtrusively(file_name) do |book_orig| 
    book_orig.each do |sheet_orig|
      file_sheet_name = dir + "/" + base_name + "_" + sheet_orig.name + "." + suffix
      # delete all existing sheets, and add the sheet    
      book = Book.open(file_sheet_name)
      book.add_sheet sheet_orig
      book.each do |sheet|
        sheet.Delete() unless sheet.name == sheet_orig.name 
      end
      book.close(:if_unsaved => :save)
      # alternative: delete all other sheets
      #book = Book.open(file_sheet_name, :force_excel => :new, :visible => true)
      #book.each do |sheet|
      #  book[sheet.Name].Delete() unless sheet.Name == sheet_orig.Name
      #end
    end
  end
end
