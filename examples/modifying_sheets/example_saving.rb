# example_saving.rb: 
# save the sheets of a book as separate workbooks

require_relative '../../lib/robust_excel_ole'
require_relative '../../spec/helpers/create_temporary_dir'
require "fileutils"

include RobustExcelOle

begin
  
  dir = File.expand_path('../../spec/data', File.dirname(__FILE__))
  workbook_name = 'workbook_new.xls'
  ws = workbook_name.split(".")
  base_name = ws[0,ws.length-1].join(".")
  suffix = ws.last  
  file_name = dir + "/" + workbook_name

  Workbook.unobtrusively(file_name) do |book_orig| 
    book_orig.each do |sheet_orig|
      file_sheet_name = dir + "/" + base_name + "_" + sheet_orig.name + "." + suffix
      Excel.current.generate_workbook(file_sheet_name)
      # delete all existing sheets, and add the sheet    
      book = Workbook.open(file_sheet_name)
      book.add_sheet sheet_orig
      book.each do |sheet|
        sheet.Delete unless sheet.name == sheet_orig.name 
      end
      book.close(:if_unsaved => :save)
      # alternative: delete all other worksheets
      #book = Workbook.open(file_sheet_name, :force => {:excel => :new}, :visible => true)
      #book.each do |sheet|
      #  book.sheet(sheet.Name).Delete() unless sheet.Name == sheet_orig.Name
      #end
    end
  end
end
