# example_copying.rb:
# each named cell is to be copied into another sheet
# unnamed cells shall not be copied
# if a sheet does not contain any named cell, then the sheet shall not be copied

require 'rubygems'
#require 'robust_excel_ole'
require File.join(File.dirname(__FILE__), '../../lib/robust_excel_ole')
require "fileutils"

include RobustExcelOle

begin
  Excel.close_all
  dir = "C:/data"
  workbook_name = 'another_workbook.xls'
  base_name = workbook_name[0,workbook_name.rindex('.')]
  suffix = workbook_name[workbook_name.rindex('.')+1,workbook_name.length]
  file_name = dir + "/" + workbook_name
  extended_file_name = dir + "/" + base_name + "_copied" + "." + suffix
  book_orig = Book.open(file_name)
  book_orig.save_as(extended_file_name, :if_exists => :overwrite) 
  book_orig.close
  Book.unobtrusively(extended_file_name) do |book|  
    book.each do |sheet|
      new_sheet = book.add_sheet 
      contains_named_cells = false
      sheet.each do |cell|
        full_name = cell.Name.Name rescue nil
        if full_name
          sheet_name, short_name = full_name.split("!") 
          cell_name = short_name ? short_name : sheet_name
          contains_named_cells = true
          new_sheet[cell.Row, cell.Column].Value = cell.Value
          new_sheet.set_name(cell_name, cell.Row,cell.Column)
        end
      end
      new_sheet.Delete() unless contains_named_cells
    end
  end
end
