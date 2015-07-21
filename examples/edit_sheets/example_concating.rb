# example_concating.rb: 
# each named cell gets the value of cell right to it appended to its own value
# the new workbook's name is extended by the suffix "_concat"

require 'rubygems'
#require 'robust_excel_ole'
require File.join(File.dirname(__FILE__), '../../lib/robust_excel_ole')
require "fileutils"

include RobustExcelOle

begin
  Excel.close_all
  dir = "C:/data"
  workbook_name = 'workbook_named.xls'
  base_name = workbook_name[0,workbook_name.rindex('.')]
  suffix = workbook_name[workbook_name.rindex('.')+1,workbook_name.length]
  file_name = dir + "/" + workbook_name
  extended_file_name = dir + "/" + base_name + "_concat" + "." + suffix
  book_orig = Book.open(file_name)
  book_orig.save_as(extended_file_name, :if_exists => :overwrite) 
  book_orig.close
  Book.unobtrusively(extended_file_name) do |book|
    book.each do |sheet|
      sheet.each do |cell|
        name = cell.Name.Name rescue nil
        if name
          cell.Value = cell.Value.to_s + cell.Offset(0,1).Value.to_s
          sheet.add_name(cell.Row, cell.Column, name)
        end
      end
    end
  end
end
