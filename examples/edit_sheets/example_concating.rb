# example_concatening.rb: 
# each named cell gets the value of cell right to it appended to its own value
# the new workbook's name is extended by the suffix "_concat"

require 'rubygems'
require 'robust_excel_ole'
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
      sheet.each do |cell_orig|
        name = cell_orig.Name.Name rescue nil
        if name
          sheet[cell_orig.Row-1, cell_orig.Column-1].Value = cell_orig.Value.to_s + cell_orig.Offset(0,1).Value.to_s
          sheet.Names.Add("Name" => name, "RefersTo" => "=" + cell_orig.Address) 
        end
      end
    end
  end
end
