# example_naming.rb: 
# each cell is named with the name equaling its value unless it is empty or not a string
# (and the contents is copied?)
# the new workbook's name is extended by the suffix "_named"

require File.join(File.dirname(__FILE__), '../../lib/robust_excel_ole')
require "fileutils"

include RobustExcelOle

begin
  Excel.close_all
  dir = "C:/data"
  # for some reason: does not work with workbook.xls
  workbook_name = 'workbook.xls'
  base_name, suffix = workbook_name.split(".")
  file_name = dir + "/" + workbook_name
  extended_file_name = dir + "/" + base_name + "_named" + "." + suffix
  Excel.current.generate_workbook(extended_file_name)
  book_new = Book.open(extended_file_name, :visible => true)
  sheet_new  = book_new[0]
  excel = Excel.new(:reuse => false, :visible => true)
  Book.unobtrusively(file_name, :if_closed => excel, :keep_open => true) do |book_orig|     
    sheet_orig = book_orig[0]
    sheet_orig.each do |cell_orig|
      contents = cell_orig.Value
      if contents && (contents.class == String)
        sheet_new.Names.Add("Name" => contents, "RefersTo" => "=" + cell_orig.Address) 
      end
    end
  end
  book_new.save
  #book_new.close
end
