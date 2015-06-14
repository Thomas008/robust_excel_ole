# example_renaming.rb: 
# 1. each cell gets the name of its value
# 2. each renamed cell gets the value of of the value right of the cell
# 3. for each name a new sheet is created.
#     its name is the name, it is a copy of the sheet, the cell B2 gets the name of the sheet

require File.join(File.dirname(__FILE__), '../../lib/robust_excel_ole')
require File.join(File.dirname(__FILE__), '../../spec/helpers/create_temporary_dir')
require "fileutils"

include RobustExcelOle

Excel.close_all
begin
  dir = create_tmpdir
  simple_file = dir + 'workbook.xls'
  simple_save_file = dir + 'x1_workbook_with_names.xls'
  Excel.current.generate_workbook(simple_save_file)
  book_new = Book.open(simple_save_file, :visible => true)
  sheet_new  = book_new[0]
  Book.unobtrusively(simple_file) do |book_orig|     
    sheet_orig = book_orig[0]
    sheet_orig.each do |cell_orig|
      p cell_orig.Address
      contents = cell_orig.Value
      p "contents: #{contents}"
      if contents
        sheet_new.Names.Add("Name" => contents, "RefersTo" => "=" + cell_orig.Address) 
        sheet_new.Cells(cell_orig.Row,cell_orig.Column).Value = cell_orig.Offset(0,1).Value
      end
    end
   #sheet_new.name = sheet_orig.Name
  end
  #book_new.close(:if_unsaved => :forget)
ensure
  #Excel.close_all
  #rm_tmp(dir)
end

  

