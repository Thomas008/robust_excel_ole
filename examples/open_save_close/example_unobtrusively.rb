# example_unobtrusively.rb:

require_relative '../../lib/robust_excel_ole'
require_relative '../../spec/helpers/create_temporary_dir'
require "fileutils"

include RobustExcelOle

Excel.kill_all
begin
  dir = create_tmpdir
  simple_file = dir + 'workbook.xls'
  book = Workbook.open(simple_file, :visible => true)  # open a workbook, make Excel visible
  old_sheet = book.sheet(1)
  p "1st cell: #{old_sheet[1,1]}"
  sleep 2
  Workbook.unobtrusively(simple_file) do |book|   # modify the book and keep its status unchanged
    sheet = book.sheet(1)
    sheet[1,1] = sheet[1,1] == "simple" ? "complex" : "simple"
  end
  new_sheet = book.sheet(1)
  p "1st cell: #{new_sheet[1,1]}"
  p "book saved" if book.Saved
  book.close                                 # close the workbook                      
ensure
  Excel.kill_all                            # close all workbooks, quit Excel application
  rm_tmp(dir)
end
