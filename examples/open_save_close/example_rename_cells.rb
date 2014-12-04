# example_simple.rb: 
# open a book, simple save, save_as, close

require File.join(File.dirname(__FILE__), '../../lib/robust_excel_ole')
require File.join(File.dirname(__FILE__), '../../spec/helpers/create_temporary_dir')
require "fileutils"

include RobustExcelOle

Excel.close_all
begin
  dir = create_tmpdir
  file_name = dir + 'simple.xls'
  book = Book.open(file_name)                                # open a book.  default:  :read_only => false
  book.visible true                                          # make current Excel visible
  sheet = book[0]
  #sheet.Names.Add("Wert","$A$1"

  workbook = book.workbook
  fullname = workbook.Fullname
  puts "fullname: #{fullname}"  
  workbook.Names.Add("a_name", "=$A$1")   # rename cell A1 to "a_name"
  number = workbook.Names.Count
  puts "number of name objects :#{number}"
  name_object = workbook.Names("a_name")
  name = name_object.Name 
  puts "name: #{name}"
  value = name_object.Value                   # definition of the cell
  puts "definition: #{value}"
  reference = name_object.RefersTo
  puts "reference: #{reference}"
  visible = name_object.Visible
  puts "visible: #{visible}"
  sleep 2
  workbook.Names("a_name").Name = "new_name"
  new_name_object = workbook.Names("new_name")
  puts "name: #{new_name_object.Name}"
  puts "definition: #{new_name_object.Value}"
  puts "reference: #{new_name_object.RefersTo}"
  puts "visible: #{new_name_object.Visible}"
  sleep 2
  new_name_object.RefersTo = "=$A$2"
  puts "name: #{new_name_object.Name}"
  puts "definition: #{new_name_object.Value}"
  puts "reference: #{new_name_object.RefersTo}"
  puts "visible: #{new_name_object.Visible}"
  sleep 2
  new_name_object.Visible = false
  puts "visible: #{new_name_object.Visible}"
  sleep 2

  # to do:
  # read, write contents of the cell (under the old name)
  # read, write contents in the cell under the new name

  #sheet.Cells
  # sheet[0,0].value
  # Worksheets("sheet1").Cells

  new_name_object.Delete

  sleep 2
  book.close(:if_unsaved => :forget)                        # close the book

ensure
	  Excel.close_all                                         # close workbooks, quit Excel application
    rm_tmp(dir)
end		

