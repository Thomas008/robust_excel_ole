# example_simple.rb: 
# open a book, simple save, save_as, close

require_relative '../../lib/robust_excel_ole'
require_relative '../../spec/helpers/create_temporary_dir'
require "fileutils"

include RobustExcelOle

Excel.kill_all
begin
  dir = create_tmpdir
  file_name = dir + 'workbook.xls'
  book = Workbook.open(file_name, :visible => true)       # open a workbook.  default:  :read_only => false
  sheet = book.sheet(1)
  workbook = book.ole_workbook
  fullname = workbook.Fullname
  puts "fullname: #{fullname}"  
  sheet.add_name("a_name",[1,1])   # rename cell A1 to "a_name"
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
  new_name_object.Delete
  sleep 2
  book.close(:if_unsaved => :forget)                        # close the workbook

ensure
	  Excel.kill_all                                         # close workbooks, quit Excel application
    rm_tmp(dir)
end		

