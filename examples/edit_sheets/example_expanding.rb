# example_expanding.rb:  
# create a workbook which is named like the old one, expect that the suffix "_expanded" is appended to the base name
# for each (global or local) Excel name of the workbook that refers to a range in a single sheet
# this sheet is to be copied into the new workbook
# the sheet's name shall be the name of the Excel name
# in addition to that, the cell B2 shall be named "name" and get the sheet name as its value 

# todos: global names
# save the new workbook, keep the old one unchanged (we can copy only within one workbook)
# avoid create_tmpdir

require File.join(File.dirname(__FILE__), '../../lib/robust_excel_ole')
require File.join(File.dirname(__FILE__), '../../spec/helpers/create_temporary_dir')
require "fileutils"

include RobustExcelOle

begin
  Excel.close_all
  dir = create_tmpdir
  workbook_name = 'workbook_named_filled_concat.xls'
  base_name, suffix = workbook_name.split(".")
  file_name = dir + "/" + workbook_name
  extended_file_name = dir + "/" + base_name + "_expanded" + "." + suffix
  Excel.current.generate_workbook(extended_file_name)
  book_new = Book.open(extended_file_name, :visible => true)
  sheet_new  = book_new[0]  
  sheet_orig_names = []
  excel = Excel.create
  excel.visible = true
  Book.unobtrusively(file_name, :if_closed => excel, :keep_open => true) do |book_orig|     
    p "file_name: #{file_name}"
    # todo: global names
    # for all local names
    book_orig.each do |sheet_orig|
      sheet_orig_names << sheet_orig.name 
      sheet_orig.Names.each do |excel_name|        
        full_name = excel_name.Name
        sheet_name, short_name = full_name.split("!")
        # we have to work in the same workbook, then save it as the new workbook and close it
        sheet_new = book_orig.add_sheet(sheet_orig, :as => short_name)
        sheet_new.Names.Add("Name" => "name", "RefersTo" => "=" + "$B$2")
        sheet_new[1,1].Value = short_name
      end
    end
    sheet_orig_names.each do |sheet_orig_name|
      book_orig[sheet_orig_name].Delete()
    end
  end

  book_orig.save_as(extended_file_name)
  #book_orig.close
  book_new.close

ensure
  #Excel.close_all
  #rm_tmp(dir)
end

  

