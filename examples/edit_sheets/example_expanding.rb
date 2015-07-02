# example_expanding.rb:  
# create a workbook which is named like the old one, expect that the suffix "_expanded" is appended to the base name
# for each (global or local) Excel name of the workbook that refers to a range in a single sheet
# this sheet is to be copied into the new workbook
# the sheet's name shall be the name of the Excel name
# in addition to that, the cell B2 shall be named "name" and get the sheet name as its value 

require 'rubygems'
require 'robust_excel_ole'
require "fileutils"

include RobustExcelOle

begin
  Excel.close_all
  dir = "C:/data"
  workbook_name = 'workbook_named_filled_concat.xls'
  base_name, suffix = workbook_name.split(".")
  file_name = dir + "/" + workbook_name
  extended_file_name = dir + "/" + base_name + "_expanded" + "." + suffix
  #Excel.current.generate_workbook(extended_file_name)
  #Excel.close_all
  sheet_names = []
  book_orig = Book.open(file_name)
  book_orig.save_as(extended_file_name, :if_exists => :overwrite) 
  book_orig.close
  #Excel.close_all
  excel = Excel.new(:reuse => false, :visible => true)
  Book.unobtrusively(extended_file_name, :if_closed => excel, :keep_open => true) do |book|     
    p "book: #{book}"
    book.each do |sheet|
      p "sheet.name: #{sheet.name}"
      sheet_names << sheet.name 
      names = sheet.Names
      p "names: #{names} #{names.Count}"
      sheet.Names.each do |excel_name|
        full_name = excel_name.Name
        p "full_name: #{full_name}"
        sheet_name, short_name = full_name.split("!")
        p "short_name: #{short_name}"
        sheet_new = book.add_sheet(sheet, :as => short_name)
        sheet_new.Names.Add("Name" => "name", "RefersTo" => "=" + "$B$2")
        sheet_new[1,1].Value = short_name
        #sheet_new["name"] = short_name
        book_new["name"] = short_name
      end
    end
    sheet_names.each do |sheet_name|
      book[sheet_name].Delete()
    end
  end
end
