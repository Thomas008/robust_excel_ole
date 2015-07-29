# example_expanding.rb:  
# create a workbook which is named like the old one, expect that the suffix "_expanded" is appended to the base name
# for each (global or local) Excel name of the workbook that refers to a range in a single sheet
# this sheet is to be copied into the new workbook
# the sheet's name shall be the name of the Excel name
# in addition to that, the cell B2 shall be named "name" and get the sheet name as its value 

require 'rubygems'
#require 'robust_excel_ole'
require File.join(File.dirname(__FILE__), '../../lib/robust_excel_ole')
require "fileutils"

include RobustExcelOle

begin
  dir = "C:/data"
  workbook_name = 'workbook_named_concat.xls'
  ws = workbook_name.split(".")
  base_name = ws[0,ws.length-1].join(".")
  suffix = ws.last  
  file_name = dir + "/" + workbook_name
  extended_file_name = dir + "/" + base_name + "_expanded" + "." + suffix
  FileUtils.copy file_name, extended_file_name 
  
  Book.unobtrusively(extended_file_name) do |book|     
    book.extend Enumerable
    sheet_names = book.map { |sheet| sheet.name }

    book.Names.each do |excel_name|
      full_name = excel_name.Name
      sheet_name, short_name = full_name.split("!")
      sheet = excel_name.RefersToRange.Worksheet
      sheet_name = short_name ? short_name : sheet_name
      sheet_new = book.add_sheet sheet
      begin
        sheet_new.name = sheet_name
      rescue ExcelErrorSheet => msg
        sheet_new.name = sheet_name + sheet.name if msg.message == "sheet name already exists" 
      end
      sheet_new.set_name("name", 2, 2)
      sheet_new["name"] = sheet_name   
    end
    
    sheet_names.each do |sheet_name|
      book[sheet_name].Delete()
    end
  end
end
