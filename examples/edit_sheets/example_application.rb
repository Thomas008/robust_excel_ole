# example_renaming.rb: 
# 1. each cell is named with the name equaling its value unless it is empty or not a string
#    the new workbook's name is extended by the suffix "_named"
# 2. each named cell gets the value of cell right to it appended to its own value
#    the new workbook's name is extended by the suffix "_concat"
# 3. 
# create a workbook which is named like the old one, expect that the suffix "_expanded" is appended to the base name
# for each (global or local) Excel name of the workbook that refers to a range in a single sheet
# this sheet is to be copied into the new workbook
# the sheet's name shall be the name of the Excel name
# in addition to that, the cell B2 shall be named "name" and get the sheet name as its value 


require File.join(File.dirname(__FILE__), '../../lib/robust_excel_ole')
require File.join(File.dirname(__FILE__), '../../spec/helpers/create_temporary_dir')
require "fileutils"

include RobustExcelOle

Excel.close_all
begin
  dir = create_tmpdir
  simple_file = dir + 'workbook.xls'
  simple_save_file = dir + 'x1_workbook_with_names.xls'
  File.delete @simple_save_file rescue nil
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
  end
  #book_orig = book_new
  Book.unobtrusively(simple_save_file) do |book_orig|
    p "book_orig.Name: #{book_orig.Name}"
    book_orig.each do |sheet_orig|
      p "sheet_orig.Name: #{sheet_orig.Name}"
      p "c: #{sheet_orig.Names.Count}"
      sheet_orig.Names.each do |excel_name|
        full_name = excel_name.Name
        sheet_name, short_name = full_name.split("!")
        p "short_name: #{short_name}"

        #name = cell_orig.Value ? cell_orig.Value.to_s : " "
        #p "name: #{name}"
        #book_new.add_sheet(sheet_new, :as => name)
        #book_new[name].Cells(2,2).Value = name
      end
    end
  end

  #book_new["Tabelle1"].Delete()  
  #book_new["Tabelle2"].Delete()
  #book_new["Tabelle3"].Delete()

  #book_new.close(:if_unsaved => :forget)

ensure
  #Excel.close_all
  #rm_tmp(dir)
end

  

