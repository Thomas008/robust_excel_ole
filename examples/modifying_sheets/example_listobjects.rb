# example_ranges.rb: 
# access row and column ranges of a sheet. 

require_relative '../../lib/robust_excel_ole'
require_relative '../../spec/helpers/create_temporary_dir'
require "fileutils"

include RobustExcelOle

using StringRefinement

Excel.close_all(if_unsaved: :forget)
begin
  dir = create_tmpdir
  table_file = dir + 'workbook_listobjects.xlsx'
  book = Workbook.open(table_file, :visible => true)      # open a workbook
  sheet = book.sheet(3)              # access a worksheet via the name
  table = sheet.table(1)             # access a list object (table)

  puts "table: #{table}"

  listrow1 = table[2]                # access the second list row of the table

  puts "listrow1: #{listrow1}"

  values1 = listrow1.values             # access the values of this listrow
  puts "values: #{values1}"

  listrow2 = table[{"Number" => 3, "Person" => "Angel"}]  # acess a list row via a given key hash

  puts "listrow2: #{listrow2}"

  values2 = listrow2.values
  puts "values: #{values2}"

  listrows3 = table[{"Number" => 3}, limit: 2]  # access maximal 2 listrows matching the key

  puts "listrows3: #{listrows3}"
  puts "values:"
  listrows3.map{|l| puts l.values}

  puts "deleting the values of the second row"   # deleting the contents of the second row
  table.delete_row_values(2)

  puts "deleting empty rows"   
  table.delete_empty_rows              # deleting empty rows

  puts "sorting table:"
  table.sort("Number")                 # sort table

  puts "find all cells:"               # find all cells containing a given value
  cells = table.find_cells(40)
  puts "cells: #{cells}"

  # create a new table
  table2 = Table.new(sheet, "table_name", [20,1], 3, ["Verkäufer", "Straße", "area in m²"])
  puts "table2: #{table2}"

  table_row1 = table[1]         
  puts "list_row1 of second table: #{table_row1}"
  
  table_row1.verkaeufer = "John"
  value3 = table_row1.verkaeufer
  puts "value of verkaeufer: #{value3}"

  value4 = table_row1.Verkaeufer
  puts "value of verkaeufer: #{value4}"

  table_row1.strasse = 42
  value5 = table_row1.strasse
  puts "value of strasse: #{value5}"

  table_row1.area_in_m2 = 400
  value6 = table_row1.area_in_m2
  puts "value of area in m2: #{value6}"

  sleep 5
 
  book.close(:if_unsaved => :forget)
  
ensure
  Excel.close_all(if_unsaved: forget)
  rm_tmp(dir)
end

  