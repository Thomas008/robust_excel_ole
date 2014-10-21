# example_access_sheets_and_cells.rb: 
# access sheets, print cells, rows, and columns of a sheet

require File.join(File.dirname(__FILE__), '../../lib/robust_excel_ole')
require File.join(File.dirname(__FILE__), '../../spec/helpers/create_temporary_dir')
require "fileutils"

include RobustExcelOle

Excel.close_all
begin
  dir = create_tmpdir
  simple_file = dir + 'simple.xls'
  simple_save_file = dir + 'simple_save.xls'
  File.delete simple_save_file rescue nil
  book = Book.open(simple_file)      # open a book
  sheet = book[0]                    # access a sheet via integer 
  cell = sheet[0,0]                  # access the first cell
  puts "1st cell: #{cell.value}"     # put the value of the first cell
  sheet[0,0] = "complex"             # write a value into a cell
  puts "new cell: #{sheet[0,0].value}"
  puts "all cells:"
  sheet.each do |cell|               # access all cells
    puts "#{cell.value}"             #   for each row: for every column: put the value of the cells
  end
  
  sheet_enum = proc do |enum_method|     # put each cell, each row or each column 
    i = 0
    sheet.send(enum_method) do |item|
      i = i + 1
      item_name = 
        case enum_method
        when :each        : "cell"
        when :each_row    : "row"
        when :each_column : "column"
        end 
      puts "#{item_name} #{i}: #{item.value}" # put values of the item of the sheet
    end
  end

  sheet_enum[:each]        # put cells
  sheet_enum[:each_row]    # put rows
  sheet_enum[:each_column] # put columns

  book.save                # save the book
  book.close               # close the book
  
ensure
  Excel.close_all
  rm_tmp(dir)
end

  