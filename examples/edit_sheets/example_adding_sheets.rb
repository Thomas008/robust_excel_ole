# example_adding_sheets.rb: 
# adding new and copied at various positions with various sheet names

require File.expand_path('../../lib/robust_excel_ole', File.dirname(__FILE__))
require File.join(File.dirname(File.expand_path(__FILE__)), '../../spec/helpers/create_temporary_dir')
require "fileutils"

include RobustExcelOle

Excel.close_all
begin
  dir = create_tmpdir
  simple_file = dir + 'workbook.xls'
  simple_save_file = dir + 'workbook_save.xls'
  File.delete simple_save_file rescue nil
  @book = Book.open(simple_file)      # open a book

  def show_sheets 
    @book.each do |sheet|               # access each sheet
      puts "sheet name: #{sheet.name}" #  put the sheet name
    end
  end

  puts "sheets of the book:"
  show_sheets

  puts "adding a new sheet"
  @book.add_sheet
  show_sheets

  puts "adding a new sheet with the name 'sheet_name'"
  @book.add_sheet(:as => 'sheet_name')
  show_sheets

  puts "adding a copy of the 2nd sheet"
  sheet = @book.sheet(2)
  @book.add_sheet sheet
  show_sheets

  puts "adding a copy of the 2nd sheet and name it 'second_sheet_copy'"
  @book.add_sheet(sheet, :as => 'second_sheet_copy')
  show_sheets

  puts "adding a new sheet after the 2nd sheet"
  @book.add_sheet(:after => sheet)
  show_sheets

  puts "adding a copy of the 2nd sheet after the 2nd sheet"
  @book.add_sheet(sheet, :after => sheet)
  show_sheets

  puts "adding a copy of the 4th sheet before the 7th sheet and name it 'sheet_copy'"
  @book.add_sheet(@book.sheet(4), :as => 'sheet_copy', :after => @book.sheet(7))
  show_sheets

  puts"adding a copy of the 2nd sheet and name it again 'second_sheet_copy'"
  begin
    @book.add_sheet(sheet, :as => 'second_sheet_copy')
  rescue ExcelErrorSheet => msg
    puts "error: add_sheet: #{msg.message}"
  end    

  @book.close(:if_unsaved => :forget)   # close the book without saving it
  
ensure
  Excel.close_all
  rm_tmp(dir)
end

  