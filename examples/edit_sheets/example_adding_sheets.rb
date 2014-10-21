# example_adding_sheets.rb: 
# adding new and copied at various positions with various sheet names

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
  sheet = @book[1]
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
  @book.add_sheet(@book[3], :as => 'sheet_copy', :after => @book[6])
  show_sheets

  @book.close(:if_unsaved => :forget)   # close the book without saving it
  
ensure
  Excel.close_all
  rm_tmp(dir)
end

  