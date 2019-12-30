# example_adding_sheets.rb: 
# adding new and copied at various positions with various sheet names

require_relative '../../lib/robust_excel_ole'
require_relative '../../spec/helpers/create_temporary_dir'
require "fileutils"

include RobustExcelOle

Excel.kill_all
begin
  dir = create_tmpdir
  simple_file = dir + 'workbook.xls'
  simple_save_file = dir + 'workbook_save.xls'
  File.delete simple_save_file rescue nil
  @book = Workbook.open(simple_file, :visible => true)      # open a workbook

  def show_sheets 
    @book.each do |sheet|               # access each worksheet
      puts "sheet name: #{sheet.name}" #  put the worksheet name
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
  rescue NameAlreadyExists => msg
    puts "results in an error: add_sheet: #{msg.message}"
  end    

  @book.close(:if_unsaved => :forget)   # close the book without saving it
  
ensure
  Excel.kill_all
  #Excel.close_all(:if_unsaved => :forget)
  rm_tmp(dir)
end

  