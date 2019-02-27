# example_naming.rb: 
# each cell is named with the name equaling its value unless it is empty or not a string
# the contents of each cell is copied
# the new workbook's name is extended by the suffix "_named"

require File.expand_path('../../lib/robust_excel_ole', File.dirname(__FILE__))
require File.join(File.dirname(File.expand_path(__FILE__)), '../../spec/helpers/create_temporary_dir')
require "fileutils"

include RobustExcelOle

begin
  @id2exl = ["house", "tree", "cat", "mouse", "elephant", "yes", "no"]
  column_ids = [2,4,6]
  dir = File.expand_path('../../spec/data', File.dirname(__FILE__))
  workbook_name = 'workbook.xls'
  filename = dir + "/" + workbook_name
  puts "filename: #{filename}"

  Excel.close_all if_unsaved: :forget
  book = Workbook.new filename, if_absent: :create, visible: true, if_unsaved: :accept
  sheet = book.sheet(1)
  puts "book: #{book}"
  puts "sheet: #{sheet}"

  # book.Names.Add("Âµ",  RefersToR1C1Local:"=Z")
  # book.Names.Add("Âµ_1", RefersToR1C1Local:"=Z(-1)")

  def define_columns sheet, columns_ids
    puts "define_columns:"
    first_column = sheet.range("A")
    puts "first_column: #{first_column}"

    columns_ids.each_with_index do |id,idx|
      puts "id: #{id}"
      puts "idx: #{idx}"

      nam = @id2exl[id]
      puts "nam: #{nam}"
      colnr = idx+1
      puts "colnr: #{colnr}"
      sheet[1,colnr] = nam
      puts "sheet[1,colnr]: #{sheet[1,colnr].Value}"
      sheet.add_name(nam,[nil,colnr])
      puts "sheet.Range().Address: #{sheet.Range(nam).Address}"
      #sheet.add_name(nam,"A")
      # blatt.Names.Add(Name: nam, RefersTo: "="+erste_spalte.Offset(0,idx).Address)
    end
  end

  define_columns sheet, column_ids

  
  # book.save

end

