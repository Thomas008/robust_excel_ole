#require 'robust_excel_ole'

#workbook = Workbook.open './sample_excel_files/xlsx_500_rows.xlsx'

require_relative '../lib/robust_excel_ole'

include RobustExcelOle

workbook = Workbook.open './../spec/data/workbook.xls'

puts "Found #{workbook.worksheets_count} worksheets"

workbook.each do |worksheet|
  puts "Reading: #{worksheet.name}"
  num_rows = 0

  worksheet.each do |row_values|
    a = row_values.map{ |cell| cell }
    num_rows += 1

    # uncomment to print out row values
    # puts row_cells.join " "
  end
  puts "Read #{num_rows} rows"
end

puts 'Done'
