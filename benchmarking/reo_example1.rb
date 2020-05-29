#require 'robust_excel_ole'
require File.join(File.dirname(__FILE__), '../lib/robust_excel_ole')

start_time = Time.now 

workbook = RobustExcelOle::Workbook.open './sample_excel_files/xlsx_50000_rows.xlsx'

puts "Found #{workbook.Worksheets.Count} worksheets"

workbook.each do |worksheet|
  puts "Reading: #{worksheet.name}"
  num_rows = 0

  worksheet.each_row do |row|
    row_cells = row.map{ |cell| cell.value }
    num_rows += 1

    # uncomment to print out row values
    # puts row_cells.join " "
  end
  puts "Read #{num_rows} rows"

end



end_time = Time.now
running_time = end_time - start_time
puts "time: #{running_time} sec."

puts 'Done'
