require File.join(File.dirname(__FILE__), '../lib/robust_excel_ole')

start_time = Time.now 

workbook = RobustExcelOle::Workbook.open './sample_excel_files/xlsx_500_rows.xlsx'

puts "Found #{workbook.worksheets_count} worksheets"

workbook.each do |worksheet|
  puts "Reading: #{worksheet.name}"
  num_rows = 0

  worksheet.values.each do |row_vals|
    row_cells = row_vals
    num_rows += 1
  end

  puts "Read #{num_rows} rows"

end

end_time = Time.now
running_time = end_time - start_time
puts "time: #{running_time} sec."

puts 'Done'
