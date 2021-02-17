
require 'simple_xlsx_reader'

start_time = Time.now 

# ============================================
# ===========   Read Example   ==============
# ============================================

workbook = SimpleXlsxReader.open './sample_excel_files/xlsx_500_rows.xlsx'

worksheets = workbook.sheets
puts "Found #{worksheets.size} worksheets"

worksheets.each do |worksheet|
  puts "Reading: #{worksheet.name}"
  num_rows = 0

  worksheet.rows.each do |row|
    row_cells = row
    num_rows += 1

    # uncomment to print out row values
    # puts row_cells.join ' '
  end
  puts "Read #{num_rows} rows"
end

end_time = Time.now
running_time = end_time - start_time
puts "time: #{running_time} sec."

puts 'Done'
