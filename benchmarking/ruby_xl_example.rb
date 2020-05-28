
require 'rubyXL'

start_time = Time.now 

# ============================================
# ===========   Read Example   ===============
# ============================================

workbook = RubyXL::Parser.parse './sample_excel_files/xlsx_500_rows.xlsx'

worksheets = workbook.worksheets
puts "Found #{worksheets.count} worksheets"

worksheets.each do |worksheet|
  puts "Reading: #{worksheet.sheet_name}"
  num_rows = 0

  worksheet.each do |row|
    row_cells = row.cells.map{ |cell| cell.value }
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


puts 'Done'
