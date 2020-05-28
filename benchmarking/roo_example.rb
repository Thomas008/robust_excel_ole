require 'roo'

# ============================================
# ===========   Read Example   ===============
# ============================================

start_time = Time.now

workbook = Roo::Spreadsheet.open './sample_excel_files/xlsx_500_rows.xlsx'

worksheets = workbook.sheets
puts "Found #{worksheets.count} worksheets"

worksheets.each do |worksheet|
  puts "Reading: #{worksheet}"
  num_rows = 0

  workbook.sheet(worksheet).each_row_streaming do |row|
    row_cells = row.map { |cell| cell.value }
    #puts row_cells.inspect
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
