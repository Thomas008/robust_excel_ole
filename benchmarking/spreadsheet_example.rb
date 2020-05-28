
require 'spreadsheet'

start_time = Time.now 

# ============================================
# ===========   Read Example   ===============
# ============================================

# Note: spreadsheet only supports .xls files (not .xlsx)
workbook = Spreadsheet.open './sample_excel_files/xls_500_rows.xls'

worksheets = workbook.worksheets
puts "Found #{worksheets.count} worksheets"


worksheets.each do |worksheet|
  puts "Reading: #{worksheet.name}"
  num_rows = 0

  worksheet.rows.each do |row|
    row_cells = row.to_a.map{ |v| v.methods.include?(:value) ? v.value : v }
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
