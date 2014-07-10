#'require 'lib/robust_excel_ole'
require File.join(File.dirname(__FILE__), './lib/robust_excel_ole')
book = RobustExcelOle::Book.open('./map1.xls', :read_only => false)
sheet = book[0]
cell = sheet[0,0]
puts "cell: #{cell}"
i = 0
sheet.each do |cell|
  i = i + 1
  puts "#{i}. Zelle: #{cell} Wert: #{cell.value}"
end
i = 0
sheet.each_row do |row|
  i = i + 1
  puts "#{i}. Reihe: #{row}"
end
i = 0
sheet.each_column do |column|
  i = i + 1
  puts "#{i}. Spalte: #{column}"
end
a = column_range[1]
#row = row_range[0]
#puts "row: #{r}"
book.save('./map2.xls', :if_exists => :overwrite)
book.save('./map2.xls', :if_exists => :excel)
book.save './map2.xls'
book.save
book.close
