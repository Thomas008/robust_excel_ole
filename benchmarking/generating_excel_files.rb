require_relative '../lib/robust_excel_ole'

include RobustExcelOle

begin

  filename = './large_excel_files/xlsx_10_rows.xlsx'
  Excel.kill_all
 
  Workbook.open(filename, :if_absent => :create) do |workbook|
    # default: 3 worksheets
    num_worksheets = 1
    workbook.each do |worksheet|
      puts "worksheet-number #{num_worksheets}"
      num_rows = 1
      (1..10).each do |row|
        puts "row-number: #{num_rows}"
        (1..10).each do |column|
          worksheet[num_rows,column].value = (num_rows-1)*10+column
        end
        num_rows += 1
      end
      num_worksheets += 1
    end
    workbook.save
  end

end