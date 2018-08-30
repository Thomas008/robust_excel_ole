# example_concating.rb: 
# each named cell gets the value of cell right to it appended to its own value
# the new workbook's name is extended by the suffix "_concat"

require File.expand_path('../../lib/robust_excel_ole', File.dirname(__FILE__))
require File.join(File.dirname(File.expand_path(__FILE__)), '../../spec/helpers/create_temporary_dir')
require "fileutils"

include RobustExcelOle

begin
  dir = File.expand_path('../../spec/data', File.dirname(__FILE__))
  workbook_name = 'workbook_named.xls'
  ws = workbook_name.split(".")
  base_name = ws[0,ws.length-1].join(".")
  suffix = ws.last  
  file_name = dir + "/" + workbook_name
  extended_file_name = dir + "/" + base_name + "_concat" + "." + suffix
  FileUtils.copy file_name, extended_file_name 

  Workbook.unobtrusively(extended_file_name) do |book|
    book.each do |sheet|
      sheet.each do |cell|
        name = cell.Name.Name rescue nil
        if name
          cell.Value = cell.Value.to_s + cell.Offset(0,1).Value.to_s
          sheet.set_name(cell.Row, cell.Column, name)
        end
      end
    end
  end
end
