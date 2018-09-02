# example_naming.rb: 
# each cell is named with the name equaling its value unless it is empty or not a string
# the contents of each cell is copied
# the new workbook's name is extended by the suffix "_named"

require File.expand_path('../../lib/robust_excel_ole', File.dirname(__FILE__))
require File.join(File.dirname(File.expand_path(__FILE__)), '../../spec/helpers/create_temporary_dir')
require "fileutils"

include RobustExcelOle

begin
  dir = File.expand_path('../../spec/data', File.dirname(__FILE__))
  workbook_name = 'workbook.xls'
  ws = workbook_name.split(".")
  base_name = ws[0,ws.length-1].join(".")
  suffix = ws.last  
  #base_name = File.basename(workbook_name, ".xls*")
  #workbook_name =~ /^(.*)\.[^.]$/; base_name = $1
  #base_name = workbook_name.sub(/^.*(\.[^.])$/; '')
  #base_name = workbook_name[0,workbook_name.rindex('.')]
  #suffix = workbook_name[workbook_name.rindex('.')+1,workbook_name.length]
  #suffix = workbook_name.scan(/\.[^.\/]+$/).last
  file_name = dir + "/" + workbook_name
  extended_file_name = dir + "/" + base_name + "_named" + "." + suffix
  FileUtils.copy file_name, extended_file_name 

  Workbook.unobtrusively(extended_file_name) do |book|     
    book.each do |sheet|
      sheet.each do |cell_orig|
        contents = cell_orig.Value
        sheet.add_name(contents, cell_orig.Row, cell_orig.Column) if contents && contents.is_a?(String)
      end
    end
  end
end
