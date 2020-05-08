if RUBY_PLATFORM =~ /java/
  require 'jruby-win32ole'
else
  require 'win32ole'
end
require File.join(File.dirname(__FILE__), 'robust_excel_ole/base')
require File.join(File.dirname(__FILE__), 'robust_excel_ole/vba_objects')
require File.join(File.dirname(__FILE__), 'robust_excel_ole/range_owners')
require File.join(File.dirname(__FILE__), 'robust_excel_ole/address_tool')
require File.join(File.dirname(__FILE__), 'robust_excel_ole/general')
require File.join(File.dirname(__FILE__), 'robust_excel_ole/excel')
require File.join(File.dirname(__FILE__), 'robust_excel_ole/bookstore')
require File.join(File.dirname(__FILE__), 'robust_excel_ole/workbook')
require File.join(File.dirname(__FILE__), 'robust_excel_ole/worksheet')
require File.join(File.dirname(__FILE__), 'robust_excel_ole/cell')
require File.join(File.dirname(__FILE__), 'robust_excel_ole/range')
require File.join(File.dirname(__FILE__), 'robust_excel_ole/cygwin') if RUBY_PLATFORM =~ /cygwin/
require File.join(File.dirname(__FILE__), 'robust_excel_ole/version')

include RobustExcelOle
include General
