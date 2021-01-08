if RUBY_PLATFORM =~ /java/
  require 'jruby-win32ole'
else
  require 'win32ole'
end

require_relative 'robust_excel_ole/general'
require_relative 'robust_excel_ole/base'
require_relative 'robust_excel_ole/vba_objects'
require_relative 'robust_excel_ole/range_owners'
require_relative 'robust_excel_ole/address_tool'
require_relative 'robust_excel_ole/excel'
require_relative 'robust_excel_ole/bookstore'
require_relative 'robust_excel_ole/workbook'
require_relative 'robust_excel_ole/worksheet'
require_relative 'robust_excel_ole/cell'
require_relative 'robust_excel_ole/range'
require_relative 'robust_excel_ole/list_row'
require_relative 'robust_excel_ole/list_object'
require_relative 'robust_excel_ole/cygwin' if RUBY_PLATFORM =~ /cygwin/
require_relative 'robust_excel_ole/version'

General.init_reo_for_win32ole
