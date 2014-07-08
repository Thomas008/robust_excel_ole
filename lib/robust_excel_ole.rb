require "win32ole"
require File.join(File.dirname(__FILE__), 'robust_excel_ole/book')
require File.join(File.dirname(__FILE__), 'robust_excel_ole/sheet')
require File.join(File.dirname(__FILE__), 'robust_excel_ole/cell')
require File.join(File.dirname(__FILE__), 'robust_excel_ole/range')
require File.join(File.dirname(__FILE__), 'robust_excel_ole/cygwin') if RUBY_PLATFORM =~ /cygwin/
#+#require "robust_excel_ole/version"
require File.join(File.dirname(__FILE__), 'robust_excel_ole/version')

module RobustExcelOle

end
