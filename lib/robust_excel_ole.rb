require 'win32ole'
require File.join(File.dirname(__FILE__), 'robust_excel_ole/reo_common')
require File.join(File.dirname(__FILE__), 'robust_excel_ole/range_owners')
require File.join(File.dirname(__FILE__), 'robust_excel_ole/address')
require File.join(File.dirname(__FILE__), 'robust_excel_ole/general')
require File.join(File.dirname(__FILE__), 'robust_excel_ole/excel')
require File.join(File.dirname(__FILE__), 'robust_excel_ole/bookstore')
require File.join(File.dirname(__FILE__), 'robust_excel_ole/workbook')
require File.join(File.dirname(__FILE__), 'robust_excel_ole/worksheet')
require File.join(File.dirname(__FILE__), 'robust_excel_ole/cell')
require File.join(File.dirname(__FILE__), 'robust_excel_ole/range')
require File.join(File.dirname(__FILE__), 'robust_excel_ole/cygwin') if RUBY_PLATFORM =~ /cygwin/
#+#require "robust_excel_ole/version"
require File.join(File.dirname(__FILE__), 'robust_excel_ole/version')
