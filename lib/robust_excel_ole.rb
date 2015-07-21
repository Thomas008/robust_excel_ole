require "win32ole"
require File.join(File.dirname(__FILE__), 'robust_excel_ole/excel')
require File.join(File.dirname(__FILE__), 'robust_excel_ole/book_store')
require File.join(File.dirname(__FILE__), 'robust_excel_ole/book')
require File.join(File.dirname(__FILE__), 'robust_excel_ole/sheet')
require File.join(File.dirname(__FILE__), 'robust_excel_ole/cell')
require File.join(File.dirname(__FILE__), 'robust_excel_ole/range')
require File.join(File.dirname(__FILE__), 'robust_excel_ole/cygwin') if RUBY_PLATFORM =~ /cygwin/
#+#require "robust_excel_ole/version"
require File.join(File.dirname(__FILE__), 'robust_excel_ole/version')

REO = RobustExcelOle

module RobustExcelOle

  def absolute_path(file)
    file = File.expand_path(file)
    file = RobustExcelOle::Cygwin.cygpath('-w', file) if RUBY_PLATFORM =~ /cygwin/
    WIN32OLE.new('Scripting.FileSystemObject').GetAbsolutePathName(file)
  end

  def canonize(filename)
    raise "No string given to canonize, but #{filename.inspect}" unless filename.is_a?(String)
    filename.downcase rescue nil
  end

  module_function :absolute_path, :canonize

  class VBAMethodMissingError < RuntimeError  # :nodoc: #
  end

end

class Object
  def excel
    raise ExcelErrorOpen, "provided instance is neither an Excel nor a Book"
  end

end
