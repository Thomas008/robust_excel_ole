require "win32ole"
require File.join(File.dirname(__FILE__), 'robust_excel_ole/utilities')
require File.join(File.dirname(__FILE__), 'robust_excel_ole/excel')
require File.join(File.dirname(__FILE__), 'robust_excel_ole/bookstore')
require File.join(File.dirname(__FILE__), 'robust_excel_ole/book')
require File.join(File.dirname(__FILE__), 'robust_excel_ole/sheet')
require File.join(File.dirname(__FILE__), 'robust_excel_ole/cell')
require File.join(File.dirname(__FILE__), 'robust_excel_ole/range')
require File.join(File.dirname(__FILE__), 'robust_excel_ole/cygwin') if RUBY_PLATFORM =~ /cygwin/
#+#require "robust_excel_ole/version"
require File.join(File.dirname(__FILE__), 'robust_excel_ole/version')

REO = RobustExcelOle

include Enumerable

module RobustExcelOle  

  def rot
    # allocate 4 bytes to store a pointer to the IRunningObjectTable object
    irot_ptr = 0.chr * 4      # or [0].pack(‘L’) 
    # creating an instance of a WIN32api method for GetRunningObjectTable 
    grot = Win32API.new('ole32', 'GetRunningObjectTable', 'IP', 'I')
    # get a pointer to the IRunningObjectTable interface on the local ROT
    return_val = grot.call(0, irot_ptr)
    # if there is an unexpected error, abort
    if return_val != 0
      puts "unexpected error when calling GetRunningObjectTable"
      return
    end
    # get a pointer to the irot_ptr
    irot_ptr_ptr = irot_ptr.unpack('L').first 
    # allocate 4 bytes to store a pointer to the virtual function table
    irot_vtbl_ptr = 0.chr * 4    # or irot_vtbl_ptr = [0].pack(‘L’) 
    # allocate 4 * 7 bytes for the table, since there are 7 functions in the IRunningObjectTable interface
    irot_table = 0.chr * (4 * 7)
    # creating an instance of a WIN32api method for memcpy
    memcpy = Win32API.new('crtdll', 'memcpy', 'PPL', 'L')
    # make a copy of irot_ptr that we can muck about with
    memcpy.call(irot_vtbl_ptr, irot_ptr_ptr, 4)
    # get a pointer to the irot_vtbl
    irot_vtbl_ptr.unpack('L').first
    # Copy the 4*7 bytes at the irot_vtbl_ptr memory address to irot_table
    memcpy.call(irot_table, irot_vtbl_ptr.unpack('L').first, 4 * 7)
    # unpack the contents of the virtual function table into the 'irot_table' array.
    irot_table = irot_table.unpack('L*')
    puts "Number of elements in the vtbl is: " + irot_table.length.to_s
    # EnumRunning is the 1st function in the vtbl.  
    enumRunning = Win32::API::Function.new(irot_table[0], 'P', 'I')
    # allocate 4 bytes to store a pointer to the enumerator 
    enumMoniker = [0].pack('L') # or 0.chr * 4
    # create a pointer to the enumerator
    return_val_er = enumRunning.call(enumMoniker)
  end

  def absolute_path(file)
    file = File.expand_path(file)
    file = RobustExcelOle::Cygwin.cygpath('-w', file) if RUBY_PLATFORM =~ /cygwin/
    WIN32OLE.new('Scripting.FileSystemObject').GetAbsolutePathName(file)
  end

  def canonize(filename)
    raise ExcelError, "No string given to canonize, but #{filename.inspect}" unless filename.is_a?(String)  
    normalize(filename).downcase rescue nil
  end

  def normalize(path)
    path = path.gsub('/./', '/') + '/'
    path = path.gsub(/[\/\\]+/, "/")
    nil while path.gsub!(/(\/|^)(?!\.\.?)([^\/]+)\/\.\.\//, '\1') 
    path = path.chomp("/")
    path
  end

  module_function :absolute_path, :canonize, :rot

  class VBAMethodMissingError < RuntimeError  # :nodoc: #
  end

  #module RobustExcelOle::Utilites  # :nodoc: #

  #end
end

class Object      # :nodoc: #
  def excel
    raise ExcelError, "receiver instance is neither an Excel nor a Book"
  end
end

class WIN32OLE
end

class ::String    # :nodoc: #
  def / path_part
    if empty?
      path_part
    else
      begin 
        File.join self, path_part
      rescue TypeError
        raise "Only strings can be parts of paths (given: #{path_part.inspect} of class #{path_part.class})"
      end
    end
  end

  # taken from http://apidock.com/rails/ActiveSupport/Inflector/underscore
  def underscore
    word = gsub('::', '/')
    word.gsub!(/([A-Z\d]+)([A-Z][a-z])/,'\1_\2')
    word.gsub!(/([a-z\d])([A-Z])/,'\1_\2')
    word.tr!("-", "_")
    word.downcase!
    word
  end

  # taken from http://apidock.com/rails/ActiveSupport/Inflector/constantize
  # File activesupport/lib/active_support/inflector/methods.rb, line 226
  def constantize #(camel_cased_word)
    names = self.split('::')

    # Trigger a builtin NameError exception including the ill-formed constant in the message.
    Object.const_get(self) if names.empty?

    # Remove the first blank element in case of '::ClassName' notation.
    names.shift if names.size > 1 && names.first.empty?

    names.inject(Object) do |constant, name|
      if constant == Object
        constant.const_get(name)
      else
        candidate = constant.const_get(name)
        next candidate if constant.const_defined?(name)
        next candidate unless Object.const_defined?(name)

        # Go down the ancestors to check it it's owned
        # directly before we reach Object or the end of ancestors.
        constant = constant.ancestors.inject do |const, ancestor|
          break const    if ancestor == Object
          break ancestor if ancestor.const_defined?(name)
          const
        end

        # owner is in Object, so raise
        constant.const_get(name)
      end
    end
  end
end

# taken from http://api.rubyonrails.org/v2.3.8/classes/ActiveSupport/CoreExtensions/Module.html#M000806
class Module   # :nodoc: #
  def parent_name
    unless defined? @parent_name
      @parent_name = name =~ /::[^:]+\Z/ ? $`.freeze : nil
    end
    @parent_name
  end
  def parent
    parent_name ? parent_name.constantize : Object
  end
end

