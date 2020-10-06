# -*- coding: utf-8 -*-

module General

  IS_JRUBY_PLATFORM = (RUBY_PLATFORM =~ /java/)
  ::EXPANDPATH_JRUBY_BUG   = IS_JRUBY_PLATFORM && true
  ::CONNECT_JRUBY_BUG      = IS_JRUBY_PLATFORM && true
  ::COPYSHEETS_JRUBY_BUG   = IS_JRUBY_PLATFORM && true
  ::ERRORMESSAGE_JRUBY_BUG = IS_JRUBY_PLATFORM && true
  ::CONNECT_EXCEL_JRUBY_BUG      = IS_JRUBY_PLATFORM && true
  ::RANGES_JRUBY_BUG       = IS_JRUBY_PLATFORM && true

  @private
  NetworkDrive = Struct.new(:drive_letter, :network_name) do

    def self.get_all(drives)
      ndrives = []
      count = drives.Count
      (0..(count - 1)).step(2) do |i|
        ndrives << NetworkDrive.new( drives.Item(i), drives.Item(i + 1).tr('\\','/'))
      end
      ndrives
    end

  end

  @private
  def hostnameshare2networkpath(filename)
    return filename unless filename[0,2] == "//"
    network = WIN32OLE.new('WScript.Network')
    drives = network.enumnetworkdrives
    network_drives = NetworkDrive.get_all(drives)
    f_c = filename.dup
    network_drive = network_drives.find do |d| 
      e = f_c.sub!(d.network_name,d.drive_letter)
      return e if e
    end    
    filename 
  end  

  # @private
  def absolute_path(file)    
    puts "absolute_path:"
    puts "file: #{file.inspect}"
    return file if file[0,2] == "//" 
    file[0,2] = './' if ::EXPANDPATH_JRUBY_BUG && file  =~ /[A-Z]:[^\/]/
    puts "file1: #{file.inspect}"
    file = File.expand_path(file)
    puts "file2: #{file.inspect}"
    file = RobustExcelOle::Cygwin.cygpath('-w', file) if RUBY_PLATFORM =~ /cygwin/
    puts "file3: #{file.inspect}"
    a = WIN32OLE.new('Scripting.FileSystemObject')
    puts "a: #{a.inspect}"

    WIN32OLE.new('Scripting.FileSystemObject').GetAbsolutePathName(file) #.tr('/','\\')
  end

  # @private
  def canonize(filename)
    raise TypeREOError, "No string given to canonize, but #{filename.inspect}" unless filename.is_a?(String)
    filename = hostnameshare2networkpath(filename)
    normalize(filename) if filename
  end

  # @private
  def normalize(path)  
    return unless path    
    path = path.gsub('/./', '/') + '/'
    path = path.gsub(/[\/\\]+/, '/')
    nil while path.gsub!(/(\/|^)(?!\.\.?)([^\/]+)\/\.\.\//, '\1')
    path = path.chomp('/')
    path
  end

  module_function :absolute_path, :canonize, :normalize

end

# @private
class Integer

  alias old_spaceship <=>

  def <=> other
    # p other
    if other.is_a? Array
      self <=> other.first
    else
      old_spaceship other
    end
  end

end

# @private
class Array

  alias old_spaceship <=>

  def <=> other
    # p other
    if other.is_a? Integer
      self <=> [other]
    else
      old_spaceship other
    end
  end

  def find_each_index find
    found, index, q = -1, -1, []
    while found
      found = self[index+1..-1].index(find)
      if found
        index = index + found + 1
        q << index
      end
    end
    q
  end
end

# @private
class WIN32OLE

  include Enumerable
  
  # type-lifting WIN32OLE objects to RobustExcelOle objects
  def to_reo
    class2method = [
      {Excel => :Hwnd},
      {Workbook => :FullName},
      {Worksheet => :UsedRange},
      {RobustExcelOle::Range => :Row},
      {ListObject => :ListRows}
    ]
    class2method.each do |element|
      classname = element.first.first
      method = element.first.last
      begin
        self.send(method)
        if classname == RobustExcelOle::Range && self.Rows.Count == 1 && self.Columns.Count == 1
          return Cell.new(self, self.Parent)
        else
          return classname.new(self)
        end
      rescue
        next
      end
    end
    raise TypeREOError, "given object cannot be type-lifted to a RobustExcelOle object"
  end

=begin
  def method_missing(name, *args, &blk)
    #super if name.to_s[0,1] =~ /[A-Z]/
    super if name == "Hwnd" or name =="FullName" or name == "UsedRange" or name == "Row" or name == "ListRows"
    begin
      reo_obj = self.to_reo
    rescue
      puts "error: #{$!.message}"
      raise # NoMethodError, "undefined method #{name.inspect} for #{self.inspect}"
    end
    reo_obj.send(name, *args, &blk)
  end
=end

end

# @private
class ::String 
  def / path_part
    if empty?
      path_part
    else
      if path_part.nil? || path_part.empty?
        self
      else
        begin
          File.join self, path_part
        rescue TypeError
          raise TypeError, "Only strings can be parts of paths (given: #{path_part.inspect} of class #{path_part.class})"
        end
      end
    end
  end

  # taken from http://apidock.com/rails/ActiveSupport/Inflector/underscore
  def underscore
    word = gsub('::', '/')
    word.gsub!(/([A-Z\d]+)([A-Z][a-z])/,'\1_\2')
    word.gsub!(/([a-z\d])([A-Z])/,'\1_\2')
    word.tr!('-', '_')
    word.downcase!
    word
  end

  def delete_multiple_underscores
    word = self
    while word.index('__') do
      word.gsub!('__','_')
    end    
    word
  end

  def replace_umlauts
    word = self
    word.gsub!('ä','ae')
    word.gsub!('Ä','Ae')
    word.gsub!('ö','oe')
    word.gsub!('Ö','Oe')
    word.gsub!('ü','ue')
    word.gsub!('Ü','Ue')
    #word.gsub!(/\x84/,'ae')
    #word.gsub!(/\x8E/,'Ae')
    #word.gsub!(/\x94/,'oe')
    #word.gsub!(/\x99/,'Oe')
    #word.gsub!(/\x81/,'ue')
    #word.gsub!(/\x9A/,'Ue')
    word
  end

  # taken from http://apidock.com/rails/ActiveSupport/Inflector/constantize
  # File activesupport/lib/active_support/inflector/methods.rb, line 226
  def constantize # (camel_cased_word)
    names = split('::')

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
# @private
class Module
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

module MethodHelpers

  # @private
  def respond_to?(meth_name, include_private = false) 
    if alive?
      methods.include?(meth_name.to_s)
    else
      super
    end
  end

  # @private
  def methods 
    if alive?
      (super.map { |m| m.to_s } + ole_object.ole_methods.map { |m| m.to_s }).uniq.select { |m| m =~ /^(?!\_)/ }.sort
    else
      super
    end
  end

end

REO = RobustExcelOle
