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
  def network2hostnamesharepath(filename)
    network = WIN32OLE.new('WScript.Network')
    drives = network.enumnetworkdrives
    drive_letter, filename_after_drive_letter = filename.split(':') 
    # if filename starts with a drive letter not c and this drive exists,
    # then determine the corresponding host_share_path
    default_drive = File.absolute_path(".")[0]
    if drive_letter != default_drive && drive_letter != filename  
      for i in 0 .. drives.Count-1
        next if i % 2 == 1
        if drives.Item(i).gsub(':','') == drive_letter
          hostname_share = drives.Item(i+1)  #.gsub('\\','/').gsub('//','')
          break
        end
      end
      hostname_share + filename_after_drive_letter if hostname_share
    else
      return filename
    end
  end

  # @private
  def absolute_path(file)     
    file[0,2] = './' if ::EXPANDPATH_JRUBY_BUG && file  =~ /[A-Z]:[^\/]/
    file = File.expand_path(file)
    file = RobustExcelOle::Cygwin.cygpath('-w', file) if RUBY_PLATFORM =~ /cygwin/
    WIN32OLE.new('Scripting.FileSystemObject').GetAbsolutePathName(file).tr('/','\\')
  end

  # @private
  def canonize(filename)
    raise TypeREOError, "No string given to canonize, but #{filename.inspect}" unless filename.is_a?(String)
    filename = network2hostnamesharepath(filename)
    normalize(filename).downcase
  end

  # @private
  def normalize(path)      
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

end


# @private
class WIN32OLE

  include Enumerable
  
  # type-lifting WIN32OLE objects to RobustExcelOle objects
  def to_reo
    class2method = [
      {Excel => :Hwnd},
      {Workbook => :FullName},
      {Worksheet => :Copy},
      {RobustExcelOle::Range => :Row},
      {ListObject => :ListRows}
    ]
    class2method.each do |element|
      classname = element.first.first
      method = element.first.last
      begin
        self.send(method)
        if classname == RobustExcelOle::Range && self.Rows.Count == 1 && self.Columns.Count == 1
          return Cell.new(self)
        else
          return classname.new(self)
        end
      rescue
        next
      end
    end
    raise TypeREOError, "given object cannot be type-lifted to a RobustExcelOle object"
  end

  alias method_missing_before_implicit_typelift method_missing 
  def xx_method_missing(name, *args, &blk)
    begin
      reo_obj = self.to_reo
    rescue
      puts "$!.message: #{$!.message}"
      method_missing_before_implicit_typelift(name, *args, &blk)
    end
    reo_obj.send(name, *args, &blk)
  end
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
