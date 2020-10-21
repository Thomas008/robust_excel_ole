# -*- coding: utf-8 -*-

module General

  IS_JRUBY_PLATFORM = (RUBY_PLATFORM =~ /java/)
  ::EXPANDPATH_JRUBY_BUG    = IS_JRUBY_PLATFORM && true
  ::CONNECT_JRUBY_BUG       = IS_JRUBY_PLATFORM && true
  ::COPYSHEETS_JRUBY_BUG    = IS_JRUBY_PLATFORM && true
  ::ERRORMESSAGE_JRUBY_BUG  = IS_JRUBY_PLATFORM && true
  ::CONNECT_EXCEL_JRUBY_BUG = IS_JRUBY_PLATFORM && true
  ::RANGES_JRUBY_BUG        = IS_JRUBY_PLATFORM && true

  # @private
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

  # @private
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
    return file if file[0,2] == "//" 
    file[0,2] = './' if ::EXPANDPATH_JRUBY_BUG && file  =~ /[A-Z]:[^\/]/
    file = File.expand_path(file)
    file = RobustExcelOle::Cygwin.cygpath('-w', file) if RUBY_PLATFORM =~ /cygwin/
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

  def change_current_binding(current_object)
    Pry.change_current_binding(current_object)
  end

  def class2method
    [{Excel => :Hwnd},
     {Workbook => :FullName},
     {Worksheet => :UsedRange},
     {RobustExcelOle::Range => :Row},
     {ListObject => :ListRows}]
  end


  # enable RobustExcelOle methods to Win32Ole objects
  def uplift_to_reo
    exclude_list = [:each, :inspect]
    class2method.each do |element|
      classname = element.first.first
      method = element.first.last
      classname.instance_methods(false).each do |inst_method|
        if !exclude_list.include?(inst_method)
          WIN32OLE.send(:define_method, inst_method) do |*args, &blk|  
            self.to_reo.send(inst_method, *args, &blk) 
          end
        end
      end
    end
    nil
  end

  module_function :absolute_path, :canonize, :normalize, :change_current_binding, :class2method, :uplift_to_reo

end

# @private
class Pry

  # change the current binding such that self is the current object in the pry-instance, 
  # preserve the local variables

  class << self
    attr_accessor :pry_instance
  end

  def self.change_current_binding(current_object)
    pry_instance = self.pry_instance
    old_binding = pry_instance.binding_stack.pop
    pry_instance.push_binding(current_object.__binding__)
    exclude_vars = [:__, :_, :_dir, :_dir_, :_file, :_file_, :_in_, :_out_, :_ex, :_ex_, :pry_instance]
    old_binding.local_variables.each do |var|
      pry_instance.add_sticky_local(var) {old_binding.local_variable_get(var)} unless exclude_vars.include?(var)
    end
    self.pry_instance = pry_instance
    nil
  end

  def push_initial_binding(target = nil)
    # memorize the current pry instance
    self.class.pry_instance = self 
    push_binding(target || Pry.toplevel_binding)
  end

  class REPL

    def read
      @indent.reset if pry.eval_string.empty?
      current_prompt = pry.select_prompt
      indentation = pry.config.auto_indent ? @indent.current_prefix : ''
      val = read_line("#{current_prompt}#{indentation}")
      # Return nil for EOF, :no_more_input for error, or :control_c for <Ctrl-C>
      return val unless val.is_a?(String)
      if pry.config.auto_indent
        original_val = "#{indentation}#{val}"
        indented_val = @indent.indent(val)

        if output.tty? &&
           pry.config.correct_indent &&
           Pry::Helpers::BaseHelpers.use_ansi_codes?
          # avoid repeating read line

          #output.print @indent.correct_indentation(
          #  current_prompt,
          #  indented_val,
          #  calculate_overhang(current_prompt, original_val, indented_val)
          #)
          output.flush
        end
      else
        indented_val = val
      end
      indented_val
    end
  end
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
    General.class2method.each do |element|
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
  def to_reo
    case ole_type.name
    when 'Range' then RobustExcelOle::Range.new(self)
    when '_Worksheet' then RobustExcelOle::Worksheet.new(self)
    when '_Workbook' then RobustExcelOle::Workbook.new(self)
    when '_Application' then RobustExcelOle::Excel.new(self)
    else
      self
    end
  end
=end

=begin
  alias method_missing_before_implicit_typelift method_missing 
  
  def method_missing(name, *args, &blk)
    puts "method_missing:"
    puts "name: #{name.inspect}"
    #raise NoMethodError if name.to_s == "Hwnd" or name.to_s == "FullName" or name.to_s == "UsedRange" or name.to_s == "Row" or name.to_s == "ListRows"
    begin
      reo_obj = self.to_reo
      puts "reo_obj: #{reo_obj.inspect}"
    rescue
      puts "$!.message: #{$!.message}"
      method_missing_before_implicit_typelift(name, *args, &blk)
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
