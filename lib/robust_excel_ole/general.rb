# -*- coding: utf-8 -*-
require 'pathname'

module ToReoRefinement

  refine WIN32OLE do

    # type-lifting WIN32OLE objects to RobustExcelOle objects
    def to_reo
      General.main_classes_ole_types_and_recognising_methods.each do |classname, ole_type, methods|
        if !::OLETYPE_JRUBY_BUG
          if self.ole_type.name == ole_type
            if classname != RobustExcelOle::Range
              return classname.new(self)
            elsif self.Rows.Count == 1 && self.Columns.Count == 1
              return RobustExcelOle::Cell.new(self, self.Parent)
            else
              return RobustExcelOle::Range.new(self, self.Parent)
            end
          end
        else
          begin
            recognising_method, no_method = methods
            self.send(recognising_method)
            unless no_method.nil?
              begin
                self.send(no_method[:no_method])
                next
              rescue NoMethodError
                return classname.new(self)
              end
            end
            if classname != RobustExcelOle::Range
              return classname.new(self)
            elsif self.Rows.Count == 1 && self.Columns.Count == 1
              return RobustExcelOle::Cell.new(self, self.Parent)
            else
              return RobustExcelOle::Range.new(self, self.Parent)
            end
          rescue Java::OrgRacobCom::ComFailException => msg # NoMethodError
            #if $!.message =~ /undefined method/ && 
            #  main_classes_ole_types_and_recognising_methods.any?{ |_c, _o, recognising_method| $!.message.include?(recognising_method.to_s) }
              next
            #end 
          end
        end
      end
      raise RobustExcelOle::TypeREOError, "given object cannot be type-lifted to a RobustExcelOle object"
    end

  end
end

# @private
class WIN32OLE

  include Enumerable

end

# @private
module FindAllIndicesRefinement

  refine Array do

    def find_all_indices elem
      elem = elem.encode('utf-8') if elem.respond_to?(:gsub)      
      found, index, result = -1, -1, []
      while found
        found = self[index+1..-1].index(elem)
        if found
          index = index + found + 1
          result << index
        end
      end
      result
    end

  end

end

TRANSLATION_TABLE = {
      'ä' => 'ae', 'ö' => 'oe', 'ü' => 'ue', 'Ä' => 'Ae', 'Ö' => 'Oe', 'Ü' => 'Ue',
      'ß' => 'ss', '²' => '2', '³' => '3'
    }

# @private
module StringRefinement

  refine String do

    def / path_part
      return path_part if empty?
      return self if path_part.nil? || path_part.empty?
      begin
        path_part = path_part.strip
        (path_part  =~ /^(\/|([A-Z]:\/))/i) ? path_part : (chomp('/') + '/' + path_part)
      rescue TypeError
        raise TypeError, "Only strings can be parts of paths (given: #{path_part.inspect} of class #{path_part.class})"
      end
    end    

    def replace_umlauts
      TRANSLATION_TABLE.inject(encode('utf-8')) { |word,(umlaut, replacement)| word.gsub(umlaut, replacement) }
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
end

# @private
module ParentRefinement

  using StringRefinement

  # taken from http://api.rubyonrails.org/v2.3.8/classes/ActiveSupport/CoreExtensions/Module.html#M000806
  refine Module do

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
end

class Integer

  alias old_spaceship <=>

  def <=> other
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


module General

  using ToReoRefinement

  IS_JRUBY_PLATFORM = (RUBY_PLATFORM =~ /java/)
  ::EXPANDPATH_JRUBY_BUG    = IS_JRUBY_PLATFORM && true
  ::CONNECT_JRUBY_BUG       = IS_JRUBY_PLATFORM && true
  ::COPYSHEETS_JRUBY_BUG    = IS_JRUBY_PLATFORM && true
  ::ERRORMESSAGE_JRUBY_BUG  = IS_JRUBY_PLATFORM && true
  ::CONNECT_EXCEL_JRUBY_BUG = IS_JRUBY_PLATFORM && true
  ::RANGES_JRUBY_BUG        = IS_JRUBY_PLATFORM && true
  ::OLETYPE_JRUBY_BUG       = IS_JRUBY_PLATFORM && true

  # @private
  NetworkDrive = Struct.new(:drive_letter, :network_name) do

    def self.get_all_drives
      network = WIN32OLE.new('WScript.Network')
      drives = network.enumnetworkdrives
      count = drives.Count
      # (0..(count - 1)).step(2).map{ |i| NetworkDrive.new( drives.Item(i), drives.Item(i + 1).tr('\\','/')) }      
      result = (0..(count - 1)).step(2).map { |i| 
        NetworkDrive.new( drives.Item(i), drives.Item(i + 1).tr('\\','/')) unless drives.Item(i).empty?
      }.compact
      result
    end
  end

  # @private
  def hostnameshare2networkpath(filename)
    return filename unless filename[0,2] == "//"
    hostname = filename[0,filename[3,filename.length].index('/')+3]
    filename_wo_hostname = filename[hostname.length+1,filename.length]
    abs_filename = absolute_path(filename_wo_hostname).tr('\\','/').sub('C:/','c$/')
    adapted_filename = hostname + "/" + abs_filename
    NetworkDrive.get_all_drives.each do |d|
      new_filename = filename.sub(/#{(Regexp.escape(d.network_name))}/i,d.drive_letter)
      return new_filename if new_filename != filename
      new_filename = adapted_filename.sub(/#{(Regexp.escape(d.network_name))}/i,d.drive_letter)
      return new_filename if new_filename != filename
    end
    filename
  end
  
  # @private
  def absolute_path(file)
    file = file.to_path if file.respond_to?(:to_path)
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

  # @private
  def change_current_binding(current_object)
    Pry.change_current_binding(current_object)
  end

  # @private
  def main_classes_ole_types_and_recognising_methods
    [[RobustExcelOle::Range     , 'Range'       , :Row],
     [RobustExcelOle::Worksheet , '_Worksheet'  , :UsedRange],
     [RobustExcelOle::Workbook  , '_Workbook'   , :FullName],
     [RobustExcelOle::Excel     , '_Application', :Hwnd],
     [RobustExcelOle::ListObject, 'ListObject' , :ListRows],
     [RobustExcelOle::ListRow   , 'ListRow'    , [:Creator, :no_method => :Row]]]
  end

  WIN32OLE_INSTANCE_METHODS = [
    :ole_methods, :ole_free, :ole_get_methods, :ole_put_methods, :ole_func_methods, :ole_method, :ole_method_help,
    :ole_activex_initialize, :ole_type, :ole_obj_help, :ole_typelib, :ole_query_interface, :ole_respond_to?, 
    :invoke, :_invoke, :_getproperty, :_setproperty, :setproperty, :[], :[]=, :methods, :method_missing, :each
  ]

  # @private
  # enable RobustExcelOle methods to Win32Ole objects
  def init_reo_for_win32ole
    method_occurrences = {}
    main_classes_ole_types_and_recognising_methods.each do |classname, _ole_type, _recognising_method|
      meths = (classname.instance_methods(false) - WIN32OLE_INSTANCE_METHODS - Object.methods - Enumerable.instance_methods(false) - [:Calculation=])
      meths.each do |inst_method|
        method_occurrences[inst_method] = method_occurrences[inst_method] ? :several_classes : classname
      end
    end
    method_occurrences.each do |inst_method, class_name|
      if WIN32OLE.method_defined?(inst_method)
        aliased_method = "#{inst_method}_after_reo".to_s.to_sym
        WIN32OLE.send(:alias_method, aliased_method, inst_method)
      else
        aliased_method = nil
      end
      if aliased_method || class_name == :several_classes
        WIN32OLE.send(:define_method, inst_method) do |*args, &blk|  
          begin 
            obj = to_reo                        
          rescue     
            sending_method = aliased_method ? aliased_method : inst_method.capitalize
            return self.send(sending_method, *args, &blk)
          end
          obj.send(inst_method, *args, &blk)
        end
      else       
        WIN32OLE.send(:define_method, inst_method) do |*args, &blk|  
          begin 
            obj = class_name.new(self)     
          rescue
             sending_method = aliased_method ? aliased_method : inst_method.capitalize
            return self.send(sending_method, *args, &blk)
          end
          obj.send(inst_method, *args, &blk) 
        end
      end
    end
  end

  module_function :absolute_path, :canonize, :normalize, :change_current_binding, 
                  :main_classes_ole_types_and_recognising_methods, 
                  :init_reo_for_win32ole, :hostnameshare2networkpath, :test

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
