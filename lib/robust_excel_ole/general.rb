include Enumerable

module General

  def absolute_path(file)   # :nodoc: #
    file = File.expand_path(file)
    file = RobustExcelOle::Cygwin.cygpath('-w', file) if RUBY_PLATFORM =~ /cygwin/
    WIN32OLE.new('Scripting.FileSystemObject').GetAbsolutePathName(file)
  end

  def canonize(filename)    # :nodoc: #
    raise TypeErrorREO, "No string given to canonize, but #{filename.inspect}" unless filename.is_a?(String)  
    normalize(filename).downcase
  end

  def normalize(path)       # :nodoc: #
    path = path.gsub('/./', '/') + '/'
    path = path.gsub(/[\/\\]+/, "/")
    nil while path.gsub!(/(\/|^)(?!\.\.?)([^\/]+)\/\.\.\//, '\1') 
    path = path.chomp("/")
    path
  end

  module_function :absolute_path, :canonize, :normalize

  class VBAMethodMissingError < RuntimeError  # :nodoc: #
  end

end

class WIN32OLE
end

class ::String    # :nodoc: #
  def / path_part
    if empty?
      path_part
    else
      if path_part.nil? or path_part.empty?
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

module MethodHelpers

  def respond_to?(meth_name, include_private = false)  # :nodoc: #    
    raise ObjectNotAlive, "respond_to?: #{self.class.name} not alive" unless alive?
    super
  end

  def methods   # :nodoc: # 
    (super + ole_object.ole_methods.map{|m| m.to_s}).uniq.select{|m| m =~ /^(?!\_)/}.sort
  end

end

REO = General
