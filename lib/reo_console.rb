require 'pry'
require '../robust_excel_ole/lib/robust_excel_ole'

include REO
include General

# pry mit behalten lokaler Variablen

class Object

  def pry(object = nil, hash = {})
  	puts "pry!!!"
    @local_vars = { }
    if object.nil? || Hash === object # rubocop:disable Style/CaseEquality
      Pry.start(self, object || {})
    else
      Pry.start(object, hash)
    end
  end

end

class Pry

  def self.set_local_vars local_vars
    @local_vars = local_vars
  end
  
  def self.get_local_vars
    @local_vars
  end  

  def self.binding_for(target)
    return target if Binding === target # rubocop:disable Style/CaseEquality
    return TOPLEVEL_BINDING if Pry.main == target
    __bnd = target.instance_eval{ target.__binding__ }
    @local_vars.each{ |var,value| __bnd.local_variable_set(var, value) }
    #target.__binding__
    __bnd
  end

  class REPL

    def repl
      loop do
        case val = read
        when :control_c
          output.puts ""
          pry.reset_eval_string
        when :no_more_input
          output.puts "" if output.tty?
          break
        else
          output.puts "" if val.nil? && output.tty?                    
          # determine the local variables in the binding before evaluation
          bnd = pry.binding_stack.first
          exclude_vars = [:__, :_, :_dir, :_dir_, :_file, :_file_, :_in_, :_out_, :_ex, :_ex_, :pry_instance]
          local_vars = Pry.get_local_vars
          bnd.local_variables.each{ |var| local_vars[var] = bnd.local_variable_get(var) unless exclude_vars.include?(var) }
          Pry.set_local_vars(local_vars)
          return pry.exit_value unless pry.eval(val)
        end
      end
    end
  end
end


# some pry configuration
Pry.config.windows_console_warning = false
Pry.config.color = false
Pry.config.prompt_name = "REO "

#Pry.config.history_save = true
#Pry.editor = 'notepad'  # 'subl', 'vi'

Pry.config.prompt = Pry::Prompt.new(
  "REO",
  "The RobustExcelOle Prompt. Besides the standard information it puts the current object",
  [
   proc { |target_self, nest_level, pry|
  "[#{pry.input_ring.count}] #{pry.config.prompt_name}(#{Pry.view_clip(target_self)})#{":#{nest_level}" unless nest_level.zero?}> "
 },

 proc { |target_self, nest_level, pry|
  "[#{pry.input_ring.count}] #{pry.config.prompt_name}(#{Pry.view_clip(target_self)})#{":#{nest_level}" unless nest_level.zero?}* "
 }
]
)


hooks = Pry::Hooks.new

hooks.add_hook :when_started, :hook12 do
puts 'REO console started'
puts
end
Pry.start(nil, hooks: hooks)
