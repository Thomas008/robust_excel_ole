require 'pry'
require '../robust_excel_ole/lib/robust_excel_ole'

include REO
include General

# change the current binding such that self is the current object in the pry-instance, 
# preserve the local variables

class Pry

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

end

# some pry configuration
Pry.config.windows_console_warning = false
Pry.config.color = false
Pry.config.prompt_name = "REO "

#Pry.config.history_save = true
#Pry.editor = 'notepad'  # 'subl', 'vi'

prompt_proc1 = proc { |target_self, nest_level, pry|
   "[#{pry.input_ring.count}] #{pry.config.prompt_name}(#{Pry.view_clip(target_self.inspect)})#{":#{nest_level}" unless nest_level.zero?}> "
 }

prompt_proc2 =  proc { |target_self, nest_level, pry|
  "[#{pry.input_ring.count}] #{pry.config.prompt_name}(#{Pry.view_clip(target_self.inspect)})#{":#{nest_level}" unless nest_level.zero?}* "
 }

Pry.config.prompt = if RUBY_PLATFORM =~ /java/
  [prompt_proc1, prompt_proc2]
else
  Pry::Prompt.new(
    "REO",
    "The RobustExcelOle Prompt. Besides the standard information it puts the current object",
    [prompt_proc1, prompt_proc2]
    )
end

hooks = Pry::Hooks.new

hooks.add_hook :when_started, :hook12 do
puts 'REO console started'
puts
end
Pry.start(nil, hooks: hooks)
