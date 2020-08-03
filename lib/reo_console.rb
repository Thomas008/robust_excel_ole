require 'pry'
require '../robust_excel_ole/lib/robust_excel_ole'

include REO
include General

# some pry configuration
Pry.config.windows_console_warning = false
#Pry.config.history_save = true
Pry.config.color = false
#Pry.editor = 'notepad'  # 'subl', 'vi'
#Pry.config.prompt =
#[
#->(_obj, _nest_level, _) { ">> " },
#->(*) { "  " }
#]

hooks = Pry::Hooks.new

hooks.add_hook :when_started, :hook12 do
puts 'REO console started'
puts
end
#{when_started: -> {puts 'Hello'} }
Pry.start(nil, hooks: hooks)
