require 'pry'

include REO
include General

puts 'REO console started'
puts 

# some pry configuration
Pry.config.windows_console_warning = false
Pry.config.history.should_save = true
Pry.config.color = false
#Pry.editor = 'notepad'  # 'subl', 'vi'
#Pry.config.prompt =
#  [
#    ->(_obj, _nest_level, _) { ">> " },
#    ->(*) { "  " }
#  ]

pry
