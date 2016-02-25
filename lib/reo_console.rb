#require 'lib/robust_excel_ole'
include REO
include RobustExcelOle
include General

require 'irb/completion'
require 'irb/ext/save-history'

ARGV.concat [ "--readline",
              "--prompt-mode",
              "simple" ]

# 250 entries in the list
IRB.conf[:SAVE_HISTORY] = 250

# Store results in home directory with specified file name
#IRB.conf[:HISTORY_FILE] = "#{ENV['HOME']}/.irb-history"
IRB.conf[:HISTORY_FILE] = "#{ENV['HOME']}/.reo-history"

module Readline   # :nodoc: #
  module Hist     # :nodoc: #
    LOG = IRB.conf[:HISTORY_FILE]
#    LOG = "#{ENV['HOME']}/.irb-history"

    def self.write_log(line)
      File.open(LOG, 'ab') {|f| f << "#{line}
"}
    end

    def self.start_session_log
      timestamp = proc{ Time.now.strftime("%Y-%m-%d, %H:%M:%S")}
      class <<timestamp  # :nodoc: #
        alias to_s call
      end   
      write_log(          "###### session start: #{timestamp}")
      at_exit { write_log("###### session stop:  #{timestamp}") }
    end
  end

  alias :old_readline :readline
  def readline(*args)
    ln = old_readline(*args)
    begin
      Hist.write_log(ln)
    rescue
    end
    ln
  end
end

Readline::Hist.start_session_log
puts "REO console started"
