LOG_TO_STDOUT = false     unless Object.const_defined?(:LOG_TO_STDOUT)
REO_LOG_DIR   = ""        unless Object.const_defined?(:REO_LOG_DIR)
REO_LOG_FILE  = "reo.log" unless Object.const_defined?(:REO_LOG_FILE)
  
File.delete REO_LOG_FILE rescue nil

module Utilities  # :nodoc: #

  def trace(text)
    if LOG_TO_STDOUT 
      puts text
    else
      if REO_LOG_DIR.empty?
        homes = ["HOME", "HOMEPATH"]
        home = homes.find {|h| ENV[h] != nil}
        reo_log_dir = ENV[home]
      else
        reo_log_dir = REO_LOG_DIR
      end
      File.open(reo_log_dir + "/" + REO_LOG_FILE,"a") do | file |
        file.puts text
      end
    end
  end

  module_function :trace

end