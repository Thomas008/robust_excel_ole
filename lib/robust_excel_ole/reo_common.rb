LOG_TO_STDOUT = true     unless Object.const_defined?(:LOG_TO_STDOUT)
REO_LOG_DIR   = "" unless Object.const_defined?(:REO_LOG_DIR)
REO_LOG_FILE  = "reo.log" unless Object.const_defined?(:REO_LOG_FILE)
  
File.delete REO_LOG_FILE rescue nil

class REOCommon

  def excel
    raise TypeErrorREO, "receiver instance is neither an Excel nor a Book"
  end

  def own_methods
    (self.methods - Object.methods).sort
  end

  def trace(text)
    self.class.trace(text)
  end

  def self.tr1(text)
    puts :text
  end

  def self.trace(text)
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

  def puts_hash(hash)
    self.class.puts_hash(hash)
  end

  def self.puts_hash(hash)
    hash.each do |e|
      if e[1].is_a?(Hash)
        puts "#{e[0]} =>"
        e[1].each do |f|
          puts "  #{f[0]} => #{f[1]}"
        end
      else
        puts "#{e[0]} => #{e[1]}"
      end
    end
  end

end

module RobustExcelOle

  class REOError < RuntimeError                 # :nodoc: #
  end                 

  class ExcelError < REOError                   # :nodoc: #
  end

  class WorkbookError < REOError                # :nodoc: #
  end

  class FileError < REOError                    # :nodoc: #
  end

  class NamesError < REOError                   # :nodoc: #
  end

  class MiscError < REOError                    # :nodoc: #
  end

  class ExcelDamaged < ExcelError               # :nodoc: #
  end

  class UnsavedWorkbooks < ExcelError           # :nodoc: #
  end

  class WorkbookBlocked < WorkbookError         # :nodoc: #
  end

  class WorkbookNotSaved < WorkbookError        # :nodoc: #
  end

  class WorkbookReadOnly < WorkbookError        # :nodoc: #
  end

  class WorkbookBeingUsed < WorkbookError       # :nodoc: #
  end

  class FileNotFound < FileError                # :nodoc: #
  end

  class FileNameNotGiven < FileError            # :nodoc: #
  end

  class FileAlreadyExists < FileError           # :nodoc: #
  end

  class NameNotFound < NamesError               # :nodoc: #
  end

  class NameAlreadyExists < NamesError          # :nodoc: #
  end

  class RangeNotEvaluatable < MiscError         # :nodoc: #
  end

  class OptionInvalid < MiscError               # :nodoc: #
  end

  class ObjectNotAlive < MiscError              # :nodoc: #
  end

  class TypeErrorREO < REOError                 # :nodoc: #
  end   

  class TimeOut < REOError                      # :nodoc: #
  end  

  class UnexpectedError < REOError              # :nodoc: #
  end

end

