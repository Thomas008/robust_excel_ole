# -*- coding: utf-8 -*-

LOG_TO_STDOUT = true     unless Object.const_defined?(:LOG_TO_STDOUT)
REO_LOG_DIR   = "" unless Object.const_defined?(:REO_LOG_DIR)
REO_LOG_FILE  = "reo.log" unless Object.const_defined?(:REO_LOG_FILE)
  
File.delete REO_LOG_FILE rescue nil

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

include RobustExcelOle

class REOCommon

  def excel
    raise TypeErrorREO, "receiver instance is neither an Excel nor a Book"
  end

  def own_methods
    (self.methods - Object.methods).sort
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

class RangeOwners < REOCommon

    # returns the contents of a range with given name
  # evaluates formula contents of the range is a formula
  # if no contents could be returned, then return default value, if provided, raise error otherwise
  # Excel Bug: if a local name without a qualifier is given, then by default Excel takes the first worksheet,
  #            even if a different worksheet is active
  # @param  [String]      name      the name of the range
  # @param  [Hash]        opts      the options
  # @option opts [Symbol] :default  the default value that is provided if no contents could be returned
  # @return [Variant] the contents of a range with given name
  def nameval(name, opts = {:default => nil})
    name_obj = name_object(name)
    value = begin
      name_obj.RefersToRange.Value
    rescue  WIN32OLERuntimeError
      #begin
      #  self.sheet(1).Evaluate(name_obj.Name)
      #rescue WIN32OLERuntimeError
      return opts[:default] if opts[:default]
      raise RangeNotEvaluatable, "cannot evaluate range named #{name.inspect} in #{File.basename(workbook.stored_filename).inspect rescue nil}"
      #end
    end
    if value.is_a?(Bignum)  #RobustExcelOle::XlErrName  
      return opts[:default] if opts[:default]
      raise RangeNotEvaluatable, "cannot evaluate range named #{name.inspect} in #{File.basename(workbook.stored_filename).inspect rescue nil}"
    end 
    return opts[:default] if opts[:default] && value.nil?
    value      
  end

  # sets the contents of a range
  # @param [String]  name  the name of a range
  # @param [Variant] value the contents of the range
  # @param [FixNum]  color the color when setting a value
  # @param [Hash]    opts :color [FixNum]  the color when setting the contents
  def set_nameval(name, value, opts = {:color => 0})
    begin
      cell = name_object(name).RefersToRange
      cell.Interior.ColorIndex = opts[:color] 
      workbook.modified_cells << cell unless cell_modified?(cell)
      cell.Value = value
    rescue WIN32OLERuntimeError
      raise RangeNotEvaluatable, "cannot assign value to range named #{name.inspect} in #{File.basename(workbook.stored_filename).inspect rescue nil}" 
    end
  end

  # returns the contents of a range with a locally defined name
  # evaluates the formula if the contents is a formula
  # if no contents could be returned, then return default value, if provided, raise error otherwise
  # @param  [String]      name      the name of a range
  # @param  [Hash]        opts      the options
  # @option opts [Symbol] :default  the default value that is provided if no contents could be returned
  # @return [Variant] the contents of a range with given name   
  def rangeval(name, opts = {:default => nil})
    begin
      range = self.Range(name)
    rescue WIN32OLERuntimeError
      return opts[:default] if opts[:default]
      raise NameNotFound, "name #{name.inspect} not in #{workbook.stored_filename rescue nil}"
    end
    begin
      value = range.Value
    rescue  WIN32OLERuntimeError
      return opts[:default] if opts[:default]
      raise RangeNotEvaluatable, "cannot determine value of range named #{name.inspect} in #{workbook.stored_filename rescue nil}"
    end
    return opts[:default] if (value.nil? && opts[:default])
    raise RangeNotEvaluatable, "cannot evaluate range named #{name.inspect}" if value.is_a?(Bignum)
    value
  end

  # assigns a value to a range given a locally defined name
  # @param [String]  name   the name of a range
  # @param [Variant] value  the assigned value
  # @param [Hash]    opts :color [FixNum]  the color when setting the contents
  def set_rangeval(name,value, opts = {:color => 0})
    begin
      return set_nameval(name, value, opts) if self.is_a?(Book)
      range = self.Range(name)
    rescue WIN32OLERuntimeError
      raise NameNotFound, "name #{name.inspect} not in #{workbook.stored_filename rescue nil}"
    end
    begin
      range.Interior.ColorIndex = opts[:color]
      workbook.modified_cells << range unless cell_modified?(range)
      range.Value = value
    rescue  WIN32OLERuntimeError
      raise RangeNotEvaluatable, "cannot assign value to range named #{name.inspect} in #{workbook.stored_filename rescue nil}"
    end
  end

private  

  def name_object(name)
    begin
      self.Parent.Names.Item(name)
    rescue WIN32OLERuntimeError
      begin
        self.Names.Item(name)
      rescue WIN32OLERuntimeError
        raise RobustExcelOle::NameNotFound, "name #{name.inspect} not in #{File.basename(workbook.stored_filename).inspect rescue nil}"  
      end
    end
  end

  def cell_modified?(cell)
    workbook.modified_cells.each{|c| return true if c.Name.Value == cell.Name.Value}    
    false
  end

end
