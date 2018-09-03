# -*- coding: utf-8 -*-

LOG_TO_STDOUT = true      unless Object.const_defined?(:LOG_TO_STDOUT)
REO_LOG_DIR   = ""        unless Object.const_defined?(:REO_LOG_DIR)
REO_LOG_FILE  = "reo.log" unless Object.const_defined?(:REO_LOG_FILE)
  
File.delete REO_LOG_FILE rescue nil

module RobustExcelOle

  class REOError < RuntimeError                    # :nodoc: #
  end                 

  class ExcelREOError < REOError                   # :nodoc: #
  end

  class WorkbookREOError < REOError                # :nodoc: #
  end

  class SheetREOError < REOError                   # :nodoc: #
  end

  class FileREOError < REOError                    # :nodoc: #
  end

  class NamesREOError < REOError                   # :nodoc: #
  end

  class MiscREOError < REOError                    # :nodoc: #
  end

  class ExcelDamaged < ExcelREOError               # :nodoc: #
  end

  class UnsavedWorkbooks < ExcelREOError           # :nodoc: #
  end

  class WorkbookBlocked < WorkbookREOError         # :nodoc: #
  end

  class WorkbookNotSaved < WorkbookREOError        # :nodoc: #
  end

  class WorkbookReadOnly < WorkbookREOError        # :nodoc: #
  end

  class WorkbookBeingUsed < WorkbookREOError       # :nodoc: #
  end

  class FileNotFound < FileREOError                # :nodoc: #
  end

  class FileNameNotGiven < FileREOError            # :nodoc: #
  end

  class FileAlreadyExists < FileREOError           # :nodoc: #
  end

  class NameNotFound < NamesREOError               # :nodoc: #
  end

  class NameAlreadyExists < NamesREOError          # :nodoc: #
  end

  class RangeNotEvaluatable < MiscREOError         # :nodoc: #
  end

  class OptionInvalid < MiscREOError               # :nodoc: #
  end

  class ObjectNotAlive < MiscREOError              # :nodoc: #
  end

  class TypeREOError < REOError                    # :nodoc: #
  end   

  class TimeOut < REOError                         # :nodoc: #
  end  

  class UnexpectedREOError < REOError              # :nodoc: #
  end

  class NotImplementedREOError < REOError          # :nodoc: #
  end

  class REOCommon

    def excel
      raise TypeREOError, "receiver instance is neither an Excel nor a Book"
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
    # if the name could not be found or the value could not be determined,
    #   then return default value, if provided, raise error otherwise
    # Excel Bug: if a local name without a qualifier is given, 
    #   then by default Excel takes the first worksheet,
    #   even if a different worksheet is active
    # @param  [String]      name      the name of the range
    # @param  [Hash]        opts      the options
    # @option opts [Symbol] :default  the default value that is provided if no contents could be returned
    # @return [Variant] the contents of a range with given name
    def namevalue_glob(name, opts = {:default => :__not_provided})
      name_obj = begin
        name_object(name)
      rescue NameNotFound => msg
        return opts[:default] unless opts[:default] == :__not_provided
        raise
      end
      value = begin
        name_obj.RefersToRange.Value
      rescue WIN32OLERuntimeError
        sheet = if self.is_a?(Sheet) then self
        elsif self.is_a?(Book) then self.sheet(1)
        elsif self.is_a?(Excel) then self.workbook.sheet(1)
        end
        begin
          sheet.Evaluate(name_obj.Name).Value
        rescue # WIN32OLERuntimeError
          return opts[:default] unless opts[:default] == :__not_provided
          raise RangeNotEvaluatable, "cannot evaluate range named #{name.inspect} in #{self}"
        end
      end
      if value == -2146828288 + RobustExcelOle::XlErrName  
        return opts[:default] unless opts[:default] == __not_provided
        raise RangeNotEvaluatable, "cannot evaluate range named #{name.inspect} in #{File.basename(workbook.stored_filename).inspect rescue nil}"
      end 
      return opts[:default] unless opts[:default] == :__not_provided or value.nil?
      value      
    end

    # sets the contents of a range
    # @param [String]  name  the name of a range
    # @param [Variant] value the contents of the range
    # @param [FixNum]  color the color when setting a value
    # @param [Hash]    opts :color [FixNum]  the color when setting the contents
    def set_namevalue_glob(name, value, opts = {:color => 0})
      begin
        cell = name_object(name).RefersToRange
        cell.Interior.ColorIndex = opts[:color] 
        workbook.modified_cells << cell if workbook #unless cell_modified?(cell)
        cell.Value = value
      rescue WIN32OLERuntimeError
        raise RangeNotEvaluatable, "cannot assign value to range named #{name.inspect} in #{self.inspect}" 
      end
    end

    # returns the contents of a range with a locally defined name
    # evaluates the formula if the contents is a formula
    # if the name could not be found or the range or value could not be determined,
    # then return default value, if provided, raise error otherwise
    # @param  [String]      name      the name of a range
    # @param  [Hash]        opts      the options
    # @option opts [Symbol] :default  the default value that is provided if no contents could be returned
    # @return [Variant] the contents of a range with given name   
    def namevalue(name, opts = {:default => :__not_provided})
      return namevalue_glob(name, opts) if self.is_a?(Book)
      begin
        range = self.Range(name)
      rescue WIN32OLERuntimeError
        return opts[:default] unless opts[:default] == :__not_provided
        raise NameNotFound, "name #{name.inspect} not in #{self.inspect}"
      end
      begin
        value = range.Value
      rescue  WIN32OLERuntimeError
        return opts[:default] unless opts[:default] == :__not_provided
        raise RangeNotEvaluatable, "cannot determine value of range named #{name.inspect} in #{self.inspect}"
      end
      if value == -2146828288 + RobustExcelOle::XlErrName  
        return opts[:default] unless opts[:default] == __not_provided
        raise RangeNotEvaluatable, "cannot evaluate range named #{name.inspect} in #{File.basename(workbook.stored_filename).inspect rescue nil}"
      end 
      return opts[:default] unless opts[:default] == :__not_provided or value.nil?
      value      
    end

    # assigns a value to a range given a locally defined name
    # @param [String]  name   the name of a range
    # @param [Variant] value  the assigned value
    # @param [Hash]    opts :color [FixNum]  the color when setting the contents
    def set_namevalue(name, value, opts = {:color => 0})
      begin
        return set_namevalue_glob(name, value, opts) if self.is_a?(Book)
        range = self.Range(name)
      rescue WIN32OLERuntimeError
        raise NameNotFound, "name #{name.inspect} not in #{self.inspect}"
      end
      begin
        range.Interior.ColorIndex = opts[:color]
        workbook.modified_cells << range if workbook #unless cell_modified?(range)
        range.Value = value
      rescue  WIN32OLERuntimeError
        raise RangeNotEvaluatable, "cannot assign value to range named #{name.inspect} in #{self.inspect}"
      end
    end

    def nameval(name, opts = {:default => :__not_provided})   # :deprivated: #
      namevalue_glob(name, opts)
    end

    def set_nameval(name, value, opts = {:color => 0})        # :deprivated: #
      set_namevalue_glob(name, value, opts)
    end

    def rangeval(name, opts = {:default => :__not_provided})  # :deprivated: #
      namevalue(name, opts)
    end

    def set_rangeval(name, value, opts = {:color => 0})       # :deprivated: #
      set_namevalue(name, value, opts)
    end


  private  

    def name_object(name)
      begin
        self.Names.Item(name)        
      rescue WIN32OLERuntimeError
        begin
          self.Parent.Names.Item(name)
        rescue WIN32OLERuntimeError
          raise RobustExcelOle::NameNotFound, "name #{name.inspect} not in #{self.inspect}"  
        end
      end
    end
     
    #def cell_modified?(cell)
    #  workbook.modified_cells.each{|c| return true if c.Name.Value == cell.Name.Value}    
    #  false
    #end

  end

end
