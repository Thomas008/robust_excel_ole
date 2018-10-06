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

  class RangeNotCreated < MiscREOError             # :nodoc: #
  end

  class RangeNotCopied < MiscREOError              # :nodoc: #
  end

  class OptionInvalid < MiscREOError               # :nodoc: #
  end

  class ObjectNotAlive < MiscREOError              # :nodoc: #
  end

  class TypeREOError < REOError                    # :nodoc: #
  end   

  class TimeOut < REOError                         # :nodoc: #
  end  

  class AddressInvalid < REOError                  # :nodoc: #
  end

  class UnexpectedREOError < REOError              # :nodoc: #
  end

  class NotImplementedREOError < REOError          # :nodoc: #
  end

  class REOCommon

    def excel
      raise TypeREOError, "receiver instance is neither an Excel nor a Workbook"
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

  class Address < REOCommon

    attr_reader :rows
    attr_reader :columns

    def initialize(address)
      address = [address] unless address.is_a?(Array)
      raise AddressInvalid, "more than two components" if address.size > 2
      begin
        if address.size == 1
          comp1, comp2 = address[0].split(':')          
          address_comp1 = comp1.gsub(/[A-Z]/,'')
          address_comp2 = comp1.gsub(/[0-9]/,'')
          if comp1 != address_comp2+address_comp1
            raise AddressInvalid, "address #{comp1.inspect} not in A1-format"
          end
          unless comp2.nil?
            address_comp3 = comp2.gsub(/[A-Z]/,'')
            address_comp4 = comp2.gsub(/[0-9]/,'')  
            if comp2 != address_comp4+address_comp3
              raise AddressInvalid, "address #{comp2.inspect} not in A1-format"
            end
            address_comp1 = address_comp1..address_comp3
            address_comp2 = address_comp2..address_comp4
          end
        else
          address_comp1, address_comp2 = address
        end
        address_comp1 = address_comp1 .. address_comp1 unless address_comp1.is_a?(Object::Range)
        address_comp2 = address_comp2 .. address_comp2 unless address_comp2.is_a?(Object::Range)
        @rows = address_comp1.min.to_i .. address_comp1.max.to_i
        if address_comp2.min.to_i == 0
          raise AddressInvalid, "address (#{address_comp1.inspect}, #{address_comp2.inspect}) not in A1-format" if address_comp1.min.to_i == 0                    
          @columns = str2num(address_comp2.min) .. str2num(address_comp2.max)
        else
          @columns = address_comp2.min.to_i .. address_comp2.max.to_i
        end
      rescue  
        raise AddressInvalid, "address (#{address.inspect}) not in A1- or R1C1-format"
      end
    end

  private

    def str2num(str)
      str = str.upcase
      sum = 0
      (1..str.length).each { |i| sum += (str[i-1].ord-64) * 26 ** (str.length - i) }
      sum
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
        sheet = if self.is_a?(Worksheet) then self
        elsif self.is_a?(Workbook) then self.sheet(1)
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
      return namevalue_glob(name, opts) if self.is_a?(Workbook)
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
        return set_namevalue_glob(name, value, opts) if self.is_a?(Workbook)
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

    def nameval(name, opts = {:default => :__not_provided})   # :deprecated: #
      namevalue_glob(name, opts)
    end

    def set_nameval(name, value, opts = {:color => 0})        # :deprecated: #
      set_namevalue_glob(name, value, opts)
    end

    def rangeval(name, opts = {:default => :__not_provided})  # :deprecated: #
      namevalue(name, opts)
    end

    def set_rangeval(name, value, opts = {:color => 0})       # :deprecated: #
      set_namevalue(name, value, opts)
    end

    # @params [String] name  defined range name
    # @returns [Range] a Range object
    def name2range(name)
      begin
        RobustExcelOle::Range.new(name_object(name).RefersToRange)
      rescue WIN32OLERuntimeError
        raise RangeNotCreated, "range could not be created from the defined name"
      end
    end

    # adds a name referring to a range given by the row and column
    # @param [String] name   the range name
    # @params [Address] address of the range 
    def add_name(name, addr, addr_deprecated = :__not_provided)
      addr = [addr,addr_deprecated] unless addr_deprecated == :__not_provided
      address = Address.new(addr)
      address_string = "Z" + address.rows.min.to_s + "S" + address.columns.min.to_s + 
                ":Z" + address.rows.max.to_s + "S" + address.columns.max.to_s
      begin
        self.Names.Add("Name" => name, "RefersToR1C1" => "=" + address_string)
      rescue WIN32OLERuntimeError => msg
        #trace "WIN32OLERuntimeError: #{msg.message}"
        raise RangeNotEvaluatable, "cannot add name #{name.inspect} to range #{addr.inspect}"
      end
      name
    end

    def set_name(name,row,column)     # :deprecated :#
      add_name(name,row,column)
    end

    # renames a range
    # @param [String] name     the previous range name
    # @param [String] new_name the new range name
    def rename_range(name, new_name)
      begin
        item = self.Names.Item(name)
      rescue WIN32OLERuntimeError
        raise NameNotFound, "name #{name.inspect} not in #{File.basename(self.stored_filename).inspect}"  
      end
      begin
        item.Name = new_name
      rescue WIN32OLERuntimeError
        raise UnexpectedREOError, "name error in #{File.basename(self.stored_filename).inspect}"      
      end
    end

    # deletes a name of a range
    # @param [String] name     the previous range name
    # @param [String] new_name the new range name
    def delete_name(name)
      begin
        item = self.Names.Item(name)
      rescue WIN32OLERuntimeError
        raise NameNotFound, "name #{name.inspect} not in #{File.basename(self.stored_filename).inspect}"  
      end
      begin
        item.Delete
      rescue WIN32OLERuntimeError
        raise UnexpectedREOError, "name error in #{File.basename(self.stored_filename).inspect}"      
      end
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
