# -*- coding: utf-8 -*-

module RobustExcelOle

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
    def namevalue_glob(name, opts = { :default => :__not_provided })
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
        return opts[:default] unless opts[:default] == :__not_provided
        raise RangeNotEvaluatable, "cannot evaluate range named #{name.inspect} in #{File.basename(workbook.stored_filename).inspect rescue nil}"
      end
      return opts[:default] unless (opts[:default] == :__not_provided) || value.nil?
      value
    end

    # sets the contents of a range
    # @param [String]  name  the name of a range
    # @param [Variant] value the contents of the range
    # @param [FixNum]  color the color when setting a value
    # @param [Hash]    opts :color [FixNum]  the color when setting the contents
    def set_namevalue_glob(name, value, opts = { :color => 0 })
      cell = name_object(name).RefersToRange
      cell.Interior.ColorIndex = opts[:color]
      workbook.modified_cells << cell if workbook # unless cell_modified?(cell)
      cell.Value = value
    rescue WIN32OLERuntimeError
      raise RangeNotEvaluatable, "cannot assign value to range named #{name.inspect} in #{self.inspect}"
    end

    # returns the contents of a range with a locally defined name
    # evaluates the formula if the contents is a formula
    # if the name could not be found or the range or value could not be determined,
    # then return default value, if provided, raise error otherwise
    # @param  [String]      name      the name of a range
    # @param  [Hash]        opts      the options
    # @option opts [Symbol] :default  the default value that is provided if no contents could be returned
    # @return [Variant] the contents of a range with given name
    def namevalue(name, opts = { :default => :__not_provided })
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
      return opts[:default] unless (opts[:default] == :__not_provided) || value.nil?
      value
    end

    # assigns a value to a range given a locally defined name
    # @param [String]  name   the name of a range
    # @param [Variant] value  the assigned value
    # @param [Hash]    opts :color [FixNum]  the color when setting the contents
    def set_namevalue(name, value, opts = { :color => 0 })
      begin
        return set_namevalue_glob(name, value, opts) if self.is_a?(Workbook)
        range = self.Range(name)
      rescue WIN32OLERuntimeError
        raise NameNotFound, "name #{name.inspect} not in #{self.inspect}"
      end
      begin
        range.Interior.ColorIndex = opts[:color]
        workbook.modified_cells << range if workbook # unless cell_modified?(range)
        range.Value = value
      rescue  WIN32OLERuntimeError
        raise RangeNotEvaluatable, "cannot assign value to range named #{name.inspect} in #{self.inspect}"
      end
    end

    # @private
    def nameval(name, opts = { :default => :__not_provided })   # :deprecated: #
      namevalue_glob(name, opts)
    end

    # @private
    def set_nameval(name, value, opts = { :color => 0 })        # :deprecated: #
      set_namevalue_glob(name, value, opts)
    end

    # @private
    def rangeval(name, opts = { :default => :__not_provided })  # :deprecated: #
      namevalue(name, opts)
    end

    # @private
    def set_rangeval(name, value, opts = { :color => 0 })       # :deprecated: #
      set_namevalue(name, value, opts)
    end

    # creates a range from a given defined name or address
    # range(address) does work for Worksheet objects only
    # @params [Variant] range name or address
    # @return [Range] a range
    def range(name_or_address, address2 = :__not_provided)
      begin
        if address2 == :__not_provided
          range = begin
            RobustExcelOle::Range.new(name_object(name_or_address).RefersToRange) 
          rescue NameNotFound
            nil
          end
        end
        if self.is_a?(Worksheet) && (range.nil? || (address2 != :__not_provided))
          address = name_or_address
          address = [name_or_address,address2] unless address2 == :__not_provided
          self.Names.Add('Name' => '__dummy001', 'RefersToR1C1' => '=' + Address.r1c1(address))
          range = RobustExcelOle::Range.new(name_object('__dummy001').RefersToRange)
          self.Names.Item('__dummy001').Delete
          range                    
        end
      rescue WIN32OLERuntimeError
        address2_string = address2.nil? ? "" : ", #{address2.inspect}"
        raise RangeNotCreated, "cannot create range (#{name_or_address.inspect}#{address2_string})"
      end      
      range
    end

    def name2range(name)   # :deprecated: #
      range(name)
    end

    # adds a name referring to a range given by the row and column
    # @param [String] name   the range name
    # @params [Address] address of the range
    def add_name(name, addr, addr_deprecated = :__not_provided)
      addr = [addr,addr_deprecated] unless addr_deprecated == :__not_provided
      begin
        self.Names.Add('Name' => name, 'RefersToR1C1' => '=' + Address.r1c1(addr))
      rescue WIN32OLERuntimeError => msg
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
      self.Names.Item(name)
    rescue WIN32OLERuntimeError
      begin
        self.Parent.Names.Item(name)
      rescue WIN32OLERuntimeError
        raise RobustExcelOle::NameNotFound, "name #{name.inspect} not in #{self.inspect}"
      end
    end

  end

end
