# -*- coding: utf-8 -*-

module RobustExcelOle

  class RangeOwners < VbaObjects

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
    def namevalue_global(name, opts = { default: :__not_provided })
      begin
        name_obj = begin
          get_name_object(name)
        rescue NameNotFound
          raise
        end
        ole_range = name_obj.RefersToRange
        worksheet = self if self.is_a?(Worksheet)
        value = begin
          if !::RANGES_JRUBY_BUG       
           ole_range.Value
          else
            values = RobustExcelOle::Range.new(ole_range, worksheet).v
            (values.size==1 && values.first.size==1) ? values.first.first : values
          end
        end
      rescue # WIN32OLERuntimeError, Java::OrgRacobCom::ComFailException 
        sheet = if self.is_a?(Worksheet) then self
        # chooses simply the 1st worksheet?
        elsif self.is_a?(Workbook) then self.sheet(1)
        end
        begin
          # does it result in a range?
          ole_range = sheet.Evaluate(name_obj.Name)
          if !::RANGES_JRUBY_BUG
            ole_range.Value
          else
            values = RobustExcelOle::Range.new(ole_range, worksheet).v
            (values.size==1 && values.first.size==1) ? values.first.first : values
          end
        rescue # WIN32OLERuntimeError, Java::OrgRacobCom::ComFailException 
          return opts[:default] unless opts[:default] == :__not_provided
          if name_obj.nil?
            raise NameNotFound, "cannot find name #{name.inspect}"
          else
            raise RangeNotEvaluatable, "cannot evaluate range named #{name.inspect}"
          end
        end
      end
      if value == -2146828288 + RobustExcelOle::XlErrName
        return opts[:default] unless opts[:default] == :__not_provided
        raise RangeNotEvaluatable, "cannot evaluate range named #{name.inspect} in #{File.basename(workbook.stored_filename).inspect rescue nil}"
      end
      return opts[:default] if opts[:default] != :__not_provided && !value.nil?
      value
    end

    # sets the contents of a range
    # @param [String]  name  the name of a range
    # @param [Variant] value the contents of the range
    # @option opts [Symbol] :color the color of the range when set
    def set_namevalue_global(name, value, opts = { }) 
      name_obj = begin
        get_name_object(name)
      rescue NameNotFound => msg
        raise
      end        
      ole_range = name_obj.RefersToRange
      ole_range.Interior.ColorIndex = opts[:color] unless opts[:color].nil?
      if !::RANGES_JRUBY_BUG
        ole_range.Value = value
      else
        address_r1c1 = ole_range.AddressLocal(true,true,XlR1C1)
        row, col = address_tool.as_integer_ranges(address_r1c1)
        row.each_with_index do |r,i|
          col.each_with_index do |c,j|
            ole_range.Cells(i+1,j+1).Value = (value.respond_to?(:pop) ? value[i][j] : value )
          end
        end
      end
      value
    rescue #WIN32OLERuntimeError, Java::OrgRacobCom::ComFailException
      raise RangeNotEvaluatable, "cannot assign value to range named #{name.inspect} in #{self.inspect}\n#{$!.message}"
    end

    # @private
    def nameval(name, opts = { default: :__not_provided })   # :deprecated: #
      namevalue_global(name, opts)
    end

    # @private
    def set_nameval(name, value)        # :deprecated: #
      set_namevalue_global(name, value)
    end

    # @private
    def rangeval(name, opts = { default: :__not_provided })  # :deprecated: #
      namevalue(name, opts)
    end

    # @private
    def set_rangeval(name, value)       # :deprecated: #
      set_namevalue(name, value)
    end

    # creates a range from a given defined name or address
    # @params [Variant] defined name or address, and optional a worksheet
    # @return [Range] a range
    def range(*args)
      raise RangeNotCreated, "not yet implemented"
    end

    def name2range(name)   # :deprecated: #
      range(name)
    end

    # adds a name referring to a range given by the row and column
    # @param [String] name   the range name
    # @params [Address] address of the range
    def add_name(name, addr, addr_deprecated = :__not_provided)
      addr = [addr,addr_deprecated] unless addr_deprecated == :__not_provided
      self.Names.Add(name, nil, true, nil, nil, nil, nil, nil, nil, '=' + address_tool.as_r1c1(addr))
      name
    rescue WIN32OLERuntimeError, Java::OrgRacobCom::ComFailException
      raise RangeNotEvaluatable, "cannot add name #{name.inspect} to range #{addr.inspect}\n#{$!.message}"
    end

    alias set_name add_name  # :deprecated :#

    # renames an Excel object
    # @param [String] old_name the previous name of the Excel object
    # @param [String] new_name the new name of the Excel object
    def rename_name(old_name, new_name)
      item = get_name_object(old_name)
      item.Name = new_name
    rescue RobustExcelOle::NameNotFound
      raise
    rescue WIN32OLERuntimeError, Java::OrgRacobCom::ComFailException => msg
      raise UnexpectedREOError, "name error with name #{old_name.inspect} in #{File.basename(self.stored_filename).inspect}\n#{$!.message}"
    end

    alias rename_range rename_name  # :deprecated :#

    # deletes a name of an Excel object
    # @param [String] name    the name of the Excel object
    def delete_name(name)
      item = get_name_object(name)
      item.Delete
    rescue WIN32OLERuntimeError, Java::OrgRacobCom::ComFailException
      raise UnexpectedREOError, "name error with name #{name.inspect} in #{File.basename(self.stored_filename).inspect}\n#{$!.message}"
    end

    # @return [Array] defined names
    def names
      self.Names.to_a.map(&:name)
    end

  private

    def get_name_object(name)
      self.Names.Item(name)
    rescue WIN32OLERuntimeError, Java::OrgRacobCom::ComFailException, VBAMethodMissingError
      begin
        self.Parent.Names.Item(name)
      rescue WIN32OLERuntimeError, Java::OrgRacobCom::ComFailException
        raise RobustExcelOle::NameNotFound, "name #{name.inspect} not in #{self.inspect}"
      end
    end

  end

  # @private
  class NameNotFound < NamesREOError               
  end

  # @private
  class NameAlreadyExists < NamesREOError          
  end

  # @private
  class RangeNotCreated < MiscREOError             
  end  

end
