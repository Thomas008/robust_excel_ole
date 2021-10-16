# -*- coding: utf-8 -*-

module RobustExcelOle

  # This class essentially wraps a Win32Ole Worksheet object. 
  # You can apply all VBA methods (starting with a capital letter) 
  # that you would apply for a Worksheet object. 
  # see https://docs.microsoft.com/en-us/office/vba/api/excel.worksheet#methods


  # worksheet: see https://github.com/Thomas008/robust_excel_ole/blob/master/lib/robust_excel_ole/worksheet.rb
  class Worksheet < RangeOwners

    include Enumerable

    using ToReoRefinement

    attr_reader :ole_worksheet
    attr_reader :workbook

    alias ole_object ole_worksheet

    def initialize(win32_worksheet)
      @ole_worksheet = win32_worksheet
      if @ole_worksheet.ProtectContents
        @ole_worksheet.Unprotect
        @end_row = last_row
        @end_column = last_column
        @ole_worksheet.Protect
      else
        @end_row = last_row
        @end_column = last_column
      end
    end

    def workbook
      @workbook ||= begin
        ole_workbook = self.Parent
        saved_status = ole_workbook.Saved
        ole_workbook.Saved = true unless saved_status
        @workbook = workbook_class.new(ole_workbook)
        ole_workbook.Saved = saved_status
        @workbook
      end
    end

    def excel
      workbook.excel
    end

    # sheet name
    # @returns name of the sheet
    def name
      @ole_worksheet.Name.encode('utf-8')
    rescue
      raise WorksheetREOError, "name could not be determined\n#{$!.message}"
    end

    # sets sheet name
    # @param [String] new_name the new name of the sheet
    def name= (new_name)
      @ole_worksheet.Name = new_name
    rescue WIN32OLERuntimeError, Java::OrgRacobCom::ComFailException => msg
      if msg.message =~ /800A03EC/ || msg.message =~ /Visual Basic/
        raise NameAlreadyExists, "sheet name #{new_name.inspect} already exists"
      else
        raise UnexpectedREOError, "unexpected WIN32OLERuntimeError: #{msg.message}"
      end
    end

    # value of a range given its defined name or address
    # @params [Variant] defined name or address
    # @returns [Variant] value (contents) of the range
    def [](name_or_address, address2 = :__not_provided)
      range(name_or_address, address2).value
    end

    # sets the value of a range given its defined name or address, and the value
    # @params [Variant] defined name or address of the range
    # @params [Variant] value (contents) of the range
    # @returns [Variant] value (contents) of the range
    def []=(name_or_address, value_or_address2, remaining_arg = :__not_provided) 
      if remaining_arg != :__not_provided
        name_or_address, value = [name_or_address, value_or_address2], remaining_arg
      else
        value = value_or_address2
      end
      begin
        range(name_or_address).value = value
      rescue #WIN32OLERuntimeError, Java::OrgRacobCom::ComFailException
        raise RangeNotEvaluatable, "cannot assign value to range with name or address #{name_or_address.inspect}\n#{$!.message}"
      end
    end

    # a range given a defined name or address
    # @params [Variant] defined name or address
    # @return [Range] a range
    def range(name_or_address, address2 = :__not_provided)
      if name_or_address.respond_to?(:gsub) && address2 == :__not_provided
        name = name_or_address
        range = get_name_object(name).RefersToRange rescue nil
      end
      unless range
        address = normalize_address(name_or_address, address2)
        workbook.retain_saved do
          begin
            a1_address = address_tool.as_a1(address) rescue nil
            if a1_address
              range = self.Range(a1_address)              
            else
              saved = self.Parent.Saved
              begin
                self.Names.Add('__dummy_name_object_001__',nil,true,nil,nil,nil,nil,nil,nil,'=' + address_tool.as_r1c1(address))
                range = get_name_object('__dummy_name_object_001__').RefersToRange
              ensure
                self.Names.Item('__dummy_name_object_001__').Delete
                self.Parent.Saved = saved
              end
            end
          rescue
            address2_string = (address2.nil? || address2 == :__not_provided) ? "" : ", #{address2.inspect}"
            raise RangeNotCreated, "cannot find name or address #{name_or_address.inspect}#{address2_string}"
          end
        end
      end
      range.to_reo
    end

  private

    def normalize_address(address, address2)
      address = [address,address2] unless address2 == :__not_provided     
      address = if address.is_a?(Integer) || address.is_a?(::Range)
        [address, nil]
      elsif address.is_a?(Array) && address.size == 1 && (address.first.is_a?(Integer) || address.first.is_a?(::Range))
        [address.first, nil]
      else 
        address
      end
    end

  public
    
    # returns the contents of a range with a locally defined name
    # evaluates the formula if the contents is a formula
    # if the name could not be found or the range or value could not be determined,
    # then return default value, if provided, raise error otherwise
    # @param  [String]      name      the name of a range
    # @param  [Hash]        opts      the options
    # @option opts [Symbol] :default  the default value that is provided if no contents could be returned
    # @return [Variant] the contents of a range with given name
    def namevalue(name, opts = { default: :__not_provided })
      begin
        ole_range = self.Range(name)
      rescue # WIN32OLERuntimeError, VBAMethodMissingError, Java::OrgRacobCom::ComFailException 
        return opts[:default] unless opts[:default] == :__not_provided
        raise NameNotFound, "name #{name.inspect} not in #{self.inspect}"
      end
      begin
        worksheet = self if self.is_a?(Worksheet)
        #value = ole_range.Value
        value = if !::RANGES_JRUBY_BUG
          ole_range.Value
        else
          values = RobustExcelOle::Range.new(ole_range, worksheet).v
          (values.size==1 && values.first.size==1) ? values.first.first : values
        end
      rescue WIN32OLERuntimeError, Java::OrgRacobCom::ComFailException 
        return opts[:default] unless opts[:default] == :__not_provided
        raise RangeNotEvaluatable, "cannot determine value of range named #{name.inspect} in #{self.inspect}\n#{$!.message}"
      end
      if value == -2146828288 + RobustExcelOle::XlErrName
        return opts[:default] unless opts[:default] == __not_provided
        raise RangeNotEvaluatable, "cannot evaluate range named #{name.inspect} in #{File.basename(workbook.stored_filename).inspect rescue nil}\n#{$!.message}"
      end
      return opts[:default] unless (opts[:default] == :__not_provided) || value.nil?
      value
    end

    # assigns a value to a range given a locally defined name
    # @param [String]  name   the name of a range
    # @param [Variant] value  the assigned value   
    # @option opts [Symbol] :color the color of the cell when set
    def set_namevalue(name, value, opts = { })  
      begin
        ole_range = self.Range(name)
      rescue WIN32OLERuntimeError, Java::OrgRacobCom::ComFailException, VBAMethodMissingError
        raise NameNotFound, "name #{name.inspect} not in #{self.inspect}"
      end
      begin
        ole_range.Interior.ColorIndex = opts[:color] unless opts[:color].nil?
        if !::RANGES_JRUBY_BUG
          ole_range.Value = value
        else
          address_r1c1 = ole_range.AddressLocal(true,true,XlR1C1)
          row, col = address_tool.as_integer_ranges(address_r1c1)
          row.each_with_index do |r,i|
            col.each_with_index do |c,j|
              ole_range.Cells(i+1,j+1).Value = (value.respond_to?(:pop) ? value[i][j] : value)
            end
          end
        end
        value
      rescue WIN32OLERuntimeError, Java::OrgRacobCom::ComFailException
        raise RangeNotEvaluatable, "cannot assign value to range named #{name.inspect} in #{self.inspect}\n#{$!.message}"
      end
    end

    # value of a cell, if row and column are given
    # @params row and column
    # @returns value of the cell
    def cellval(x,y)                         # :deprecated :#
      @ole_worksheet.Cells.Item(x, y).Value
    rescue
      raise RangeNotEvaluatable, "cannot read cell (#{x.inspect},#{y.inspect})\n#{$!.message}"
    end

    # sets the value of a cell, if row, column and color of the cell are given
    # @params [Integer] x,y row and column
    # @option opts [Symbol] :color the color of the cell when set
    def set_cellval(x,y,value, opts = { }) # option opts is deprecated
      cell = @ole_worksheet.Cells.Item(x, y)
      cell.Interior.ColorIndex = opts[:color] unless opts[:color].nil?
      cell.Value = value
    rescue # WIN32OLERuntimeError, Java::OrgRacobCom::ComFailException
      raise RangeNotEvaluatable, "cannot assign value #{value.inspect} to cell (#{y.inspect},#{x.inspect})\n#{$!.message}"
    end

    # @return [Array] a 2-dimensional array that contains the values in each row of the used range
    def values
      @ole_worksheet.UsedRange.Value
    end

    # @return [Enumerator] traversing the rows values
    def each
      if block_given?
        @ole_worksheet.UsedRange.Rows.lazy.each do |ole_row|
          row_value = ole_row.Value
          yield (row_value.nil? ? [] : row_value.first)  
        end
      else
        to_enum(:each).lazy
      end
    end

    # @return [Enumerator] traversing the rows
    def each_row(offset = 0)
      if block_given?
        offset += 1
        1.upto(@end_row) do |row|
          next if row < offset
          yield RobustExcelOle::Range.new(@ole_worksheet.Range(@ole_worksheet.Cells(row, 1), @ole_worksheet.Cells(row, @end_column)), self)
        end
      else
        to_enum(:each_row).lazy
      end
    end

    def each_row_with_index(offset = 0)    # :nodoc: #   # :deprecated :#
      each_row(offset) do |row_range|
        yield RobustExcelOle::Range.new(row_range, self), (row_range.Row - 1 - offset)
      end
    end

    # @return [Enumerator] traversing the columns
    def each_column(offset = 0)
      if block_given?
        offset += 1
        1.upto(@end_column) do |column|
          next if column < offset
          yield RobustExcelOle::Range.new(@ole_worksheet.Range(@ole_worksheet.Cells(1, column), @ole_worksheet.Cells(@end_row, column)), self)
        end
      else
        to_enum(:each_column).lazy
      end
    end

    def each_column_with_index(offset = 0)    # :nodoc: #    # :deprecated :#
      each_column(offset) do |column_range|
        yield RobustExcelOle::Range.new(column_range, self), (column_range.Column - 1 - offset)
      end
    end

    # @return [Enumerator] traversing the cells
    def each_cell
      if block_given?
        each_row do |row_range|
          row_range.lazy.each do |cell|
            yield cell
          end
        end
      else
        to_enum(:each_cell).lazy
      end
    end

    def each_cell_with_index(offset = 0)   # :nodoc: #  # :deprecated :#
      i = offset
      each_row do |row_range|
        row_range.each do |cell|
          yield cell, i
          i += 1
        end
      end
    end

    def each_rowvalue  # :deprecated: #
      values.each do |row_values|
        yield row_values
      end
    end

    def each_rowvalue_with_index(offset = 0)    # :deprecated: #
      i = offset
      values.each do |row_values|
        yield row_values, i
        i += 1
      end
    end

    alias each_value each_rowvalue   # :deprecated: #
     
    def row_range(row, integer_range = nil)
      integer_range ||= 1..@end_column
      RobustExcelOle::Range.new(@ole_worksheet.Range(@ole_worksheet.Cells(row, integer_range.min), @ole_worksheet.Cells(row, integer_range.max)), self)
    end

    def col_range(col, integer_range = nil)
      integer_range ||= 1..@end_row
      RobustExcelOle::Range.new(@ole_worksheet.Range(@ole_worksheet.Cells(integer_range.min, col), @ole_worksheet.Cells(integer_range.max, col)), self)
    end

    def == other_worksheet
      other_worksheet.is_a?(Worksheet) && 
      self.workbook == other_worksheet.workbook &&
      self.Name == other_worksheet.Name
    end

    
    # @params [Variant] table (listobject) name or number 
    # @return [ListObject] a table (listobject)
    def table(number_or_name)
      listobject_class.new(@ole_worksheet.ListObjects.Item(number_or_name))
    rescue
      raise WorksheetREOError, "table #{number_or_name} not found"
    end

    # @private
    # returns true, if the worksheet object responds to VBA methods, false otherwise
    def alive?
      @ole_worksheet.UsedRange
      true
    rescue
      # trace $!.message
      false
    end

    # last_row, last_column:
    # the last row and last column in a worksheet can be determined with help of
    # UsedRange.SpecialCells and UsedRange.Rows/Columns
    # both values can differ in certain cases:
    # - if the worksheet contains a table, then UsedRange starts at the table, not in the first cell
    #   therefore we use SpecialCells.
    # - if the worksheet contains merged cells, then SpecialCells considers the merged cells only,
    #   therefor we use UsedRange here.

    # @private
    def last_row
      special_last_row = @ole_worksheet.UsedRange.SpecialCells(RobustExcelOle::XlLastCell).Row
      used_last_row = @ole_worksheet.UsedRange.Rows.Count
      [special_last_row, used_last_row].max
    end

    # @private
    def last_column
      special_last_column = @ole_worksheet.UsedRange.SpecialCells(RobustExcelOle::XlLastCell).Column
      used_last_column = @ole_worksheet.UsedRange.Columns.Count
      [special_last_column, used_last_column].max
    end

    using ParentRefinement
    using StringRefinement

    # @private
    def self.workbook_class  
      @workbook_class ||= begin
        module_name = self.parent_name
        "#{module_name}::Workbook".constantize        
      rescue NameError => e
        Workbook
      end
    end

    # @private
    def workbook_class        
      self.class.workbook_class
    end

    # @private
    def self.listobject_class  
      @listobject_class ||= begin
        module_name = self.parent_name
        "#{module_name}::ListObject".constantize        
      rescue NameError => e
        ListObject
      end
    end

    # @private
    def listobject_class        
      self.class.listobject_class
    end

    # @private
    def to_s    
      "#<Worksheet: #{(workbook.nil? ? "not alive " : (name + ' ' + File.basename(workbook.stored_filename)))}>"
    end

    # @private
    def inspect  
      to_s
    end

    include MethodHelpers

  private

    def method_missing(name, *args)
      super unless name.to_s[0,1] =~ /[A-Z]/
      if ::ERRORMESSAGE_JRUBY_BUG 
        begin
          @ole_worksheet.send(name, *args)
        rescue Java::OrgRacobCom::ComFailException 
          raise VBAMethodMissingError, "unknown VBA property or method #{name.inspect}"
        end
      else
        begin
          @ole_worksheet.send(name, *args)
        rescue NoMethodError 
          raise VBAMethodMissingError, "unknown VBA property or method #{name.inspect}"
        end
      end
    end
    
  end

  public

  Sheet = Worksheet

end
