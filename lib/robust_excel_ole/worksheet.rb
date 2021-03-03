# -*- coding: utf-8 -*-

module RobustExcelOle

  # This class essentially wraps a Win32Ole Worksheet object. 
  # You can apply all VBA methods (starting with a capital letter) 
  # that you would apply for a Worksheet object. 
  # see https://docs.microsoft.com/en-us/office/vba/api/excel.worksheet#methods


  # worksheet: see https://github.com/Thomas008/robust_excel_ole/blob/master/lib/robust_excel_ole/worksheet.rb
  class Worksheet < RangeOwners

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
        raise NameAlreadyExists, "sheet name #{new_name.inspect} already exists\n#{$!.message}"
      else
        raise UnexpectedREOError, "unexpected WIN32OLERuntimeError: #{msg.message}"
      end
    end

    # returns a cell given the defined name or row and column
    # @params row, column, or name
    # @returns cell, if row and column are given
    def [] p1, p2 = :__not_provided
      if p2 != :__not_provided
        x, y = p1, p2
        xy = "#{x}_#{y}"
        @cells = { }
        begin
          @cells[xy] = RobustExcelOle::Cell.new(@ole_worksheet.Cells.Item(x, y), @worksheet)
        rescue
          raise RangeNotEvaluatable, "cannot read cell (#{x.inspect},#{y.inspect})\n#{$!.message}"
        end
      else
        name = p1
        begin
          namevalue_global(name)
        rescue REOError
          namevalue(name)
        end
      end
    end

    # sets the value of a cell
    # @params row and column, or defined name
    def []= (p1, p2, p3 = :__not_provided)
      if p3 != :__not_provided
        x, y, value = p1, p2, p3
        set_cellval(x,y,value)
      else
        name, value = p1, p2
        begin
          set_namevalue_global(name, value)
        rescue REOError
          begin
            workbook.set_namevalue_global(name, value)
          rescue REOError
            set_namevalue(name, value)
          end
        end
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
    def namevalue(name, opts = { default: :__not_provided })
      begin
        ole_range = self.Range(name)
      rescue # WIN32OLERuntimeError, VBAMethodMissingError, Java::OrgRacobCom::ComFailException 
        return opts[:default] unless opts[:default] == :__not_provided
        raise NameNotFound, "name #{name.inspect} not in #{self.inspect}\n#{$!.message}"
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
        raise NameNotFound, "name #{name.inspect} not in #{self.inspect}\n#{$!.message}"
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
    def cellval(x,y)
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

    # provides a 2-dimensional array that contains the values in each row
    def values
      @ole_worksheet.UsedRange.Value
    end

    # enumerator for accessing cells
    def each
      each_row do |row_range|
        row_range.each do |cell|
          yield cell
        end
      end
    end

    def each_with_index(offset = 0)
      i = offset
      each_row do |row_range|
        row_range.each do |cell|
          yield cell, i
          i += 1
        end
      end
    end

    # enumerator for accessing rows
    def each_row(offset = 0)
      offset += 1
      1.upto(@end_row) do |row|
        next if row < offset
        yield RobustExcelOle::Range.new(@ole_worksheet.Range(@ole_worksheet.Cells(row, 1), @ole_worksheet.Cells(row, @end_column)), self)
      end
    end

    def each_row_with_index(offset = 0)
      each_row(offset) do |row_range|
        yield RobustExcelOle::Range.new(row_range, self), (row_range.Row - 1 - offset)
      end
    end

    # enumerator for accessing columns
    def each_column(offset = 0)
      offset += 1
      1.upto(@end_column) do |column|
        next if column < offset
        yield RobustExcelOle::Range.new(@ole_worksheet.Range(@ole_worksheet.Cells(1, column), @ole_worksheet.Cells(@end_row, column)), self)
      end
    end

    def each_column_with_index(offset = 0)
      each_column(offset) do |column_range|
        yield RobustExcelOle::Range.new(column_range, self), (column_range.Column - 1 - offset)
      end
    end

    def each_rowvalue  # :deprecated: #
      values.each do |row_values|
        yield row_values
      end
    end

    def each_value   # :deprecated: #
      each_rowvalue
    end

    def each_rowvalue_with_index(offset = 0)    # :deprecated: #
      i = offset
      values.each do |row_values|
        yield row_values, i
        i += 1
      end
    end

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

    # creates a range from a given defined name or address
    # @params [Variant] defined name or address
    # @return [Range] a range
    def range(name_or_address, address2 = :__not_provided)
      if name_or_address.respond_to?(:gsub) && address2 == :__not_provided
        name = name_or_address
        range = RobustExcelOle::Range.new(get_name_object(name).RefersToRange, self) rescue nil
      end
      unless range
        address = name_or_address
        address = [name_or_address,address2] unless address2 == :__not_provided         
        workbook.retain_saved do
          begin
            self.Names.Add('__dummy001',nil,true,nil,nil,nil,nil,nil,nil,'=' + address_tool.as_r1c1(address))          
            range = RobustExcelOle::Range.new(get_name_object('__dummy001').RefersToRange, self)
            self.Names.Item('__dummy001').Delete
          rescue
            address2_string = address2.nil? ? "" : ", #{address2.inspect}"
            raise RangeNotCreated, "cannot create range (#{name_or_address.inspect}#{address2_string})\n#{$!.message}"
          end
        end
      end
      range
    end

    # @params [Variant] table (listobject) name or number 
    # @return [ListObject] a table (listobject)
    def table(number_or_name)
      listobject_class.new(@ole_worksheet.ListObjects.Item(number_or_name))
    rescue
      raise WorksheetREOError, "table #{number_or_name} not found\n#{$!.message}"
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
      "#<Worksheet: " + (workbook.nil? ? "not alive " : (name + ' ' + File.basename(workbook.stored_filename)).to_s) + ">"
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

    def last_row
      special_last_row = begin 
        @ole_worksheet.UsedRange.SpecialCells(RobustExcelOle::XlLastCell).Row
      rescue 
        nil
      end
      used_last_row = @ole_worksheet.UsedRange.Rows.Count

      special_last_row && special_last_row >= used_last_row ? special_last_row : used_last_row

    end

    def last_column
      special_last_column = @ole_worksheet.UsedRange.SpecialCells(RobustExcelOle::XlLastCell).Column
      used_last_column = @ole_worksheet.UsedRange.Columns.Count

      special_last_column >= used_last_column ? special_last_column : used_last_column
    end

  end

  public

  Sheet = Worksheet

end
