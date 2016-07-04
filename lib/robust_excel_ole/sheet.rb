# -*- coding: utf-8 -*-

module RobustExcelOle

  class Sheet < REOCommon
    attr_reader :worksheet

    def initialize(win32_worksheet)
      @worksheet = win32_worksheet
      if @worksheet.ProtectContents
        @worksheet.Unprotect
        @end_row = last_row
        @end_column = last_column
        @worksheet.Protect
      else
        @end_row = last_row
        @end_column = last_column
      end
    end

    def workbook
      book_class.new(self.Parent)
    end

    # returns name of the sheet
    def name
      @worksheet.Name
    end

    # name the sheet
    # @param [String] new_name the new name of the sheet 
    def name= (new_name)
      begin
        @worksheet.Name = new_name
      rescue WIN32OLERuntimeError => msg
        if msg.message =~ /800A03EC/ 
          raise ExcelErrorSheet, "sheet name #{new_name.inspect} already exists"
        else
          trace "#{msg.message}"
          raise ExcelErrorSheetUnknown
        end
      end
    end

    # returns the value of a cell, if row and column are given
    # returns the value of a range if its name is given 
    def [] p1, p2 = :__not_provided
      if p2 != :__not_provided  
        y, x = p1, p2
        yx = "#{y}_#{x}"
        @cells = { }
        @cells[yx] = RobustExcelOle::Cell.new(@worksheet.Cells.Item(y, x))
      else
        name = p1
        begin
          nameval(name) 
        rescue SheetError
          begin
            book_class.new(self.Parent).nameval(name)
          rescue ExcelError
            rangeval(name)
          end
        end
      end
    end

    # sets the value of a cell, if row and column are given
    # sets the value of a range if its name is given
    def []= (p1, p2, p3 = :__not_provided)
      if p3 != :__not_provided
        y, x, value = p1, p2, p3
        @worksheet.Cells.Item(y, x).Value = value
      else
        name, value = p1, p2
        begin
          set_nameval(name, value) 
        rescue SheetError
          begin
            workbook.set_nameval(name, value)
          rescue ExcelError
            set_rangeval(name, value)
          end
        end
      end
    end

    def each
      each_row do |row_range|
        row_range.each do |cell|
          yield cell
        end
      end
    end

    def each_row(offset = 0)
      offset += 1
      1.upto(@end_row) do |row|
        next if row < offset
        yield RobustExcelOle::Range.new(@worksheet.Range(@worksheet.Cells(row, 1), @worksheet.Cells(row, @end_column)))
      end
    end

    def each_row_with_index(offset = 0)
      each_row(offset) do |row_range|
        yield RobustExcelOle::Range.new(row_range), (row_range.row - 1 - offset)
      end
    end

    def each_column(offset = 0)
      offset += 1
      1.upto(@end_column) do |column|
        next if column < offset
        yield RobustExcelOle::Range.new(@worksheet.Range(@worksheet.Cells(1, column), @worksheet.Cells(@end_row, column)))
      end
    end

    def each_column_with_index(offset = 0)
      each_column(offset) do |column_range|
        yield RobustExcelOle::Range.new(column_range), (column_range.column - 1 - offset)
      end
    end

    def row_range(row, range = nil)
      range ||= 1..@end_column
      RobustExcelOle::Range.new(@worksheet.Range(@worksheet.Cells(row , range.min ), @worksheet.Cells(row , range.max )))
    end

    def col_range(col, range = nil)
      range ||= 1..@end_row
      RobustExcelOle::Range.new(@worksheet.Range(@worksheet.Cells(range.min , col ), @worksheet.Cells(range.max , col )))
    end

    # returns the contents of a range
    # evaluates the formula if the contents is a formula
    # if no contents could be returned, then return default value, if provided, raise error otherwise
    # @param [String] name  the name of a range
    # @param [Hash]   opts  the options
    # @option opts [Variant] :default default value (default: nil)
    # @raise SheetError if name is not defined or if value of the range cannot be evaluated  
    def nameval(name, opts = {:default => nil})
      begin
        name_obj = self.Names.Item(name)
      rescue WIN32OLERuntimeError
        return opts[:default] if opts[:default]
        raise SheetError, "name #{name.inspect} not in #{self.Name}"
      end
      begin
        value = name_obj.RefersToRange.Value
      rescue  WIN32OLERuntimeError
        begin
          value = self.Evaluate(name_obj.Name)
        rescue WIN32OLERuntimeError
          return opts[:default] if opts[:default]
          raise SheetError, "cannot evaluate name #{name.inspect} in #{self.Name}"
        end
      end
      if value == -2146826259
        return opts[:default] if opts[:default]
        raise SheetError, "cannot evaluate name #{name.inspect} in #{self.Name}"
      end 
      return opts[:default] if (value.nil? && opts[:default])
      value      
    end
    
    # assigns a value to a range
    # @param [String]  name   the name of a range
    # @param [Variant] value  the assigned value
    # @raise SheetError if name is not in the sheet or the value cannot be assigned
    def set_nameval(name,value)
      begin
        name_obj = self.Names.Item(name)
      rescue WIN32OLERuntimeError
        raise SheetError, "name #{name.inspect} not in #{self.name}"
      end
      begin
        name_obj.RefersToRange.Value = value
      rescue  WIN32OLERuntimeError
        raise SheetError, "cannot assign value to range named #{name.inspect} in #{self.name}"
      end
    end

    # returns the contents of a range with a defined local name
    # evaluates the formula if the contents is a formula
    # if no contents could be returned, then return default value, if provided, raise error otherwise
    # @param  [String]      name      the name of a range
    # @param  [Hash]        opts      the options
    # @option opts [Symbol] :default  the default value that is provided if no contents could be returned
    # @raise  SheetError if range name is not definied in the worksheet or if range value could not be evaluated
    # @return [Variant] the contents of a range with given name   
    def rangeval(name, opts = {:default => nil})
      begin
        range = self.Range(name)
      rescue WIN32OLERuntimeError
        return opts[:default] if opts[:default]
        raise SheetError, "name #{name.inspect} not in #{self.name}"
      end
      begin
        value = range.Value
      rescue  WIN32OLERuntimeError
        return opts[:default] if opts[:default]
        raise SheetError, "cannot determine value of range named #{name.inspect} in #{self.name}"
      end
      return opts[:default] if (value.nil? && opts[:default])
      value
    end

    # assigns a value to a range given a defined local name
    # @param [String]  name   the name of a range
    # @param [Variant] value  the assigned value
    # @raise SheetError if name is not in the sheet or the value cannot be assigned
    def set_rangeval(name,value)
      begin
        range = self.Range(name)
      rescue WIN32OLERuntimeError
        raise SheetError, "name #{name.inspect} not in #{self.name}"
      end
      begin
        range.Value = value
      rescue  WIN32OLERuntimeError
        raise SheetError, "cannot assign value to range named #{name.inspect} in #{self.name}"
      end
    end
    
    # assigns a name to a range (a cell) given by an address
    # @param [String] name   the range name
    # @param [Fixnum] row    the row
    # @param [Fixnum] column the column
    def set_name(name,row,column)
      begin
        old_name = self[row,column].Name.Name rescue nil
        if old_name
          self[row,column].Name.Name = name
        else
          address = "Z" + row.to_s + "S" + column.to_s 
          self.Names.Add("Name" => name, "RefersToR1C1" => "=" + address)
        end
      rescue WIN32OLERuntimeError => msg
        trace "WIN32OLERuntimeError: #{msg.message}"
        raise SheetError, "cannot add name #{name.inspect} to cell with row #{row.inspect} and column #{column.inspect}"
      end
    end

    def respond_to?(name, include_private = false)  # :nodoc: #    
      super
    end

    def methods   # :nodoc: # 
      super
    end

    def self.book_class   # :nodoc: #
      @book_class ||= begin
        module_name = self.parent_name
        "#{module_name}::Book".constantize
      rescue NameError => e
        book
      end
    end

    def book_class        # :nodoc: #
      self.class.book_class
    end

    private

    def method_missing(name, *args)    # :nodoc: #
      if name.to_s[0,1] =~ /[A-Z]/ 
        begin
          @worksheet.send(name, *args)
        rescue WIN32OLERuntimeError => msg
          if msg.message =~ /unknown property or method/
            raise VBAMethodMissingError, "unknown VBA property or method #{name.inspect}"
          else 
            raise msg
          end
        end
      else  
        super 
      end
    end


    def last_row
      special_last_row = @worksheet.UsedRange.SpecialCells(RobustExcelOle::XlLastCell).Row
      used_last_row = @worksheet.UsedRange.Rows.Count

      special_last_row >= used_last_row ? special_last_row : used_last_row
    end

    def last_column
      special_last_column = @worksheet.UsedRange.SpecialCells(RobustExcelOle::XlLastCell).Column
      used_last_column = @worksheet.UsedRange.Columns.Count

      special_last_column >= used_last_column ? special_last_column : used_last_column
    end    
  end

  public
  
  class SheetError < RuntimeError    # :nodoc: #
  end

end
