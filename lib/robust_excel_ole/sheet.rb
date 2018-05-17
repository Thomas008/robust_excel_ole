# -*- coding: utf-8 -*-

module RobustExcelOle

  class Sheet < RangeOwners

    attr_reader :ole_worksheet
    attr_reader :workbook

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
      @workbook = book_class.new(self.Parent)
    end

    # returns name of the sheet
    def name
      @ole_worksheet.Name
    end

    # name the sheet
    # @param [String] new_name the new name of the sheet 
    def name= (new_name)
      begin
        @ole_worksheet.Name = new_name
      rescue WIN32OLERuntimeError => msg
        if msg.message =~ /800A03EC/ 
          raise NameAlreadyExists, "sheet name #{new_name.inspect} already exists"
        else
          raise UnexpectedError, "unexpected WIN32OLERuntimeError: #{msg.message}"
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
        begin
          @cells[yx] = RobustExcelOle::Cell.new(@ole_worksheet.Cells.Item(y, x))
        rescue
          raise RangeNotEvaluatable, "cannot read cell (#{p1.inspect},#{p2.inspect})"
        end
      else
        name = p1
        begin
          nameval(name) 
        rescue REOError
          rangeval(name)
        end
      end
    end

    # sets the value of a cell, if row and column are given
    # sets the value of a range if its name is given
    def []= (p1, p2, p3 = :__not_provided)
      if p3 != :__not_provided
        y, x, value = p1, p2, p3
        begin
          cell = @ole_worksheet.Cells.Item(y, x)
          cell.Value = value
          cell.Interior.ColorIndex = 42 # aqua-marin, 4-green
        rescue WIN32OLERuntimeError
          raise RangeNotEvaluatable, "cannot assign value #{p3.inspect} to cell (#{p1.inspect},#{p2.inspect})"
        end
      else
        name, value = p1, p2
        begin
          set_nameval(name, value, :color => 42) # aqua-marin, 4-green
        rescue REOError
          begin
            workbook.set_nameval(name, value)
          rescue REOError
            set_rangeval(name, value)
          end
        end
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
        #trace "WIN32OLERuntimeError: #{msg.message}"
        raise RangeNotEvaluatable, "cannot add name #{name.inspect} to cell with row #{row.inspect} and column #{column.inspect}"
      end
    end


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
          i+=1
        end
      end
    end

    def each_row(offset = 0)
      offset += 1
      1.upto(@end_row) do |row|
        next if row < offset
        yield RobustExcelOle::Range.new(@ole_worksheet.Range(@ole_worksheet.Cells(row, 1), @ole_worksheet.Cells(row, @end_column)))
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
        yield RobustExcelOle::Range.new(@ole_worksheet.Range(@ole_worksheet.Cells(1, column), @ole_worksheet.Cells(@end_row, column)))
      end
    end

    def each_column_with_index(offset = 0)
      each_column(offset) do |column_range|
        yield RobustExcelOle::Range.new(column_range), (column_range.column - 1 - offset)
      end
    end

    def row_range(row, range = nil)
      range ||= 1..@end_column
      RobustExcelOle::Range.new(@ole_worksheet.Range(@ole_worksheet.Cells(row , range.min ), @ole_worksheet.Cells(row , range.max )))
    end

    def col_range(col, range = nil)
      range ||= 1..@end_row
      RobustExcelOle::Range.new(@ole_worksheet.Range(@ole_worksheet.Cells(range.min , col ), @ole_worksheet.Cells(range.max , col )))
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

    def to_s    # :nodoc: #
      "#<Sheet: " + "#{"not alive " unless @workbook.alive?}" + "#{name}" + " #{File.basename(@workbook.stored_filename)} >"
    end

    def inspect    # :nodoc: #
      self.to_s
    end

  private

    def method_missing(name, *args)    # :nodoc: #
      if name.to_s[0,1] =~ /[A-Z]/ 
        begin
          @ole_worksheet.send(name, *args)
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
      special_last_row = @ole_worksheet.UsedRange.SpecialCells(RobustExcelOle::XlLastCell).Row
      used_last_row = @ole_worksheet.UsedRange.Rows.Count

      special_last_row >= used_last_row ? special_last_row : used_last_row
    end

    def last_column
      special_last_column = @ole_worksheet.UsedRange.SpecialCells(RobustExcelOle::XlLastCell).Column
      used_last_column = @ole_worksheet.UsedRange.Columns.Count

      special_last_column >= used_last_column ? special_last_column : used_last_column
    end    
  end
  
end
