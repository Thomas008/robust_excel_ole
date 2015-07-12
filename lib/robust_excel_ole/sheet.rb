# -*- coding: utf-8 -*-
module RobustExcelOle
  class Sheet
    attr_reader :sheet
    include Enumerable

    def initialize(win32_worksheet)
      @sheet = win32_worksheet
      if @sheet.ProtectContents
        @sheet.Unprotect
        @end_row = last_row
        @end_column = last_column
        @sheet.Protect
      else
        @end_row = last_row
        @end_column = last_column
      end
    end

    def name
      @sheet.Name
    end

    def name= (new_name)
      @sheet.Name = new_name
    end

    def [] y, x
      yx = "#{y+1}_#{x+1}"
      @cells ||= { }
      @cells[yx] ||= RobustExcelOle::Cell.new(@sheet.Cells.Item(y+1, x+1))
    end

    def []= (y, x, value)
      @sheet.Cells.Item(y+1, x+1).Value = value
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
        yield RobustExcelOle::Range.new(@sheet.Range(@sheet.Cells(row, 1), @sheet.Cells(row, @end_column)))
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
        yield RobustExcelOle::Range.new(@sheet.Range(@sheet.Cells(1, column), @sheet.Cells(@end_row, column)))
      end
    end

    def each_column_with_index(offset = 0)
      each_column(offset) do |column_range|
        yield RobustExcelOle::Range.new(column_range), (column_range.column - 1 - offset)
      end
    end

    def row_range(row, range = nil)
      range ||= 0..@end_column - 1
      RobustExcelOle::Range.new(@sheet.Range(@sheet.Cells(row + 1, range.min + 1), @sheet.Cells(row + 1, range.max + 1)))
    end

    def col_range(col, range = nil)
      range ||= 0..@end_row - 1
      RobustExcelOle::Range.new(@sheet.Range(@sheet.Cells(range.min + 1, col + 1), @sheet.Cells(range.max + 1, col + 1)))
    end

    # returns the contents of a range with given name
    # if no contents could returned, then return default value, if a default value was provided
    #                                raise an error, otherwise
    def nvalue(name, opts = {:default => nil})
      begin
        item = self.Names.Item(name)
      rescue WIN32OLERuntimeError
        return opts[:default] if opts[:default]
        raise SheetErrorNValue, "name #{name} not in sheet"
      end
      begin
        value = item.RefersToRange.Value
      rescue  WIN32OLERuntimeError
        return opts[:default] if opts[:default]
        raise SheetErrorNValue, "RefersToRange of name #{name}"
      end
      value
    end

    # adds a name to a range given by an address
    # row, column are oriented at 0
    def add_name(row,column,name)
      begin
        self[row,column].Name.Name = name
      rescue WIN32OLERuntimeError => msg
        #puts "WIN32OLERuntimeError: #{msg.message}"
        raise SheetError, "cannot add name #{name} to cell with row #{row} and column #{column}"
      end
    end

    def method_missing(id, *args)  # :nodoc: #
      @sheet.send(id, *args)
    end

    private
    def last_row
      special_last_row = @sheet.UsedRange.SpecialCells(RobustExcelOle::XlLastCell).Row
      used_last_row = @sheet.UsedRange.Rows.Count

      special_last_row >= used_last_row ? special_last_row : used_last_row
    end

    def last_column
      special_last_column = @sheet.UsedRange.SpecialCells(RobustExcelOle::XlLastCell).Column
      used_last_column = @sheet.UsedRange.Columns.Count

      special_last_column >= used_last_column ? special_last_column : used_last_column
    end    
  end

  public
  
  class SheetError < RuntimeError    # :nodoc: #
  end

  class SheetErrorNValue < WIN32OLERuntimeError  # :nodoc: #
  end

end
