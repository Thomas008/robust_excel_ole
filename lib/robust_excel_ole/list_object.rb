# -*- coding: utf-8 -*-

module RobustExcelOle

  class ListRow
  end

  # This class essentially wraps a Win32Ole ListObject. 
  # You can apply all VBA methods (starting with a capital letter) 
  # that you would apply for a ListObject. 
  # See https://docs.microsoft.com/en-us/office/vba/api/excel.listobject#methods

  class ListObject < VbaObjects

    attr_reader :ole_table

    # constructs a list object (or table).
    # @param [Variable] worksheet_or_ole_listobject  a worksheet or a Win32Ole list object
    # @param [Variable] table_name_or_number         a table name or table number
    # @param [Array]    position                     a position of the upper left corner
    # @param [Integer]  rows_count                   number of rows
    # @param [Variable] columns_count_or_names       number of columns or array of column names
    # @return [ListObject] a ListObject object
    def initialize(worksheet_or_ole_listobject,
                   table_name_or_number = "",
                   position = [1,1],
                   rows_count = 1, 
                   columns_count_or_names = 1)
                   
      if (worksheet_or_ole_listobject.ListRows rescue nil)
        @ole_table = worksheet_or_ole_listobject
      else
        @worksheet = worksheet_or_ole_listobject.to_reo
        @ole_table = @worksheet.ListObjects.Item(table_name_or_number) rescue nil
      end
      unless @ole_table
        columns_count = 
          columns_count_or_names.is_a?(Integer) ? columns_count_or_names : columns_count_or_names.length 
        column_names = columns_count_or_names.respond_to?(:first) ? columns_count_or_names : []
        begin
          listobjects = @worksheet.ListObjects
          @ole_table = listobjects.Add(XlSrcRange, 
                                       @worksheet.range([position[0]..position[0]+rows_count-1,
                                                  position[1]..position[1]+columns_count-1]).ole_range,
                                       XlYes)
          @ole_table.Name = table_name_or_number
          @ole_table.HeaderRowRange.Value = [column_names] unless column_names.empty?
        rescue WIN32OLERuntimeError => msg # , Java::OrgRacobCom::ComFailException => msg
          raise TableError, "error #{$!.message}"
        end
      end

      ole_table = @ole_table
      @row_class = Class.new(ListRow) do

        @@ole_table = ole_table

        def initialize(row_number)
          @ole_listrow = @@ole_table.ListRows.Item(row_number)
        end

        # values of the row
        # @return [Array] values of the row
        def values
          begin
            @ole_listrow.Range.Value.first
          rescue WIN32OLERuntimeError
            raise TableError, "could not read values"
          end
        end

        # sets the values of the row
        # @param [Array] values of the row
        def set_values values
          begin
            @ole_listrow.Range.Value = [values]
          rescue WIN32OLERuntimeError
            raise TableError, "could not set values #{values.inspect}"
          end
        end

        # deletes the values of the row
        def delete_values
          begin
            @ole_listrow.Range.Value = [[].fill(nil,0..(@@ole_table.ListColumns.Count)-1)]
            nil
          rescue WIN32OLERuntimeError
            raise TableError, "could not delete values"
          end
        end
       
        def method_missing(name, *args)
          name_before_last_equal = name.to_s.split('=').first
          column_names = @@ole_table.HeaderRowRange.Value.first
          method_names = column_names.map{|c| c.underscore.gsub(/[^[\w\d]]/, '_')}
          column_name = column_names[method_names.index(name_before_last_equal)]
          if column_name
            ole_cell = @@ole_table.Application.Intersect(
              @ole_listrow.Range, @@ole_table.ListColumns(column_name).Range)
            define_getting_setting_method(ole_cell,name.to_s)            
            self.send(name, *args)
          else
            super
          end
        end

      private

        def define_getting_setting_method(ole_cell,name_str)
          if name_str[-1] != '='
            self.class.define_method(name_str) do
              ole_cell.Value
            end
          else
            self.class.define_method(name_str) do |value|
              ole_cell.Value = value
            end
          end
        end
      end

      # accesses a table row object
      # @param [Integer]  a row number (>= 1)
      # @return [ListRow] a object of dynamically constructed class with superclass ListRow 
      def [] row_number
        @row_class.new(row_number)
      end

    end

    # @return [Array] a list of column names
    def column_names
      @ole_table.HeaderRowRange.Value.first
    end

    # inserts a column
    # @param [Integer] position of the new column
    # @param [String]  name of the column
    def insert_column(position = 1, column_name = "")
      begin
        @ole_table.ListColumns.Add(position)
        rename_column(position,column_name)
      rescue WIN32OLERuntimeError, TableError
        raise TableError, "could not insert column at position #{position.inspect} with name #{column_name.inspect}"
      end
    end

    # deletes a column
    # @param [Variant] column number or column name
    def delete_column(column_number_or_name)
      begin
        @ole_table.ListColumns.Item(column_number_or_name).Delete
      rescue WIN32OLERuntimeError
        raise TableError, "could not delete column #{column_number_or_name.inspect}"
      end
    end

    # inserts a row
    # @param [Integer] position of the new row
    def insert_row(position = 1)
      begin
        @ole_table.ListRows.Add(position)
        position
      rescue WIN32OLERuntimeError
        raise TableError, "could not insert row at position #{position.inspect}"
      end
    end

    # deletes a row
    # @param [Integer] position of the old row
    def delete_row(row_number)
      begin
        @ole_table.ListRows.Item(row_number).Delete
      rescue WIN32OLERuntimeError
        raise TableError, "could not delete row #{row_number.inspect}"
      end
    end

    # deletes the contents of a column
    # @param [Variant] column number or column name
    def delete_column_values(column_number_or_name)
      begin
        column_name = @ole_table.ListColumns.Item(column_number_or_name).Range.Value.first
        @ole_table.ListColumns.Item(column_number_or_name).Range.Value = [column_name] + [].fill([nil],0..(@ole_table.ListRows.Count-1))
        nil
      rescue WIN32OLERuntimeError
        raise TableError, "could not delete contents of column #{column_number_or_name.inspect}"
      end
    end

    # deletes the contents of a row
    # @param [Integer] row number
    def delete_row_values(row_number)
      begin
        @ole_table.ListRows.Item(row_number).Range.Value = [[].fill(nil,0..(@ole_table.ListColumns.Count-1))]
        nil
      rescue WIN32OLERuntimeError
        raise TableError, "could not delete contents of row #{row_number.inspect}"
      end
    end

    # renames a row
    # @param [String] previous name or number of the column
    # @param [String] new name of the column   
    def rename_column(name_or_number, new_name)
      begin
        column_names = @ole_table.HeaderRowRange.Value.first
        position = name_or_number.respond_to?(:abs) ? name_or_number : (column_names.index(name_or_number) + 1)
        column_names[position-1] = new_name
        @ole_table.HeaderRowRange.Value = [column_names]
        new_name
      rescue
        raise TableError, "could not rename column #{name_or_number.inspect} to #{new_name.inspect}"
      end
    end

    # contents of a row
    # @param [Integer] row number
    # @return [Array] contents of a row
    def row_values(row_number)
      begin
        @ole_table.ListRows.Item(row_number).Range.Value.first
      rescue WIN32OLERuntimeError
        raise TableError, "could not read the values of row #{row_number.inspect}"
      end
    end

    # sets the contents of a row
    # @param [Integer] row number
    # @param [Array]   values of the row
    def set_row_values(row_number, values)
      begin
        @ole_table.ListRows.Item(row_number).Range.Value = [values]
      rescue WIN32OLERuntimeError
        raise TableError, "could not set the values of row #{row_number.inspect}"
      end
    end

    # @return [Array] contents of a column
    def column_values(column_number_or_name)
      begin
        @ole_table.ListColumns.Item(column_number_or_name).Range.Value[1,@ole_table.ListRows.Count].flatten
      rescue WIN32OLERuntimeError
        raise TableError, "could not read the values of column #{column_number_or_name.inspect}"
      end
    end

    # sets the contents of a column
    # @param [Integer] column name or column number
    # @param [Array]   contents of the column
    def set_column_values(column_number_or_name, values)
      begin
        column_name = @ole_table.ListColumns.Item(column_number_or_name).Range.Value.first
        @ole_table.ListColumns.Item(column_number_or_name).Range.Value = column_name + values.map{|v| [v]}
        values
      rescue WIN32OLERuntimeError
        raise TableError, "could not read the values of column #{column_number_or_name.inspect}"
      end
    end

    # deletes rows that have an empty contents
    def delete_empty_rows
      listrows = @ole_table.ListRows
      nil_array = [[].fill(nil,0..(@ole_table.ListColumns.Count-1))]
      i = 1
      while i <= listrows.Count do 
        row = listrows.Item(i)
        if row.Range.Value == nil_array
          row.Delete
        else
          i = i+1
        end
      end
    end

    # @private
    def to_s    
      @ole_table.Name.to_s
    end

    # @private
    def inspect    
      "#<ListObject:" + "#{@ole_table.Name}" + 
      " size:#{@ole_table.ListRows.Count}x#{@ole_table.ListColumns.Count}" +
      " worksheet:#{@ole_table.Parent.Name}" + " workbook:#{@ole_table.Parent.Parent.Name}" + ">"
    end

  private

    def method_missing(name, *args) 
      if name.to_s[0,1] =~ /[A-Z]/
        if ::ERRORMESSAGE_JRUBY_BUG
          begin
            @ole_table.send(name, *args)
          rescue Java::OrgRacobCom::ComFailException 
            raise VBAMethodMissingError, "unknown VBA property or method #{name.inspect}"
          end
        else
          begin
            @ole_table.send(name, *args)
          rescue NoMethodError 
            raise VBAMethodMissingError, "unknown VBA property or method #{name.inspect}"
          end
        end
      else
        super
      end
    end
  end

  # @private
  class TableError < WorksheetREOError
  end

  Table = ListObject
  TableRow = ListRow

end
