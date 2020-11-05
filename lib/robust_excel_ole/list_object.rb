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

    alias ole_object ole_table

    using RefinedArray

    # constructs a list object (or table).
    # @param [Variable] worksheet_or_listobject      a worksheet or a list object
    # @param [Variable] table_name_or_number         a table name or table number
    # @param [Array]    position                     a position of the upper left corner
    # @param [Integer]  rows_count                   number of rows
    # @param [Variable] columns_count_or_names       number of columns or array of column names
    # @return [ListObject] a ListObject object
    def initialize(worksheet_or_listobject,
                   table_name_or_number = "_table_name",
                   position = [1,1],
                   rows_count = 1, 
                   columns_count_or_names = 1)

      # ole_table is being assigned to the first parameter, if this parameter is a ListObject
      # otherwise the first parameter could be a worksheet, and get the ole_table via the ListObject name or number
      @ole_table = if worksheet_or_listobject.respond_to?(:ListRows)
        worksheet_or_listobject.ole_table
      else
        begin
          worksheet_or_listobject.send(:ListRows)
          worksheet_or_listobject
        rescue
          @worksheet = worksheet_or_listobject.to_reo
          puts "@worksheet: #{@worksheet.inspect}"
          @worksheet.ListObjects.Item(table_name_or_number) rescue nil
        end
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

        # returns the value of the cell with given column name or number
        # @param [Variant]  column number or column name
        # @return [Variant] value of the cell 
        def [] column_number_or_name
          begin
            ole_cell = @@ole_table.Application.Intersect(
              @ole_listrow.Range, @@ole_table.ListColumns.Item(column_number_or_name).Range)
            ole_cell.Value
          rescue WIN32OLERuntimeError
            raise TableRowError, "could not determine the value at column #{column_number_or_name}"
          end
        end


        def []=(column_number_or_name, value)
          begin
            ole_cell = @@ole_table.Application.Intersect(
              @ole_listrow.Range, @@ole_table.ListColumns.Item(column_number_or_name).Range)
            ole_cell.Value = value
          rescue WIN32OLERuntimeError
            raise TableRowError, "could not assign value #{value.inspect} to cell at column #{column_number_or_name}"
          end
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
            updated_values = self.values
            updated_values[0,values.length] = values
            @ole_listrow.Range.Value = [updated_values]
            values
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
          name_str = name.to_s
          core_name = name_str[-1]!='=' ? name_str : name_str[0..-2]
          column_names = @@ole_table.HeaderRowRange.Value.first
          column_name = column_names.find do |c|
            c == core_name ||  
            c.gsub(/\W/,'') == core_name ||
            c.replace_umlauts == core_name ||
            c.gsub(/\W/,'').replace_umlauts == core_name ||
            c.gsub(/\W/,'').replace_umlauts.underscore.gsub(/[^[\w\d]]/, '_').delete_multiple_underscores == core_name           
          end
          if column_name
            method_name = core_name.gsub(/\W/,'') + (name_str[-1]!='=' ? "" : "=")
            define_and_call_method(column_name,method_name,*args)
          else
            super(name, *args)
          end
        end

      private

        def define_and_call_method(column_name,method_name,*args)
          ole_cell = @@ole_table.Application.Intersect(
              @ole_listrow.Range, @@ole_table.ListColumns.Item(column_name).Range)
          define_getting_setting_method(ole_cell,method_name)            
          self.send(method_name, *args)
        end
       
        def define_getting_setting_method(ole_cell,name)
          if name[-1] != '='
            self.class.define_method(name) do
              ole_cell.Value
            end
          else
            self.class.define_method(name) do |value|
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
      begin
        @ole_table.HeaderRowRange.Value.first
      rescue WIN32OLERuntimeError
        raise TableError, "could not determine column names"
      end
    end

    # adds a row
    # @param [Integer] position of the new row
    # @param [Array]   values of the column
    def add_row(position = nil, contents = nil)
      begin
        @ole_table.ListRows.Add(position)
        set_row_values(position, contents) if contents
      rescue WIN32OLERuntimeError
        raise TableError, ("could not add row" + (" at position #{position.inspect}" if position))
      end
    end

    # adds a column    
    # @param [String]  name of the column
    # @param [Integer] position of the new column
    # @param [Array]   values of the column
    def add_column(column_name = nil, position = nil, contents = nil)
      begin
        new_column = @ole_table.ListColumns.Add(position)
        new_column.Name = column_name if column_name
        set_column_values(column_name, contents) if contents
      rescue WIN32OLERuntimeError, TableError
        raise TableError, ("could not add column"+ ("at position #{position.inspect} with name #{column_name.inspect}" if position))
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

    # deletes a column
    # @param [Variant] column number or column name
    def delete_column(column_number_or_name)
      begin
        @ole_table.ListColumns.Item(column_number_or_name).Delete
      rescue WIN32OLERuntimeError
        raise TableError, "could not delete column #{column_number_or_name.inspect}"
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

    # renames a row
    # @param [String] previous name or number of the column
    # @param [String] new name of the column   
    def rename_column(name_or_number, new_name)
      begin
        @ole_table.ListColumns.Item(name_or_number).Name = new_name
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
        updated_values = row_values(row_number)
        updated_values[0,values.length] = values
        @ole_table.ListRows.Item(row_number).Range.Value = [updated_values]
        values
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
        updated_values = column_values(column_number_or_name)
        updated_values[0,values.length] = values
        column_name = @ole_table.ListColumns.Item(column_number_or_name).Range.Value.first
        @ole_table.ListColumns.Item(column_number_or_name).Range.Value = column_name + updated_values.map{|v| [v]}
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

    # deletes columns that have an empty contents
    def delete_empty_columns
      listcolumns = @ole_table.ListColumns
      nil_array = [].fill([nil],0..(@ole_table.ListRows.Count-1))
      i = 1
      while i <= listcolumns.Count do 
        column = listcolumns.Item(i)
        if column.Range.Value[1..-1] == nil_array
          column.Delete
        else
          i = i+1
        end
      end
    end

    # finds all cells containing a given value
    # @param[Variant] value to find
    # @return [Array] win32ole cells containing the given value
    def find_cells(value)
      listrows = @ole_table.ListRows      
      result = []
      (1..listrows.Count).each do |row_number|
        row_values(row_number).find_all_indices(value).each do |col_number|
          result << @ole_table.Application.Intersect(listrows.Item(row_number).Range, 
                                                     @ole_table.ListColumns.Item(col_number+1).Range).to_reo
        end
      end
      result
    end

    # sorts the rows of the list object according to the given column
    # @param [Variant] column number or name
    # @option opts [Symbol]   sort order
    def sort(column_number_or_name, sort_order = :ascending)
      key_range = @ole_table.ListColumns.Item(column_number_or_name).Range
      @ole_table.Sort.SortFields.Clear
      sort_order_option = sort_order == :ascending ? XlAscending : XlDescending
      @ole_table.Sort.SortFields.Add(key_range, XlSortOnValues,sort_order_option,XlSortNormal)
      @ole_table.Sort.Apply
    end

    # @private
    # returns true, if the list object responds to VBA methods, false otherwise
    def alive?
      @ole_table.ListRows
      true
    rescue
      # trace $!.message
      false
    end

    # @private
    def to_s    
      @ole_table.Name.to_s
    end

    # @private
    def inspect    
      "#<ListObject:" + "#{@ole_table.Name}" + 
      " #{@ole_table.ListRows.Count}x#{@ole_table.ListColumns.Count}" +
      " #{@ole_table.Parent.Name}" + " #{@ole_table.Parent.Parent.Name}" + ">"
    end

    include MethodHelpers

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

  # @private
  class TableRowError < WorksheetREOError
  end

  Table = ListObject
  TableRow = ListRow

end
