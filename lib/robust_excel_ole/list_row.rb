# -*- coding: utf-8 -*-

module RobustExcelOle

  using StringRefinement

  class ListRow    

    def initialize(row_number)
      @ole_listrow = ole_table.ListRows.Item(row_number)
    end

    # returns the value of the cell with given column name or number
    # @param [Variant]  column number or column name
    # @return [Variant] value of the cell 
    def [] column_number_or_name
      begin
        ole_cell = ole_table.Application.Intersect(
          @ole_listrow.Range, ole_table.ListColumns.Item(column_number_or_name).Range)
        ole_cell.Value
      rescue WIN32OLERuntimeError
        raise TableRowError, "could not determine the value at column #{column_number_or_name}"
      end
    end

    # sets the value of the cell with given column name or number
    # @param [Variant] column number or column name
    # @param [Variant] value of the cell
    def []=(column_number_or_name, value)
      begin
        ole_cell = ole_table.Application.Intersect(
          @ole_listrow.Range, ole_table.ListColumns.Item(column_number_or_name).Range)
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
        @ole_listrow.Range.Value = [[].fill(nil,0..(ole_table.ListColumns.Count)-1)]
        nil
      rescue WIN32OLERuntimeError
        raise TableError, "could not delete values"
      end
    end        

    def method_missing(name, *args)
      name_str = name.to_s
      core_name = name_str[-1]!='=' ? name_str : name_str[0..-2]
      column_names = ole_table.HeaderRowRange.Value.first
      column_name = column_names.find do |c|
        c == core_name ||
        c.gsub(/\W/,'_') == core_name ||
        c.underscore == core_name ||
        c.underscore.gsub(/\W/,'_') == core_name ||
        c.replace_umlauts.gsub(/\W/,'_') == core_name ||
        c.replace_umlauts.underscore.gsub(/\W/,'_') == core_name 
      end         
      if column_name
        appended_eq = (name_str[-1]!='=' ? "" : "=")
        method_name = core_name.replace_umlauts.underscore + appended_eq 
        define_and_call_method(column_name,method_name,*args)
      else
        super(name, *args)
      end
    end

    # @private
    def to_s    
      inspect  
    end

    # @private
    def inspect    
      "#<ListRow: " + "index:#{@ole_listrow.Index}" + " size:#{ole_table.ListColumns.Count}" + " #{ole_table.Name}" + ">"
    end

  private

    def define_and_call_method(column_name,method_name,*args)
      ole_cell = ole_table.Application.Intersect(
          @ole_listrow.Range, ole_table.ListColumns.Item(column_name).Range)
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
  
  TableRow = ListRow

end