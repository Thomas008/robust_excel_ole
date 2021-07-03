# -*- coding: utf-8 -*-

module RobustExcelOle

  using StringRefinement

  class ListRow < VbaObjects

    attr_reader :ole_tablerow

    alias ole_object ole_tablerow

    def initialize(rownumber_or_oletablerow)
      @ole_tablerow = if rownumber_or_oletablerow.is_a?(ListRow)
        rownumber_or_oletablerow.ole_tablerow
      else
        begin
          rownumber_or_oletablerow.Parent.send(:ListRows)
          rownumber_or_oletablerow
        rescue
          ole_table.ListRows.Item(rownumber_or_oletablerow)
        end
      end
    end

    # returns the value of the cell with given column name or number
    # @param [Variant]  column number or column name
    # @return [Variant] value of the cell 
    def [] column_number_or_name
      column_number_or_name = column_number_or_name.to_s if column_number_or_name.is_a?(Symbol)
      ole_cell = ole_table.Application.Intersect(
        @ole_tablerow.Range, ole_table.ListColumns.Item(column_number_or_name).Range)
      value = ole_cell.Value
      value.respond_to?(:gsub) ? value.encode('utf-8') : value
    rescue WIN32OLERuntimeError
      raise TableRowError, "could not determine the value at column #{column_number_or_name}\n#{$!.message}"
    end

    # sets the value of the cell with given column name or number
    # @param [Variant] column number or column name
    # @param [Variant] value of the cell
    def []=(column_number_or_name, value)
      begin
        column_number_or_name = column_number_or_name.to_s if column_number_or_name.is_a?(Symbol)
        ole_cell = ole_table.Application.Intersect(
          @ole_tablerow.Range, ole_table.ListColumns.Item(column_number_or_name).Range)
        ole_cell.Value = value
      rescue WIN32OLERuntimeError
        raise TableRowError, "could not assign value #{value.inspect} to cell at column #{column_number_or_name}\n#{$!.message}"
      end
    end

    # values of the row
    # @return [Array] values of the row
    def values
      value = @ole_tablerow.Range.Value
      return value if value==[nil]
      value = if !value.respond_to?(:pop)
        [value]
      elsif value.first.respond_to?(:pop)
        value.first
      end
      value.map{|v| v.respond_to?(:gsub) ? v.encode('utf-8') : v}
    rescue WIN32OLERuntimeError
      raise TableError, "could not read values\n#{$!.message}"
    end

    # sets the values of the row
    # @param [Array] values of the row
    def values= values
      updated_values = self.values
      updated_values[0,values.length] = values
      @ole_tablerow.Range.Value = [updated_values]
      values
    rescue WIN32OLERuntimeError
      raise TableError, "could not set values #{values.inspect}\n#{$!.message}"
    end

    # key-value pairs of the row
    # @return [Hash] key-value pairs of the row
    def keys_values
      ole_table.column_names.zip(values).to_h
    end

    alias set_values values=
    alias to_a values
    alias to_h keys_values

    # deletes the values of the row
    def delete_values
      @ole_tablerow.Range.Value = [[].fill(nil,0..(ole_table.ListColumns.Count)-1)]
      nil
    rescue WIN32OLERuntimeError
      raise TableError, "could not delete values\n#{$!.message}"
    end

    def == other_listrow
      other_listrow.is_a?(ListRow) && other_listrow.values == self.values
    end

    # @private
    def workbook
      @workbook ||= workbook_class.new(ole_table.Parent.Parent)
    end

    # @private
    def self.workbook_class  
      @workbook_class ||= begin
        module_name = parent_name
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
    def column_names
      ole_table.HeaderRowRange.Value.first
    end

  private

    def valid_similar_names meth_name
      [
        meth_name,
        meth_name.gsub(/\W/,'_'),
        meth_name.underscore,
        meth_name.underscore.gsub(/\W/,'_'),
        meth_name.replace_umlauts.gsub(/\W/,'_'),
        meth_name.replace_umlauts.underscore.gsub(/\W/,'_')
      ].uniq
    end

  public

    # @private
    def methods
      @methods ||= begin
        arr = column_names.map{ |c| valid_similar_names(c) }.flatten
        (arr + arr.map{|m| m + '='}).map{|m| m.to_sym} + super
      end
    end

    # @private
    def respond_to?(meth_name)
      methods.include?(meth_name)
    end

    # @private
    def to_s    
      inspect  
    end

    # @private
    def inspect    
      "#<ListRow: index:#{@ole_tablerow.Index} size:#{ole_table.ListColumns.Count} #{ole_table.Name}>"
    end

  private

    def method_missing(meth_name, *args)
      # this should not happen:
      raise(TableRowError, "internal error: ole_table not defined") unless self.class.method_defined?(:ole_table)
      if respond_to?(meth_name)
        core_name = meth_name.to_s.chomp('=')
        column_name = column_names.find{ |c| valid_similar_names(c).include?(core_name) }
        define_and_call_method(column_name, meth_name, *args) if column_name
      else
        super(meth_name, *args)
      end
    end

    def define_and_call_method(column_name, method_name, *args)
      #column_name = column_name.force_encoding('cp850')
      ole_cell = ole_table.Application.Intersect(
          @ole_tablerow.Range, ole_table.ListColumns.Item(column_name).Range)
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

  # @private
  class TableRowError < WorksheetREOError
  end
  
  TableRow = ListRow

end
