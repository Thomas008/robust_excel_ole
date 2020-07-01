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
    attr_reader :table

    def initialize(worksheet,
                   rows_count = 1, 
                   columns_count_or_names = 1, 
                   position = [1,1],
                   table_name = "")
      # vba representation
      columns_count = 
        columns_count_or_names.is_a?(Integer) ? columns_count_or_names : columns_count_or_names.length 
      column_names = columns_count_or_names.respond_to?(:first) ? columns_count_or_names : []
      @worksheet = worksheet                # ? worksheet : worksheet_class.new(self.Parent)
      begin
        listobjects = @worksheet.ListObjects
        @ole_table = listobjects.Add(XlSrcRange, 
                                     @worksheet.range([position[0]..position[0]+rows_count-1,
                                                position[1]..position[1]+columns_count-1]).ole_range,
                                     XlYes)
        @ole_table.Name = table_name
        @ole_table.HeaderRowRange.Value = [column_names] unless column_names.empty?
      rescue WIN32OLERuntimeError => msg # , Java::OrgRacobCom::ComFailException => msg
        raise TableError, "error #{$!.message}"
      end
      # reo representation
      ole_table = @ole_table
      @table = []
      row_class = Class.new(ListRow) do
        @@ole_table = ole_table

        def initialize(row_number)
          @row_number = row_number         
        end
        
        def method_missing(name, *args)
          ole_row_range = @@ole_table.ListRows.Item(@row_number)
          column_names = @@ole_table.HeaderRowRange.Value.first
          name_before_last_equal = name.to_s.split('=').first
          columns = column_names.map{|c| c.underscore}
          if columns.include?(name_before_last_equal)
            index = columns.index(name_before_last_equal)
            if name.to_s[-1] != '='
              self.send(:define_singleton_method, name) do
                ole_row_range.Range.Value.first[index]
              end
            else
              self.class.send(:define_singleton_method, name) do |value|
                values_array = ole_row_range.Range.Value 
                values_array.first[index] = value
                ole_row_range.Range.Value = values_array
              end
            end
            self.send(name, *args)
          else
            super
          end
        end
      end     
      (1..rows_count).each do |row_number|
        row_object = row_class.new(row_number)
        @table << row_object
      end
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
