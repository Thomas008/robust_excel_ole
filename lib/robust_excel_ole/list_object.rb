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

    def column_names
      @ole_table.HeaderRowRange.Value.first
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
