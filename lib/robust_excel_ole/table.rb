# -*- coding: utf-8 -*-

module RobustExcelOle

  class TableRow

    attr_reader :ole_row_range

    def initialize(row_number, ole_table)
      @ole_row_range = ole_table.ListRows.Item(row_number)
      column_names = ole_table.HeaderRowRange.Value.first
      column_names.each_with_index do |column_name, i|
        column_method_name = column_name.downcase
        define_singleton_method(column_method_name) do
          value = @ole_row_range.Range.Value.first[i-1]
          value
        end
        define_singleton_method(column_method_name + '=') do |value|
          @ole_row_range.Range.Value.first[i-1] = value
          value
        end
      end
    end
  end



  # This class essentially wraps a Win32Ole ListObject. 
  # You can apply all VBA methods (starting with a capital letter) 
  # that you would apply for a ListObject. 
  # See https://docs.microsoft.com/en-us/office/vba/api/excel.listobject#methods

  class Table < VbaObjects

    attr_reader :ole_table
    attr_reader :table

    def initialize(rows_count = 1, 
                   columns_count_or_names = 1, 
                   position = [1,1],
                   table_name = "",
                   worksheet = nil)
      columns_count = 
        columns_count_or_names.is_a?(Integer) ? columns_count_or_names : columns_count_or_names.length 
      column_names = columns_count_or_names.respond_to?(:first) ? columns_count_or_names : []
      @worksheet = worksheet                # ? worksheet : worksheet_class.new(self.Parent)
      begin
        listobjects = @worksheet.ListObjects
        @ole_table = listobjects.Add(XlSrcRange, 
                                     @worksheet.range([position[0]..rows_count,position[1]..columns_count]).ole_range,
                                     XlYes)
        @ole_table.Name = table_name
        @ole_table.HeaderRowRange.Value = [column_names] unless column_names.nil?
      rescue WIN32OLERuntimeError => msg # , Java::OrgRacobCom::ComFailException => msg
        raise TableError, "error #{$!.message}"
      end
      # reo-representation
      @table = []
      (1..rows_count).each do |row_number|
        row_object = TableRow.new(row_number, @ole_table)
        @table << row_object
      end
    end

    def each
      @ole_range.each_with_index do |ole_cell, index|
        yield cell(index){ole_cell}
      end
    end

    def [] index
      cell(index) {
        @ole_range.Cells.Item(index + 1)
      }
    end

  private

    def cell(index)
      @cells ||= []
      @cells[index + 1] ||= RobustExcelOle::Cell.new(yield,@worksheet)
    end

  public

    # returns flat array of the values of a given range
    # @params [Range] a range
    # @returns [Array] the values
    def values(range = nil)
      #result = map { |x| x.Value }.flatten
      result_unflatten = if !::RANGES_JRUBY_BUG
        map { |x| x.v }
      else
        self.v
      end
      result = result_unflatten.flatten
      if range
        relevant_result = []
        result.each_with_index { |row_or_column, i| relevant_result << row_or_column if range.include?(i) }
        relevant_result
      else
        result
      end
    end

    def v
      begin
        if !::RANGES_JRUBY_BUG
          self.Value
        else
          address_r1c1 = self.AddressLocal(true,true,XlR1C1)
          row, col = address_tool.as_integer_ranges(address_r1c1)
          values = []
          row.each do |r|
            values_col = []
            col.each{ |c| values_col << worksheet.Cells(r,c).Value}
            values << values_col
          end
          values
        end
      rescue WIN32OLERuntimeError, Java::OrgRacobCom::ComFailException => msg
        raise RangeNotEvaluatable, 'cannot read value'
      end 

    end

    def v=(value)
      begin
        if !::RANGES_JRUBY_BUG
          ole_range.Value = value
        else
          address_r1c1 = ole_range.AddressLocal(true,true,XlR1C1)
          row, col = address_tool.as_integer_ranges(address_r1c1)
          row.each_with_index do |r,i|
            col.each_with_index do |c,j|
              ole_range.Cells(i+1,j+1).Value = (value.respond_to?(:first) ? value[i][j] : value)
            end
          end
        end
        value
      rescue WIN32OLERuntimeError, Java::OrgRacobCom::ComFailException => msg  
        raise RangeNotEvaluatable, "cannot assign value to range #{address_r1c1.inspect}"
      end
    end

    alias_method :value, :v
    alias_method :value=, :v=

    # copies a range
    # @params [Address or Address-Array] address or upper left position of the destination range
    # @options [Worksheet] the destination worksheet
    # @options [Hash] options: :transpose, :values_only
    def copy(dest_address1, sheet_or_dest_address2 = :__not_provided, options_or_sheet = :__not_provided, not_provided_or_options = :__not_provided)
      dest_address = if sheet_or_dest_address2.is_a?(Object::Range) or sheet_or_dest_address2.is_a?(Integer)
        [dest_address1,sheet_or_dest_address2] 
      else
        dest_address1
      end
      dest_sheet = if sheet_or_dest_address2.is_a?(Worksheet)
        sheet_or_dest_address2
      else
        if options_or_sheet.is_a?(Worksheet)
          options_or_sheet
        else
          @worksheet
        end
      end
      options = if options_or_sheet.is_a?(Hash)
        options_or_sheet 
      else
        if not_provided_or_options.is_a?(Hash)
          not_provided_or_options
        else
          { }
        end
      end
      rows, columns = address_tool.as_integer_ranges(dest_address)
      dest_address_is_position = (rows.min == rows.max && columns.min == columns.max)
      dest_range_address = if (not dest_address_is_position) 
          [rows.min..rows.max,columns.min..columns.max]
        else
          if (not options[:transpose])
            [rows.min..rows.min+self.Rows.Count-1,
             columns.min..columns.min+self.Columns.Count-1]
          else
            [rows.min..rows.min+self.Columns.Count-1,
             columns.min..columns.min+self.Rows.Count-1]
          end
        end
      dest_range = dest_sheet.range(dest_range_address)
      begin
        if options[:values_only]
          # dest_range.Value = options[:transpose] ? self.Value.transpose : self.Value
          dest_range.v = options[:transpose] ? self.v.transpose : self.v
        else
          if dest_range.worksheet.workbook.excel == @worksheet.workbook.excel 
            if options[:transpose]
              self.Copy              
              #dest_range.PasteSpecial('transpose' => true) 
              dest_range.PasteSpecial(XlPasteAll,XlPasteSpecialOperationNone,false,true)
            else
              #self.Copy('destination' => dest_range.ole_range)
              self.Copy(dest_range.ole_range)
            end            
          else
            if options[:transpose]
              added_sheet = @worksheet.workbook.add_sheet
              self.copy_special(dest_address, added_sheet, :transpose => true)
              added_sheet.range(dest_range_address).copy_special(dest_address,dest_sheet)
              @worksheet.workbook.excel.with_displayalerts(false) {added_sheet.Delete}
            else
              self.Copy
              #dest_sheet.Paste('destination' => dest_range.ole_range)
              dest_sheet.Paste(dest_range.ole_range)
            end
          end
        end
      rescue WIN32OLERuntimeError, Java::OrgRacobCom::ComFailException => msg
        raise RangeNotCopied, 'cannot copy range'
      end
    end

    # becomes copy
    # copies a range
    # @params [Address or Address-Array] address or upper left position of the destination range
    # @options [Worksheet] the destination worksheet
    # @options [Hash] options: :transpose, :values_only
    def copy_special(dest_address, dest_sheet = :__not_provided, options = { })
      rows, columns = address_tool.as_integer_ranges(dest_address)
      dest_sheet = @worksheet if dest_sheet == :__not_provided
      dest_address_is_position = (rows.min == rows.max && columns.min == columns.max)
      dest_range_address = if (not dest_address_is_position) 
          [rows.min..rows.max,columns.min..columns.max]
        else
          if (not options[:transpose])
            [rows.min..rows.min+self.Rows.Count-1,
             columns.min..columns.min+self.Columns.Count-1]
          else
            [rows.min..rows.min+self.Columns.Count-1,
             columns.min..columns.min+self.Rows.Count-1]
          end
        end
      dest_range = dest_sheet.range(dest_range_address)
      begin
        if options[:values_only]
          dest_range.Value = options[:transpose] ? self.Value.transpose : self.Value
        else
          if dest_range.worksheet.workbook.excel == @worksheet.workbook.excel     
            if options[:transpose]
              self.Copy
              #dest_range.PasteSpecial('transpose' => true) 
              dest_range.PasteSpecial(XlPasteAll,XlPasteSpecialOperationNone,false,true)
            else
              #self.Copy('destination' => dest_range.ole_range)
              self.Copy(dest_range.ole_range)
            end            
          else
            if options[:transpose]
              added_sheet = @worksheet.workbook.add_sheet
              self.copy_special(dest_address, added_sheet, :transpose => true)
              added_sheet.range(dest_range_address).copy_special(dest_address,dest_sheet)
              @worksheet.workbook.excel.with_displayalerts(false) {added_sheet.Delete}
            else
              self.Copy
              #dest_sheet.Paste('destination' => dest_range.ole_range)
              dest_sheet.Paste(dest_range.ole_range)
            end
          end
        end
      rescue WIN32OLERuntimeError, Java::OrgRacobCom::ComFailException => msg 
        raise RangeNotCopied, 'cannot copy range'
      end
    end

    def == other_range
      other_range.is_a?(Range) &&
        self.worksheet == other_range.worksheet
        self.Address == other_range.Address 
    end

    # @private
    def excel
      @worksheet.workbook.excel
    end

    # @private
    def self.worksheet_class   
      @worksheet_class ||= begin
        module_name = parent_name
        "#{module_name}::Worksheet".constantize
      rescue NameError => e
        Worksheet
      end
    end

    # @private
    def worksheet_class 
      self.class.worksheet_class
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

end
