# -*- coding: utf-8 -*-
module RobustExcelOle

  # This class essentially wraps a Win32Ole Range object. 
  # You can apply all VBA methods (starting with a capital letter) 
  # that you would apply for a Range object. 
  # See https://docs.microsoft.com/en-us/office/vba/api/excel.worksheet#methods

  class Range < REOCommon
    include Enumerable
    attr_reader :ole_range
    attr_reader :worksheet

    def initialize(win32_range)
      @ole_range = win32_range
      @worksheet = worksheet_class.new(self.Parent)
    end

    def each
      @ole_range.each do |row_or_column|
        yield RobustExcelOle::Cell.new(row_or_column)
      end
    end

    # returns flat array of the values of a given range
    # @params [Range] a range
    # @returns [Array] the values
    def values(range = nil)
      result = map { |x| x.Value }.flatten
      if range
        relevant_result = []
        result.each_with_index { |row_or_column, i| relevant_result << row_or_column if range.include?(i) }
        relevant_result
      else
        result
      end
    end

    def v
      self.Value
    end

    def [] index
      @cells = []
      @cells[index + 1] = RobustExcelOle::Cell.new(@ole_range.Cells.Item(index + 1))
    end

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
      address = Address.new(dest_address)
      dest_sheet = @worksheet if dest_sheet == :__not_provided
      dest_address_is_position = (address.rows.min == address.rows.max && address.columns.min == address.columns.max)
      dest_range_address = if (not dest_address_is_position) 
          [address.rows.min..address.rows.max,address.columns.min..address.columns.max]
        else
          if (not options[:transpose])
            [address.rows.min..address.rows.min+self.Rows.Count-1,
             address.columns.min..address.columns.min+self.Columns.Count-1]
          else
            [address.rows.min..address.rows.min+self.Columns.Count-1,
             address.columns.min..address.columns.min+self.Rows.Count-1]
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
              dest_range.PasteSpecial(:transpose => true) 
            else
              self.Copy(:destination => dest_range.ole_range)
            end            
          else
            if options[:transpose]
              added_sheet = @worksheet.workbook.add_sheet
              self.copy_special(dest_address, added_sheet, :transpose => true)
              added_sheet.range(dest_range_address).copy_special(dest_address,dest_sheet)
              @worksheet.workbook.excel.with_displayalerts(false) {added_sheet.Delete}
            else
              self.Copy
              dest_sheet.Paste(:destination => dest_range.ole_range)
            end
          end
        end
      rescue WIN32OLERuntimeError
        raise RangeNotCopied, 'cannot copy range'
      end
    end

    # becomes copy
    # copies a range
    # @params [Address or Address-Array] address or upper left position of the destination range
    # @options [Worksheet] the destination worksheet
    # @options [Hash] options: :transpose, :values_only
    def copy_special(dest_address, dest_sheet = :__not_provided, options = { })
      address = Address.new(dest_address)
      dest_sheet = @worksheet if dest_sheet == :__not_provided
      dest_address_is_position = (address.rows.min == address.rows.max && address.columns.min == address.columns.max)
      dest_range_address = if (not dest_address_is_position) 
          [address.rows.min..address.rows.max,address.columns.min..address.columns.max]
        else
          if (not options[:transpose])
            [address.rows.min..address.rows.min+self.Rows.Count-1,
             address.columns.min..address.columns.min+self.Columns.Count-1]
          else
            [address.rows.min..address.rows.min+self.Columns.Count-1,
             address.columns.min..address.columns.min+self.Rows.Count-1]
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
              dest_range.PasteSpecial(:transpose => true) 
            else
              self.Copy(:destination => dest_range.ole_range)
            end            
          else
            if options[:transpose]
              added_sheet = @worksheet.workbook.add_sheet
              self.copy_special(dest_address, added_sheet, :transpose => true)
              added_sheet.range(dest_range_address).copy_special(dest_address,dest_sheet)
              @worksheet.workbook.excel.with_displayalerts(false) {added_sheet.Delete}
            else
              self.Copy
              dest_sheet.Paste(:destination => dest_range.ole_range)
            end
          end
        end
      rescue WIN32OLERuntimeError
        raise RangeNotCopied, 'cannot copy range'
      end
    end

=begin
    # copies a range
    # @params [Address] address of the destination range
    # @options [Worksheet] the worksheet in which to copy
    def copy(address, sheet = :__not_provided, third_argument_deprecated = :__not_provided)
      if third_argument_deprecated != :__not_provided
        address = [address,sheet]
        sheet = third_argument_deprecated
      end
      address = Address.new(address)
      sheet = @worksheet if sheet == :__not_provided
      destination_range = sheet.range([address.rows.min..address.rows.max,
                                       address.columns.min..address.columns.max]).ole_range
      if sheet.workbook.excel == @worksheet.workbook.excel
        begin
          self.Copy(:destination => destination_range)
        rescue WIN32OLERuntimeError
          raise RangeNotCopied, 'cannot copy range'
        end
      else
        #self.Select
        self.Copy
        sheet.Paste(:destination => destination_range) 
      end
    end
=end

    def self.worksheet_class # :nodoc:
      @worksheet_class ||= begin
        module_name = parent_name
        "#{module_name}::Worksheet".constantize
      rescue NameError => e
        Worksheet
      end
    end

    def worksheet_class # :nodoc:
      self.class.worksheet_class
    end

    private

    def method_missing(name, *args) # :nodoc:
      #if name.to_s[0,1] =~ /[A-Z]/
        begin
          @ole_range.send(name, *args)
        rescue WIN32OLERuntimeError => msg
          if msg.message =~ /unknown property or method/
            raise VBAMethodMissingError, "unknown VBA property or method #{name.inspect}"
          else
            raise msg
          end
        end
   #   else
   #     super
   #   end
    end
  end
end
