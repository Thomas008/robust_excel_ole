# -*- coding: utf-8 -*-
module RobustExcelOle
  
  class Range < REOCommon
    include Enumerable
    attr_reader :ole_range
    attr_reader :worksheet

    def initialize(win32_range)
      @ole_range = win32_range
      @worksheet = sheet_class.new(self.Parent)
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
      result = self.map{|x| x.value}.flatten
      if range 
        relevant_result = []
        result.each_with_index{ |row_or_column, i| relevant_result << row_or_column if range.include?(i) }
        relevant_result 
      else
        result
      end
    end

    def [] index
      @cells = []
      @cells[index + 1] = RobustExcelOle::Cell.new(@ole_range.Cells.Item(index + 1))
    end

    # copies a range
    # @params [Address] address of a range
    # @options [Sheet] the worksheet in which to copy      
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
          raise RangeNotCopied, "cannot copy range"
        end
      else
        self.Select
        self.Copy
        sheet.Paste(destination_range)
      end
    end

    def self.sheet_class    # :nodoc: #
      @sheet_class ||= begin
        module_name = self.parent_name
        "#{module_name}::Sheet".constantize
      rescue NameError => e
        Sheet
      end
    end

    def sheet_class        # :nodoc: #
      self.class.sheet_class
    end

  private
    def method_missing(name, *args)    # :nodoc: #
      if name.to_s[0,1] =~ /[A-Z]/ 
        begin
          @ole_range.send(name, *args)
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

  end
end
