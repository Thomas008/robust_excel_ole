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

    # values of a range, as array
    # @params [Range] a range
    # @returns [Array] the values
    def values(range = nil)
#+#      result = self.map(&:value).flatten
      result = self.map{|x| x.value}.flatten
#+#      range ? result.each_with_index.select{ |row_or_column, i| range.include?(i) }.map{ |i| i[0] } : result
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
    # @params [Fixnum,Range] row or range of the rows 
    # @params [Fixnum,Range] column or range of columns 
    # @options [Sheet] the worksheet in which to copy      
    def copy(int_range1, int_range2, sheet = :__not_provided)
      int_range1 = int_range1 .. int_range1 if int_range1.is_a?(Fixnum)
      int_range2 = int_range2 .. int_range2 if int_range2.is_a?(Fixnum)
      sheet = @worksheet if sheet == :__not_provided
      if sheet.workbook.excel == @worksheet.workbook.excel 
        begin
          self.Copy(:destination => sheet.range(int_range1.min..int_range1.max,
                                                int_range2.min..int_range2.max).ole_range)
        rescue WIN32OLERuntimeError
          raise RangeNotCopied, "cannot copy range"
        end
      else
        self.Select
        self.Copy
        sheet.Paste(sheet.range(int_range1.min..int_range1.max,int_range2.min..int_range2.max).ole_range)
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

    def method_missing(id, *args)  # :nodoc: #
      @ole_range.send(id, *args)
    end
  end
end
