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
    # @params [Fixnum] row, column of the destination cell or
    #                  row, column of the top left cell and the bottum down cell of a rectangular range
    # @options [Sheet] the worksheet in which to copy      
    def copy(r1, c1, r2 = :__not_provided, c2 = :__not_provided, sheet = :__not_provided)
      if r2.is_a?(RobustExcelOle::Sheet)
        sheet = r2
        r2 = :__not_provided
      end
      if r2 == :__not_provided
        r2 = r1
        c2 = c1
      end
      sheet = @worksheet if sheet == :__not_provided
      begin
        self.Copy(:destination => sheet.range(r1,c1,r2,c2).ole_range)
      rescue WIN32OLERuntimeError
        raise RangeNotCopied, "cannot copy range to (#{r1.inspect},#{c1.inspect}),(#{r2.inspect},#{c2.inspect})"
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
