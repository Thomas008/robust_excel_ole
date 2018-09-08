# -*- coding: utf-8 -*-
module RobustExcelOle
  class Range < REOCommon
    include Enumerable
    attr_reader :ole_range

    def initialize(win32_range)
      @ole_range = win32_range
    end

    def each
      @ole_range.each do |row_or_column|
        yield RobustExcelOle::Cell.new(row_or_column)
      end
    end

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

    def method_missing(id, *args)  # :nodoc: #
      @ole_range.send(id, *args)
    end
  end
end
