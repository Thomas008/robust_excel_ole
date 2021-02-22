# -*- coding: utf-8 -*-

require_relative 'range'

module RobustExcelOle

  class Cell < Range
    #attr_reader :ole_cell    

    def initialize(win32_cell, worksheet)      
      super
      ole_cell
    end

    def value
      self.Value
    end

    def value=(value)
      self.Value = value
    end

    alias_method :v, :value
    alias_method :v=, :value=

    # @private
    def ole_cell
      @ole_range = @ole_range.MergeArea.Item(1,1) if @ole_range.MergeCells
    end

    # @private
    def to_s    
      "#<Cell: (#{@ole_range.Row},#{@ole_range.Column})>"
    end

    # @private
    def inspect 
      self.to_s[0..-2] + " #{@ole_range.Parent.Name}" + ">" 
    end

  private

    # @private
    def method_missing(name, *args) 
      super unless name.to_s[0,1] =~ /[A-Z]/
      if ::ERRORMESSAGE_JRUBY_BUG
        begin
          @ole_range.send(name, *args)
        rescue Java::OrgRacobCom::ComFailException 
          raise VBAMethodMissingError, "unknown VBA property or method #{name.inspect}"
        end
      else
        begin
          @ole_range.send(name, *args)
        rescue NoMethodError 
          raise VBAMethodMissingError, "unknown VBA property or method #{name.inspect}"
        end
      end
    end
  end
end
