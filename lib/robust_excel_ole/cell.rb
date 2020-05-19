# -*- coding: utf-8 -*-

require File.join(File.dirname(__FILE__), './range')

module RobustExcelOle

  class Cell < Range
    attr_reader :ole_cell    

    def initialize(win32_cell)
      @ole_cell = win32_cell.MergeCells ? win32_cell.MergeArea.Item(1,1) : win32_cell
      #ole_cell(win32_cell)
      super
    end

    def v
      self.Value
    end

    def v=(value)
      self.Value = value
    end

    #def ole_cell(win32_cell)
    #  win32_cell.MergeArea.Item(1,1) if win32_cell.MergeCells
    #end

  private

    # @private
    def method_missing(name, *args) 
      if name.to_s[0,1] =~ /[A-Z]/
        if ::ERRORMESSAGE_JRUBY_BUG
          begin
            @ole_cell.send(name, *args)
          rescue Java::OrgRacobCom::ComFailException 
            raise VBAMethodMissingError, "unknown VBA property or method #{name.inspect}"
          end
        else
          begin
            @ole_cell.send(name, *args)
          rescue NoMethodError 
            raise VBAMethodMissingError, "unknown VBA property or method #{name.inspect}"
          end
        end
      else
        super
      end
    end
  end
end
