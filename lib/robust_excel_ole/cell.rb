# -*- coding: utf-8 -*-

module RobustExcelOle
  class Cell < REOCommon
    attr_reader :cell

    def initialize(win32_cell)
      @cell = win32_cell.MergeCells ? win32_cell.MergeArea.Item(1,1) : win32_cell
    end

    def v
      self.Value
    end

    def v=(value)
      self.Value = value
    end




    # @private
    def method_missing(name, *args) 
      #if name.to_s[0,1] =~ /[A-Z]/
      if RUBY_PLATFORM =~ /java/  
        begin
          @cell.send(name, *args)
        rescue Java::OrgRacobCom::ComFailException 
          raise VBAMethodMissingError, "unknown VBA property or method #{name.inspect}"
        end
      else
        begin
          @cell.send(name, *args)
        rescue NoMethodError 
          raise VBAMethodMissingError, "unknown VBA property or method #{name.inspect}"
        end
      end
     # else
     #   super
     # end
    end
  end
end
