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

    # @private
    def method_missing(name, *args) 
      #if name.to_s[0,1] =~ /[A-Z]/
        begin
          @cell.send(name, *args)
        rescue WIN32OLERuntimeError => msg
          if msg.message =~ /unknown property or method/
            raise VBAMethodMissingError, "unknown VBA property or method #{name.inspect}"
          else
            raise msg
          end
        end
     # else
     #   super
     # end
    end
  end
end
