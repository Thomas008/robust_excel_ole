# -*- coding: utf-8 -*-

module RobustExcelOle

  class ExcelObjects < REOCommon

    def to_reo
      self
    end

    # @private
    def address_tool
      excel.address_tool
    end

  end

  # @private
  class RangeNotEvaluatable < MiscREOError         
  end

  # @private
  class OptionInvalid < MiscREOError               
  end

  # @private
  class ObjectNotAlive < MiscREOError              
  end

end
