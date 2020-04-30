# -*- coding: utf-8 -*-

module RobustExcelOle

  class ExcelObjects < REOCommon

    def to_reo
      self
    end

  end

  # @private
  class RangeNotCreated < MiscREOError             
  end

  # @private
  class RangeNotCopied < MiscREOError              
  end

  # @private
  class NameNotFound < NamesREOError               
  end

  # @private
  class NameAlreadyExists < NamesREOError          
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
