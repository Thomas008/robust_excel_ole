# -*- coding: utf-8 -*-

LOG_TO_STDOUT = true      unless Object.const_defined?(:LOG_TO_STDOUT)
REO_LOG_DIR   = ''.freeze        unless Object.const_defined?(:REO_LOG_DIR)
REO_LOG_FILE  = 'reo.log'.freeze unless Object.const_defined?(:REO_LOG_FILE)

File.delete REO_LOG_FILE rescue nil

unless "any string".respond_to?(:end_with?)
  class String 
    def end_with?(*suffixes)
      suffixes.any? do |suffix|
        self[-suffix.size .. -1] == suffix
      end
    end
  end
end

module RobustExcelOle

  # @private
  class REOError < RuntimeError                    
  end

  # @private
  class ExcelREOError < REOError                   
  end

  # @private
  class WorkbookREOError < REOError                
  end

  # @private
  class WorksheetREOError < REOError                   
  end

  # @private
  class FileREOError < REOError                    
  end

  # @private
  class NamesREOError < REOError                   
  end

  # @private
  class MiscREOError < REOError                    
  end

  # @private
  class ExcelDamaged < ExcelREOError               
  end

  # @private
  class UnsavedWorkbooks < ExcelREOError           
  end

  # @private
  class WorkbookBlocked < WorkbookREOError         
  end

  # @private
  class WorkbookNotSaved < WorkbookREOError        
  end

  # @private
  class WorkbookReadOnly < WorkbookREOError        
  end

  # @private
  class WorkbookBeingUsed < WorkbookREOError       
  end

  # @private
  class FileNotFound < FileREOError                
  end

  # @private
  class FileNameNotGiven < FileREOError            
  end

  # @private
  class FileAlreadyExists < FileREOError           
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
  class RangeNotCreated < MiscREOError             
  end

  # @private
  class RangeNotCopied < MiscREOError              
  end

  # @private
  class OptionInvalid < MiscREOError               
  end

  # @private
  class ObjectNotAlive < MiscREOError              
  end

  # @private
  class TypeREOError < REOError                    
  end

  # @private
  class TimeOut < REOError                         
  end

  # @private
  class AddressInvalid < REOError                  
  end

  # @private
  class UnexpectedREOError < REOError              
  end

  # @private
  class NotImplementedREOError < REOError          
  end

  class REOCommon

    # @private
    def excel
      raise TypeREOError, 'receiver instance is neither an Excel nor a Workbook'
    end

    # @private
    def own_methods
      (self.methods - Object.methods).sort
    end

    # @private
    def self.tr1(_text)
      puts :text
    end

    # @private
    def self.trace(text)
      if LOG_TO_STDOUT
        puts text
      else
        if REO_LOG_DIR.empty?
          homes = ['HOME', 'HOMEPATH']
          home = homes.find { |h| !ENV[h].nil? }
          reo_log_dir = ENV[home]
        else
          reo_log_dir = REO_LOG_DIR
        end
        File.open(reo_log_dir + '/' + REO_LOG_FILE,'a') do |file|
          file.puts text
        end
      end
    end

    # @private
    def self.puts_hash(hash)
      hash.each do |e|
        if e[1].is_a?(Hash)
          puts "#{e[0]} =>"
          e[1].each do |f|
            puts "  #{f[0]} => #{f[1]}"
          end
        else
          puts "#{e[0]} => #{e[1]}"
        end
      end
    end

  end

end
