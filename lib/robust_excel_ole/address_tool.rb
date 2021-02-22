# -*- coding: utf-8 -*-

module RobustExcelOle

  class AddressTool < Base

    def initialize(address_string)
      r1c1_letters = address_string.gsub(/[0-9]/,'')
      @row_letter = r1c1_letters[0..0]
      @col_letter = r1c1_letters[1..1]
    end

    # address formats that are valid:
    #   r1c1-format: e.g. "Z3S1", "Z3S1:Z5S2", "Z[3]S1", "Z3S[-1]:Z[5]S1", "Z[3]", "S[-2]"
    #                      infinite ranges are not possible, e.g. "Z3:Z5", "S2:S5", "Z2", "S3", "Z[2]" 
    #   integer_ranges-fromat: e.g. [3,1], [3,"A"], [3..5,1..2], [3..5, "A".."B"], 
    #                               [3..4, nil], [nil, 2..4], [2,nil], [nil,4]
    #   a1-format: e.g. "A3", "A3:B5", "A:B", "3:5", "A", "3"

    def as_r1c1(address)
      transform_address(address,:r1c1)
    end

    def as_a1(address)
      transform_address(address,:a1)
    end

    # valid address formats: e.g. [3,1], [3,"A"], [3..5,1..2], [3..5, "A".."B"], 
    #                             [3..4, nil], [nil, 2..4], [2,nil], [nil,4]
    def as_integer_ranges(address)
      transform_address(address,:int_range)
    end

  private

    def transform_address(address, format)
      address = address.is_a?(Array) ? address : [address]
      raise AddressInvalid, "address #{address.inspect} has more than two components" if address.size > 2
      begin
        if address.size == 1
          comp1, comp2 = address[0].split(':')
          a1_expr = /^(([A-Z]+[0-9]+)|([A-Z]+$)|([0-9]+))$/
          is_a1 = comp1 =~ a1_expr && (comp2.nil? || comp2 =~ a1_expr)
          r1c1_expr = /^(([A-Z]\[?-?[0-9]+\]?[A-Z]\[?-?[0-9]+\]?)|([A-Z]\[?-?[0-9]+\]?)|([A-Z]\[?-?[0-9]+\]?))$/
          is_r1c1 = comp1 =~ r1c1_expr && (comp2.nil? || comp2 =~ r1c1_expr) && (not is_a1) 
          raise AddressInvalid, "address #{address.inspect} not in A1- or r1c1-format" unless (is_a1 || is_r1c1)
          return address[0].gsub('[','(').gsub(']',')') if (is_a1 && format==:a1) || (is_r1c1 && format==:r1c1)         
          given_format = (is_a1) ? :a1 : :r1c1
          row_comp1, col_comp1 = analyze(comp1,given_format)
          row_comp2, col_comp2 = analyze(comp2,given_format) unless comp2.nil?
          address_comp1 = comp2 && (not row_comp1.nil?) ? (row_comp1 .. row_comp2) : row_comp1
          address_comp2 = comp2 && (not col_comp1.nil?) ? (col_comp1 .. col_comp2) : col_comp1          
        else
          address_comp1, address_comp2 = address      
        end
        address_comp1 = address_comp1..address_comp1 if (address_comp1.is_a?(Integer) || address_comp1.is_a?(String) || address_comp1.is_a?(Array))
        address_comp2 = address_comp2..address_comp2 if (address_comp2.is_a?(Integer) || address_comp2.is_a?(String) || address_comp2.is_a?(Array)) 
        rows = unless address_comp1.nil? || address_comp1.begin == 0         
          address_comp1.begin..address_comp1.end
        end
        columns = unless address_comp2.nil?
          if address_comp2.begin.is_a?(String) #address_comp2.begin.to_i == 0
            col_range = str2num(address_comp2.begin)..str2num(address_comp2.end)
            col_range==(0..0) ? nil : col_range
          else
            address_comp2.begin..address_comp2.end
          end          
        end
      rescue
        raise AddressInvalid, "address (#{address.inspect}) format not correct"
      end
      if format==:r1c1
        r1c1_string(@row_letter,rows,:min) + r1c1_string(@col_letter,columns,:min) + ":" + 
        r1c1_string(@row_letter,rows,:max) + r1c1_string(@col_letter,columns,:max)
      elsif format==:int_range
        [rows,columns]
      else
        raise NotImplementedREOError, "not implemented"
      end
    end  

    def r1c1_string(letter,int_range,type)
      return "" if int_range.nil? || int_range.begin.nil?
      parameter = type == :min ? int_range.begin : int_range.end
      is_relative = parameter.is_a?(Array)
      parameter = parameter.first if is_relative
      letter + (is_relative ? "(" : "") + parameter.to_s + (is_relative ? ")" : "")       
    end 

    def analyze(comp,format)
      row_comp, col_comp = if format==:a1 
        [comp.gsub(/[A-Z]/,''), comp.gsub(/[0-9]/,'')]
      else
        a,b = comp.split(@row_letter)
        c,d = b.split(@col_letter)
        b.nil? ? ["",b] : (d.nil? ? [c,""] : [c,d])  
      end
      def s2n(s)
        s!="" ? (s[0] == "[" ? [s.gsub(/\[|\]/,'').to_i] : (s.to_i!=0 ? s.to_i : s)) : nil
      end
      [s2n(row_comp), s2n(col_comp)]
    end

    def str2num(str)
      str.tr("A-Z","0-9A-P").to_i(26) + (26**str.size-1)/25
    end

  end

  # @private
  class AddressInvalid < REOError                  
  end

end
