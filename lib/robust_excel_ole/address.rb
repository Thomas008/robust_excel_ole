# -*- coding: utf-8 -*-

module RobustExcelOle

  class Address < REOCommon

    def self.r1c1(address)
      transform_address(address,:r1c1)
    end

    def self.a1(address)
      transform_address(address,:a1)
    end

    def self.int_range(address)
      transform_address(address,:int_range)
    end

  private
    # @private
    # possible formats: a1    e.g. "A3", "A3:B5", "A:B", "3:5", "A", "3"
    #                   r1c1, e.g. "Z3S1", "Z3S1:Z5S2", 
    #{                  infinite r1c1-formats are not possible: ("Z3:Z5", "S2:S5", "Z2", "S3")
    #                   ranges, e.g. [3,1], [3,"A"], [3..5,1..2], [3..5, "A".."B"], [3..4, nil], [nil, 2..4], [2,nil], [nil,4]
    def self.transform_address(address, format)
      address = address.is_a?(Array) ? address : [address]
      raise AddressInvalid, "address #{address.inspect} has more than two components" if address.size > 2
      begin
        if address.size == 1
          comp1, comp2 = address[0].split(':')
          is_a1 = comp1 =~ /^(([A-Z]+[0-9]+)|([A-Z]+$)|([0-9]+))$/ && 
              (comp2.nil? || comp2 =~ /^(([A-Z]+[0-9]+)|([A-Z]+)|([0-9]+))$/ )
          is_r1c1 = comp1 =~ /^((Z[0-9]+S[0-9]+)|(Z[0-9])|(S[0-9]+))$/ &&
              (comp2.nil? || comp2 =~ /^((Z[0-9]+S[0-9]+)|(Z[0-9])|(S[0-9]+))$/)
          raise AddressInvalid, "address #{address.inspect} not in A1- or r1c1-format" unless (is_a1 || is_r1c1)
          return address[0] if (is_a1 && format==:a1) || (is_r1c1 && format==:r1c1)         
          given_format = (is_a1) ? :a1 : :r1c1
          row_comp1, col_comp1 = analyze(comp1,given_format)
          row_comp2, col_comp2 = analyze(comp2,given_format) unless comp2.nil?
          address_comp1 = comp2 ? (row_comp1 .. row_comp2) : row_comp1
          address_comp2 = comp2 ? (col_comp1 .. col_comp2) : col_comp1          
        else
          address_comp1, address_comp2 = address      
        end
        address_comp1 = address_comp1..address_comp1 if (address_comp1.nil? || address_comp1.is_a?(Integer) || address_comp1.is_a?(String)) #unless address_comp1.is_a?(Range)
        address_comp2 = address_comp2..address_comp2 if (address_comp2.nil? || address_comp2.is_a?(Integer) || address_comp2.is_a?(String)) #unless address_comp2.is_a?(Range)
        raise if address_comp1.begin.to_i==0 && (not address_comp1.begin.nil?) && (not address_comp1.begin.empty?)
        rows = unless address_comp1.begin.to_i==0
          address_comp1.begin.to_i..address_comp1.end.to_i 
        end
        columns = unless address_comp2.begin.nil?
          if address_comp2.begin.to_i == 0
            col_range = str2num(address_comp2.begin)..str2num(address_comp2.end)
            col_range==(0..0) ? nil : col_range
          else
            address_comp2.begin.to_i..address_comp2.end.to_i
          end
        end
      rescue 
        raise AddressInvalid, "address (#{address.inspect}) format not correct"
      end
      if format==:r1c1
        def self.r(a,b,c); b ? "#{a}#{(c==:min ? b.min : b.max)}" : ""; end 
        r("Z",rows,:min) + r("S",columns,:min) + ":" + r("Z",rows,:max) + r("S",columns,:max)
      elsif format==:int_range
        [rows,columns]
      else
        raise NotImplementedREOError, "not implemented"
      end
    end

    # @private
    def self.analyze(comp,format)
      if format==:a1 
        [comp.gsub(/[A-Z]/,''), comp.gsub(/[0-9]/,'')]
      else
        a,b = comp.split('Z')
        c,d = b.split('S')
        b.nil? ? ["",b] : (d.nil? ? [c,""] : [c,d])  
      end
    end

    # @private
    def self.str2num(str)
      #return if str.empty?
      str = str.upcase
      sum = 0
      (1..str.length).each { |i| sum += (str[i - 1].ord - 64) * 26**(str.length - i) }
      sum
    end

  end

end
