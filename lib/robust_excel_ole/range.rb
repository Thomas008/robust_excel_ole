# -*- coding: utf-8 -*-

module RobustExcelOle

  # This class essentially wraps a Win32Ole Range object. 
  # You can apply all VBA methods (starting with a capital letter) 
  # that you would apply for a Range object. 
  # See https://docs.microsoft.com/en-us/office/vba/api/excel.worksheet#methods

  class Range < VbaObjects

    include Enumerable
    
    attr_reader :ole_range
    attr_reader :worksheet

    alias ole_object ole_range


    def initialize(win32_range, worksheet = nil)
      @ole_range = win32_range
      @worksheet = worksheet ? worksheet.to_reo : worksheet_class.new(self.Parent)
    end

    def rows
      @rows ||= (1..@ole_range.Rows.Count)
    end

    def columns
      @columns ||= (1..@ole_range.Columns.Count)
    end

    def each
      @ole_range.each_with_index do |ole_cell, index|
        yield cell(index){ole_cell}
      end
    end

    def [] index
      cell(index) {
        @ole_range.Cells.Item(index + 1)
      }
    end

  private

    def cell(index)
      @cells ||= []
      @cells[index + 1] ||= RobustExcelOle::Cell.new(yield,@worksheet)
    end

  public

    # returns flat array of the values of a given range
    # @params [Range] a range
    # @returns [Array] the values
    def values(range = nil)
      result_unflatten = if !::RANGES_JRUBY_BUG
        map { |x| x.v }
      else
        self.v
      end
      result = result_unflatten.flatten
      if range
        relevant_result = []
        result.each_with_index { |row_or_column, i| relevant_result << row_or_column if range.include?(i) }
        relevant_result
      else
        result
      end
    end

    def v
      begin
        if !::RANGES_JRUBY_BUG
          self.Value
        else
          values = []
          rows.each do |r|
            values_col = []
            columns.each{ |c| values_col << worksheet.Cells(r,c).Value}
            values << values_col
          end
          values
        end
      rescue WIN32OLERuntimeError, Java::OrgRacobCom::ComFailException => msg
        raise RangeNotEvaluatable, 'cannot read value'
      end 

    end

    def v=(value)
      begin
        if !::RANGES_JRUBY_BUG
          ole_range.Value = value
        else
          rows.each_with_index do |r,i|
            columns.each_with_index do |c,j|
              ole_range.Cells(i+1,j+1).Value = (value.respond_to?(:first) ? value[i][j] : value)
            end
          end
        end
        value
      rescue WIN32OLERuntimeError, Java::OrgRacobCom::ComFailException => msg  
        raise RangeNotEvaluatable, "cannot assign value to range #{address_r1c1.inspect}"
      end
    end

    alias_method :value, :v
    alias_method :value=, :v=

    # copies a range
    # @params [Address or Address-Array] address or upper left position of the destination range
    # @options [Worksheet] the destination worksheet
    # @options [Hash] options: :transpose, :values_only
    def copy(dest_address, *remaining_args)
      dest_sheet = @worksheet
      options = { }
      remaining_args.each do |arg|
        case arg
        when Object::Range, Integer then dest_address = [dest_address,arg]
        when Worksheet, WIN32OLE    then dest_sheet = arg.to_reo
        when Hash                   then options = arg
        else raise RangeNotCopied, "cannot copy range: argument error: #{remaining_args.inspect}"
        end
      end
      begin
        rows, columns = address_tool.as_integer_ranges(dest_address)
        dest_address_is_position = (rows.min == rows.max && columns.min == columns.max)
        dest_range_address = if (not dest_address_is_position) 
          [rows.min..rows.max,columns.min..columns.max]
        else
          if (not options[:transpose])
            [rows.min..rows.min+self.Rows.Count-1, columns.min..columns.min+self.Columns.Count-1]
          else
            [rows.min..rows.min+self.Columns.Count-1, columns.min..columns.min+self.Rows.Count-1]
          end
        end
        dest_range = dest_sheet.range(dest_range_address)
        if options[:values_only]
          dest_range.v = options[:transpose] ? self.v.transpose : self.v
        else
          if dest_range.worksheet.workbook.excel == @worksheet.workbook.excel 
            if options[:transpose]
              self.Copy              
              dest_range.PasteSpecial(XlPasteAll,XlPasteSpecialOperationNone,false,true)
            else
              self.Copy(dest_range.ole_range)
            end            
          else
            if options[:transpose]
              added_sheet = @worksheet.workbook.add_sheet
              self.copy(dest_address, added_sheet, :transpose => true)
              added_sheet.range(dest_range_address).copy(dest_address,dest_sheet)
              @worksheet.workbook.excel.with_displayalerts(false) {added_sheet.Delete}
            else
              self.Copy
              dest_sheet.Paste(dest_range.ole_range)
            end
          end
        end
        dest_range
      rescue WIN32OLERuntimeError, Java::OrgRacobCom::ComFailException => msg
        raise RangeNotCopied, 'cannot copy range'
      end
    end

    def == other_range
      other_range.is_a?(Range) &&
        self.worksheet == other_range.worksheet
        self.Address == other_range.Address 
    end

    # @private
    def excel
      @worksheet.workbook.excel
    end

    # @private
    # returns true, if the Range object responds to VBA methods, false otherwise
    def alive?
      @ole_range.Row
      true
    rescue
      # trace $!.message
      false
    end

    # @private
    def to_s
      "#<REO::Range: " + "#{@ole_range.Address('External' => true).gsub(/\$/,'')} " + ">"
      # "#<REO::Range: " + "#{@ole_range.Address.gsub(/\$/,'')} " + "#{worksheet.Name} " + ">"
    end

    # @private
    def inspect
      to_s # [0..-2] + "#{worksheet.workbook.Name} " + ">"
    end

    # @private
    def self.worksheet_class   
      @worksheet_class ||= begin
        module_name = parent_name
        "#{module_name}::Worksheet".constantize
      rescue NameError => e
        Worksheet
      end
    end

    # @private
    def worksheet_class 
      self.class.worksheet_class
    end

    include MethodHelpers

  private

    def method_missing(name, *args) 
      if name.to_s[0,1] =~ /[A-Z]/
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
      else
        super
      end
    end
  end

  # @private
  class RangeNotCopied < MiscREOError              
  end

end
