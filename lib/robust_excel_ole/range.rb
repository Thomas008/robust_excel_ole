# -*- coding: utf-8 -*-

module RobustExcelOle

  # This class essentially wraps a Win32Ole Range object. 
  # You can apply all VBA methods (starting with a capital letter) 
  # that you would apply for a Range object. 
  # See https://docs.microsoft.com/en-us/office/vba/api/excel.range#methods

  class Range < VbaObjects

    include Enumerable
    
    attr_reader :ole_range
    attr_reader :worksheet

    alias ole_object ole_range

    using ToReoRefinement

    def initialize(win32_range, worksheet = nil)
      @ole_range = begin
        win32_range.send(:Areas)
        win32_range
      rescue
        raise TypeREOError, "given win32ole object is not a range"
      end
      @worksheet = (worksheet ? worksheet : self.Parent).to_reo
    end

    def rows
      @rows ||= (1..@ole_range.Rows.Count)
    end

    def columns
      @columns ||= (1..@ole_range.Columns.Count)
    end

    def each
      if block_given?
        @ole_range.lazy.each_with_index do |ole_cell, index|
          yield cell(index){ole_cell}
        end
      else
        to_enum(:each).lazy
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

    # returns value of a given range restricted to used range
    # @returns [Array] values of the range (as a nested array)    
    def value
      value = begin
        if !::RANGES_JRUBY_BUG    
          intersection_range = ole_range.Application.Intersect(ole_range, worksheet.Range(
            worksheet.Cells(1,1),worksheet.Cells(worksheet.last_row,worksheet.last_column)))
          begin
            intersection_range.Value 
          rescue 
            nil
          end
        else
          # optimization is possible here
          rows_used_range = [rows, last_row].min
          columns_used_rage = [columns, last_column].min
          values = rows_used_range.map{|r| columns_used_range.map {|c| worksheet.Cells(r,c).Value} }
          (values.size==1 && values.first.size==1) ? values.first.first : values
        end
      rescue
        raise RangeNotEvaluatable, "cannot evaluate range #{self.inspect}\n#{$!.message}"
      end
      if value == -2146828288 + RobustExcelOle::XlErrName
        raise RangeNotEvaluatable, "cannot evaluate range #{self.inspect}"
      end
      value
    end

    # sets the values if the range
    # @param [Variant] value
    def value=(value)
      if !::RANGES_JRUBY_BUG
        ole_range.Value = value
      else
        rows.each_with_index do |r,i|
          columns.each_with_index do |c,j|
            ole_range.Cells(i+1,j+1).Value = (value.respond_to?(:pop) ? value[i][j] : value)
          end
        end
      end
      value
    rescue WIN32OLERuntimeError, Java::OrgRacobCom::ComFailException => msg  
      raise RangeNotEvaluatable, "cannot assign value to range #{self.inspect}\n#{$!.message}"
    end

    alias v value
    alias v= value=

    # sets the values if the range with a given color
    # @param [Variant] value
    # @option opts [Symbol] :color the color of the cell when set
    def set_value(value, opts = { })
      if !::RANGES_JRUBY_BUG
        ole_range.Value = value
      else
        rows.each_with_index do |r,i|
          columns.each_with_index do |c,j|
            ole_range.Cells(i+1,j+1).Value = (value.respond_to?(:pop) ? value[i][j] : value)
          end
        end
      end
      ole_range.Interior.ColorIndex = opts[:color] unless opts[:color].nil?
      value
    rescue WIN32OLERuntimeError, Java::OrgRacobCom::ComFailException => msg  
      raise RangeNotEvaluatable, "cannot assign value to range #{self.inspect}\n#{$!.message}"
    end

    # copies a range
    # @params [Address or Address-Array] address or upper left position of the destination range
    # @options [Worksheet] the destination worksheet
    # @options [Hash] options: :transpose, :values_only
    def copy(dest_address, *remaining_args)
      dest_sheet = @worksheet
      options = { }
      remaining_args.each do |arg|
        case arg
        when ::Range, Integer then dest_address = [dest_address,arg]
        when Worksheet, WIN32OLE    then dest_sheet = arg.to_reo
        when Hash                   then options = arg
        else raise RangeNotCopied, "cannot copy range: argument error: #{remaining_args.inspect}"
        end
      end
      dest_range_address = destination_range(dest_address, dest_sheet, options)
      dest_range = dest_sheet.range(dest_range_address)
      if options[:values_only]
        dest_range.v = !options[:transpose] ? self.v : self.v.transpose
      else
        copy_ranges(dest_address, dest_range, dest_range_address, dest_sheet, options)
      end
      dest_range
    rescue WIN32OLERuntimeError, Java::OrgRacobCom::ComFailException => msg
      raise RangeNotCopied, "cannot copy range\n#{$!.message}"
    end

  private

    def destination_range(dest_address, dest_sheet, options)
      rows, columns = address_tool.as_integer_ranges(dest_address)
      dest_address_is_position = (rows.min == rows.max && columns.min == columns.max)
      if !dest_address_is_position
        [rows.min..rows.max,columns.min..columns.max]
      else
        ole_rows, ole_columns = self.Rows, self.Columns
        [rows.min..rows.min + (options[:transpose] ? ole_columns : ole_rows).Count - 1, 
         columns.min..columns.min + (options[:transpose] ? ole_rows : ole_columns).Count - 1]
      end
    end

    def copy_ranges(dest_address, dest_range, dest_range_address, dest_sheet, options)
      workbook = @worksheet.workbook
      if dest_range.worksheet.workbook.excel == workbook.excel 
        if options[:transpose]
          self.Copy              
          dest_range.PasteSpecial(XlPasteAll,XlPasteSpecialOperationNone,false,true)
        else
          self.Copy(dest_range.ole_range)
        end            
      else
        if options[:transpose]
          added_sheet = workbook.add_sheet
          copy(dest_address, added_sheet, transpose: true)
          added_sheet.range(dest_range_address).copy(dest_address,dest_sheet)
          workbook.excel.with_displayalerts(false) {added_sheet.Delete}
        else
          self.Copy
          dest_sheet.Paste(dest_range.ole_range)
        end
      end
    end
     
  public

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
    def workbook
      @workbook ||= @worksheet.workbook
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
      "#<REO::Range: #{@ole_range.Address(External: true).gsub(/\$/,'')} >"
    end

    # @private
    def inspect
      to_s
    end

    using ParentRefinement
    using StringRefinement

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

  # @private
  class RangeNotCopied < MiscREOError              
  end

end
