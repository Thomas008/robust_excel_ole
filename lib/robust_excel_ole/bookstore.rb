# -*- coding: utf-8 -*-

module RobustExcelOle

  # @private
  class Bookstore < Base
    
    def initialize
      @filename2books ||= Hash.new { |hash, key| hash[key] = [] }
      @hidden_excel_instance = nil
    end

    # returns an excel
    def excel
      fn2books = @filename2books.first
      return nil if fn2books.nil? || fn2books.empty? || fn2books[1].empty?  
      book = fn2books[1].first.__getobj__
      book.excel
    end

    # returns a workbook with the given filename, if it was open once
    # @param [String] filename  the file name
    # @param [Hash]   options   the options
    # @option option [Boolean] :prefer_writable
    # @option option [Boolean] :prefer_excel
    # prefers open workbooks to closed workbooks, and among them, prefers more recently opened workbooks
    # excludes hidden Excel instance
    # options: :prefer_writable   returns the writable workbook, if it is open (default: true)
    #                             otherwise returns the workbook according to the preference order mentioned above
    #          :prefer_excel      returns the workbook in the given Excel instance, if it exists,
    #                             otherwise proceeds according to prefer_writable
    def fetch(filename, options = { :prefer_writable => true })
      return nil unless filename
      filename = General.absolute_path(filename)
      filename_key = General.canonize(filename)
      weakref_books = @filename2books[filename_key]
      return nil if weakref_books.nil? || weakref_books.empty?

      result = open_book = closed_book = nil
      weakref_books = weakref_books.map { |wr_book| wr_book if wr_book.weakref_alive? }.compact
      @filename2books[filename_key] = weakref_books
      weakref_books.each do |wr_book|
        if !wr_book.weakref_alive?
          # trace "warn: this should never happen"
          begin
            @filename2books[filename_key].delete(wr_book)
          rescue
            trace "Warning: deleting dead reference failed: file: #{filename.inspect}"
          end
        else
          book = wr_book.__getobj__
          next if book.excel == try_hidden_excel

          if options[:prefer_excel] && book.excel == options[:prefer_excel]
            result = book
            break
          end
          if book.alive?
            open_book = book
            break if book.writable && options[:prefer_writable]
          else
            closed_book = book
          end
        end
      end
      result ||= (open_book || closed_book)
      result
    end

    # stores a workbook
    # @param [Workbook] book a given book
    def store(book)
      filename_key = General.canonize(book.filename)
      if book.stored_filename
        old_filename_key = General.canonize(book.stored_filename)
        # deletes the weak reference to the book
        @filename2books[old_filename_key].delete(book)
      end
      @filename2books[filename_key] |= [WeakRef.new(book)]      
    end

    # creates and returns a separate Excel instance with Visible and DisplayAlerts equal false
    # @private
    def hidden_excel 
      unless @hidden_excel_instance && @hidden_excel_instance.weakref_alive? && @hidden_excel_instance.__getobj__.alive?
        @hidden_excel_instance = WeakRef.new(Excel.create)
      end
      @hidden_excel_instance.__getobj__
    end

    # returns all stored books
    def books
      result = []
      if @filename2books
        @filename2books.each do |_filename,books|
          next if books.empty?

          books.each do |wr_book|
            result << wr_book.__getobj__ if wr_book.weakref_alive?
          end
        end
      end
      result
    end

  private

    def try_hidden_excel 
      @hidden_excel_instance.__getobj__ if @hidden_excel_instance && @hidden_excel_instance.weakref_alive? && @hidden_excel_instance.__getobj__.alive?
    end

  public

    # prints the book store
    # @private
    def print_filename2books
      #puts "@filename2books:"
      if @filename2books
        @filename2books.each do |filename,books|
          #puts " filename: #{filename}"
          #puts " books:"
          if books.empty?
            #puts " []"
          else
            books.each do |book|
              if book.weakref_alive?
                #puts "book.filename: #{book.filename}"
              else # this should never happen
                #puts "weakref not alive"
              end
            end
          end
        end
      else
        #puts "nil"
      end
    end
  end

  # @private
  class BookstoreError < WIN32OLERuntimeError  
  end

end
