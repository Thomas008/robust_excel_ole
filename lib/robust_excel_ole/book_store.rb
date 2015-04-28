
# -*- coding: utf-8 -*-

module RobustExcelOle

  class BookStore

    def initialize
      @filename2books = Hash.new {|hash, key| hash[key] = [] }
    end

    # returns a book with the given filename, if it was open once
    # preference order: writable book, readonly unsaved book, readonly book (the last one), closed book
    # options:  
    #  :readonly_excel => <instance> -> return the book that was open in the given excel instance, 
    #                                 even if it is not writable, if such a book exists
    #                          prefer the writable book as described above, otherwise 
    def fetch(filename, options = { })
      p "fetch:"
      print
      filename_key = RobustExcelOle::canonize(filename)
      p "filename_key: #{filename_key}"
      weakref_books = @filename2books[filename_key]
      p "weakref_books: #{weakref_books}" rescue nil
      return nil unless weakref_books
      result = readonly_book = readonly_unsaved_book = closed_book = nil      
      weakref_books.each do |wr_book|
        p "wr_book: #{wr_book}" rescue nil
        if (not wr_book.weakref_alive?)
          p "weakref not alive" 
          @filename2books[filename_key].delete(wr_book)
        else
          p "weakref alive"
          if options[:readonly_excel] && wr_book.excel == options[:readonly_excel]
            p "wr_book.excel == options :read_only"
            result = wr_book
            p "result: #{result}"
            break 
          end
          if wr_book.alive?
            p "wr_book alive"
            if (not wr_book.readonly)
              p "writable"
              result = wr_book
              p "result: #{result}"
              break unless options[:readonly_excel]
            else
              p "wr_book readonly"
              p "wr_book: saved: #{wr_book.saved}"
              wr_book.saved ? readonly_book = wr_book : readonly_unsaved_book = wr_book
              p "readonly_book: #{readonly_book}"
              p "readonly_unsaved_book: #{readonly_unsaved_book}"
            end
          else
            p "wr_book not alive"
            closed_book = wr_book
          end
        end
      end
      result = result ? result : (readonly_unsaved_book ? readonly_unsaved_book : (readonly_book ? readonly_book : closed_book))
      p "result: #{result}"      
      p "end:"
      print
      result.__getobj__ if result
    end

    # stores a book
    def store(book)
      p "store:"
      print
      filename_key = RobustExcelOle::canonize(book.filename)      
      if book.stored_filename
        old_filename_key = RobustExcelOle::canonize(book.stored_filename)
        # deletes the weak reference to the book
        @filename2books[old_filename_key].delete(book)
      end
      @filename2books[filename_key] |= [WeakRef.new(book)]
      book.stored_filename = book.filename
      p "end:"
      print
    end

    # prints the book store
    def print
      p "@filename2books:"
      if @filename2books
        @filename2books.each do |filename,books|
          p " filename: #{filename}"
          p " books:"
          p " []" if books == []
          books.each do |book|
            if book.weakref_alive?
              p "#{book}"
              p "excel: #{book.excel}"
              p "alive: #{book.alive?}"
            else
              p "weakref not alive"
            end
          end
        end
      end
    end

  end

  class BookStoreError < WIN32OLERuntimeError
  end
 
end
