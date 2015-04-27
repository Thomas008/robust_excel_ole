
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
      filename_key = RobustExcelOle::canonize(filename)
      weakref_books = @filename2books[filename_key]
      return nil unless weakref_books
      readonly_book = readonly_unsaved_book = closed_book = result = nil      
      weakref_books.each do |book|
        if (not book.weakref_alive?) 
          @filename2books[filename_key].delete(book)
        else
          return book if options[:readonly_excel] && book.excel == options[:readonly_excel]
          if book.alive?
            if (not book.readonly)
              if options[:readonly_excel]
                result = book
              else
                return book
              end
            else
              book.saved ? readonly_book = book : readonly_unsaved_book = book
            end
          else
            closed_book = book
          end
        end
      end
      result ? result : (readonly_unsaved_book ? readonly_unsaved_book : (readonly_book ? readonly_book : closed_book))
    end

    # stores a book
    def store(book)
      filename_key = RobustExcelOle::canonize(book.filename)      
      if book.stored_filename
        old_filename_key = RobustExcelOle::canonize(book.stored_filename)
        #@filename2books[old_filename_key].delete(book)
        @filename2books[old_filename_key].delete(WeakRef.new(book))
      end
      #@filename2books[filename_key] |= [book] 
      @filename2books[filename_key] |= [WeakRef.new(book)]
      book.stored_filename = book.filename
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
            p "#{book}"
            p "excel: #{book.excel}"
            p "alive: #{book.alive?}"
          end
        end
      end
    end

  end

  class BookStoreError < WIN32OLERuntimeError
  end
 
end
