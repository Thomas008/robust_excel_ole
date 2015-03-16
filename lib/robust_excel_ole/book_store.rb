
# -*- coding: utf-8 -*-

module RobustExcelOle

  class BookStore

    def initialize
      @filename2books = Hash.new {|hash, key| hash[key] = [] }
    end

    # returns a book with the given filename, if it was open once
    # preference order: writable book, readonly unsaved book, readonly book (the last one), closed book
    # options: 
    #   :readonly  declares whether a writable book is prefered, or a book for read_only is sufficient
    #             false (default)  -> prefer a writable book,  true -> read_only book is sufficient 

    def fetch(filename, options = { })
      filename_key = RobustExcelOle::canonize(filename)
      readonly_book = readonly_unsaved_book = closed_book = result = nil
      books = @filename2books[filename_key]
      return nil  unless books
      books.each do |book|
        if book.alive?
          if (not book.ReadOnly)
            return book
          else
            book.Saved ? readonly_book = book : readonly_unsaved_book = book
          end
        else
          closed_book = book
        end
      end
      result = readonly_unsaved_book ? readonly_unsaved_book : (readonly_book ? readonly_book : closed_book)
      result
    end

    # stores a book
    def store(book)
      filename_key = RobustExcelOle::canonize(book.filename)      
      if book.stored_filename
        old_filename_key = RobustExcelOle::canonize(book.stored_filename)
        @filename2books[old_filename_key].delete(book)
      end
      @filename2books[filename_key] |= [book]
      book.stored_filename = book.filename
    end

    # prints the book store
    def print
      p "@bookstore:"
      @filename2books.each do |filename,books|
        p " filename: #{filename}"
        p " books:"
        books.each do |book|
          p "#{book}"
        end
      end
    end

  end

  class BookStoreError < RuntimeError
  end
 
end
