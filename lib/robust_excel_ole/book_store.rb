
# -*- coding: utf-8 -*-

module RobustExcelOle

  class BookStore

    @@filename2book = {}

    def initialize
      @@filename2book = {}
    end

    # returns a book with the given filename, if it was open once
    # preference order: writable book, readonly unsaved book, readonly book (the last one), closed book
    # options: 
    #   :readonly  declares whether a writable book is prefered, or a book for read_only is sufficient
    #             false (default)  -> prefer a writable book,  true -> read_only book is sufficient 
    #     
    
    # ??? what if the changes happen AFTER storing the book?
    # when fetching: ask not the properties of the stored book but of the properties that the
    # storaged book has now

    def self.fetch(filename, options = { })
      #p "fetch:"
      #print
      filename_key = RobustExcelOle::canonize(filename)
      #p "filename_key: #{filename_key}"
      readonly_book = readonly_unsaved_book = closed_book = result = nil
      books = @@filename2book[filename_key]
      #p "books: #{books}"
      return nil  unless books
      books.each do |book|
        #p "book: #{book}"
        if book.alive?
          #p "book alive"
          if (not book.ReadOnly)
            #p "book writable"
            return book
          else
            #p "book read_only"
            book.Saved ? readonly_book = book : readonly_unsaved_book = book
          end
        else
          #p "book closed"
          closed_book = book
        end
      end
      result = readonly_unsaved_book ? readonly_unsaved_book : (readonly_book ? readonly_book : closed_book)
      #p "book: #{result}"
      result
    end

    # stores a book
    def self.store(book)
      #p "store:"
      #p "filename: #{book.filename}"
      filename_key = RobustExcelOle::canonize(book.filename)      
      #p "filename_key: #{filename_key}"
      if @@filename2book[filename_key]
        @@filename2book[filename_key] << book unless @@filename2book[filename_key].include?(book)
      else
        @@filename2book[filename_key] = [book]
      end
      #print
    end

    # prints the book store
    def self.print
      p "bookstore:"
      @@filename2book.each do |filename,books|
        p " filename: #{filename}"
        p " books:"
        books.each do |book|
          p "#{book}"
        end
      end
    end


    private :print


  end
 
end
