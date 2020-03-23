# -*- coding: utf-8 -*-

module RobustExcelOle

  class Bookstore < REOCommon
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
      if weakref_books.empty? || weakref_books.nil?
        weakref_books = workbooks_considering_networkpaths(filename_key) 
      end
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
      book.stored_filename = book.filename
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

    # @private
    def workbooks_considering_networkpaths(filename)  
      network = WIN32OLE.new('WScript.Network')
      drives = network.enumnetworkdrives
      drive_letter, filename_after_drive_letter = filename.split(':')   
      found_filename = nil
      # if filename starts with a drive letter not c and this drive exists,
      # then if there is the corresponding host_share_path in the bookstore, 
      # then take the corresponding workbooks      
      # otherwise (there is an usual file path) find in the bookstore the workbooks of which filenames 
      # ends with the latter part of the given filename (after the drive letter)
      current_path = File.absolute_path(".")
      default_hard_drive = current_path[0,current_path.index(':')].downcase
      if drive_letter != default_hard_drive && drive_letter != filename  
        for i in 0 .. drives.Count-1
          next if i % 2 == 1
          if drives.Item(i).gsub(':','').downcase == drive_letter
            hostname_share = drives.Item(i+1).gsub('\\','/').gsub('//','').downcase
            break
          end
        end        
        @filename2books.each do |stored_filename,_|
          if hostname_share && stored_filename
            if stored_filename[0] == '/'
              index_hostname = stored_filename[1,stored_filename.length].index('/')+2
              index_hostname_share = stored_filename[index_hostname,stored_filename.length].index('/')
              hostname_share_in_stored_filename = stored_filename[1,index_hostname+index_hostname_share-1] 
              if hostname_share_in_stored_filename == hostname_share
                found_filename = stored_filename
                break
              end    
            elsif found_filename.nil? && stored_filename.end_with?(filename_after_drive_letter)
              found_filename = stored_filename
            end         
          end
        end
      elsif filename[0] == '/'
        # if filename starts with a host name and share, and this is an existing host name share path,
        # (then if there are workbooks with the corresponding drive letter),
        # and there are workbooks having the same ending, 
        # then take these workbooks,
        # otherwise (there is an usual file path) find in the bookstore the workbooks of which filenames
        # ends with the latter part of the given filename (after the drive letter)
        index_hostname = filename[1,filename.length].index('/')+2
        index_hostname_share = filename[index_hostname,filename.length].index('/')
        hostname_share_in_filename = filename[1,index_hostname+index_hostname_share-1] 
        filename_after_hostname_share = filename[index_hostname+index_hostname_share+1, filename.length]
        require 'socket'
        hostname = Socket.gethostname
        if hostname_share_in_filename[0,hostname_share_in_filename.index('/')] == hostname.downcase
          for i in 0 .. drives.Count-1
            next if i % 2 == 1
            hostname_share = drives.Item(i+1).gsub('\\','/').gsub('//','').downcase
            if hostname_share == hostname_share_in_filename
              drive_letter = drives.Item(i).gsub(':','').downcase
              break
            end
          end
          @filename2books.each do |stored_filename,_|
            if stored_filename
              if drive_letter && stored_filename.start_with?(drive_letter.downcase) && stored_filename.end_with?(filename_after_hostname_share)
                found_filename = stored_filename
                break
              elsif found_filename.nil? && stored_filename.end_with?(filename_after_hostname_share)
                found_filename = stored_filename
              end
            end
          end
        end
      else
        # if filename is an usual file path,
        # then find in the bookstore a workbook of which filename starts with
        # a drive letter or a host name
        @filename2books.each do |stored_filename,_|
          if stored_filename
            drive_letter, _ = stored_filename.split(':')               
            first_str = stored_filename[0,stored_filename.rindex('/')]
            stored_filename_end = stored_filename[first_str.rindex('/')+1,stored_filename.length]
            if stored_filename_end && filename.end_with?(stored_filename_end)
              found_filename = stored_filename
              break unless @filename2books[found_filename].empty?
            end
          end
        end
      end
      @filename2books[found_filename]
    end

    # @private
    def try_hidden_excel 
      @hidden_excel_instance.__getobj__ if @hidden_excel_instance && @hidden_excel_instance.weakref_alive? && @hidden_excel_instance.__getobj__.alive?
    end

  public

    # prints the book store
    # @private
    def print_filename2books
      #trace "@filename2books:"
      if @filename2books
        @filename2books.each do |filename,books|
          #trace " filename: #{filename}"
          #trace " books:"
          if books.empty?
            #trace " []"
          else
            books.each do |book|
              if book.weakref_alive?
                #trace "#{book}"
              else # this should never happen
                #trace "weakref not alive"
              end
            end
          end
        end
      else
        #trace "nil"
      end
    end
  end

  # @private
  class BookstoreError < WIN32OLERuntimeError  
  end

end
