# -*- coding: utf-8 -*-

require File.join(File.dirname(__FILE__), './../spec_helper')


$VERBOSE = nil

include RobustExcelOle
include General

unless Object.method_defined?(:require_relative)
  def require_relative path
    require File.expand_path(path, File.dirname(__FILE__))  
  end
end

require_relative "book_open_spec"
require_relative "book_close_spec"
require_relative "book_save_spec"
require_relative "book_misc_spec"
require_relative "book_sheet_spec"
require_relative "book_unobtr_spec"
require_relative "book_subclass_spec"
