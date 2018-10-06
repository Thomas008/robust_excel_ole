# -*- coding: utf-8 -*-

require File.join(File.dirname(__FILE__), './../spec_helper')

$VERBOSE = nil

require File.expand_path("workbook_open_spec", File.dirname(__FILE__)) 
require File.expand_path("workbook_close_spec", File.dirname(__FILE__))
require File.expand_path("workbook_save_spec", File.dirname(__FILE__))
require File.expand_path("workbook_misc_spec", File.dirname(__FILE__))
require File.expand_path("workbook_sheet_spec", File.dirname(__FILE__))
require File.expand_path("workbook_unobtr_spec", File.dirname(__FILE__))
require File.expand_path("workbook_subclass_spec", File.dirname(__FILE__))

=begin
$VERBOSE = nil

include General

unless Object.method_defined?(:require_relative)
  def require_relative path
    require File.expand_path(path, File.dirname(__FILE__))  
  end
end

require_relative "workbook_open_spec"
require_relative "workbook_close_spec"
require_relative "workbook_save_spec"
require_relative "workbook_misc_spec"
require_relative "workbook_sheet_spec"
require_relative "workbook_unobtr_spec"
require_relative "workbook_subclass_spec"
=end