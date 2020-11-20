# -*- coding: utf-8 -*-
require 'rspec'
require 'tmpdir'
require 'fileutils'
require_relative '../lib/robust_excel_ole'

# @private
module RobustExcelOle::SpecHelpers

  # @private
  def create_tmpdir  
    tmpdir = Dir.mktmpdir
    FileUtils.cp_r(File.join(File.dirname(__FILE__), 'data'), tmpdir)
    tmpdir + '/data'
  end

  # @private
  def rm_tmp(tmpdir)  
    FileUtils.remove_entry_secure(File.dirname(tmpdir))
  end

  # This method is almost copy of wycats's implementation.
  # http://pochi.hatenablog.jp/entries/2010/03/24
  # @private
  def capture(stream) 
    begin
      stream = stream.to_s
      eval "$#{stream} = StringIO.new"
      yield
      result = eval("$#{stream}").string
    ensure
      eval("$#{stream} = #{stream.upcase}")
    end
    result
  end
end

RSpec.configure do |config|
  config.include RobustExcelOle::SpecHelpers
end
