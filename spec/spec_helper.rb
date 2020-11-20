# -*- coding: utf-8 -*-
require "rspec"
require 'tmpdir'
require "fileutils"
require_relative '../lib/robust_excel_ole'

# @private
module RobustExcelOle::SpecHelpers

  def create_tmpdir     
    tmpdir = Dir.mktmpdir
    FileUtils.cp_r(File.join(File.dirname(__FILE__), 'data'), tmpdir)
    tmpdir + '/data'
  end

  def rm_tmp(tmpdir)   
    FileUtils.rm_f(File.dirname(tmpdir))
  end

  # This method is almost copy of wycats's implementation.
  # http://pochi.hatenablog.jp/entries/2010/03/24
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
  config.expect_with :rspec do |expectations|
    expectations.syntax = [:should, :expect]
  end
end
