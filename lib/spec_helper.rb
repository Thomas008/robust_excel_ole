# -*- coding: utf-8 -*-
require "rspec"
require 'tmpdir'
require "fileutils"
require File.join(File.dirname(__FILE__), '../lib/general')

module RobustExcelOle::SpecHelpers   # :nodoc: #
  def create_tmpdir   # :nodoc: #
    tmpdir = Dir.mktmpdir
    FileUtils.cp_r(File.join(File.dirname(__FILE__), 'data'), tmpdir)
    tmpdir + '/data'
  end

  def rm_tmp(tmpdir)  # :nodoc: #
    FileUtils.remove_entry_secure(File.dirname(tmpdir))
  end

  # This method is almost copy of wycats's implementation.
  # http://pochi.hatenablog.jp/entries/2010/03/24
  def capture(stream) # :nodoc: #
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
