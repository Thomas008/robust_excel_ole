# -*- coding: utf-8 -*-

require_relative 'spec_helper'

$VERBOSE = nil

include RobustExcelOle
include General

describe "on cygwin",  :if => RUBY_PLATFORM =~ /cygwin/ do
  describe ".cygpath" do
    context "cygwin path is '/cygdrive/c/Users'" do
      context "with '-w' options" do
        it { RobustExcelOle::Cygwin.cygpath('-w', '/cygdrive/c/Users').should eq 'C:\\Users' }
      end

      context "with '-wa' options" do
        it { RobustExcelOle::Cygwin.cygpath('-wa', '/cygdrive/c/Users').should eq 'C:\\Users' }
      end

      context "with '-ws' options" do
        it { RobustExcelOle::Cygwin.cygpath('-ws', '/cygdrive/c/Users').should eq 'C:\\Users' }
      end
    end

    context "windows path is 'C:\\Users'" do
      context "with '-p option" do
        it { RobustExcelOle::Cygwin.cygpath('-p', 'C:\\Users').should eq '/cygdrive/c/Users'}
      end
    end

    context "cygwin path is './'" do
      context "with '-p' options" do
        it { RobustExcelOle::Cygwin.cygpath('-p', './').should eq './' }
      end

     # context "with '-pa' options" do
     #   it { RobustExcelOle::Cygwin.cygpath('-pa', './').should eq File.expand_path('./') + '/' }
     # end

      #context "with '-ps' options" do
      #  it { RobustExcelOle::Cygwin.cygpath('-ps', './').should eq './' }
      #end
    end

  end

end

