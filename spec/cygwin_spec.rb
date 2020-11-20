# -*- coding: utf-8 -*-
require_relative 'spec_helper'

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
      context "with '-u option" do
        it { RobustExcelOle::Cygwin.cygpath('-u', 'C:\\Users').should eq '/cygdrive/c/Users'}
      end
    end

    context "cygwin path is './'" do
      context "with '-u' options" do
        it { RobustExcelOle::Cygwin.cygpath('-u', './').should eq './' }
      end

      context "with '-ua' options" do
        it { RobustExcelOle::Cygwin.cygpath('-ua', './').should eq File.expand_path('./') + '/' }
      end

      context "with '-us' options" do
        it { RobustExcelOle::Cygwin.cygpath('-us', './').should eq './' }
      end
    end

  end

end
