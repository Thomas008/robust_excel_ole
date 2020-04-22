# -*- coding: utf-8 -*-

require File.join(File.dirname(__FILE__), './spec_helper')
require File.expand_path( '../../lib/robust_excel_ole/reo_common', __FILE__)

$VERBOSE = nil

include General
include RobustExcelOle

module RobustExcelOle

  describe AddressTool do

    before(:all) do
      excel = Excel.new(:reuse => true)
      open_books = excel == nil ? 0 : excel.Workbooks.Count
      puts "*** open books *** : #{open_books}" if open_books > 0
      Excel.kill_all
      @dir = create_tmpdir
      @simple_file = @dir + '/workbook.xls'
      @workbook = Workbook.open(@simple_file)
      @address_tool = @workbook.excel.address_tool
    end

    after(:all) do
      @book.close
    end

    it "should transform relative r1c1-reference into r1c1-format" do
      address_tool.r1c1("Z1S[2]:Z[-1]S4").should == "Z1S(2):Z(-1)S4"
      address_tool.r1c1("Z[1]S2:Z3S[4]").should == "Z(1)S2:Z3S(4)"
      address_tool.r1c1("Z1S[2]").should == "Z1S(2)"
      address_tool.r1c1("Z[-1]S4").should == "Z(-1)S4"
      address_tool.r1c1("Z[3]").should == "Z(3)"
      address_tool.r1c1("S[-2]").should == "S(-2)"
    end

    # test for 1.8.6
    it "should transform relative int_range-reference into r1c1-format" do
      address_tool.r1c1([1..2,[3]..4]).should == "Z1S(3):Z2S4"
      address_tool.r1c1([[1]..2,3..4]).should == "Z(1)S3:Z2S4"
      address_tool.r1c1([[1]..2,nil]).should == "Z(1):Z2"
      address_tool.r1c1([nil,[1]..2]).should == "S(1):S2"
      address_tool.r1c1([nil,[1]]).should == "S(1):S(1)"
    end

    it "should transform relative int_range-reference into r1c1-format" do
      address_tool.r1c1([1..[-2],[3]..4]).should == "Z1S(3):Z(-2)S4"
      address_tool.r1c1([[1]..2,3..[4]]).should == "Z(1)S3:Z2S(4)"
      address_tool.r1c1([1..[-2],nil]).should == "Z1:Z(-2)"
      address_tool.r1c1([nil,[-1]..2]).should == "S(-1):S2"
      address_tool.r1c1([[3]..[3],nil]).should == "Z(3):Z(3)"
      address_tool.r1c1([nil,[-2]..[-2]]).should == "S(-2):S(-2)"
      address_tool.r1c1([[3],nil]).should == "Z(3):Z(3)"
    end

    it "should transform relative r1c1-reference into r1c1-format" do
      address_tool.int_range("Z1S[2]:Z[3]S4").should == [1..[3],[2]..4]
      address_tool.int_range("Z[1]S2:Z3S[4]").should == [[1]..3,2..[4]]
      address_tool.int_range("Z1S[2]").should == [1..1,[2]..[2]]
      address_tool.int_range("Z[3]S4").should == [[3]..[3],4..4]
    end

    it "should transform a1-format" do
      address_tool.a1("A2").should == "A2"
      address_tool.r1c1("A2").should == "Z2S1:Z2S1"
      address_tool.int_range("A2").should == [2..2,1..1]
    end

    it "should transform several-letter-a1-format" do
      address_tool.a1("ABO15").should == "ABO15"
      address_tool.r1c1("ABO15").should == "Z15S743:Z15S743"
      address_tool.int_range("ABO15").should == [15..15,743..743]
    end

    it "should transform complex a1-format" do
      address_tool.a1("A2:B3").should == "A2:B3"
      address_tool.r1c1("A2:B3").should == "Z2S1:Z3S2"
      address_tool.int_range("A2:B3").should == [2..3,1..2]
      address_tool.a1("S1:DP2").should == "S1:DP2"
      address_tool.r1c1("S1:DP2").should == "Z1S19:Z2S120"
      address_tool.int_range("S1:DP2").should == [1..2,19..120]
    end

    it "should transform infinite a1-format" do
      address_tool.a1("A:B").should == "A:B"
      address_tool.r1c1("A:B").should == "S1:S2"
      address_tool.int_range("A:B").should == [nil,1..2]
      address_tool.a1("1:3").should == "1:3"
      address_tool.r1c1("1:3").should == "Z1:Z3"
      address_tool.int_range("1:3").should == [1..3,nil]
      address_tool.a1("B").should == "B"
      address_tool.r1c1("B").should == "S2:S2"
      address_tool.int_range("B").should == [nil,2..2]
      address_tool.a1("3").should == "3"
      address_tool.r1c1("3").should == "Z3:Z3"
      address_tool.int_range("3").should == [3..3,nil]
    end

    it "should transform r1c1-format" do
      address_tool.r1c1("Z2S1").should == "Z2S1"
      address_tool.int_range("Z2S1").should == [2..2,1..1]
      expect{
        address_tool.a1("Z2S1")
      }.to raise_error(NotImplementedREOError)
    end

    it "should transform complex r1c1-format" do
      address_tool.r1c1("Z2S1:Z3S2").should == "Z2S1:Z3S2"
      address_tool.int_range("Z2S1:Z3S2").should == [2..3,1..2]
    end

    it "should transform int_range format" do
      address_tool.int_range([2..2,1..1]).should == [2..2,1..1]
      address_tool.r1c1([2..2,1..1]).should == "Z2S1:Z2S1"
      expect{
        address_tool.a1([2..2,1..1])
      }.to raise_error(NotImplementedREOError)
    end

    it "should transform simple int_range format" do
      address_tool.int_range([2,1]).should == [2..2,1..1]
      address_tool.r1c1([2,1]).should == "Z2S1:Z2S1"
    end

    it "should transform complex int_range format" do
      address_tool.int_range([2,"A"]).should == [2..2,1..1]
      address_tool.r1c1([2,"A"]).should == "Z2S1:Z2S1"
      address_tool.int_range([2,"A".."B"]).should == [2..2,1..2]
      address_tool.r1c1([2,"A".."B"]).should == "Z2S1:Z2S2"
      address_tool.int_range([1..2,"C"]).should == [1..2,3..3]
      address_tool.r1c1([1..2,"C"]).should == "Z1S3:Z2S3"
      address_tool.int_range([1..2,"C".."E"]).should == [1..2,3..5]
      address_tool.r1c1([1..2,"C".."E"]).should == "Z1S3:Z2S5"
      address_tool.int_range([2,3..5]).should == [2..2,3..5]
      address_tool.r1c1([2,3..5]).should == "Z2S3:Z2S5"
      address_tool.int_range([1..2,3..5]).should == [1..2,3..5]
      address_tool.r1c1([1..2,3..5]).should == "Z1S3:Z2S5"
    end

    it "should transform infinite int_range format" do
      address_tool.int_range([nil,1..2]).should == [nil,1..2]
      address_tool.r1c1([nil,1..2]).should == "S1:S2"
      address_tool.int_range([1..3,nil]).should == [1..3,nil]
      address_tool.r1c1([1..3,nil]).should == "Z1:Z3"
      address_tool.int_range([nil,2]).should == [nil,2..2]
      address_tool.r1c1([nil,2]).should == "S2:S2"
      address_tool.int_range([3,nil]).should == [3..3,nil]
      address_tool.r1c1([3,nil]).should == "Z3:Z3"
    end
  
    it "should raise an error" do
      expect{
        address_tool.a1("1A")
      }.to raise_error(address_toolInvalid, /format not correct/)
      expect{
        address_tool.r1c1("A1B")
      }.to raise_error(address_toolInvalid, /format not correct/)
      #expect{
      #  address_tool.int_range(["A".."B","C".."D"])
      #}.to raise_error(address_toolInvalid, /format not correct/)
      #expect{
      #  address_tool.int_range(["A".."B",1..2])
      #}.to raise_error(address_toolInvalid, /format not correct/)
      #expect{
      #  address_tool.int_range(["A".."B",nil])
      #}.to raise_error(address_toolInvalid, /format not correct/)
      expect{
        address_tool.int_range(["A",1,2])
      }.to raise_error(address_toolInvalid, /more than two components/)
    end

  end
end
