# -*- encoding: utf-8 -*-

require 'helper'
require 'write_xlsx/chart'

class TestWriteALatin < Minitest::Test
  def setup
    @chart = Writexlsx::Chart.new('Bar')
  end

  def test_write_a_latin
    expected = '<a:latin typeface="Arial" pitchFamily="34" charset="0"/>'

    result = @chart.__send__(
      'write_a_latin',
      [
        %w[typeface Arial],
        ['pitchFamily', 34],
        ['charset', 0]
      ]
    )
    assert_equal(expected, result.join)
  end
end
