# -*- coding: utf-8 -*-

require 'helper'
require 'write_xlsx/chart'

class TestWriteMarkerValue < Minitest::Test
  def setup
    @chart = Writexlsx::Chart.new('Bar')
  end

  def test_write_marker_value
    expected = '<c:marker val="1"/>'
    @chart.instance_variable_set(:@default_marker, 'none')
    result = @chart.__send__('write_marker_value')
    assert_equal(expected, result.join)
  end
end
