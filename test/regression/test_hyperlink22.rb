# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionHyperlink22 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_hyperlink22
    @xlsx = 'hyperlink22.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.write_url('A1', 'external:\\\\Vboxsvr\share\foo bar.xlsx')

    workbook.close

    compare_for_regression
  end
end
