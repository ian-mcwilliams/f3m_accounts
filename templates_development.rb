require 'rubyXL'
require 'date'

require_relative 'excel'

file = Excel.new

file.workbook = RubyXL::Workbook.new

timestamp = DateTime.now.strftime('%y%m%d_%H%M%S')
file.filepath = "tempxlsx/file_#{timestamp}.xlsx"

#
file.workbook.worksheets[0].sheet_name = "A1.Cash"
file.workbook.add_worksheet('A2.AR')
file.workbook.add_worksheet('L1.VAT')
file.workbook.add_worksheet('L2.CT')

(0..3).to_a.each do |item|
  file.worksheet = file.workbook[item]

# Default formatting for sheets
  file.worksheet.change_column_width(1, 45)

  (0..4).to_a.each do |item|
    file.worksheet.change_column_font_name(item, 'Consolas')
  end

  (0..7).to_a.each do |item1|
    (0..4).to_a.each do |item2|
      file.worksheet.add_cell(item1, item2, '')
    end
  end

  (0..7).to_a.each do |item1|
    (0..4).to_a.each do |item2|
      file.worksheet.sheet_data[item1][item2].change_font_size(11)
    end
  end

  %w[left].each_with_index do |item, index|
    file.worksheet.change_column_horizontal_alignment(index, item)
  end

# Title
  file.worksheet.add_cell(0, 5, '') # for reference: additional cell to allow formating of merged cells

# file.worksheet.add_cell(0, 0, 'CASH')
  file.worksheet.sheet_data[0][5].change_border(:left, 'thin')
  file.worksheet.merge_cells(0, 0, 0, 4) # 0, 0 = A1, 1, 1, = B2
  file.worksheet.sheet_data[0][0].change_fill('c0c0c0')
  file.worksheet.sheet_data[0][0].change_horizontal_alignment('center')
  file.worksheet.sheet_data[0][0].change_font_size(12)
  file.worksheet.sheet_data[0][0].change_font_bold(true)

# Headings
  %w[Date Description Dr Cr Balance].each_with_index do |item, index|
    file.worksheet.add_cell(1, index, item).change_horizontal_alignment('center')
    file.worksheet.sheet_data[1][index].change_fill('c0c0c0')
    file.worksheet.sheet_data[1][index].change_border(:top, 'thin')
    file.worksheet.sheet_data[1][index].change_border(:right, 'thin')
    file.worksheet.sheet_data[1][index].change_border(:bottom, 'thin')
  end
# Balance Formula
  file.worksheet.add_cell(2, 4, '', 'C3-D3').set_number_format '0.00'
  file.worksheet.add_cell(3, 4, '', 'E3+C4-D4').set_number_format '0.00'

# Footer Formulae
  file.worksheet.add_cell(5, 1, 'Totals')
  file.worksheet.add_cell(5, 2, '', '=sum(c2:c5)').set_number_format '0.00'
  file.worksheet.add_cell(5, 3, '', '=sum(d2:d5)').set_number_format '0.00'
  file.worksheet.add_cell(6, 1, 'Balance')
  file.worksheet.add_cell(6, 2, '', '=if(c6-d6>0,c6-d6,0)').set_number_format '0.00'
  file.worksheet.add_cell(6, 3, '', '=if(d6-c6>0,d6-c6,0)').set_number_format '0.00'
  file.worksheet.add_cell(6, 4, '', '=c7-d7').set_number_format '0.00'

# Body and Footer Formatting

  (0..4).to_a.each do |item|
    file.worksheet.sheet_data[4][item].change_border(:top, 'thin')
    file.worksheet.sheet_data[4][item].change_border(:bottom, 'thin')
    file.worksheet.sheet_data[7][item].change_border(:top, 'thin')
    file.worksheet.sheet_data[7][item].change_border(:bottom, 'thin')
  end

  (0..4).to_a.each do |item|
    file.worksheet.sheet_data[4][item].change_fill('808080')
    file.worksheet.sheet_data[7][item].change_fill('808080')
  end

  (5..6).to_a.each do |item1|
    (0..4).to_a.each do |item2|
      file.worksheet.sheet_data[item1][item2].change_fill('c0c0c0')
    end
  end
end

file.save_filerequire 'rubyXL'
require 'date'

require_relative 'excel'

file = Excel.new

file.workbook = RubyXL::Workbook.new

timestamp = DateTime.now.strftime('%y%m%d_%H%M%S')
file.filepath = "tempxlsx/file_#{timestamp}.xlsx"

#
file.workbook.worksheets[0].sheet_name = "A1.Cash"
file.workbook.add_worksheet('A2.AR')
file.workbook.add_worksheet('L1.VAT')
file.workbook.add_worksheet('L2.CT')

(0..3).to_a.each do |item|
  file.worksheet = file.workbook[item]

# Default formatting for sheets
  file.worksheet.change_column_width(1, 45)

  (0..4).to_a.each do |item|
    file.worksheet.change_column_font_name(item, 'Consolas')
  end

  (0..7).to_a.each do |item1|
    (0..4).to_a.each do |item2|
      file.worksheet.add_cell(item1, item2, '')
    end
  end

  (0..7).to_a.each do |item1|
    (0..4).to_a.each do |item2|
      file.worksheet.sheet_data[item1][item2].change_font_size(11)
    end
  end

  %w[left].each_with_index do |item, index|
    file.worksheet.change_column_horizontal_alignment(index, item)
  end

# Title
  file.worksheet.add_cell(0, 5, '') # for reference: additional cell to allow formating of merged cells

# file.worksheet.add_cell(0, 0, 'CASH')
  file.worksheet.sheet_data[0][5].change_border(:left, 'thin')
  file.worksheet.merge_cells(0, 0, 0, 4) # 0, 0 = A1, 1, 1, = B2
  file.worksheet.sheet_data[0][0].change_fill('c0c0c0')
  file.worksheet.sheet_data[0][0].change_horizontal_alignment('center')
  file.worksheet.sheet_data[0][0].change_font_size(12)
  file.worksheet.sheet_data[0][0].change_font_bold(true)

# Headings
  %w[Date Description Dr Cr Balance].each_with_index do |item, index|
    file.worksheet.add_cell(1, index, item).change_horizontal_alignment('center')
    file.worksheet.sheet_data[1][index].change_fill('c0c0c0')
    file.worksheet.sheet_data[1][index].change_border(:top, 'thin')
    file.worksheet.sheet_data[1][index].change_border(:right, 'thin')
    file.worksheet.sheet_data[1][index].change_border(:bottom, 'thin')
  end
# Balance Formula
  file.worksheet.add_cell(2, 4, '', 'C3-D3').set_number_format '0.00'
  file.worksheet.add_cell(3, 4, '', 'E3+C4-D4').set_number_format '0.00'

# Footer Formulae
  file.worksheet.add_cell(5, 1, 'Totals')
  file.worksheet.add_cell(5, 2, '', '=sum(c2:c5)').set_number_format '0.00'
  file.worksheet.add_cell(5, 3, '', '=sum(d2:d5)').set_number_format '0.00'
  file.worksheet.add_cell(6, 1, 'Balance')
  file.worksheet.add_cell(6, 2, '', '=if(c6-d6>0,c6-d6,0)').set_number_format '0.00'
  file.worksheet.add_cell(6, 3, '', '=if(d6-c6>0,d6-c6,0)').set_number_format '0.00'
  file.worksheet.add_cell(6, 4, '', '=c7-d7').set_number_format '0.00'

# Body and Footer Formatting

  (0..4).to_a.each do |item|
    file.worksheet.sheet_data[4][item].change_border(:top, 'thin')
    file.worksheet.sheet_data[4][item].change_border(:bottom, 'thin')
    file.worksheet.sheet_data[7][item].change_border(:top, 'thin')
    file.worksheet.sheet_data[7][item].change_border(:bottom, 'thin')
  end

  (0..4).to_a.each do |item|
    file.worksheet.sheet_data[4][item].change_fill('808080')
    file.worksheet.sheet_data[7][item].change_fill('808080')
  end

  (5..6).to_a.each do |item1|
    (0..4).to_a.each do |item2|
      file.worksheet.sheet_data[item1][item2].change_fill('c0c0c0')
    end
  end
end

file.save_file