require 'rubyXL'

class Excel
  attr_accessor :workbook

  def initialize(source: nil)
    if source.class == Hash
      @workbook = source
    elsif source.class == String
      read_file(source)
    end
  end

  def save_file
    rubyxl.write(@workbook[:filepath])
  end

  def rubyxl
    rubyxl = RubyXL::Workbook.new
    first_worksheet = true
    @workbook.each do |key, worksheet|
      next if [:filepath].include?(key)
      if first_worksheet
        rubyxl.worksheets[0].sheet_name = key
        first_worksheet = false
      else
        rubyxl.add_worksheet(key)
      end
      rubyxl_cells(rubyxl[key], worksheet)
    end
    rubyxl
  end

  def rubyxl_cells(rubyxl_worksheet, worksheet)
    worksheet.each do |cell_key, attributes|
      (0..253).to_a.each do |item|
        rubyxl_worksheet.change_column_font_name(item, 'Consolas')
        rubyxl_worksheet.change_column_font_size(item, 11)
      end
      rubyxl_worksheet.change_column_width(1, 40)
      rubyxl_worksheet.change_column_width(2, 20)
      rubyxl_worksheet.change_column_width(3, 20)
      rubyxl_worksheet.change_row_font_size(0, 13)
      rubyxl_worksheet.change_row_font_size(1, 12)

      row_index, column_index = RubyXL::Reference.ref2ind(cell_key)
      rubyxl_worksheet.add_cell(row_index, column_index, attributes[:value])
      rubyxl_worksheet.merge_cells(0, 0, 0, 4) if attributes[:merge]
      rubyxl_worksheet.merge_cells(4, 0, 4, 4) if attributes[:merge]
      rubyxl_worksheet.merge_cells(7, 0, 7, 4) if attributes[:merge]
      rubyxl_worksheet.add_cell(row_index, column_index, attributes[:sum]) if attributes[:sum]
      rubyxl_worksheet[row_index][column_index].set_number_format(attributes[:format]) if attributes[:format]
      rubyxl_worksheet[row_index][column_index].change_fill(attributes[:fill]) if attributes[:fill]
      rubyxl_worksheet[row_index][column_index].change_horizontal_alignment(attributes[:align]) if attributes[:align]
      rubyxl_worksheet[row_index][column_index].set_number_format(attributes[:format]) if attributes[:format]
      rubyxl_worksheet[row_index][column_index].change_font_bold(attributes[:bold]) if attributes[:bold]

      # change_border_all???
      rubyxl_worksheet[row_index][column_index].change_border('top' , attributes[:border_top]) if attributes[:border_top]
      rubyxl_worksheet[row_index][column_index].change_border('bottom' , attributes[:border_bottom]) if attributes[:border_bottom]
      rubyxl_worksheet[row_index][column_index].change_border('left' , attributes[:border_left]) if attributes[:border_left]
      rubyxl_worksheet[row_index][column_index].change_border('right' , attributes[:border_right]) if attributes[:border_right]
    end
  end

  def read_file(path)
    rubyxl_workbook = RubyXL::Parser.parse(path)
    @workbook = rubyxl_to_hash(rubyxl_workbook)
  end
end
