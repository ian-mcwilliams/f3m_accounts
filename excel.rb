require 'rubyXL'

class Excel
  attr_accessor :workbook, :filepath, :worksheet

  def initialize(source: nil)
    if source.class == Hash
      @workbook = source
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
      row_index, column_index = cell_key_to_coordinates(cell_key)
      puts column_index
      rubyxl_worksheet.add_cell(row_index, column_index, attributes[:value])
      rubyxl_worksheet[row_index][column_index].change_fill(attributes[:fill])
    end
  end

  def cell_key_to_coordinates(cell_key)
    row_start_index = cell_key =~ /\d+/
    column_string = cell_key[0..row_start_index - 1]
    value = ('A'..'Z').map.with_index.to_h
    column_index = column_string.chars.inject(0) { |sum, current| sum * 26 + value[current] + 1 } - 1
    row_index = cell_key[row_start_index..-1].to_i - 1
    [row_index, column_index]
  end
end
