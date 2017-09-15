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
      row_index, column_index = RubyXL::Reference.ref2ind(cell_key)

      index_b, index_a = RubyXL::Reference.ref2ind(attributes[:merge]) if attributes[:merge]
      rubyxl_worksheet.merge_cells(row_index, column_index, index_a, index_b) if attributes[:merge]

      if attributes[:formula]
        rubyxl_worksheet.add_cell(row_index, column_index, '', attributes[:formula]).set_number_format '0.00'
      else
        rubyxl_worksheet.add_cell(row_index, column_index, attributes[:value])
      end

      rubyxl_worksheet[row_index][column_index].change_contents(attributes[:sum], rubyxl_worksheet[row_index][column_index].formula) if attributes[:sum]
      rubyxl_worksheet.change_column_width(column_index, attributes[:width]) if attributes[:width]

      rubyxl_worksheet.change_row_font_name(row_index, attributes[:name]) if attributes[:name]
      rubyxl_worksheet.change_row_font_size(row_index, attributes[:size])  if attributes[:size]


      rubyxl_worksheet[row_index][column_index].set_number_format(attributes[:format]) if attributes[:format]
      rubyxl_worksheet[row_index][column_index].change_fill(attributes[:fill]) if attributes[:fill]
      rubyxl_worksheet[row_index][column_index].change_horizontal_alignment(attributes[:align]) if attributes[:align]
      rubyxl_worksheet[row_index][column_index].set_number_format(attributes[:format]) if attributes[:format]
      rubyxl_worksheet[row_index][column_index].change_font_bold(attributes[:bold]) if attributes[:bold]

      # needs a: change_border_all
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

  def rubyxl_to_hash(rubyxl_workbook)
    workbook_hash = {}
    rubyxl_workbook.each do |worksheet|
      next unless worksheet.sheet_name == 'E4. Net Sales'
      worksheet_hash = {row_count: worksheet.count, column_count: 1, cells: {}}
      worksheet_to_hash(worksheet, worksheet_hash)
      process_sheet_to_populated_block(worksheet_hash)
      workbook_hash[worksheet.sheet_name] = worksheet_hash
    end
    workbook_hash
  end

  def worksheet_to_hash(worksheet, worksheet_hash)
    worksheet.each_with_index do |row, row_index|
      cells = row&.cells
      if cells.nil?
        cell_hash = {}
        cell_key = RubyXL::Reference.ind2ref(row_index, 0)
        worksheet_hash[cell_key] = cell_hash
      else
        row&.cells.each_with_index do |cell, column_index|
          cell_hash = {}
          cell_key = RubyXL::Reference.ind2ref(row_index, column_index)
          worksheet_hash[:cells][cell_key] = cell_hash
          worksheet_hash[:column_count] = column_index + 1 if column_index + 1 > worksheet_hash[:column_count]
        end
      end
    end
  end

  def process_sheet_to_populated_block(worksheet_hash)
    worksheet_hash[:row_count].times do |row_index|
      worksheet_hash[:column_count].times do |column_index|
        cell_key = RubyXL::Reference.ind2ref(row_index, column_index)
        worksheet_hash[cell_key] = {} unless worksheet_hash[cell_key]
      end
    end
  end
end
