require 'rubyXL'

class Excel
  attr_accessor :hash_workbook

  def initialize(source: nil)
    if source.class == Hash
      @hash_workbook = source
    elsif source.class == String
      read_file(source)
    end
  end

  def save_file(filepath)
    # rubyxl_workbook = hash_workbook_to_rubyxl_workbook
    # rubyxl_workbook.write(@hash_workbook[:filepath])
    hash_workbook_to_rubyxl_workbook.write(filepath)
  end

  def hash_workbook_to_rubyxl_workbook
    rubyxl_workbook = RubyXL::Workbook.new
    first_worksheet = true
    @hash_workbook.each do |hash_key, hash_value|
      if first_worksheet
        rubyxl_workbook.worksheets[0].sheet_name = hash_key
        first_worksheet = false
      else
        rubyxl_workbook.add_worksheet(hash_key)
      end
      hash_worksheet_to_rubyxl_worksheet(hash_value, rubyxl_workbook[hash_key])
    end
    rubyxl_workbook
  end

  def hash_worksheet_to_rubyxl_worksheet(hash_worksheet, rubyxl_worksheet)

    hash_worksheet[:cells].each do |hash_cell_key, hash_cell|
      hash_cell_to_rubyxl_cell(hash_cell_key, hash_cell, rubyxl_worksheet)

      # rubyxl_worksheet.change_column_width(column_index, hash_cell[:width]) if hash_cell[:width]
      #
      # rubyxl_worksheet.change_row_font_name(row_index, hash_cell[:name]) if hash_cell[:name]
      # rubyxl_worksheet.change_row_font_size(row_index, hash_cell[:size])  if hash_cell[:size]
    end
  end

  def hash_cell_to_rubyxl_cell(hash_cell_key, hash_cell, rubyxl_worksheet)
    row_index, column_index = RubyXL::Reference.ref2ind(hash_cell_key)

    index_b, index_a = RubyXL::Reference.ref2ind(hash_cell[:merge]) if hash_cell[:merge]
    rubyxl_worksheet.merge_cells(row_index, column_index, index_a, index_b) if hash_cell[:merge]
    if hash_cell[:formula]
      rubyxl_worksheet.add_cell(row_index, column_index, '', hash_cell[:formula]).set_number_format '0.00'
    else
      rubyxl_worksheet.add_cell(row_index, column_index, hash_cell[:value])
    end

    rubyxl_worksheet[row_index][column_index].change_contents(hash_cell[:sum], rubyxl_worksheet[row_index][column_index].formula) if hash_cell[:sum]



    rubyxl_worksheet[row_index][column_index].set_number_format(hash_cell[:format]) if hash_cell[:format]
    rubyxl_worksheet[row_index][column_index].change_fill(hash_cell[:fill]) if hash_cell[:fill]
    rubyxl_worksheet[row_index][column_index].change_horizontal_alignment(hash_cell[:align]) if hash_cell[:align]
    rubyxl_worksheet[row_index][column_index].set_number_format(hash_cell[:format]) if hash_cell[:format]
    rubyxl_worksheet[row_index][column_index].change_font_bold(hash_cell[:bold]) if hash_cell[:bold]

    if hash_cell[:border_all]
      rubyxl_worksheet[row_index][column_index].change_border('top' , hash_cell[:border_all])
      rubyxl_worksheet[row_index][column_index].change_border('bottom' , hash_cell[:border_all])
      rubyxl_worksheet[row_index][column_index].change_border('left' , hash_cell[:border_all])
      rubyxl_worksheet[row_index][column_index].change_border('right' , hash_cell[:border_all])
    end
  end
  def read_file(path)
    rubyxl_workbook = RubyXL::Parser.parse(path)
    @hash_workbook = rubyxl_to_hash(rubyxl_workbook)
  end

  def rubyxl_to_hash(rubyxl_workbook)
    workbook_hash = {}
    rubyxl_workbook.each do |worksheet|
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
