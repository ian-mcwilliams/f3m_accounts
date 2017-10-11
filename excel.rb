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

    hash_worksheet[:row_count] = 0
    hash_worksheet[:column_count] = 0
    hash_worksheet[:cells].keys.each do |hash_cell_key|
      row_index, column_index = RubyXL::Reference.ref2ind(hash_cell_key)
      hash_worksheet[:row_count] = row_index + 1 if row_index + 1 > hash_worksheet[:row_count]
      hash_worksheet[:column_count] = column_index + 1 if column_index + 1 > hash_worksheet[:column_count]
    end

    process_sheet_to_populated_block(hash_worksheet)
    hash_row_column(hash_worksheet)

    hash_worksheet[:cells].each do |hash_cell_key, combined_hash_cell|
      combined_hash_cell = hash_worksheet[:worksheet].merge(combined_hash_cell)
      hash_cell_to_rubyxl_cell(hash_cell_key, combined_hash_cell, rubyxl_worksheet)
    end
  end

  def hash_cell_to_rubyxl_cell(hash_cell_key, combined_hash_cell, rubyxl_worksheet)
    row_index, column_index = RubyXL::Reference.ref2ind(hash_cell_key)

    index_b, index_a = RubyXL::Reference.ref2ind(combined_hash_cell[:merge])
    rubyxl_worksheet.merge_cells(row_index, column_index, index_a, index_b) if combined_hash_cell[:merge]
    if combined_hash_cell[:formula]
      rubyxl_worksheet.add_cell(row_index, column_index, '', combined_hash_cell[:formula]).set_number_format combined_hash_cell[:dp_2]
    else
      rubyxl_worksheet.add_cell(row_index, column_index, combined_hash_cell[:value])
      rubyxl_worksheet.change_row_fill(row_index, combined_hash_cell[:fill])  if combined_hash_cell[:fill]
    end

    rubyxl_worksheet.change_column_width(column_index, combined_hash_cell[:width])  if combined_hash_cell[:width]

    # rubyxl_worksheet[row_index][column_index].change_contents(hash_cell[:sum], rubyxl_worksheet[row_index][column_index].formula) if hash_cell[:sum]
    # rubyxl_worksheet[row_index][column_index].set_number_format(hash_cell[:format]) if hash_cell[:format]
    rubyxl_worksheet[row_index][column_index].change_font_name(combined_hash_cell[:font_style]) if combined_hash_cell[:font_style]
    rubyxl_worksheet[row_index][column_index].change_font_size(combined_hash_cell[:font_size]) if combined_hash_cell[:font_size]
    rubyxl_worksheet[row_index][column_index].change_fill(combined_hash_cell[:fill]) if combined_hash_cell[:fill]
    rubyxl_worksheet[row_index][column_index].change_horizontal_alignment(combined_hash_cell[:align]) if combined_hash_cell[:align]
    rubyxl_worksheet[row_index][column_index].change_font_bold(combined_hash_cell[:bold]) if combined_hash_cell[:bold]

    if combined_hash_cell[:border_all]
      rubyxl_worksheet[row_index][column_index].change_border('top' , combined_hash_cell[:border_all])
      rubyxl_worksheet[row_index][column_index].change_border('bottom' , combined_hash_cell[:border_all])
      rubyxl_worksheet[row_index][column_index].change_border('left' , combined_hash_cell[:border_all])
      rubyxl_worksheet[row_index][column_index].change_border('right' , combined_hash_cell[:border_all])
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

  def process_sheet_to_populated_block(hash_worksheet)
    hash_worksheet[:row_count].times do |row_index|
      hash_worksheet[:column_count].times do |column_index|
        cell_key = RubyXL::Reference.ind2ref(row_index, column_index)
        hash_worksheet[:cells][cell_key] = {} unless hash_worksheet[:cells][cell_key]
      end
    end
  end

  def hash_row_column(hash_worksheet)
    hash_row = Hash.new{|hsh,row| hsh[row] = [] }
    hash_column = Hash.new{|hsh,column| hsh[column] = [] }
    hash_worksheet[:row_count].times do |row_index|
      row_key = RubyXL::Reference.ind2ref(row_index)
      row_key = /\d+/.match(row_key)
      hash_row['rows'].push row_index
      hash_row['rows'].push row_key
    end
    hash_worksheet[:column_count].times do |column_index|
      column_key = RubyXL::Reference.ind2ref(0, column_index)
      column_key = /\D+/.match(column_key)
      hash_column['columns'].push column_index
      hash_column['columns'].push column_key
    end
    puts [hash_row]
    puts [hash_column]
  end
end
