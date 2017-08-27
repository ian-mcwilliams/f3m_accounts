require 'date'
require_relative 'excel'

timestamp = DateTime.now.strftime('%y%m%d_%H%M%S')
filepath = "tempxlsx/file_#{timestamp}.xlsx"

workbook = {
    filepath: filepath,
    'A1.Cash' => {
        'A1' => {
            value: 'abcde',
            fill: 'c0c0c0'
        }
    }
}

file = Excel.new(source: workbook)
file.save_file
