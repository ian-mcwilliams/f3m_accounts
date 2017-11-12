require 'date'
require_relative 'excel'
require 'awesome_print'

timestamp = DateTime.now.strftime('%y%m%d_%H%M%S')
filepath = "tempxlsx/file_#{timestamp}.xlsx"

sheet_names = %w[A1.Cash] # A2.AR L1.CT]

hash_workbook = {}
sheet_names.each do |sheet_name|
  current_worksheet = {
      sheet_name => {
          worksheet: {
              font_style: 'Consolas',
              font_size: 11,
              dp_2: '0.00',
              border_all: 'thin'
          },
          rows: {
              '1' => {
                  font_style: 'Arial',
                  fill: 'c0c0c0',
                  align: 'center',
                  bold: true,
                  border_all: 'thin'
              },
              '2' => {
                  font_style: 'Arial',
                  fill: 'c0c0c0',
                  align: 'center',
                  bold: true,
                  border_all: 'thin'
              },
              '6' => {
                  fill: 'c0c0c0',
                  align: 'right',
                  border_all: 'none'
              },
              '7' => {
                  fill: 'c0c0c0',
                  align: 'right',
                  border_all: 'none'
              },
              '5' => {
                  fill: '808080',
                  border_all: 'none'
              },
              '8' => {
                  fill: '808080',
                  border_all: 'none'
              }
          },
          columns: {

          },
          cells: {
              'A1' => {
                  value: sheet_name,
                  font_size: 13,
                  merge: 'A5'
              },
              'A2' => {
                  value: 'Date',
                  font_size: 12
              },
              'B2' => {
                  value: 'Description',
                  width: 40
              },
              'B6' => {
                  value: 'Totals'
              },
              'B7' => {
                  value: 'Balance'
              },
              'C2' => {
                  value: 'Dr'
              },
              'C6' => {
                  formula: '=sum(C2:C5)',
                  border_all: 'thin'
              },
              'C7' => {
                  formula: '=IF(C6-D6>0, C6-D6, 0)',
                  border_all: 'thin'
              },
              'D2' => {
                  value: 'Cr'
              },
              'D6' => {
                  formula: '=SUM(D2:D5)',
                  border_all: 'thin'
              },
              'D7' => {
                  formula: '=IF(D6-C6>0, D6-C6, 0)',
                  border_all: 'thin'
              },
              'E2' => {
                  value: 'Balance',
              },
              'E3' => {
                  formula: '=IF(sum(C3:D3)=0,sum(C3:D3), C3-D3)'
              },
              'E6' => {
                  border_all: 'thin'
              },
              'E7' => {
                  formula: '=C7-D7',
                  border_all: 'thin'
              },
          }
      }
  }
  hash_workbook.update(current_worksheet)
end
# ap hash_workbook
excel_object = Excel.new(source: hash_workbook)
excel_object.save_file(filepath)


# if combined_hash_cell[:balance] || combined_hash_cell[:cr_balance]
#   balance = hash_cell_key.sub((/\d+/), row_index.to_s)
#   dr = hash_cell_key.sub(/\D+/, "C")
#   cr = hash_cell_key.sub(/\D+/, "D")
#   if combined_hash_cell[:cr_balance]
#     balance = "=#{balance}+#{cr}-#{dr}"
#   else
#     balance = "=#{balance}+#{dr}-#{cr}"
#   end
#   rubyxl_worksheet[row_index][column_index].change_contents(combined_hash_cell[:balance].to_s, balance)
# end