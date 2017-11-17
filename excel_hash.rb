require 'date'
require_relative 'excel'
require 'awesome_print'
require_relative 'sample_worksheet'


sheet_names = %w[A1.Cash A2.AR L1.CT]

excel_object = Excel.new(source: sheet_names)

timestamp = DateTime.now.strftime('%y%m%d_%H%M%S')
filepath = "tempxlsx/file_#{timestamp}.xlsx"

sheet_names.each do |sheet_name|
  excel_object.hash_workbook[sheet_name].update(sample_worksheet(sheet_name))
end

# ap excel_object.hash_workbook
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