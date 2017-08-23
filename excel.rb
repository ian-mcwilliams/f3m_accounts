class Excel
  attr_accessor :workbook, :filepath
  @workbook = nil
  @filepath = nil
  def save_file
    @workbook.write(@filepath)
  end
end
