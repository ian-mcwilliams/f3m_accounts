class Excel
  attr_accessor :workbook, :filepath, :worksheet
  @workbook = nil
  @filepath = nil
  @worksheet = nil
  def save_file
    @workbook.write(@filepath)
  end
end
