class Excel
  attr_accessor :workbook, :filepath, :worksheet

  def initialize(source: nil)
    if source
      # add code to read in from file
    else
      @workbook = RubyXL::Workbook.new
      @filepath = nil
      @worksheet = nil
    end
  end

  def save_file
    @workbook.write(@filepath)
  end
end
