require 'rubyXL'

class Excel
  attr_accessor :workbook

  def initialize(source: nil)
    if source
      read_file(source) if source.is_a?(String)
    else
      @workbook = RubyXL::Workbook.new
    end
  end

  def save_file
    @workbook.write(@filepath)
  end

  def read_file(path)
    rubyxl_workbook = RubyXL::Parser.parse(path)
    @workbook = rubyxl_to_hash(rubyxl_workbook)
  end
end
