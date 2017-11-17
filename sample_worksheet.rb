def sample_worksheet(sheet_name)
  hash_worksheet = {}
  hash_worksheet[:worksheet] = {
      font_style: 'Consolas',
      font_size: 11,
      dp_2: '0.00',
      border_all: 'thin'
  }
  set_worksheet_rows(hash_worksheet)
  hash_worksheet[:columns] = {}
  hash_worksheet[:cells] = {
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
      }
  }
  hash_worksheet
end

def set_worksheet_rows(hash_worksheet)
  hash_worksheet[:rows] = {}
  %w[1 2].each do |row_key|
    hash_worksheet[:rows][row_key] = {
        font_style: 'Arial',
        fill: 'c0c0c0',
        align: 'center',
        bold: true,
        border_all: 'thin'
    }
  end
  %w[6 7].each do |row_key|
    hash_worksheet[:rows][row_key] = {
        fill: 'c0c0c0',
        align: 'right',
        border_all: 'none'
    }
  end
  %w[5 8].each do |row_key|
    hash_worksheet[:rows][row_key] = {
        fill: '808080',
        border_all: 'none'
    }
  end
end
