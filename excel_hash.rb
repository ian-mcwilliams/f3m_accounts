require 'date'
require_relative 'excel'

timestamp = DateTime.now.strftime('%y%m%d_%H%M%S')
filepath = "tempxlsx/file_#{timestamp}.xlsx"

workbook = {
    filepath: filepath,
    # Sheet1
    'A1.Cash' => {
        'A1' => {
            value: 'Cash',
            font: 'consolas',
            fill: 'c0c0c0',
            align: 'center',
            bold: true,
            merge: ''
        },
        'A2' => {
            value: 'Date',
            align: 'center',
            fill: 'c0c0c0',
            border_top: 'thin',
            border_bottom: 'thin',
            border_left: 'thin',
            border_right: 'thin'
        },
        'B2' => {
            value: 'Description',
            align: 'center',
            fill: 'c0c0c0',
            border_top: 'thin',
            border_bottom: 'thin',
            border_left: 'thin',
            border_right: 'thin'
        },
        'C2' => {
            value: 'Dr',
            align: 'center',
            fill: 'c0c0c0',
            border_top: 'thin',
            border_bottom: 'thin',
            border_left: 'thin',
            border_right: 'thin'
        },
        'D2' => {
            value: 'Cr',
            align: 'center',
            fill: 'c0c0c0',
            border_top: 'thin',
            border_bottom: 'thin',
            border_left: 'thin',
            border_right: 'thin'
        },
        'E2' => {
            value: 'Balance',
            align: 'center',
            fill: 'c0c0c0',
            border_top: 'thin',
            border_bottom: 'thin',
            border_left: 'thin',
            border_right: 'thin'
        },
        'E3' => {
            sum: '=C3-D3',
            format: '0.00',
            align: 'right',
            font: 'consolas',
            fill: 'FFFFFF'
        },
        'E4' => {
            sum: '=E3+C4-D4',
            format: '0.00',
            align: 'right',
            font: 'consolas',
            fill: 'FFFFFF'
        },
        'A5' => {
            fill: '808080'
        },
        'B5' => {
            fill: '808080'
        },
        'C5' => {
            fill: '808080'
        },
        'D5' => {
            fill: '808080'
        },
        'E5' => {
            fill: '808080'
        },
        'A6' => {
            fill: 'c0c0c0'
        },
        'B6' => {
            value: 'Totals',
            align: 'right',
            fill: 'c0c0c0'
        },
        'C6' => {
            value: '=SUM(C2:C5)',
            align: 'right',
            fill: 'c0c0c0'
        },
        'D6' => {
            value: '=SUM(D2:D5)',
            align: 'right',
            fill: 'c0c0c0'
        },
        'E6' => {
            fill: 'c0c0c0'
        },
        'A7' => {
            fill: 'c0c0c0'
        },
        'B7' => {
            value: 'Balance',
            align: 'right',
            fill: 'c0c0c0'
        },
        'C7' => {
            value: '=IF(C6-D6>0, C6-D6, 0)',
            align: 'right',
            fill: 'c0c0c0'
        },
        'D7' => {
            value: '=IF(D6-C6>0, D6-C6, 0)',
            align: 'right',
            fill: 'c0c0c0'
        },
        'E7' => {
            value: '=C7-D7',
            align: 'right',
            fill: 'c0c0c0'
        },
        'A8' => {
            fill: '808080'
        }
    },
    # Sheet2
    'A2.AR' => {
        'A1' => {
            value: 'Acccounts Receivable',
            font: 'consolas',
            fill: 'c0c0c0',
            align: 'center',
            bold: true,
            merge: ''
        },
        'A2' => {
            value: 'Date',
            align: 'center',
            fill: 'c0c0c0',
            border_top: 'thin',
            border_bottom: 'thin',
            border_left: 'thin',
            border_right: 'thin'
        },
        'B2' => {
            value: 'Description',
            align: 'center',
            fill: 'c0c0c0',
            border_top: 'thin',
            border_bottom: 'thin',
            border_left: 'thin',
            border_right: 'thin'
        },
        'C2' => {
            value: 'Dr',
            align: 'center',
            fill: 'c0c0c0',
            border_top: 'thin',
            border_bottom: 'thin',
            border_left: 'thin',
            border_right: 'thin'
        },
        'D2' => {
            value: 'Cr',
            align: 'center',
            fill: 'c0c0c0',
            border_top: 'thin',
            border_bottom: 'thin',
            border_left: 'thin',
            border_right: 'thin'
        },
        'E2' => {
            value: 'Balance',
            align: 'center',
            fill: 'c0c0c0',
            border_top: 'thin',
            border_bottom: 'thin',
            border_left: 'thin',
            border_right: 'thin'
        },
        'E3' => {
            sum: '=C3-D3',
            format: '0.00',
            align: 'right',
            font: 'consolas',
            fill: 'FFFFFF'
        },
        'E4' => {
            sum: '=E3+C4-D4',
            format: '0.00',
            align: 'right',
            font: 'consolas',
            fill: 'FFFFFF'
        },
        'A5' => {
            fill: '808080'
        },
        'B5' => {
            fill: '808080'
        },
        'C5' => {
            fill: '808080'
        },
        'D5' => {
            fill: '808080'
        },
        'E5' => {
            fill: '808080'
        },
        'A6' => {
            fill: 'c0c0c0'
        },
        'B6' => {
            value: 'Totals',
            align: 'right',
            fill: 'c0c0c0'
        },
        'C6' => {
            value: '=SUM(C2:C5)',
            align: 'right',
            fill: 'c0c0c0'
        },
        'D6' => {
            value: '=SUM(D2:D5)',
            align: 'right',
            fill: 'c0c0c0'
        },
        'E6' => {
            fill: 'c0c0c0'
        },
        'A7' => {
            fill: 'c0c0c0'
        },
        'B7' => {
            value: 'Balance',
            align: 'right',
            fill: 'c0c0c0'
        },
        'C7' => {
            value: '=IF(C6-D6>0, C6-D6, 0)',
            align: 'right',
            fill: 'c0c0c0'
        },
        'D7' => {
            value: '=IF(D6-C6>0, D6-C6, 0)',
            align: 'right',
            fill: 'c0c0c0'
        },
        'E7' => {
            value: '=C7-D7',
            align: 'right',
            fill: 'c0c0c0'
        },
        'A8' => {
            fill: '808080'
        }    },    # Sheet3
    'L1.CT' => {
        'A1' => {
            value: 'Corporation Tax',
            font: 'consolas',
            fill: 'c0c0c0',
            align: 'center',
            bold: true,
            merge: ''
        },
        'A2' => {
            value: 'Date',
            align: 'center',
            fill: 'c0c0c0',
            border_top: 'thin',
            border_bottom: 'thin',
            border_left: 'thin',
            border_right: 'thin'
        },
        'B2' => {
            value: 'Description',
            align: 'center',
            fill: 'c0c0c0',
            border_top: 'thin',
            border_bottom: 'thin',
            border_left: 'thin',
            border_right: 'thin'
        },
        'C2' => {
            value: 'Dr',
            align: 'center',
            fill: 'c0c0c0',
            border_top: 'thin',
            border_bottom: 'thin',
            border_left: 'thin',
            border_right: 'thin'
        },
        'D2' => {
            value: 'Cr',
            align: 'center',
            fill: 'c0c0c0',
            border_top: 'thin',
            border_bottom: 'thin',
            border_left: 'thin',
            border_right: 'thin'
        },
        'E2' => {
            value: 'Balance',
            align: 'center',
            fill: 'c0c0c0',
            border_top: 'thin',
            border_bottom: 'thin',
            border_left: 'thin',
            border_right: 'thin'
        },
        'E3' => {
            sum: '=C3-D3',
            format: '0.00',
            align: 'right',
            font: 'consolas',
            fill: 'FFFFFF'
        },
        'E4' => {
            sum: '=E3+C4-D4',
            format: '0.00',
            align: 'right',
            font: 'consolas',
            fill: 'FFFFFF'
        },
        'A5' => {
            fill: '808080'
        },
        'B5' => {
            fill: '808080'
        },
        'C5' => {
            fill: '808080'
        },
        'D5' => {
            fill: '808080'
        },
        'E5' => {
            fill: '808080'
        },
        'A6' => {
            fill: 'c0c0c0'
        },
        'B6' => {
            value: 'Totals',
            align: 'right',
            fill: 'c0c0c0'
        },
        'C6' => {
            value: '=SUM(C2:C5)',
            align: 'right',
            fill: 'c0c0c0'
        },
        'D6' => {
            value: '=SUM(D2:D5)',
            align: 'right',
            fill: 'c0c0c0'
        },
        'E6' => {
            fill: 'c0c0c0'
        },
        'A7' => {
            fill: 'c0c0c0'
        },
        'B7' => {
            value: 'Balance',
            align: 'right',
            fill: 'c0c0c0'
        },
        'C7' => {
            value: '=IF(C6-D6>0, C6-D6, 0)',
            align: 'right',
            fill: 'c0c0c0'
        },
        'D7' => {
            value: '=IF(D6-C6>0, D6-C6, 0)',
            align: 'right',
            fill: 'c0c0c0'
        },
        'E7' => {
            value: '=C7-D7',
            align: 'right',
            fill: 'c0c0c0'
        },
        'A8' => {
            fill: '808080'
        }
    }
}

file = Excel.new(source: workbook)
file.save_file