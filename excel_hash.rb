require 'date'
require_relative 'excel'

timestamp = DateTime.now.strftime('%y%m%d_%H%M%S')
filepath = "tempxlsx/file_#{timestamp}.xlsx"

workbook = {
    filepath: filepath,
    # Sheet1: Master Sheet
    'A1.Cash' => {
        'A1' => {
            value: 'Cash',
            name: 'consolas',
            fill: 'c0c0c0',
            align: 'center',
            bold: true,
            merge: 'A5',
            size: 13
        },
        'A2' => {
            value: 'Date',
            name: 'consolas',
            align: 'center',
            fill: 'c0c0c0',
            border_all: 'thin',
            size: 12
        },
        'A3' => {
            name: 'consolas',
            size: 11
        },
        'A4' => {
            name: 'consolas',
            size: 11
        },
        'A5' => {
            fill: '808080',
            merge: 'E5'
        },
        'A6' => {
            name: 'consolas',
            size: 11,
            fill: 'c0c0c0'
        },
        'A7' => {
            name: 'consolas',
            size: 11,
            fill: 'c0c0c0'
        },
        'A8' => {
            fill: '808080',
            merge: 'H5'
        },
        'B2' => {
            value: 'Description',
            align: 'center',
            fill: 'c0c0c0',
            border_all: 'thin',
            width: 40
        },
        'B5' => {
            fill: '808080'
        },
        'B6' => {
            value: 'Totals',
            align: 'right',
            fill: 'c0c0c0'
        },
        'B7' => {
            value: 'Balance',
            align: 'right',
            fill: 'c0c0c0'
        },
        'C2' => {
            value: 'Dr',
            align: 'center',
            fill: 'c0c0c0',
            border_all: 'thin',
            width: 20
        },
        'C5' => {
            fill: '808080'
        },
        'C6' => {
            formula: '=sum(C2:C5)',
            align: 'right',
            fill: 'c0c0c0'
        },
        'C7' => {
            formula: '=IF(C6-D6>0, C6-D6, 0)',
            align: 'right',
            fill: 'c0c0c0'
        },
        'D2' => {
            value: 'Cr',
            align: 'center',
            fill: 'c0c0c0',
            border_all: 'thin',
            width: 20
        },
        'D5' => {
            fill: '808080'
        },
        'D6' => {
            formula: '=SUM(D2:D5)',
            align: 'right',
            fill: 'c0c0c0'
        },
        'D7' => {
            formula: '=IF(D6-C6>0, D6-C6, 0)',
            align: 'right',
            fill: 'c0c0c0'
        },
        'E2' => {
            value: 'Balance',
            align: 'center',
            fill: 'c0c0c0',
            border_all: 'thin'
        },
        'E3' => {
            formula: '=(C3-D3)',
            format: '0.00',
            align: 'right',
            font: 'consolas',
            fill: 'FFFFFF'
        },
        'E4' => {
            formula: '=E3+C4-D4',
            format: '0.00',
            align: 'right',
            font: 'consolas',
            fill: 'FFFFFF'
        },
        'E5' => {
            fill: '808080'
        },
        'E6' => {
            fill: 'c0c0c0'
        },
        'E7' => {
            formula: '=C7-D7',
            align: 'right',
            fill: 'c0c0c0'
        }
    },
    # Sheet2
    'A2.AR' => {
        'A1' => {
            value: 'Accounts Receivable',
            name: 'consolas',
            fill: 'c0c0c0',
            align: 'center',
            bold: true,
            merge: 'A5',
            size: 13
        },
        'A2' => {
            value: 'Date',
            name: 'consolas',
            align: 'center',
            fill: 'c0c0c0',
            border_all: 'thin',
            size: 12
        },
        'A3' => {
            name: 'consolas',
            size: 11
        },
        'A4' => {
            name: 'consolas',
            size: 11
        },
        'A5' => {
            fill: '808080',
            merge: 'E5'
        },
        'A6' => {
            name: 'consolas',
            size: 11,
            fill: 'c0c0c0'
        },
        'A7' => {
            name: 'consolas',
            size: 11,
            fill: 'c0c0c0'
        },
        'A8' => {
            fill: '808080',
            merge: 'H5'
        },
        'B2' => {
            value: 'Description',
            align: 'center',
            fill: 'c0c0c0',
            border_all: 'thin',
            width: 40
        },
        'B5' => {
            fill: '808080'
        },
        'B6' => {
            value: 'Totals',
            align: 'right',
            fill: 'c0c0c0'
        },
        'B7' => {
            value: 'Balance',
            align: 'right',
            fill: 'c0c0c0'
        },
        'C2' => {
            value: 'Dr',
            align: 'center',
            fill: 'c0c0c0',
            border_all: 'thin',
            width: 20
        },
        'C5' => {
            fill: '808080'
        },
        'C6' => {
            formula: '=sum(C2:C5)',
            align: 'right',
            fill: 'c0c0c0'
        },
        'C7' => {
            formula: '=IF(C6-D6>0, C6-D6, 0)',
            align: 'right',
            fill: 'c0c0c0'
        },
        'D2' => {
            value: 'Cr',
            align: 'center',
            fill: 'c0c0c0',
            border_all: 'thin',
            width: 20
        },
        'D5' => {
            fill: '808080'
        },
        'D6' => {
            formula: '=SUM(D2:D5)',
            align: 'right',
            fill: 'c0c0c0'
        },
        'D7' => {
            formula: '=IF(D6-C6>0, D6-C6, 0)',
            align: 'right',
            fill: 'c0c0c0'
        },
        'E2' => {
            value: 'Balance',
            align: 'center',
            fill: 'c0c0c0',
            border_all: 'thin'
        },
        'E3' => {
            formula: '=(C3-D3)',
            format: '0.00',
            align: 'right',
            font: 'consolas',
            fill: 'FFFFFF'
        },
        'E4' => {
            formula: '=E3+C4-D4',
            format: '0.00',
            align: 'right',
            font: 'consolas',
            fill: 'FFFFFF'
        },
        'E5' => {
            fill: '808080'
        },
        'E6' => {
            fill: 'c0c0c0'
        },
        'E7' => {
            formula: '=C7-D7',
            align: 'right',
            fill: 'c0c0c0'
        }
    },
    # Sheet3
    'L1.CT' => {
        'A1' => {
            value: 'Corporation Tax',
            name: 'consolas',
            fill: 'c0c0c0',
            align: 'center',
            bold: true,
            merge: 'A5',
            size: 13
        },
        'A2' => {
            value: 'Date',
            name: 'consolas',
            align: 'center',
            fill: 'c0c0c0',
            border_all: 'thin',
            size: 12
        },
        'A3' => {
            name: 'consolas',
            size: 11
        },
        'A4' => {
            name: 'consolas',
            size: 11
        },
        'A5' => {
            fill: '808080',
            merge: 'E5'
        },
        'A6' => {
            name: 'consolas',
            size: 11,
            fill: 'c0c0c0'
        },
        'A7' => {
            name: 'consolas',
            size: 11,
            fill: 'c0c0c0'
        },
        'A8' => {
            fill: '808080',
            merge: 'H5'
        },
        'B2' => {
            value: 'Description',
            align: 'center',
            fill: 'c0c0c0',
            border_all: 'thin',
            width: 40
        },
        'B5' => {
            fill: '808080'
        },
        'B6' => {
            value: 'Totals',
            align: 'right',
            fill: 'c0c0c0'
        },
        'B7' => {
            value: 'Balance',
            align: 'right',
            fill: 'c0c0c0'
        },
        'C2' => {
            value: 'Dr',
            align: 'center',
            fill: 'c0c0c0',
            border_all: 'thin',
            width: 20
        },
        'C5' => {
            fill: '808080'
        },
        'C6' => {
            formula: '=sum(C2:C5)',
            align: 'right',
            fill: 'c0c0c0'
        },
        'C7' => {
            formula: '=IF(C6-D6>0, C6-D6, 0)',
            align: 'right',
            fill: 'c0c0c0'
        },
        'D2' => {
            value: 'Cr',
            align: 'center',
            fill: 'c0c0c0',
            border_all: 'thin',
            width: 20
        },
        'D5' => {
            fill: '808080'
        },
        'D6' => {
            formula: '=SUM(D2:D5)',
            align: 'right',
            fill: 'c0c0c0'
        },
        'D7' => {
            formula: '=IF(D6-C6>0, D6-C6, 0)',
            align: 'right',
            fill: 'c0c0c0'
        },
        'E2' => {
            value: 'Balance',
            align: 'center',
            fill: 'c0c0c0',
            border_all: 'thin'
        },
        'E3' => {
            formula: '=(C3-D3)',
            format: '0.00',
            align: 'right',
            font: 'consolas',
            fill: 'FFFFFF'
        },
        'E4' => {
            formula: '=E3+C4-D4',
            format: '0.00',
            align: 'right',
            font: 'consolas',
            fill: 'FFFFFF'
        },
        'E5' => {
            fill: '808080'
        },
        'E6' => {
            fill: 'c0c0c0'
        },
        'E7' => {
            formula: '=C7-D7',
            align: 'right',
            fill: 'c0c0c0'
        }
    }
}

file = Excel.new(source: workbook)
file.save_file