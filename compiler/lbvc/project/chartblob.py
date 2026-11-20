"""chartblob.py - the OleObjectBlob of an MSChart (MSCHRT20.OCX) in a .frx (phase 9.2).

VB6 keeps the whole design-time state of an MSChart in the .frx; the .frm has only Height, Left,
OleObjectBlob, TabIndex, Top and Width (no _ExtentX / _ExtentY / _Version). Without the blob VB6 cannot
load the control ("Cannot load control", error 31037), so the designer gives every MSChart one.

  [u32 'LB\t\0'][u32 n: the bytes after this 24-byte head][u32 Left][u32 Top][u32 Width][u32 Height] (twips)
  [u32 0x12344321][u32 8][u32 ExtentX][u32 ExtentY] (HIMETRIC) and 34 bytes that do not change,
  then the chart as a stream of objects (MFC CArchive style):
  [u32 object id] - a new object (they count from 0), an id seen before (a reference), 0xFFFFFFFF (none);
  a new object: [u32 class id]; a new class: [u16 len][name\0][u16 schema]; then its own data and the
  objects of its base classes / members.

A new chart has 183 objects of 33 classes (VtChart, VtDataGrid, VtMatrix ... VtWindow): CHART below, one
entry per object (its class, its bytes), the default state of a new chart as VB6 saves it. new_blob()
puts the size of the control in it:
  - the head (Left, Top, Width, Height) and the extents (HIMETRIC),
  - the end of object 30 (the area of the chart): Width - 15, Height - 15 (floats, twips),
  - the end of object 114 (the plot): 240, 240, Width - 255, Height - 255 (floats, twips),
and our own values in the data grid (5 rows x 4 columns, objects 7-10, 12-15 ... 27-30; VB6 puts random
numbers there). The places of the labels (objects 66, 82, 98) stay 0: MSChart works them out when it draws
(VB6 saves 0 there for a small chart).

read() / decode(): the chart of any blob VB6 wrote, object by object, and its settings as the runtime keeps them
(the compiler gives them to the control, cgen.py; runtime/forms/ocx_chart.inc OcxChart::load).

Checked against .frx files that VB6 saved (DEV-9/mc1: a new chart 1455 x 735 and one 6000 x 4005): apart from
the values of the grid and a few bytes VB6 does not set (uninitialized memory in objects 62, 66, 78, 94),
new_blob() gives them byte for byte.
"""
import struct

MAGIC = b'LB\t\0'
HEAD = 24
SIG = 0x12344321
# the bytes after the extents that do not change
CONST = bytes.fromhex('bd640605000006000000ffffffffffff000001000100000000000100010000000000')

# the default chart: (class, bytes of the object), objects 0 - 182, in the order of the stream
CHART = (
    ('VtChart', '0000000000000000080056744368617274000600'),       # 0
    ('VtDataGrid',  # 1
     '01000000010000000b005674446174614772696400010002000000090056744d6174726978000100030000000d005674'
     '436f6c6c656374696f6e0001001d0004000000090056744f626a65637400010006000500ffffffff'),
    ('VtString', '020000000500000009005674537472696e670001000200433100060000000800567456616c756500010004000000'),       # 2
    ('VtString', '0300000005000000020043320006000000'),       # 3
    ('VtObject', '04000000040000000500000002004333000600000004000000'),       # 4
    ('VtString', '05000000050000000200433400'),       # 5
    ('VtObject', '0600000004000000060000000500000002005231000600000004000000'),       # 6
    ('VtInteger', '07000000070000000a005674496e7465676572000100290000000600000004000000'),       # 7
    ('VtInteger', '0800000007000000430000000600000004000000'),       # 8
    ('VtInteger', '0900000007000000220000000600000004000000'),       # 9
    ('VtInteger', '0a00000007000000000000000600000004000000'),       # 10
    ('VtString', '0b0000000500000002005232000600000004000000'),       # 11
    ('VtInteger', '0c00000007000000450000000600000004000000'),       # 12
    ('VtInteger', '0d00000007000000180000000600000004000000'),       # 13
    ('VtInteger', '0e000000070000004e0000000600000004000000'),       # 14
    ('VtInteger', '0f000000070000003a0000000600000004000000'),       # 15
    ('VtString', '100000000500000002005233000600000004000000'),       # 16
    ('VtInteger', '11000000070000003e0000000600000004000000'),       # 17
    ('VtInteger', '1200000007000000400000000600000004000000'),       # 18
    ('VtInteger', '1300000007000000050000000600000004000000'),       # 19
    ('VtInteger', '14000000070000002d0000000600000004000000'),       # 20
    ('VtString', '150000000500000002005234000600000004000000'),       # 21
    ('VtInteger', '1600000007000000510000000600000004000000'),       # 22
    ('VtInteger', '17000000070000001b0000000600000004000000'),       # 23
    ('VtInteger', '18000000070000003d0000000600000004000000'),       # 24
    ('VtInteger', '19000000070000005b0000000600000004000000'),       # 25
    ('VtString', '1a0000000500000002005235000600000004000000'),       # 26
    ('VtInteger', '1b000000070000005f0000000600000004000000'),       # 27
    ('VtInteger', '1c000000070000002a0000000600000004000000'),       # 28
    ('VtInteger', '1d000000070000001b0000000600000004000000'),       # 29
    ('VtInteger', '1e000000070000002400000006000000040000000100010004000500000000000000000000000000b44400003444'),       # 30
    ('VtBackdrop',  # 31
     '1f000000080000000b0056744261636b64726f7000010001000100010002000000000002000000000000000000c84200'
     '00c842000001000200000000000200ffffff000000a0410000'),
    ('VtFill',  # 32
     '20000000090000000700567446696c6c000100010001000200ffffff00020000000000000001000200ffffff00020080'
     'c0ff000000'),
    ('VtPicture', '210000000a0000000a005674506963747572650001000100ffffffffffff0400000000000400000004000000'),       # 33
    ('VtChartFootnote', '220000000b000000100056744368617274466f6f746e6f74650001000c0000000c005674436861727454657874000100'),       # 34
    ('VtTextBlock', '230000000d0000000c00567454657874426c6f636b00020001000000020002000000ffffffffffff'),       # 35
    ('VtFont', '240000000e00000007005674466f6e74000100'),       # 36
    ('VtString',  # 37
     '25000000050000000d004d532053616e7320536572696600060000000400000000000441000000000200000000000400'
     '0000000064000000000000000000000000000000000000000000'),
    ('VtString',  # 38
     '260000000500000000000600000004000000000000000000000000000000000000000000000000000000000004000000'
     '0f0000000f005674436861727453656374696f6e000100f0006400010000000000000000000000000000000000000005'
     '00'),
    ('VtBackdrop',  # 39
     '270000000800000001000100010002000000000002000000000000000000c8420000c842000001000200000000000200'
     'ffffff000000a0410000'),
    ('VtFill', '2800000009000000010001000200ffffff00020000000000000001000200ffffff00020080c0ff000000'),       # 40
    ('VtPicture', '290000000a0000000100ffffffffffff040000000000040000000400000004000000'),       # 41
    ('VtChartLegend', '2a000000100000000e00567443686172744c6567656e64000100'),       # 42
    ('VtFont',  # 43
     '2b0000000e00000025000000000004410000000002000000000004000000010000000200020000000f000000f000f000'
     '01000000000000000000000000000000000000000400'),
    ('VtBackdrop',  # 44
     '2c0000000800000001000100010002000000000002000000000000000000c8420000c842000001000200000000000200'
     'ffffff000000a0410000'),
    ('VtFill', '2d00000009000000010001000200ffffff00020000000000000001000200ffffff00020080c0ff000000'),       # 45
    ('VtPicture', '2e0000000a0000000100ffffffffffff040000000000040000000400000004000000'),       # 46
    ('VtChartPlot', '2f000000110000000c0056744368617274506c6f74000400'),       # 47
    ('VtArray',  # 48
     '300000001200000008005674417272617900020000000100030000000000040000000000803f0000803f0000003f0000'
     '803f0000803f0000b443010001000200ffffff0002000000000000000100010000000000020000000000000000000000'
     '404000000000010000000200ffffff000200000000000000010001000000000002000000000000000000000040400100'
     '0300000070420000f0410000003f0000003f0000004000000000'),
    ('VtLight3', '3100000013000000090056744c6967687433000100'),       # 49
    ('VtArray', '32000000120000000100010003000000010004000000'),       # 50
    ('VtInfLight3',  # 51
     '33000000140000000c005674496e664c696768743300010000000000000000000000803f0000803f040000009a99193e'
     '0000003fffff04000000'),
    ('VtAxis',  # 52
     '340000001500000007005674417869730003000000ffff0a000000010001000000000002000000000000000000000040'
     '4001000100000000000200000000000000000000004040010001000000000002000000000000000000000040400000f0'
     '4200007042'),
    ('VtTextBlock', '350000000d00000001000000020002000000ffff'),       # 53
    ('VtBackdrop',  # 54
     '360000000800000001000100010002000000000002000000000000000000c8420000c842000001000200000000000200'
     'ffffff000000a0410000'),
    ('VtFill', '3700000009000000010001000200ffffff00020000000000000001000200ffffff00020080c0ff000000'),       # 55
    ('VtPicture', '380000000a0000000100ffffffffffff0400000000000400000004000000'),       # 56
    ('VtFont',  # 57
     '390000000e00000025000000000004410000000002000000000004000000ffff64000000000000000000000000000000'
     'bc0000000100'),
    ('VtString', '3a000000050000000000060000000400000000000000bc000000000000000000000000000000bc000000000004000000'),       # 58
    ('VtArray', '3b000000120000000100010003000000010004000000'),       # 59
    ('VtAxisScaleBlock', '3c0000001600000011005674417869735363616c65426c6f636b000100'),       # 60
    ('VtArray',  # 61
     '3d000000120000000500010003000000000004000000ffffffffffffffffffffffffffffffffffffffff000000001700'
     '00000d00567456616c7565426c6f636b0001001a000000ffffffffffff'),
    ('VtString', '3e00000005000000020052350006000000040000000000000d00000001000000020002000000ffff'),       # 62
    ('VtBackdrop',  # 63
     '3f0000000800000001000100010002000000000002000000000000000000c8420000c842000001000200000000000200'
     'ffffff000000a0410000'),
    ('VtFill', '4000000009000000010001000200ffffff00020000000000000001000200ffffff00020080c0ff000000'),       # 64
    ('VtPicture', '410000000a0000000100ffffffffffff0400000000000400000004000000'),       # 65
    ('VtFont',  # 66
     '420000000e00000025000000000004410000000002000000000004000000ffff640000000000000000000000d2000000'
     'bc00000001003e000000d2000000bc0000000000000000000000d2000000bc000000000004000000000000000000ffff'
     '00000000010001000100000000000000000000000000ffff00000000040001000000000000000000000004000000'),
    ('VtAxis',  # 67
     '43000000150000000000ffff0a0000000100010000000000020000000000000000000000404001000100000000000200'
     '000000000000000000004040010001000000000002000000000000000000000040400000f04200007042'),
    ('VtTextBlock', '440000000d00000001000000020002000200ffff'),       # 68
    ('VtBackdrop',  # 69
     '450000000800000001000100010002000000000002000000000000000000c8420000c842000001000200000000000200'
     'ffffff000000a0410000'),
    ('VtFill', '4600000009000000010001000200ffffff00020000000000000001000200ffffff00020080c0ff000000'),       # 70
    ('VtPicture', '470000000a0000000100ffffffffffff0400000000000400000004000000'),       # 71
    ('VtFont',  # 72
     '480000000e00000025000000000004410000000002000000000004000000ffff640000000000000000000000bc000000'
     '000000000100'),
    ('VtString', '49000000050000000000060000000400000000000000bc0000000000000000000000bc00000000000000000004000000'),       # 73
    ('VtArray', '4a000000120000000100010003000000010004000000'),       # 74
    ('VtAxisScaleBlock', '4b00000016000000'),       # 75
    ('VtArray',  # 76
     '4c000000120000000500010003000000000004000000ffffffffffffffffffffffffffffffffffffffff000001001700'
     '0000'),
    ('VtDouble', '4d0000001800000009005674446f75626c650001000000000000005940ffff0600000004000000ffffffffffff'),       # 77
    ('VtString', '4e0000000500000003003130300006000000040000000000000d00000001000000010002000000ffff'),       # 78
    ('VtBackdrop',  # 79
     '4f0000000800000001000100010002000000000002000000000000000000c8420000c842000001000200000000000200'
     'ffffff000000a0410000'),
    ('VtFill', '5000000009000000010001000200ffffff00020000000000000001000200ffffff00020080c0ff000000'),       # 80
    ('VtPicture', '510000000a0000000100ffffffffffff0400000000000400000004000000'),       # 81
    ('VtFont',  # 82
     '520000000e00000025000000000004410000000002000000000004000000ffff64000000000000000000000013010000'
     'bc00000001004e00000013010000bc000000000000000000000013010000bc000000000004000000000000000000ffff'
     '00000000010001000100000000000000000000000000ffff00000000040001000000000000000000010004000000'),
    ('VtAxis',  # 83
     '53000000150000000000ffff0a0000000100010000000000020000000000000000000000404001000100000000000200'
     '000000000000000000004040010001000000000002000000000000000000000040400000f04200007042'),
    ('VtTextBlock', '540000000d00000001000000020002000200ffff'),       # 84
    ('VtBackdrop',  # 85
     '550000000800000001000100010002000000000002000000000000000000c8420000c842000001000200000000000200'
     'ffffff000000a0410000'),
    ('VtFill', '5600000009000000010001000200ffffff00020000000000000001000200ffffff00020080c0ff000000'),       # 86
    ('VtPicture', '570000000a0000000100ffffffffffff0400000000000400000004000000'),       # 87
    ('VtFont',  # 88
     '580000000e00000025000000000004410000000002000000000004000000ffff640000000000000000000000bc000000'
     '000000000100'),
    ('VtString', '59000000050000000000060000000400000000000000bc0000000000000000000000bc00000000000000000004000000'),       # 89
    ('VtArray', '5a000000120000000100010003000000010004000000'),       # 90
    ('VtAxisScaleBlock', '5b00000016000000'),       # 91
    ('VtArray',  # 92
     '5c000000120000000500010003000000000004000000ffffffffffffffffffffffffffffffffffffffff000001001700'
     '0000'),
    ('VtDouble', '5d000000180000000000000000005940ffff0600000004000000ffffffffffff'),       # 93
    ('VtString', '5e0000000500000003003130300006000000040000000000000d00000001000000000002000000ffff'),       # 94
    ('VtBackdrop',  # 95
     '5f0000000800000001000100010002000000000002000000000000000000c8420000c842000001000200000000000200'
     'ffffff000000a0410000'),
    ('VtFill', '6000000009000000010001000200ffffff00020000000000000001000200ffffff00020080c0ff000000'),       # 96
    ('VtPicture', '610000000a0000000100ffffffffffff0400000000000400000004000000'),       # 97
    ('VtFont',  # 98
     '620000000e00000025000000000004410000000002000000000004000000ffff64000000000000000000000013010000'
     'bc00000001005e00000013010000bc000000000000000000000013010000bc000000000004000000000000000000ffff'
     '00000000010001000100000000000000000000000000ffff00000000040001000000000000000000010004000000'),
    ('VtAxis',  # 99
     '63000000150000000000ffff0a0000000100010000000000020000000000000000000000404001000100000000000200'
     '000000000000000000004040010001000000000002000000000000000000000040400000f04200000000'),
    ('VtTextBlock', '640000000d00000001000000020002000000ffff'),       # 100
    ('VtBackdrop',  # 101
     '650000000800000001000100010002000000000002000000000000000000c8420000c842000001000200000000000200'
     'ffffff000000a0410000'),
    ('VtFill', '6600000009000000010001000200ffffff00020000000000000001000200ffffff00020080c0ff000000'),       # 102
    ('VtPicture', '670000000a0000000100ffffffffffff0400000000000400000004000000'),       # 103
    ('VtFont',  # 104
     '680000000e00000025000000000004410000000002000000000004000000000064000000000000000000000000000000'
     '000000000000'),
    ('VtString', '690000000500000000000600000004000000000000000000000000000000000000000000000000000000000004000000'),       # 105
    ('VtArray',  # 106
     '6a000000120000000000010003000000000004000000000000000000ffff000000000100010001000000000000000000'
     '00000000ffff00000000040001000000000000000000020004000000ffff0000b4420000b442c3f5a83e0000803f1f85'
     '2b3f0400cdcccc3dffff'),
    ('VtSurfaceDesc',  # 107
     '6b000000190000000e0056745375726661636544657363000100010001000200ffffff00020000000000000001000100'
     '000000000200000000000000000000004040000000000000'),
    ('VtArray', '6c000000120000000000010003000000000004000000'),       # 108
    ('VtAxis',  # 109
     '6d00000015000000000000000a0000000100000000000000020000000000000000000000404001000000000000000200'
     '000000000000000000004040010000000000000002000000000000000000000040400000000000000000'),
    ('VtTextBlock', '6e0000000d00000001000000020002000000ffffffffffff'),       # 110
    ('VtFont',  # 111
     '6f0000000e00000025000000000004410000000002000000000004000000000064000000000000000000000000000000'
     '000000000000ffffffff000000000000000000000000000000000000000000000000000004000000'),
    ('VtArray',  # 112
     '700000001200000000000100030000000000040000000000000000000000000000000000000000000000000000000000'
     '00000000000000000000020000000000000000000000000004000000ffff'),
    ('VtCLineItem',  # 113
     '710000001a0000000c005674434c696e654974656d0001000000000000000000010001000200ff000000020000000000'
     '000001000900000000000200ff0000000000a04100004040ffffffff04000000'),
    ('VtCLineItem',  # 114
     '720000001a0000000000000000000000010001000200ffff0000020000000000000001000900000000000200ffff0000'
     '0000a04100004040ffffffff04000000000000000000010001000100000004000000ffffffff00000100000001000100'
     '00000f000000f000f0000100ffff0000704300007043000096440000f0430100'),
    ('VtBackdrop',  # 115
     '730000000800000001000100010002000000000002000000000000000000c8420000c842000001000200000000000200'
     'ffffff000000a0410000'),
    ('VtFill', '7400000009000000010001000200ffffff00020000000000000001000200ffffff00020080c0ff000000'),       # 116
    ('VtPicture', '750000000a0000000100ffffffffffff040000000000040000000400000004000000'),       # 117
    ('VtArray', '76000000120000000400010003000000040004000000'),       # 118
    ('VtSeries',  # 119
     '770000001b0000000900567453657269657300040001000100000002000200ff00000000002042000040400100010000'
     '0000000200ff0000800000000000004040010001000200ff0000000200000000000000ffffffff'),
    ('VtArray',  # 120
     '78000000120000000000010003000000000004000000ffff000000000000010001000000000002000000000000000000'
     '000040400c00000000000000000000000000000000000000000000000000000000000000000000000100000000000000'
     '020000000080000080bf00004040'),
    ('VtString', '790000000500000002004331000600000004000000'),       # 121
    ('VtSeriesLabel', '7a0000001c0000000e0056745365726965734c6162656c0001000d00000001000000020002000000ffffffffffff'),       # 122
    ('VtFont',  # 123
     '7b0000000e00000025000000000040410000000002000000000004000000000064000000000000000000000000000000'
     '00000000000079000000000000000000000000000000000000000000000000000000000004000000'),
    ('VtTextBlock', '7c0000000d00000001000000000002000000ffffffffffff'),       # 124
    ('VtFont',  # 125
     '7d0000000e00000025000000000020410000000002000000000004000000000064000000000000000000000000000000'
     '000000000000790000000000000000000000000000000000000000000000000000000000040000000100'),
    ('VtTextFormat', '7e0000001d0000000d00567454657874466f726d6174000100040000000000'),       # 126
    ('VtString', '7f0000000500000000000600000004000000'),       # 127
    ('VtTextFormat', '800000001d000000040000000000'),       # 128
    ('VtString', '810000000500000000000600000004000000'),       # 129
    ('VtSeriesLabel', '820000001c0000000d00000001000000020002000000ffffffffffff'),       # 130
    ('VtFont',  # 131
     '830000000e00000025000000000004410000000002000000000004000000000064000000000000000000000000000000'
     '000000000000ffffffff0000000000000000000000000000000000000000000000000000040000000000ffff02000100'
     '0100000000000200ff00000000002042000040400000704300000200ff000000'),
    ('VtPicture',  # 132
     '840000000a00000001000000ffffffff040000000100000004000000803f0200ff0000800200ff00008000001b000000'
     '0100050006000700080009000a0000000100ffffffffffffffff1800ffff0b00ffffffffffffffffffffffffffffffff'
     'ffffffff0500020002000200060008000000000000000000000004000000'),
    ('VtSeries',  # 133
     '850000001b0000000100010000000200020000ff000000002042000040400100010000000000020000ff008000000000'
     '0000404001000100020000ff00000200000000000000ffffffff'),
    ('VtArray',  # 134
     '86000000120000000000010003000000000004000000ffff000000000000010001000000000002000000000000000000'
     '000040400c00000000000000010001000100000000000000000000000000000000000000000000000100000000000000'
     '020000000080000080bf00004040'),
    ('VtString', '870000000500000002004332000600000004000000'),       # 135
    ('VtSeriesLabel', '880000001c0000000d00000001000000020002000000ffffffffffff'),       # 136
    ('VtFont',  # 137
     '890000000e00000025000000000040410000000002000000000004000000000064000000000000000000000000000000'
     '00000000000087000000000000000000000000000000000000000000000000000000000004000000'),
    ('VtTextBlock', '8a0000000d00000001000000000002000000ffffffffffff'),       # 138
    ('VtFont',  # 139
     '8b0000000e00000025000000000020410000000002000000000004000000000064000000000000000000000000000000'
     '000000000000870000000000000000000000000000000000000000000000000000000000040000000100'),
    ('VtTextFormat', '8c0000001d000000040000000000'),       # 140
    ('VtString', '8d0000000500000000000600000004000000'),       # 141
    ('VtTextFormat', '8e0000001d000000040000000000'),       # 142
    ('VtString', '8f0000000500000000000600000004000000'),       # 143
    ('VtSeriesLabel', '900000001c0000000d00000001000000020002000000ffffffffffff'),       # 144
    ('VtFont',  # 145
     '910000000e00000025000000000004410000000002000000000004000000000064000000000000000000000000000000'
     '000000000000ffffffff0000000000000000000000000000000000000000000000000000040000000000ffff02000100'
     '010000000000020000ff00000000204200004040000070430000020000ff0000'),
    ('VtPicture',  # 146
     '920000000a00000001000000ffffffff040000000200000004000000803f020000ff0080020000ff008001001b000000'
     '0100050006000700080009000a0000000100ffffffffffffffff1800ffff0b00ffffffffffffffffffffffffffffffff'
     'ffffffff0500020002000200060008000000000000000000000004000000'),
    ('VtSeries',  # 147
     '930000001b000000010001000000020002000000ff000000204200004040010001000000000002000000ff8000000000'
     '000040400100010002000000ff000200000000000000ffffffff'),
    ('VtArray',  # 148
     '94000000120000000000010003000000000004000000ffff000000000000010001000000000002000000000000000000'
     '000040400c00000000000000020002000200000000000000000000000000000000000000000000000100000000000000'
     '020000000080000080bf00004040'),
    ('VtString', '950000000500000002004333000600000004000000'),       # 149
    ('VtSeriesLabel', '960000001c0000000d00000001000000020002000000ffffffffffff'),       # 150
    ('VtFont',  # 151
     '970000000e00000025000000000040410000000002000000000004000000000064000000000000000000000000000000'
     '00000000000095000000000000000000000000000000000000000000000000000000000004000000'),
    ('VtTextBlock', '980000000d00000001000000000002000000ffffffffffff'),       # 152
    ('VtFont',  # 153
     '990000000e00000025000000000020410000000002000000000004000000000064000000000000000000000000000000'
     '000000000000950000000000000000000000000000000000000000000000000000000000040000000100'),
    ('VtTextFormat', '9a0000001d000000040000000000'),       # 154
    ('VtString', '9b0000000500000000000600000004000000'),       # 155
    ('VtTextFormat', '9c0000001d000000040000000000'),       # 156
    ('VtString', '9d0000000500000000000600000004000000'),       # 157
    ('VtSeriesLabel', '9e0000001c0000000d00000001000000020002000000ffffffffffff'),       # 158
    ('VtFont',  # 159
     '9f0000000e00000025000000000004410000000002000000000004000000000064000000000000000000000000000000'
     '000000000000ffffffff0000000000000000000000000000000000000000000000000000040000000000ffff02000100'
     '01000000000002000000ff00000020420000404000007043000002000000ff00'),
    ('VtPicture',  # 160
     'a00000000a00000001000000ffffffff040000000300000004000000803f02000000ff8002000000ff8002001b000000'
     '0100050006000700080009000a0000000100ffffffffffffffff1800ffff0b00ffffffffffffffffffffffffffffffff'
     'ffffffff0500020002000200060008000000000000000000000004000000'),
    ('VtSeries',  # 161
     'a10000001b00000001000100000002000200ffff0000000020420000404001000100000000000200ffff008000000000'
     '00004040010001000200ffff00000200000000000000ffffffff'),
    ('VtArray',  # 162
     'a2000000120000000000010003000000000004000000ffff000000000000010001000000000002000000000000000000'
     '000040400c00000000000000030003000300000000000000000000000000000000000000000000000100000000000000'
     '020000000080000080bf00004040'),
    ('VtString', 'a30000000500000002004334000600000004000000'),       # 163
    ('VtSeriesLabel', 'a40000001c0000000d00000001000000020002000000ffffffffffff'),       # 164
    ('VtFont',  # 165
     'a50000000e00000025000000000040410000000002000000000004000000000064000000000000000000000000000000'
     '000000000000a3000000000000000000000000000000000000000000000000000000000004000000'),
    ('VtTextBlock', 'a60000000d00000001000000000002000000ffffffffffff'),       # 166
    ('VtFont',  # 167
     'a70000000e00000025000000000020410000000002000000000004000000000064000000000000000000000000000000'
     '000000000000a30000000000000000000000000000000000000000000000000000000000040000000100'),
    ('VtTextFormat', 'a80000001d000000040000000000'),       # 168
    ('VtString', 'a90000000500000000000600000004000000'),       # 169
    ('VtTextFormat', 'aa0000001d000000040000000000'),       # 170
    ('VtString', 'ab0000000500000000000600000004000000'),       # 171
    ('VtSeriesLabel', 'ac0000001c0000000d00000001000000020002000000ffffffffffff'),       # 172
    ('VtFont',  # 173
     'ad0000000e00000025000000000004410000000002000000000004000000000064000000000000000000000000000000'
     '000000000000ffffffff0000000000000000000000000000000000000000000000000000040000000000ffff02000100'
     '0100000000000200ffff000000002042000040400000704300000200ffff0000'),
    ('VtPicture',  # 174
     'ae0000000a00000001000000ffffffff040000000400000004000000803f0200ffff00800200ffff008003001b000000'
     '0100050006000700080009000a0000000100ffffffffffffffff1800ffff0b00ffffffffffffffffffffffffffffffff'
     'ffffffff0500020002000200060008000000000000000000000004000000'),
    ('VtChartTitle', 'af0000001e0000000d00567443686172745469746c650001000c000000'),       # 175
    ('VtTextBlock', 'b00000000d00000001000000020002000000ffffffffffff'),       # 176
    ('VtFont',  # 177
     'b10000000e00000025000000000004410000000002000000000004000000000064000000000000000000000000000000'
     '000000000000'),
    ('VtString',  # 178
     'b20000000500000000000600000004000000000000000000000000000000000000000000000000000000000004000000'
     '0f000000f000640001000000000000000000000000000000000000000100'),
    ('VtBackdrop',  # 179
     'b30000000800000001000100010002000000000002000000000000000000c8420000c842000001000200000000000200'
     'ffffff000000a0410000'),
    ('VtFill', 'b400000009000000010001000200ffffff00020000000000000001000200ffffff00020080c0ff000000'),       # 180
    ('VtPicture', 'b50000000a0000000100ffffffffffff040000000000040000000400000004000000'),       # 181
    ('VtList',  # 182
     'b60000001f000000070056744c697374000200030000000000040000000000010000000000ffffffffffffffffffffff'
     'ffffff010000000000200000000900567457696e646f77000300040000000000ffff00000000'),
)

# our values in the grid of a new chart: 5 rows (R1 - R5) x 4 columns (C1 - C4), whole numbers 0 - 99
DEFAULT_DATA = ((30, 50, 20, 40), (45, 25, 60, 35), (55, 40, 35, 70), (40, 65, 50, 45), (70, 35, 45, 60))
DATA_OBJECTS = (7, 12, 17, 22, 27)          # the first value of each row (VtInteger); the row label before it
AREA_OBJECT, PLOT_OBJECT = 30, 114


def _extent(twips):
    return int(round(twips * 2540 / 1440))


def _value_at(k, b):
    """the place of the value in a VtInteger object: after its id and class (and the class, if new)"""
    if b[4:8] == b'\x07\0\0\0' and b[8:10] == b'\x0a\0':     # the first VtInteger: the class follows
        return 8 + 2 + 10 + 2
    return 8


def objects():
    """the objects of the default chart as bytearrays"""
    return [bytearray(bytes.fromhex(h)) for _, h in CHART]


def new_blob(width=4095, height=3015, left=0, top=0, data=DEFAULT_DATA):
    """the OleObjectBlob of a new MSChart (the default chart: 2D bar, 5 x 4) of this size"""
    obs = objects()
    for r, first in enumerate(DATA_OBJECTS):
        for c in range(4):
            b = obs[first + c]
            at = _value_at(first + c, b)
            b[at:at + 4] = struct.pack('<i', int(data[r][c]))
    a = obs[AREA_OBJECT]
    a[-8:] = struct.pack('<ff', width - 15, height - 15)
    p = obs[PLOT_OBJECT]
    p[-18:-2] = struct.pack('<ffff', 240, 240, width - 255, height - 255)
    body = struct.pack('<IIII', SIG, 8, _extent(width), _extent(height)) + CONST + b''.join(obs)
    return MAGIC + struct.pack('<IIIII', len(body), int(left), int(top), int(width), int(height)) + body


def is_chart_blob(data):
    data = bytes(data or b'')
    return len(data) >= HEAD + 16 and data[:4] == MAGIC and struct.unpack_from('<I', data, HEAD)[0] == SIG


def geometry(data):
    """(Left, Top, Width, Height) of the head of the blob, in twips"""
    return struct.unpack_from('<iiii', bytes(data), 8)


def _ends():
    """where objects 30 and 114 end in a blob of the default chart"""
    at, ends = HEAD + 16 + len(CONST), {}
    for k, (_, h) in enumerate(CHART):
        at += len(h) // 2
        ends[k] = at
    return at, ends[AREA_OBJECT], ends[PLOT_OBJECT]


def set_geometry(data, left, top, width, height):
    """the blob with this place and size: its head and extents and, in a chart of the default layout
    (5 x 4, the size of a new one), the area and the plot as new_blob() puts them; the rest as it is (MSChart
    works out its own layout when it draws). A blob that is not an MSChart one: a new one."""
    if not is_chart_blob(data):
        return new_blob(width, height, left, top)
    b = bytearray(data)
    struct.pack_into('<iiii', b, 8, int(left), int(top), int(width), int(height))
    struct.pack_into('<II', b, HEAD + 8, _extent(width), _extent(height))
    size, area, plot = _ends()
    if len(b) == size and b[plot - 18:plot - 10] == struct.pack('<ff', 240, 240):
        struct.pack_into('<ff', b, area - 8, width - 15, height - 15)
        struct.pack_into('<ffff', b, plot - 18, 240, 240, width - 255, height - 255)
    return bytes(b)



# ---------------------------------------------------------------- reading a chart (phase 9.2)
# The stream after CONST is the chart as objects. Each class writes its own data and calls its base class, which
# writes [u32 class id] (with the name if new) and its data; members that are objects are written as object
# references. Checked against 520 charts that VB6 wrote (DEV-9 mc1 - mc4: the .frx of the designer and the
# PropertyBag of a running chart, one setting changed in each): the reader comes to the end of every one.
# decode() gives the settings as the runtime keeps them: the properties of the control by name (ChartType ...),
# the data grid ("#cell", "#rowlabel" ...) and the parts by path ("plot.axis(1).valuescale.maximum").

class ChartError(Exception):
    pass


class _Obj:
    __slots__ = ('cls', 'f')

    def __init__(self, cls):
        self.cls, self.f = cls, {}


class _Reader:
    def __init__(self, data, pos):
        self.d, self.p, self.objs, self.classes = data, pos, [], []

    def raw(self, n):
        if self.p + n > len(self.d):
            raise ChartError('the chart ends at %d' % self.p)
        v = self.d[self.p:self.p + n]
        self.p += n
        return v

    def u16(self):
        return struct.unpack('<H', self.raw(2))[0]

    def u32(self):
        return struct.unpack('<I', self.raw(4))[0]

    def cstr(self):
        n = self.u16()
        if not n:
            return ''
        v = self.raw(n).decode('cp1252', 'replace')
        if self.raw(1) != b'\0':
            raise ChartError('a string without its end at %d' % self.p)
        return v

    def classdef(self, cid):
        if cid < len(self.classes):
            return self.classes[cid]
        if cid != len(self.classes):
            raise ChartError('class %d at %d' % (cid, self.p))
        n = self.u16()
        name = self.raw(n)[:-1].decode('ascii', 'replace')
        self.u16()                              # the schema
        self.classes.append(name)
        return name

    def obj(self, want=None):
        oid = self.u32()
        if oid == 0xFFFFFFFF:
            return None
        if oid < len(self.objs):
            return self.objs[oid]
        if oid != len(self.objs):
            raise ChartError('object %d at %d' % (oid, self.p - 4))
        o = _Obj(self.classdef(self.u32()))
        self.objs.append(o)
        if want and o.cls not in want:
            raise ChartError('%s where %s' % (o.cls, want))
        fn = _CLASSES.get(o.cls)
        if fn is None:
            raise ChartError('the class %s' % o.cls)
        fn(self, o)
        return o

    def base(self, o, name):
        n = self.classdef(self.u32())
        if n != name:
            raise ChartError('%s where %s' % (n, name))
        _CLASSES[name](self, o)


def _r_value(r, o):
    r.base(o, 'VtObject')


def _r_string(r, o):
    o.f['s'] = r.cstr()
    r.base(o, 'VtValue')


def _r_integer(r, o):
    o.f['v'] = struct.unpack('<i', r.raw(4))[0]
    r.base(o, 'VtValue')


def _r_double(r, o):
    o.f['v'] = struct.unpack('<d', r.raw(8))[0]
    r.raw(2)
    r.base(o, 'VtValue')


def _r_collection(r, o):
    r.u16()                                     # the objects that are not none
    r.base(o, 'VtObject')


def _r_matrix(r, o):
    r.base(o, 'VtCollection')
    rows, cols = r.u16(), r.u16()
    o.f['dims'] = (rows, cols)
    o.f['cells'] = [r.obj() for _ in range(rows * cols)]


def _r_datagrid(r, o):
    r.base(o, 'VtMatrix')
    o.f['labels'] = (r.u16(), r.u16())          # column label levels, row label levels
    o.f['size'] = (r.u16(), r.u16())            # columns, rows


def _r_array(r, o):
    n = r.u16()
    r.u16()
    r.base(o, 'VtCollection')
    o.f['items'] = [r.obj() for _ in range(n)]


def _r_backdrop(r, o):
    o.f['raw'] = r.raw(50)
    o.f['fill'] = r.obj(['VtFill'])
    r.base(o, 'VtObject')


def _r_fill(r, o):
    o.f['raw'] = r.raw(34)
    o.f['pic'] = r.obj(['VtPicture'])
    o.f['style'] = r.raw(2)
    r.base(o, 'VtObject')


def _r_picture(r, o):
    r.raw(4)
    r.obj()                                     # the picture itself
    r.base(o, 'VtObject')


def _r_text(r, o):
    o.f['tb'] = r.obj(['VtTextBlock'])
    r.base(o, 'VtChartSection')


def _r_section(r, o):
    o.f['loc'] = r.raw(26)
    o.f['backdrop'] = r.obj(['VtBackdrop'])
    r.base(o, 'VtObject')


def _r_textblock(r, o):
    o.f['layout'] = r.raw(12)
    o.f['backdrop'] = r.obj(['VtBackdrop'])
    o.f['font'] = r.obj(['VtFont'])
    r.raw(26)                                   # where it was drawn
    o.f['text'] = r.obj(['VtString'])
    r.raw(30)


def _r_font(r, o):
    o.f['name'] = r.obj(['VtString'])
    o.f['raw'] = r.raw(16)


def _r_legend(r, o):
    o.f['font'] = r.obj(['VtFont'])
    o.f['layout'] = r.raw(12)
    r.base(o, 'VtChartSection')


def _r_plot(r, o):
    r.obj(['VtArray'])
    o.f['raw'] = r.raw(136)
    o.f['light'] = r.obj(['VtLight3'])
    o.f['axes'] = [r.obj(['VtAxis']) for _ in range(4)]
    o.f['raw2'] = r.raw(30)
    r.obj(['VtSurfaceDesc'])
    r.obj(['VtCLineItem'])
    r.obj(['VtCLineItem'])
    o.f['raw3'] = r.raw(34)
    r.base(o, 'VtChartSection')
    o.f['series'] = r.obj(['VtArray'])


def _r_light(r, o):
    o.f['sources'] = r.obj(['VtArray'])
    o.f['raw'] = r.raw(10)
    r.base(o, 'VtObject')


def _r_inflight(r, o):
    o.f['raw'] = r.raw(16)
    r.base(o, 'VtObject')


def _r_axis(r, o):
    o.f['h'] = r.raw(8)
    o.f['pens'] = r.raw(66)
    o.f['tick'] = r.raw(8)
    o.f['title'] = r.obj(['VtTextBlock'])
    r.obj(['VtArray'])                          # the labels it drew
    o.f['t'] = r.raw(34)
    o.f['scale'] = r.raw(24) if o.f['h'][6] & 1 else None
    o.f['t2'] = r.raw(20)


def _r_scaleblock(r, o):
    r.obj(['VtArray'])
    r.raw(4)
    r.base(o, 'VtValueBlock')


def _r_valueblock(r, o):
    r.obj()
    r.obj()
    r.raw(2)
    r.obj(['VtString'])
    r.raw(3)
    r.base(o, 'VtTextBlock')


def _r_surface(r, o):
    r.raw(46)
    r.obj(['VtArray'])
    r.obj(['VtAxis'])
    r.raw(2)


def _r_lineitem(r, o):
    r.raw(48)
    r.obj()
    r.base(o, 'VtObject')


def _r_series(r, o):
    o.f['raw'] = r.raw(62)
    r.obj()
    r.obj(['VtArray'])                          # its data points
    o.f['raw2'] = r.raw(88)
    o.f['legend'] = r.obj(['VtString'])
    r.obj(['VtSeriesLabel'])
    o.f['dpl'] = r.obj(['VtTextBlock'])         # DataPoints(-1).DataPointLabel and the LegendText
    o.f['component'] = r.raw(2)
    o.f['fmt1'] = r.obj(['VtTextFormat'])
    o.f['fmt2'] = r.obj(['VtTextFormat'])
    o.f['label2'] = r.obj(['VtSeriesLabel'])
    o.f['raw3'] = r.raw(40)
    r.obj(['VtPicture'])
    o.f['tail'] = r.raw(106)


def _r_serieslabel(r, o):
    r.base(o, 'VtTextBlock')


def _r_textformat(r, o):
    r.raw(6)
    o.f['s'] = r.obj(['VtString'])


def _r_list(r, o):
    o.f['raw'] = r.raw(38)
    r.base(o, 'VtWindow')


def _r_window(r, o):
    o.f['wraw'] = r.raw(12)


def _r_chart(r, o):
    o.f['grid'] = r.obj(['VtDataGrid'])
    r.raw(18)                                   # the area of the chart
    o.f['backdrop'] = r.obj(['VtBackdrop'])
    o.f['footnote'] = r.obj(['VtChartFootnote'])
    o.f['legend'] = r.obj(['VtChartLegend'])
    o.f['plot'] = r.obj(['VtChartPlot'])
    o.f['title'] = r.obj(['VtChartTitle'])
    o.f['list'] = r.obj(['VtList'])


_CLASSES = {
    'VtObject': lambda r, o: None, 'VtValue': _r_value, 'VtString': _r_string, 'VtInteger': _r_integer,
    'VtDouble': _r_double, 'VtCollection': _r_collection, 'VtMatrix': _r_matrix, 'VtDataGrid': _r_datagrid,
    'VtArray': _r_array, 'VtBackdrop': _r_backdrop, 'VtFill': _r_fill, 'VtPicture': _r_picture,
    'VtChartFootnote': lambda r, o: r.base(o, 'VtChartText'), 'VtChartTitle': lambda r, o: r.base(o, 'VtChartText'),
    'VtChartText': _r_text, 'VtChartSection': _r_section,
    'VtTextBlock': _r_textblock, 'VtFont': _r_font, 'VtChartLegend': _r_legend, 'VtChartPlot': _r_plot,
    'VtLight3': _r_light, 'VtInfLight3': _r_inflight, 'VtAxis': _r_axis, 'VtAxisScaleBlock': _r_scaleblock,
    'VtValueBlock': _r_valueblock, 'VtSurfaceDesc': _r_surface, 'VtCLineItem': _r_lineitem, 'VtSeries': _r_series,
    'VtSeriesLabel': _r_serieslabel, 'VtTextFormat': _r_textformat, 'VtList': _r_list, 'VtWindow': _r_window,
    'VtChart': _r_chart,
}


def read(data):
    """the chart of an OleObjectBlob: (the head of the control (34 bytes), the root object); ChartError"""
    data = bytes(data)
    if not is_chart_blob(data):
        raise ChartError('not the OleObjectBlob of an MSChart')
    at = HEAD + 16
    r = _Reader(data, at + len(CONST))
    root = r.obj(['VtChart'])
    return data[at:at + len(CONST)], root


def _i16(b, at):
    return struct.unpack_from('<h', b, at)[0]


def _f32(b, at):
    return round(struct.unpack_from('<f', b, at)[0], 6)


def _f64(b, at):
    return struct.unpack_from('<d', b, at)[0]


def _bool(b, at):
    return _i16(b, at) != 0


class _Out:
    def __init__(self):
        self.items = []

    def put(self, key, value):
        self.items.append((key, value))

    def color(self, key, b, at):
        """a VtColor: [u16][R][G][B][flags: 0x80 Automatic]"""
        self.put(key + '.red', b[at + 2])
        self.put(key + '.green', b[at + 3])
        self.put(key + '.blue', b[at + 4])
        self.put(key + '.automatic', bool(b[at + 5] & 0x80))

    def pen(self, key, b, at):
        self.put(key + '.style', _i16(b, at + 2))
        self.put(key + '.join', _i16(b, at + 4))
        self.put(key + '.cap', _i16(b, at + 6))
        self.color(key + '.vtcolor', b, at + 8)
        self.put(key + '.width', _f32(b, at + 14))
        self.put(key + '.limit', _f32(b, at + 18))

    def brush(self, key, b, at):
        self.put(key + '.style', _i16(b, at))
        self.color(key + '.fillcolor', b, at + 2)
        self.color(key + '.patterncolor', b, at + 8)
        self.put(key + '.index', _i16(b, at + 14))

    def backdrop(self, key, bd):
        if bd is None:
            return
        b = bd.f['raw']
        self.brush(key + '.shadow.brush', b, 4)
        self.put(key + '.shadow.offset.x', _f32(b, 20))
        self.put(key + '.shadow.offset.y', _f32(b, 24))
        self.put(key + '.shadow.style', _i16(b, 28))
        self.color(key + '.frame.framecolor', b, 32)
        self.color(key + '.frame.spacecolor', b, 38)
        self.put(key + '.frame.width', _f32(b, 44))
        self.put(key + '.frame.style', _i16(b, 48))
        fill = bd.f['fill']
        self.put(key + '.fill.style', _i16(fill.f['style'], 0))
        self.brush(key + '.fill.brush', fill.f['raw'], 2)

    def font(self, key, fo, name='MS Sans Serif'):
        if fo is None:
            return
        n = fo.f['name']
        self.put(key + '.name', n.f['s'] if n is not None else name)
        b = fo.f['raw']
        self.put(key + '.size', _f32(b, 0))
        self.put(key + '.style', _i16(b, 4))
        self.put(key + '.effect', _i16(b, 6))
        self.color(key + '.vtcolor', b, 8)

    def layout(self, key, b, at):
        self.put(key + '.wordwrap', _bool(b, at + 2))
        self.put(key + '.horzalignment', _i16(b, at + 4))
        self.put(key + '.vertalignment', _i16(b, at + 6))
        self.put(key + '.orientation', _i16(b, at + 8))

    def location(self, key, loc, visible=True):
        if visible:
            self.put(key + '.visible', _bool(loc, 6))
        for i, k in enumerate(('min.x', 'min.y', 'max.x', 'max.y')):
            self.put(key + '.rect.' + k, _f32(loc, 8 + 4 * i))
        self.put(key + '.locationtype', _i16(loc, 24))

    def textblock(self, key, tb):
        self.layout(key + '.textlayout', tb.f['layout'], 0)
        self.backdrop(key + '.backdrop', tb.f['backdrop'])
        self.font(key + '.vtfont', tb.f['font'])


def decode(data):
    """the settings of a chart in its OleObjectBlob: [(key, value)], value an int, a float, a bool or a str.
    Keys without a dot are properties of the control (ChartType ...), "#..." the data grid, the others the
    parts by path in lower case (title.vtfont.size, plot.axis(1).valuescale.maximum ...). ChartError if it
    cannot be read."""
    head, ch = read(data)
    out = _Out()
    put = out.put
    # the control
    put('AutoIncrement', _bool(head, 8))
    put('Repaint', _bool(head, 14))
    lst = ch.f['list'].f
    lr, wr = lst['raw'], lst['wraw']
    put('ChartType', _i16(lr, 32))
    put('AllowDynamicRotation', _bool(lr, 22))
    put('AllowSelections', _bool(lr, 28))
    put('AllowSeriesSelection', _bool(lr, 30))
    put('TextLengthType', _i16(lr, 34))
    put('DrawMode', _i16(lr, 36))
    put('AllowDithering', _bool(wr, 4))
    put('DoSetCursor', _bool(wr, 6))
    # the data grid: rows of the matrix = column label levels + rows, its columns = row label levels + columns
    g = ch.f['grid'].f
    clc, rlc = g['labels']
    ncol, nrow = g['size']
    R, C = g['dims']
    put('#size', '%d,%d,%d,%d' % (rlc, clc, nrow, ncol))
    for i in range(R):
        for j in range(C):
            o = g['cells'][i * C + j]
            if i < clc and j >= rlc:            # the top row: the highest level
                if o is not None:
                    put('#collabel(%d,%d)' % (clc - i, j - rlc + 1), o.f.get('s', ''))
            elif j < rlc and i >= clc:
                if o is not None:
                    put('#rowlabel(%d,%d)' % (rlc - j, i - clc + 1), o.f.get('s', ''))
            elif i >= clc and j >= rlc:
                key = '#data(%d,%d)' % (i - clc + 1, j - rlc + 1)
                if o is None or 'v' not in o.f:
                    put(key, '')
                else:
                    put(key, float(o.f['v']))
    # the backdrop of the chart, the title, the footnote, the legend
    out.backdrop('backdrop', ch.f['backdrop'])
    for name in ('title', 'footnote'):
        t = ch.f[name].f
        tb = t['tb']
        put(name + '.text', tb.f['text'].f['s'] if tb.f['text'] is not None else '')
        out.textblock(name, tb)
        out.location(name + '.location', t['loc'])
        out.backdrop(name + '.backdrop', t['backdrop'])
    lg = ch.f['legend'].f
    put('ShowLegend', _bool(lg['loc'], 6))
    out.layout('legend.textlayout', lg['layout'], 2)
    out.font('legend.vtfont', lg['font'])
    out.location('legend.location', lg['loc'], visible=False)
    out.backdrop('legend.backdrop', lg['backdrop'])
    # the plot
    pl = ch.f['plot'].f
    p = pl['raw']
    put('plot.widthtoheightratio', _f32(p, 0))
    put('plot.depthtoheightratio', _f32(p, 4))
    put('plot.bargap', _f32(p, 8))
    put('plot.xgap', _f32(p, 12))
    put('plot.zgap', _f32(p, 16))
    put('plot.plotbase.baseheight', _f32(p, 20))
    out.brush('plot.plotbase.brush', p, 26)
    out.pen('plot.plotbase.pen', p, 42)
    put('plot.wall.width', _f32(p, 64))
    out.brush('plot.wall.brush', p, 70)
    out.pen('plot.wall.pen', p, 86)
    put('plot.view3d.rotation', _f32(p, 112))
    put('plot.view3d.elevation', _f32(p, 116))
    p2, p3 = pl['raw2'], pl['raw3']
    put('plot.clockwise', _bool(p2, 0))
    put('plot.startingangle', _f32(p2, 2))
    put('plot.uniformaxis', _bool(p3, 18))
    put('plot.projection', _i16(p3, 22))
    put('plot.subplotlabelposition', _i16(p3, 24))
    put('plot.sort', _i16(p3, 26))
    loc = pl['loc']
    put('plot.autolayout', _i16(loc, 24) != 8)
    for i, k in enumerate(('min.x', 'min.y', 'max.x', 'max.y')):
        put('plot.locationrect.' + k, _f32(loc, 8 + 4 * i))
    out.backdrop('plot.backdrop', pl['backdrop'])
    li = pl['light'].f
    put('plot.light.ambientintensity', _f32(li['raw'], 0))
    put('plot.light.edgeintensity', _f32(li['raw'], 4))
    put('plot.light.edgevisible', _bool(li['raw'], 8))
    for k, s in enumerate(li['sources'].f['items']):
        if s is not None:
            b = s.f['raw']
            base = 'plot.light.lightsources(%d)' % (k + 1)
            for i, n in enumerate(('x', 'y', 'z', 'intensity')):
                put(base + '.' + n, _f32(b, 4 * i))
    # the axes: X 0, Y 1, Y2 2, Z 3
    for a, ax in enumerate(pl['axes']):
        f = ax.f
        base = 'plot.axis(%d)' % a
        h, t, t2 = f['h'], f['t'], f['t2']
        put(base + '.axisscale.hide', _bool(h, 0))
        put(base + '.axistitle.visible', _i16(h, 2) != 1)
        put(base + '.axisscale.logbase', _i16(h, 4))
        auto = not (h[6] & 1)
        put(base + ('.valuescale.auto' if a in (1, 2) else '.categoryscale.auto'), auto)
        out.pen(base + '.pen', f['pens'], 0)
        out.pen(base + '.axisgrid.majorpen', f['pens'], 22)
        out.pen(base + '.axisgrid.minorpen', f['pens'], 44)
        put(base + '.tick.length', _f32(f['tick'], 0))
        tb = f['title']
        put(base + '.axistitle.text', tb.f['text'].f['s'] if tb.f['text'] is not None else '')
        out.textblock(base + '.axistitle', tb)
        put(base + '.intersection.auto', _i16(t, 20) == 0)
        put(base + '.intersection.point', _f64(t, 22))
        put(base + '.categoryscale.labeltick', _bool(t, 32))
        sc = f['scale']
        if sc is not None:
            if a in (1, 2):
                put(base + '.valuescale.minimum', _f64(sc, 0))
                put(base + '.valuescale.maximum', _f64(sc, 8))
                put(base + '.valuescale.majordivision', _i16(sc, 16))
                put(base + '.valuescale.minordivision', _i16(sc, 18))
            else:
                put(base + '.categoryscale.divisionsperlabel', _i16(sc, 20))
                put(base + '.categoryscale.divisionspertick', _i16(sc, 22))
        put(base + '.axisscale.type', _i16(t2, 0))
        put(base + '.axisscale.percentbasis', _i16(t2, 2))
        put(base + '.tick.style', _i16(t2, 4))
    # the series
    stacked = False
    combo = _i16(lr, 32) in (8, 9)
    for k, se in enumerate(pl['series'].f['items']):
        if se is None:
            continue
        f = se.f
        base = 'plot.seriescollection(%d)' % (k + 1)
        b, b2, b3, tl = f['raw'], f['raw2'], f['raw3'], f['tail']
        out.pen(base + '.pen', b, 0)
        out.pen(base + '.datapoints(-1).edgepen', b, 22)
        out.brush(base + '.datapoints(-1).brush', b, 46)
        put(base + '.showline', _bool(b2, 0))
        out.pen(base + '.guidelinepen', b2, 8)
        put(base + '.statline.flag', _i16(b2, 32))
        put(base + '.position.excluded', _bool(b2, 34))
        put(base + '.position.hidden', _bool(b2, 36))
        put(base + '.position.order', _i16(b2, 42))
        put(base + '.position.stackorder', _i16(b2, 44))
        stacked = stacked or _i16(b2, 44) != 0
        out.color(base + '.statline.vtcolor', b2, 74)
        put(base + '.statline.width', _f32(b2, 80))
        put(base + '.secondaryaxis', _bool(b3, 0))
        put(base + '.seriesmarker.auto', _bool(b3, 2))
        put(base + '.seriesmarker.show', _bool(b3, 32))
        put(base + '.datapoints(-1).marker.visible', _bool(b3, 32))
        out.pen(base + '.datapoints(-1).marker.pen', b3, 6)
        put(base + '.datapoints(-1).marker.size', _f32(b3, 28))
        put(base + '.datapoints(-1).marker.style', _i16(tl, 0))
        if combo:
            put(base + '.seriestype', _i16(tl, 42))
        dpl = base + '.datapoints(-1).datapointlabel'
        put(dpl + '.locationtype', _i16(tl, 98))
        put(dpl + '.linestyle', _i16(tl, 100))
        put(dpl + '.component', _i16(f['component'], 0))
        put(dpl + '.valueformat', f['fmt1'].f['s'].f['s'] if f['fmt1'] and f['fmt1'].f['s'] else '')
        put(dpl + '.percentformat', f['fmt2'].f['s'].f['s'] if f['fmt2'] and f['fmt2'].f['s'] else '')
        out.font(dpl + '.vtfont', f['label2'].f['font'] if f['label2'] else None)
        lt = f['dpl'].f['text'] if f['dpl'] is not None else None
        if lt is not None:
            put(base + '.legendtext', lt.f['s'])
    put('Stacking', stacked)
    return out.items
