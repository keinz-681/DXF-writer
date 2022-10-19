Attribute VB_Name = "DXFwriter"
Sub DXFwriter1() 
    stime
    Dim n As Integer, originpath, generatepath As String
    Dim txtLine(20) As String
    
    'set filepath
    
    ChDir DesktopFilepath(filePath)

    originpath = Application.GetOpenFilename(filefilter:="DXFdata(*.dxf;),")
    If (originpath = False) Then
        Exit Sub
    End If
    generatepath = Application.GetSaveAsFilename("output", filefilter:="dxf(*.dxf;),", Title:="生成先のファイルを指定", buttontext:="おっぱい")
    If (generatepath = originpath) Then
        MsgBox "よくないので強制終了"
        Exit Sub
        
    End If
    
    'generatepath = DesktopFilepath(filePath) & "take3.dxf"
    
    Debug.Print originpath 'check path
    Debug.Print generatepath
    
    Open originpath For Input As #1
    
    Do While Not EOF(1)
    Line Input #1, txtLine(0)
    If (txtLine(0) = "ENTITIES") Then
        
        Sheets("Sheet1").Select
        Do While Not EOF(1)
        i = i + 1
        'Debug.Print i
        Line Input #1, txtLine(0)
        Cells(i, 1).Value = txtLine(0)
        Loop
    End If
    'Debug.Print txtLine(0)
    Loop
    Close #1
    
    Open generatepath For Output As #1
        head
    Close #1
    
    Dim a As Integer
    
    Open generatepath For Append As #4
    
    Do While (1)
        a = a + 1
        txtLine(20) = Cells(a, 1).Value
        If (txtLine(20) = "_0-0_") Then
            txtLine(20) = txtLine(20) & "ORIGIN"
            'Debug.Print txtLine(20)
            
        End If
        
        If (txtLine(20) = "_0-1_") Then
            txtLine(20) = txtLine(20) & "CAM01"
            'Debug.Print txtLine(20)
        
        End If
        If (txtLine(20) = "_0-2_") Then
            txtLine(20) = txtLine(20) & "CAM02"
            'Debug.Print txtLine(20)
        
        End If

        If (txtLine(20) = "_0-3_") Then
            txtLine(20) = txtLine(20) & "CAM03"
            'Debug.Print txtLine(20)
        
        End If

        If (txtLine(20) = "_0-4_") Then
            txtLine(20) = txtLine(20) & "CAM04"
            'Debug.Print txtLine(20)
        
        End If

        If (txtLine(20) = "_0-5_") Then
            txtLine(20) = txtLine(20) & "CAM05"
            'Debug.Print txtLine(20)
        
        End If

        If (txtLine(20) = "_0-6_") Then
            txtLine(20) = txtLine(20) & "CAM06"
            'Debug.Print txtLine(20)
        
        End If

        If (txtLine(20) = "_0-7_") Then
            txtLine(20) = txtLine(20) & "CAM07"
            'Debug.Print txtLine(20)
        
        End If

        
        'Debug.Print a
        'Debug.Print txtLine(20)
        Print #4, txtLine(20)
        If (txtLine(20) = "EOF") Then
            Exit Do
        End If
    Loop
    
    Close #4
    Debug.Print "終わりでーす"
End Sub
Sub head()
    Print #1, "0"           '1
    Print #1, "SECTION"     '2
    Print #1, "2"           '3
    Print #1, "HEADER"      '4
    Print #1, "9"           '5
    Print #1, "$ACADVER"    '6
    Print #1, "1"           '7
    Print #1, "AC1009"      '8
    Print #1, "9"           '9
    Print #1, "$INSBASE"    '10
    Print #1, "10"          '11
    Print #1, "0"           '12
    Print #1, "20"          '13
    Print #1, "0"           '14
    Print #1, "30"          '15
    Print #1, "0"           '16
    Print #1, "9"           '17
    Print #1, "$EXTMIN" '18
    Print #1, "10"      '19
    Print #1, "0"       '20
    Print #1, "20"      '21
    Print #1, "0"       '22
    Print #1, "9"       '23
    Print #1, "$EXTMAX" '24
    Print #1, "10"      '25
    Print #1, "297"     '26
    Print #1, "20"      '27
    Print #1, "210"     '28
    Print #1, "9"       '29
    Print #1, "$LIMMIN" '30
    Print #1, "10"      '31
    Print #1, "0"       '32
    Print #1, "20"      '33
    Print #1, "0"       '34
    Print #1, "9"       '35
    Print #1, "$LIMMAX" '36
    Print #1, "10"      '37
    Print #1, "297"     '38
    Print #1, "20"      '39
    Print #1, "210"     '40
    Print #1, "9"       '41
    Print #1, "$LTSCALE"    '42
    Print #1, "40"  '43
    Print #1, "1"   '44
    Print #1, "0"   '45
    Print #1, "ENDSEC"  '46
    Print #1, "0"   '47
    Print #1, "SECTION" '48
    Print #1, "2"   '49
    Print #1, "TABLES"  '50
    Print #1, "0"   '51
    Print #1, "TABLE"   '52
    Print #1, "2"   '53
    Print #1, "LTYPE"   '54
    Print #1, "70"  '55
    Print #1, "9"   '56
    Print #1, "0"   '57
    Print #1, "LTYPE"   '58
    Print #1, "2"   '59
    Print #1, "CONTINUOUS"  '60
    Print #1, "70"  '61
    Print #1, "64"  '62
    Print #1, "3"   '63
    Print #1, "実線"    '64
    Print #1, "72"  '65
    Print #1, "65"  '66
    Print #1, "73"  '67
    Print #1, "0"   '68
    Print #1, "40"  '69
    Print #1, "0"   '70
    Print #1, "0"   '71
    Print #1, "LTYPE"   '72
    Print #1, "2"   '73
    Print #1, "DASHED1" '74
    Print #1, "70"  '75
    Print #1, "64"  '76
    Print #1, "3"   '77
    Print #1, "--  --  --  --  --  --  --  --  "    '78
    Print #1, "72"  '79
    Print #1, "65"  '80
    Print #1, "73"  '81
    Print #1, "2"   '82
    Print #1, "40"  '83
    Print #1, "2.5" '84
    Print #1, "49"  '85
    Print #1, "1.25"    '86
    Print #1, "49"  '87
    Print #1, "-1.25"   '88
    Print #1, "0"   '89
    Print #1, "LTYPE"   '90
    Print #1, "2"   '91
    Print #1, "DASHED2" '92
    Print #1, "70"  '93
    Print #1, "64"  '94
    Print #1, "3"   '95
    Print #1, "----    ----    ----    ----    "    '96
    Print #1, "72"  '97
    Print #1, "65"  '98
    Print #1, "73"  '99
    Print #1, "2"   '100
    Print #1, "40"  '101
    Print #1, "5"   '102
    Print #1, "49"  '103
    Print #1, "2.5" '104
    Print #1, "49"  '105
    Print #1, "-2.5"    '106
    Print #1, "0"   '107
    Print #1, "LTYPE"   '108
    Print #1, "2"   '109
    Print #1, "DASHED3" '110
    Print #1, "70"  '111
    Print #1, "64"  '112
    Print #1, "3"   '113
    Print #1, "------  ------  ------  ------  "    '114
    Print #1, "72"  '115
    Print #1, "65"  '116
    Print #1, "73"  '117
    Print #1, "2"   '118
    Print #1, "40"  '119
    Print #1, "5"   '120
    Print #1, "49"  '121
    Print #1, "3.75"    '122
    Print #1, "49"  '123
    Print #1, "-1.25"   '124
    Print #1, "0"   '125
    Print #1, "LTYPE"   '126
    Print #1, "2"   '127
    Print #1, "CENTER1" '128
    Print #1, "70"  '129
    Print #1, "64"  '130
    Print #1, "3"   '131
    Print #1, "----- - ----- - ----- - ----- - "    '132
    Print #1, "72"  '133
    Print #1, "65"  '134
    Print #1, "73"  '135
    Print #1, "4"   '136
    Print #1, "40"  '137
    Print #1, "10"  '138
    Print #1, "49"  '139
    Print #1, "6.25"    '140
    Print #1, "49"  '141
    Print #1, "-1.25"   '142
    Print #1, "49"  '143
    Print #1, "1.25"    '144
    Print #1, "49"  '145
    Print #1, "-1.25"   '146
    Print #1, "0"   '147
    Print #1, "LTYPE"   '148
    Print #1, "2"   '149
    Print #1, "CENTER2" '150
    Print #1, "70"  '151
    Print #1, "64"  '152
    Print #1, "3"   '153
    Print #1, "----------  --  ----------  --  "    '154
    Print #1, "72"  '155
    Print #1, "65"  '156
    Print #1, "73"  '157
    Print #1, "4"   '158
    Print #1, "40"  '159
    Print #1, "20"  '160
    Print #1, "49"  '161
    Print #1, "12.5"    '162
    Print #1, "49"  '163
    Print #1, "-2.5"    '164
    Print #1, "49"  '165
    Print #1, "2.5" '166
    Print #1, "49"  '167
    Print #1, "-2.5"    '168
    Print #1, "0"   '169
    Print #1, "LTYPE"   '170
    Print #1, "2"   '171
    Print #1, "PHANTOM1"    '172
    Print #1, "70"  '173
    Print #1, "64"  '174
    Print #1, "3"   '175
    Print #1, "--- - - --- - - --- - - --- - - "    '176
    Print #1, "72"  '177
    Print #1, "65"  '178
    Print #1, "73"  '179
    Print #1, "6"   '180
    Print #1, "40"  '181
    Print #1, "10"  '182
    Print #1, "49"  '183
    Print #1, "3.25"    '184
    Print #1, "49"  '185
    Print #1, "-1.25"   '186
    Print #1, "49"  '187
    Print #1, "1.25"    '188
    Print #1, "49"  '189
    Print #1, "-1.25"   '190
    Print #1, "49"  '191
    Print #1, "1.25"    '192
    Print #1, "49"  '193
    Print #1, "-1.25"   '194
    Print #1, "0"   '195
    Print #1, "LTYPE"   '196
    Print #1, "2"   '197
    Print #1, "PHANTOM2"    '198
    Print #1, "70"  '199
    Print #1, "64"  '200
    Print #1, "3"   '201
    Print #1, "--------  -  -  --------  -  -  "    '202
    Print #1, "72"  '203
    Print #1, "65"  '204
    Print #1, "73"  '205
    Print #1, "6"   '206
    Print #1, "40"  '207
    Print #1, "20"  '208
    Print #1, "49"  '209
    Print #1, "10"  '210
    Print #1, "49"  '211
    Print #1, "-2.5"    '212
    Print #1, "49"  '213
    Print #1, "1.25"    '214
    Print #1, "49"  '215
    Print #1, "-2.5"    '216
    Print #1, "49"  '217
    Print #1, "1.25"    '218
    Print #1, "49"  '219
    Print #1, "-2.5"    '220
    Print #1, "0"   '221
    Print #1, "LTYPE"   '222
    Print #1, "2"   '223
    Print #1, "DOT" '224
    Print #1, "70"  '225
    Print #1, "64"  '226
    Print #1, "3"   '227
    Print #1, "-   -   -   -   -   -   -   -   "    '228
    Print #1, "72"  '229
    Print #1, "65"  '230
    Print #1, "73"  '231
    Print #1, "2"   '232
    Print #1, "40"  '233
    Print #1, "2.5" '234
    Print #1, "49"  '235
    Print #1, "0.625"   '236
    Print #1, "49"  '237
    Print #1, "-1.875"  '238
    Print #1, "0"   '239
    Print #1, "LTYPE"   '240
    Print #1, "2"   '241
    Print #1, "DUMMY"   '242
    Print #1, "70"  '243
    Print #1, "64"  '244
    Print #1, "3"   '245
    Print #1, "ダミー"  '246
    Print #1, "72"  '247
    Print #1, "65"  '248
    Print #1, "73"  '249
    Print #1, "0"   '250
    Print #1, "40"  '251
    Print #1, "0"   '252
    Print #1, "0"   '253
    Print #1, "LTYPE"   '254
    Print #1, "2"   '255
    Print #1, "RAND1"   '256
    Print #1, "70"  '257
    Print #1, "64"  '258
    Print #1, "3"   '259
    Print #1, "ランダム線1" '260
    Print #1, "72"  '261
    Print #1, "65"  '262
    Print #1, "73"  '263
    Print #1, "4"   '264
    Print #1, "40"  '265
    Print #1, "10"  '266
    Print #1, "49"  '267
    Print #1, "1"   '268
    Print #1, "49"  '269
    Print #1, "-2"  '270
    Print #1, "49"  '271
    Print #1, "3"   '272
    Print #1, "49"  '273
    Print #1, "-4"  '274
    Print #1, "0"   '275
    Print #1, "LTYPE"   '276
    Print #1, "2"   '277
    Print #1, "RAND2"   '278
    Print #1, "70"  '279
    Print #1, "64"  '280
    Print #1, "3"   '281
    Print #1, "ランダム線2" '282
    Print #1, "72"  '283
    Print #1, "65"  '284
    Print #1, "73"  '285
    Print #1, "4"   '286
    Print #1, "40"  '287
    Print #1, "14"  '288
    Print #1, "49"  '289
    Print #1, "2"   '290
    Print #1, "49"  '291
    Print #1, "-3"  '292
    Print #1, "49"  '293
    Print #1, "4"   '294
    Print #1, "49"  '295
    Print #1, "-5"  '296
    Print #1, "0"   '297
    Print #1, "LTYPE"   '298
    Print #1, "2"   '299
    Print #1, "RAND3"   '300
    Print #1, "70"  '301
    Print #1, "64"  '302
    Print #1, "3"   '303
    Print #1, "ランダム線3" '304
    Print #1, "72"  '305
    Print #1, "65"  '306
    Print #1, "73"  '307
    Print #1, "4"   '308
    Print #1, "40"  '309
    Print #1, "18"  '310
    Print #1, "49"  '311
    Print #1, "3"   '312
    Print #1, "49"  '313
    Print #1, "-4"  '314
    Print #1, "49"  '315
    Print #1, "5"   '316
    Print #1, "49"  '317
    Print #1, "-6"  '318
    Print #1, "0"   '319
    Print #1, "LTYPE"   '320
    Print #1, "2"   '321
    Print #1, "RAND4"   '322
    Print #1, "70"  '323
    Print #1, "64"  '324
    Print #1, "3"   '325
    Print #1, "ランダム線4" '326
    Print #1, "72"  '327
    Print #1, "65"  '328
    Print #1, "73"  '329
    Print #1, "4"   '330
    Print #1, "40"  '331
    Print #1, "22"  '332
    Print #1, "49"  '333
    Print #1, "4"   '334
    Print #1, "49"  '335
    Print #1, "-5"  '336
    Print #1, "49"  '337
    Print #1, "6"   '338
    Print #1, "49"  '339
    Print #1, "-7"  '340
    Print #1, "0"   '341
    Print #1, "LTYPE"   '342
    Print #1, "2"   '343
    Print #1, "RAND5"   '344
    Print #1, "70"  '345
    Print #1, "64"  '346
    Print #1, "3"   '347
    Print #1, "ランダム線5" '348
    Print #1, "72"  '349
    Print #1, "65"  '350
    Print #1, "73"  '351
    Print #1, "4"   '352
    Print #1, "40"  '353
    Print #1, "26"  '354
    Print #1, "49"  '355
    Print #1, "5"   '356
    Print #1, "49"  '357
    Print #1, "-6"  '358
    Print #1, "49"  '359
    Print #1, "7"   '360
    Print #1, "49"  '361
    Print #1, "-8"  '362
    Print #1, "0"   '363
    Print #1, "LTYPE"   '364
    Print #1, "2"   '365
    Print #1, "LONG1"   '366
    Print #1, "70"  '367
    Print #1, "64"  '368
    Print #1, "3"   '369
    Print #1, "----------  --  ----------  --  "    '370
    Print #1, "72"  '371
    Print #1, "65"  '372
    Print #1, "73"  '373
    Print #1, "4"   '374
    Print #1, "40"  '375
    Print #1, "40"  '376
    Print #1, "49"  '377
    Print #1, "32.5"    '378
    Print #1, "49"  '379
    Print #1, "-2.5"    '380
    Print #1, "49"  '381
    Print #1, "2.5" '382
    Print #1, "49"  '383
    Print #1, "-2.5"    '384
    Print #1, "0"   '385
    Print #1, "LTYPE"   '386
    Print #1, "2"   '387
    Print #1, "LONG2"   '388
    Print #1, "70"  '389
    Print #1, "64"  '390
    Print #1, "3"   '391
    Print #1, "--------  -  -  --------  -  -  "    '392
    Print #1, "72"  '393
    Print #1, "65"  '394
    Print #1, "73"  '395
    Print #1, "6"   '396
    Print #1, "40"  '397
    Print #1, "40"  '398
    Print #1, "49"  '399
    Print #1, "30"  '400
    Print #1, "49"  '401
    Print #1, "-2.5"    '402
    Print #1, "49"  '403
    Print #1, "1.25"    '404
    Print #1, "49"  '405
    Print #1, "-2.5"    '406
    Print #1, "49"  '407
    Print #1, "1.25"    '408
    Print #1, "49"  '409
    Print #1, "-2.5"    '410
    Print #1, "0"   '411
    Print #1, "LTYPE"   '412
    Print #1, "2"   '413
    Print #1, "LONG3"   '414
    Print #1, "70"  '415
    Print #1, "64"  '416
    Print #1, "3"   '417
    Print #1, "+++++++++++++++..+++++++++++++++"    '418
    Print #1, "72"  '419
    Print #1, "65"  '420
    Print #1, "73"  '421
    Print #1, "2"   '422
    Print #1, "40"  '423
    Print #1, "40"  '424
    Print #1, "49"  '425
    Print #1, "37.5"    '426
    Print #1, "49"  '427
    Print #1, "-2.5"    '428
    Print #1, "0"   '429
    Print #1, "LTYPE"   '430
    Print #1, "2"   '431
    Print #1, "LONG4"   '432
    Print #1, "70"  '433
    Print #1, "64"  '434
    Print #1, "3"   '435
    Print #1, "+++++++++++++++..+++++++++++++++"    '436
    Print #1, "72"  '437
    Print #1, "65"  '438
    Print #1, "73"  '439
    Print #1, "2"   '440
    Print #1, "40"  '441
    Print #1, "80"  '442
    Print #1, "49"  '443
    Print #1, "75"  '444
    Print #1, "49"  '445
    Print #1, "-5"  '446
    Print #1, "0"   '447
    Print #1, "LTYPE"   '448
    Print #1, "2"   '449
    Print #1, "LONG5"   '450
    Print #1, "70"  '451
    Print #1, "64"  '452
    Print #1, "3"   '453
    Print #1, "----------  --  ----------  --  "    '454
    Print #1, "72"  '455
    Print #1, "65"  '456
    Print #1, "73"  '457
    Print #1, "4"   '458
    Print #1, "40"  '459
    Print #1, "60"  '460
    Print #1, "49"  '461
    Print #1, "37.5"    '462
    Print #1, "49"  '463
    Print #1, "-7.5"    '464
    Print #1, "49"  '465
    Print #1, "7.5" '466
    Print #1, "49"  '467
    Print #1, "-7.5"    '468
    Print #1, "0"   '469
    Print #1, "LTYPE"   '470
    Print #1, "2"   '471
    Print #1, "dashed"  '472
    Print #1, "70"  '473
    Print #1, "64"  '474
    Print #1, "3"   '475
    Print #1, "-------   -------------   ------"    '476
    Print #1, "72"  '477
    Print #1, "65"  '478
    Print #1, "73"  '479
    Print #1, "2"   '480
    Print #1, "40"  '481
    Print #1, "7.5" '482
    Print #1, "49"  '483
    Print #1, "6"   '484
    Print #1, "49"  '485
    Print #1, "-1.5"    '486
    Print #1, "0"   '487
    Print #1, "LTYPE"   '488
    Print #1, "2"   '489
    Print #1, "dashed_spaced"   '490
    Print #1, "70"  '491
    Print #1, "64"  '492
    Print #1, "3"   '493
    Print #1, "----        --------        ----"    '494
    Print #1, "72"  '495
    Print #1, "65"  '496
    Print #1, "73"  '497
    Print #1, "2"   '498
    Print #1, "40"  '499
    Print #1, "12"  '500
    Print #1, "49"  '501
    Print #1, "6"   '502
    Print #1, "49"  '503
    Print #1, "-6"  '504
    Print #1, "0"   '505
    Print #1, "LTYPE"   '506
    Print #1, "2"   '507
    Print #1, "long_dashed_dotted"  '508
    Print #1, "70"  '509
    Print #1, "64"  '510
    Print #1, "3"   '511
    Print #1, "------------- --- --------------"    '512
    Print #1, "72"  '513
    Print #1, "65"  '514
    Print #1, "73"  '515
    Print #1, "4"   '516
    Print #1, "40"  '517
    Print #1, "15.25"   '518
    Print #1, "49"  '519
    Print #1, "12"  '520
    Print #1, "49"  '521
    Print #1, "-1.5"    '522
    Print #1, "49"  '523
    Print #1, "0.25"    '524
    Print #1, "49"  '525
    Print #1, "-1.5"    '526
    Print #1, "0"   '527
    Print #1, "LTYPE"   '528
    Print #1, "2"   '529
    Print #1, "long_dashed_double-dotted"   '530
    Print #1, "70"  '531
    Print #1, "64"  '532
    Print #1, "3"   '533
    Print #1, "----------- --- --- ------------"    '534
    Print #1, "72"  '535
    Print #1, "65"  '536
    Print #1, "73"  '537
    Print #1, "6"   '538
    Print #1, "40"  '539
    Print #1, "17"  '540
    Print #1, "49"  '541
    Print #1, "12"  '542
    Print #1, "49"  '543
    Print #1, "-1.5"    '544
    Print #1, "49"      '545
    Print #1, "0.25"    '546
    Print #1, "49"      '547
    Print #1, "-1.5"    '548
    Print #1, "49"      '549
    Print #1, "0.25"    '550
    Print #1, "49"      '551
    Print #1, "-1.5"    '552
    Print #1, "0"       '553
    Print #1, "LTYPE"   '554
    Print #1, "2"       '555
    Print #1, "long_dashed_triplicate-dotted"   '556
    Print #1, "70"  '557
    Print #1, "64"  '558
    Print #1, "3"   '559
    Print #1, "--------- --- --- --- ----------"    '560
    Print #1, "72"  '561
    Print #1, "65"  '562
    Print #1, "73"  '563
    Print #1, "8"   '564
    Print #1, "40"  '565
    Print #1, "18.75"   '566
    Print #1, "49"  '567
    Print #1, "12"  '568
    Print #1, "49"  '569
    Print #1, "-1.5"    '570
    Print #1, "49"      '571
    Print #1, "0.25"    '572
    Print #1, "49"      '573
    Print #1, "-1.5"    '574
    Print #1, "49"      '575
    Print #1, "0.25"    '576
    Print #1, "49"      '577
    Print #1, "-1.5"    '578
    Print #1, "49"      '579
    Print #1, "0.25"    '580
    Print #1, "49"      '581
    Print #1, "-1.5"    '582
    Print #1, "0"       '583
    Print #1, "LTYPE"   '584
    Print #1, "2"       '585
    Print #1, "dotted"  '586
    Print #1, "70"  '587
    Print #1, "64"  '588
    Print #1, "3"   '589
    Print #1, "   -       -       -       -    "    '590
    Print #1, "72"  '591
    Print #1, "65"  '592
    Print #1, "73"  '593
    Print #1, "2"   '594
    Print #1, "40"  '595
    Print #1, "1.75"    '596
    Print #1, "49"      '597
    Print #1, "0.25"    '598
    Print #1, "49"      '599
    Print #1, "-1.5"    '600
    Print #1, "0"       '601
    Print #1, "LTYPE"   '602
    Print #1, "2"       '603
    Print #1, "chain"   '604
    Print #1, "70"  '605
    Print #1, "64"  '606
    Print #1, "3"   '607
    Print #1, "-------------- -- --------------"    '608
    Print #1, "72"  '609
    Print #1, "65"  '610
    Print #1, "73"  '611
    Print #1, "4"   '612
    Print #1, "40"  '613
    Print #1, "18.5"    '614
    Print #1, "49"      '615
    Print #1, "12"      '616
    Print #1, "49"      '617
    Print #1, "-1.5"    '618
    Print #1, "49"      '619
    Print #1, "3.5"     '620
    Print #1, "49"      '621
    Print #1, "-1.5"    '622
    Print #1, "0"       '623
    Print #1, "LTYPE"   '624
    Print #1, "2"       '625
    Print #1, "chain_double_dash"   '626
    Print #1, "70"  '627
    Print #1, "64"  '628
    Print #1, "3"   '629
    Print #1, "------------ -- -- -------------"    '630
    Print #1, "72"  '631
    Print #1, "65"  '632
    Print #1, "73"  '633
    Print #1, "6"   '634
    Print #1, "40"  '635
    Print #1, "23.5"    '636
    Print #1, "49"  '637
    Print #1, "12"  '638
    Print #1, "49"  '639
    Print #1, "-1.5"    '640
    Print #1, "49"  '641
    Print #1, "3.5" '642
    Print #1, "49"  '643
    Print #1, "-1.5"    '644
    Print #1, "49"  '645
    Print #1, "3.5" '646
    Print #1, "49"  '647
    Print #1, "-1.5"    '648
    Print #1, "0"   '649
    Print #1, "LTYPE"   '650
    Print #1, "2"   '651
    Print #1, "dashed_dotted"   '652
    Print #1, "70"  '653
    Print #1, "64"  '654
    Print #1, "3"   '655
    Print #1, "-----------     -     ----------"    '656
    Print #1, "72"  '657
    Print #1, "65"  '658
    Print #1, "73"  '659
    Print #1, "4"   '660
    Print #1, "40"  '661
    Print #1, "9.25"    '662
    Print #1, "49"  '663
    Print #1, "6"   '664
    Print #1, "49"  '665
    Print #1, "-1.5"    '666
    Print #1, "49"  '667
    Print #1, "0.25"    '668
    Print #1, "49"  '669
    Print #1, "-1.5"    '670
    Print #1, "0"   '671
    Print #1, "LTYPE"   '672
    Print #1, "2"   '673
    Print #1, "double-dashed_dotted"    '674
    Print #1, "70"  '675
    Print #1, "64"  '676
    Print #1, "3"   '677
    Print #1, "------   -----------   -   -----"    '678
    Print #1, "72"  '679
    Print #1, "65"  '680
    Print #1, "73"  '681
    Print #1, "6"   '682
    Print #1, "40"  '683
    Print #1, "16.75"   '684
    Print #1, "49"  '685
    Print #1, "6"   '686
    Print #1, "49"  '687
    Print #1, "-1.5"    '688
    Print #1, "49"  '689
    Print #1, "6"   '690
    Print #1, "49"  '691
    Print #1, "-1.5"    '692
    Print #1, "49"  '693
    Print #1, "0.25"    '694
    Print #1, "49"  '695
    Print #1, "-1.5"    '696
    Print #1, "0"   '697
    Print #1, "LTYPE"   '698
    Print #1, "2"   '699
    Print #1, "dashed_double-dotted"    '700
    Print #1, "70"  '701
    Print #1, "64"  '702
    Print #1, "3"   '703
    Print #1, "---------    -    -    ---------"    '704
    Print #1, "72"  '705
    Print #1, "65"  '706
    Print #1, "73"  '707
    Print #1, "6"   '708
    Print #1, "40"  '709
    Print #1, "11"  '710
    Print #1, "49"  '711
    Print #1, "6"   '712
    Print #1, "49"  '713
    Print #1, "-1.5"    '714
    Print #1, "49"  '715
    Print #1, "0.25"    '716
    Print #1, "49"  '717
    Print #1, "-1.5"    '718
    Print #1, "49"  '719
    Print #1, "0.25"    '720
    Print #1, "49"  '721
    Print #1, "-1.5"    '722
    Print #1, "0"   '723
    Print #1, "LTYPE"   '724
    Print #1, "2"   '725
    Print #1, "double-dashed_double-dotted" '726
    Print #1, "70"  '727
    Print #1, "64"  '728
    Print #1, "3"   '729
    Print #1, "-----   ---------   -   -   ----"    '730
    Print #1, "72"  '731
    Print #1, "65"  '732
    Print #1, "73"  '733
    Print #1, "8"   '734
    Print #1, "40"  '735
    Print #1, "18.5"    '736
    Print #1, "49"  '737
    Print #1, "6"   '738
    Print #1, "49"  '739
    Print #1, "-1.5"    '740
    Print #1, "49"  '741
    Print #1, "6"   '742
    Print #1, "49"  '743
    Print #1, "-1.5"    '744
    Print #1, "49"  '745
    Print #1, "0.25"    '746
    Print #1, "49"  '747
    Print #1, "-1.5"    '748
    Print #1, "49"  '749
    Print #1, "0.25"    '750
    Print #1, "49"  '751
    Print #1, "-1.5"    '752
    Print #1, "0"   '753
    Print #1, "LTYPE"   '754
    Print #1, "2"   '755
    Print #1, "dashed_triplicate-dotted"    '756
    Print #1, "70"  '757
    Print #1, "64"  '758
    Print #1, "3"   '759
    Print #1, "-------    -    -    -    ------"    '760
    Print #1, "72"  '761
    Print #1, "65"  '762
    Print #1, "73"  '763
    Print #1, "8"   '764
    Print #1, "40"  '765
    Print #1, "12.75"   '766
    Print #1, "49"  '767
    Print #1, "6"   '768
    Print #1, "49"  '769
    Print #1, "-1.5"    '770
    Print #1, "49"  '771
    Print #1, "0.25"    '772
    Print #1, "49"  '773
    Print #1, "-1.5"    '774
    Print #1, "49"  '775
    Print #1, "0.25"    '776
    Print #1, "49"  '777
    Print #1, "-1.5"    '778
    Print #1, "49"  '779
    Print #1, "0.25"    '780
    Print #1, "49"  '781
    Print #1, "-1.5"    '782
    Print #1, "0"   '783
    Print #1, "LTYPE"   '784
    Print #1, "2"   '785
    Print #1, "double-dashed_triplicate-dotted" '786
    Print #1, "70"  '787
    Print #1, "64"  '788
    Print #1, "3"   '789
    Print #1, "-----  ---------  -  -  -  -----"    '790
    Print #1, "72"      '791
    Print #1, "65"      '792
    Print #1, "73"      '793
    Print #1, "10"      '794
    Print #1, "40"      '795
    Print #1, "20.25"   '796
    Print #1, "49"      '797
    Print #1, "6"       '798
    Print #1, "49"      '799
    Print #1, "-1.5"    '800
    Print #1, "49"      '801
    Print #1, "6"       '802
    Print #1, "49"      '803
    Print #1, "-1.5"    '804
    Print #1, "49"      '805
    Print #1, "0.25"    '806
    Print #1, "49"      '807
    Print #1, "-1.5"    '808
    Print #1, "49"      '809
    Print #1, "0.25"    '810
    Print #1, "49"      '811
    Print #1, "-1.5"    '812
    Print #1, "49"      '813
    Print #1, "0.25"    '814
    Print #1, "49"      '815
    Print #1, "-1.5"    '816
    Print #1, "0"       '817
    Print #1, "LTYPE"   '818
    Print #1, "2"       '819
    Print #1, "undefined"   '820
    Print #1, "70"  '821
    Print #1, "64"  '822
    Print #1, "3"   '823
    Print #1, "--------------------------------"    '824
    Print #1, "72"  '825
    Print #1, "65"  '826
    Print #1, "73"  '827
    Print #1, "0"   '828
    Print #1, "40"  '829
    Print #1, "0"   '830
    Print #1, "0"   '831
    Print #1, "LTYPE"   '832
    Print #1, "2"   '833
    Print #1, "_"   '834
    Print #1, "70"  '835
    Print #1, "64"  '836
    Print #1, "3"   '837
    Print #1, "                                "    '838
    Print #1, "72"  '839
    Print #1, "65"  '840
    Print #1, "73"  '841
    Print #1, "0"   '842
    Print #1, "40"  '843
    Print #1, "0"   '844
    Print #1, "0"   '845
    Print #1, "LTYPE"   '846
    Print #1, "2"   '847
    Print #1, "__18"    '848
    Print #1, "70"  '849
    Print #1, "64"  '850
    Print #1, "3"   '851
    Print #1, "                                "    '852
    Print #1, "72"  '853
    Print #1, "65"  '854
    Print #1, "73"  '855
    Print #1, "0"   '856
    Print #1, "40"  '857
    Print #1, "0"   '858
    Print #1, "0"   '859
    Print #1, "LTYPE"   '860
    Print #1, "2"   '861
    Print #1, "__19"    '862
    Print #1, "70"  '863
    Print #1, "64"  '864
    Print #1, "3"   '865
    Print #1, "                                "    '866
    Print #1, "72"  '867
    Print #1, "65"  '868
    Print #1, "73"  '869
    Print #1, "0"   '870
    Print #1, "40"  '871
    Print #1, "0"   '872
    Print #1, "0"   '873
    Print #1, "LTYPE"   '874
    Print #1, "2"   '875
    Print #1, "__20"    '876
    Print #1, "70"  '877
    Print #1, "64"  '878
    Print #1, "3"   '879
    Print #1, "                                "    '880
    Print #1, "72"  '881
    Print #1, "65"  '882
    Print #1, "73"  '883
    Print #1, "0"   '884
    Print #1, "40"  '885
    Print #1, "0"   '886
    Print #1, "0"   '887
    Print #1, "LTYPE"   '888
    Print #1, "2"   '889
    Print #1, "__21"    '890
    Print #1, "70"  '891
    Print #1, "64"  '892
    Print #1, "3"   '893
    Print #1, "                                "    '894
    Print #1, "72"  '895
    Print #1, "65"  '896
    Print #1, "73"  '897
    Print #1, "0"   '898
    Print #1, "40"  '899
    Print #1, "0"   '900
    Print #1, "0"   '901
    Print #1, "LTYPE"   '902
    Print #1, "2"   '903
    Print #1, "__22"    '904
    Print #1, "70"  '905
    Print #1, "64"  '906
    Print #1, "3"   '907
    Print #1, "                                "    '908
    Print #1, "72"  '909
    Print #1, "65"  '910
    Print #1, "73"  '911
    Print #1, "0"   '912
    Print #1, "40"  '913
    Print #1, "0"   '914
    Print #1, "0"   '915
    Print #1, "LTYPE"   '916
    Print #1, "2"   '917
    Print #1, "__23"    '918
    Print #1, "70"  '919
    Print #1, "64"  '920
    Print #1, "3"   '921
    Print #1, "                                "    '922
    Print #1, "72"  '923
    Print #1, "65"  '924
    Print #1, "73"  '925
    Print #1, "0"   '926
    Print #1, "40"  '927
    Print #1, "0"   '928
    Print #1, "0"   '929
    Print #1, "LTYPE"   '930
    Print #1, "2"   '931
    Print #1, "__24"    '932
    Print #1, "70"  '933
    Print #1, "64"  '934
    Print #1, "3"   '935
    Print #1, "                                "    '936
    Print #1, "72"  '937
    Print #1, "65"  '938
    Print #1, "73"  '939
    Print #1, "0"   '940
    Print #1, "40"  '941
    Print #1, "0"   '942
    Print #1, "0"   '943
    Print #1, "LTYPE"   '944
    Print #1, "2"   '945
    Print #1, "__25"    '946
    Print #1, "70"  '947
    Print #1, "64"  '948
    Print #1, "3"   '949
    Print #1, "                                "    '950
    Print #1, "72"  '951
    Print #1, "65"  '952
    Print #1, "73"  '953
    Print #1, "0"   '954
    Print #1, "40"  '955
    Print #1, "0"   '956
    Print #1, "0"   '957
    Print #1, "LTYPE"   '958
    Print #1, "2"   '959
    Print #1, "__26"    '960
    Print #1, "70"  '961
    Print #1, "64"  '962
    Print #1, "3"   '963
    Print #1, "                                "    '964
    Print #1, "72"  '965
    Print #1, "65"  '966
    Print #1, "73"  '967
    Print #1, "0"   '968
    Print #1, "40"  '969
    Print #1, "0"   '970
    Print #1, "0"   '971
    Print #1, "LTYPE"   '972
    Print #1, "2"   '973
    Print #1, "__27"    '974
    Print #1, "70"  '975
    Print #1, "64"  '976
    Print #1, "3"   '977
    Print #1, "                                "    '978
    Print #1, "72"  '979
    Print #1, "65"  '980
    Print #1, "73"  '981
    Print #1, "0"   '982
    Print #1, "40"  '983
    Print #1, "0"   '984
    Print #1, "0"   '985
    Print #1, "LTYPE"   '986
    Print #1, "2"       '987
    Print #1, "__28"    '988
    Print #1, "70"      '989
    Print #1, "64"      '990
    Print #1, "3"       '991
    Print #1, "                                "    '992
    Print #1, "72"      '993
    Print #1, "65"      '994
    Print #1, "73"      '995
    Print #1, "0"       '996
    Print #1, "40"      '997
    Print #1, "0"       '998
    Print #1, "0"       '999
    Print #1, "LTYPE"   '1000
    Print #1, "2"       '1001
    Print #1, "__29"    '1002
    Print #1, "70"      '1003
    Print #1, "64"      '1004
    Print #1, "3"       '1005
    Print #1, "                                "    '1006
    Print #1, "72"      '1007
    Print #1, "65"      '1008
    Print #1, "73"      '1009
    Print #1, "0"       '1010
    Print #1, "40"      '1011
    Print #1, "0"       '1012
    Print #1, "0"       '1013
    Print #1, "LTYPE"   '1014
    Print #1, "2"       '1015
    Print #1, "__30"    '1016
    Print #1, "70"  '1017
    Print #1, "64"  '1018
    Print #1, "3"   '1019
    Print #1, "                                "    '1020
    Print #1, "72"  '1021
    Print #1, "65"  '1022
    Print #1, "73"  '1023
    Print #1, "0"   '1024
    Print #1, "40"  '1025
    Print #1, "0"   '1026
    Print #1, "0"   '1027
    Print #1, "LTYPE"   '1028
    Print #1, "2"   '1029
    Print #1, "__31"    '1030
    Print #1, "70"  '1031
    Print #1, "64"  '1032
    Print #1, "3"   '1033
    Print #1, "                                "    '1034
    Print #1, "72"  '1035
    Print #1, "65"  '1036
    Print #1, "73"  '1037
    Print #1, "0"   '1038
    Print #1, "40"  '1039
    Print #1, "0"   '1040
    Print #1, "0"   '1041
    Print #1, "LTYPE"   '1042
    Print #1, "2"   '1043
    Print #1, "__32"    '1044
    Print #1, "70"  '1045
    Print #1, "64"  '1046
    Print #1, "3"   '1047
    Print #1, "                                "    '1048
    Print #1, "72"  '1049
    Print #1, "65"  '1050
    Print #1, "73"  '1051
    Print #1, "0"   '1052
    Print #1, "40"  '1053
    Print #1, "0"   '1054
    Print #1, "0"   '1055
    Print #1, "ENDTAB"  '1056
    Print #1, "0"   '1057
    Print #1, "TABLE"   '1058
    Print #1, "2"   '1059
    Print #1, "STYLE"   '1060
    Print #1, "5"   '1061
    Print #1, "3"   '1062
    Print #1, "100" '1063
    Print #1, "AcDbSymbolTable" '1064
    Print #1, "70"  '1065
    Print #1, "1"   '1066
    Print #1, "0"   '1067
    Print #1, "STYLE"   '1068
    Print #1, "5"   '1069
    Print #1, "10"  '1070
    Print #1, "100" '1071
    Print #1, "AcDbSymbolTableRecord"   '1072
    Print #1, "100" '1073
    Print #1, "AcDbTextStyleTableRecord"    '1074
    Print #1, "2"   '1075
    Print #1, "STANDARD"    '1076
    Print #1, "70"  '1077
    Print #1, "0"   '1078
    Print #1, "40"  '1079
    Print #1, "0"   '1080
    Print #1, "41"  '1081
    Print #1, "1"   '1082
    Print #1, "50"  '1083
    Print #1, "0"   '1084
    Print #1, "71"  '1085
    Print #1, "0"   '1086
    Print #1, "42"  '1087
    Print #1, "0.2" '1088
    Print #1, "3"   '1089
    Print #1, "txt" '1090
    Print #1, "4"   '1091
    Print #1, "bigfont.shx" '1092
    Print #1, "0"   '1093
    Print #1, "STYLE"   '1094
    Print #1, "5"   '1095
    Print #1, "26"  '1096
    Print #1, "100" '1097
    Print #1, "AcDbSymbolTableRecord"   '1098
    Print #1, "100" '1099
    Print #1, "AcDbTextStyleTableRecord"    '1100
    Print #1, "2"   '1101
    Print #1, "TATEGAKI"    '1102
    Print #1, "70"  '1103
    Print #1, "68"  '1104
    Print #1, "40"  '1105
    Print #1, "0"   '1106
    Print #1, "41"  '1107
    Print #1, "1"   '1108
    Print #1, "50"  '1109
    Print #1, "0"   '1110
    Print #1, "71"  '1111
    Print #1, "0"   '1112
    Print #1, "42"  '1113
    Print #1, "1"   '1114
    Print #1, "3"   '1115
    Print #1, "txt" '1116
    Print #1, "4"   '1117
    Print #1, "bigfont.shx" '1118
    Print #1, "0"   '1119
    Print #1, "ENDTAB"  '1120
    Print #1, "0"   '1121
    Print #1, "TABLE"   '1122
    Print #1, "2"   '1123
    Print #1, "LAYER"   '1124
    Print #1, "70"  '1125
    Print #1, "8"   '1126
    Print #1, "0"   '1127
    Print #1, "LAYER"   '1128
    Print #1, "2"   '1129
    Print #1, "_0-0_ORIGIN" '1130
    Print #1, "70"  '1131
    Print #1, "64"  '1132
    Print #1, "62"  '1133
    Print #1, "7"   '1134
    Print #1, "6"   '1135
    Print #1, "CONTINUOUS"  '1136
    Print #1, "0"   '1137
    Print #1, "LAYER"   '1138
    Print #1, "2"   '1139
    Print #1, "_0-1_CAM01"  '1140
    Print #1, "70"  '1141
    Print #1, "64"  '1142
    Print #1, "62"  '1143
    Print #1, "7"   '1144
    Print #1, "6"   '1145
    Print #1, "CONTINUOUS"  '1146
    Print #1, "0"   '1147
    Print #1, "LAYER"   '1148
    Print #1, "2"   '1149
    Print #1, "_0-2_CAM02"  '1150
    Print #1, "70"  '1151
    Print #1, "64"  '1152
    Print #1, "62"  '1153
    Print #1, "7"   '1154
    Print #1, "6"   '1155
    Print #1, "CONTINUOUS"  '1156
    Print #1, "0"   '1157
    Print #1, "LAYER"   '1158
    Print #1, "2"       '1159
    Print #1, "_0-3_CAM03"  '1160
    Print #1, "70"      '1161
    Print #1, "64"      '1162
    Print #1, "62"      '1163
    Print #1, "7"       '1164
    Print #1, "6"       '1165
    Print #1, "CONTINUOUS"  '1166
    Print #1, "0"       '1167
    Print #1, "LAYER"   '1168
    Print #1, "2"       '1169
    Print #1, "_0-4_CAM04"  '1170
    Print #1, "70"      '1171
    Print #1, "64"      '1172
    Print #1, "62"      '1173
    Print #1, "7"       '1174
    Print #1, "6"       '1175
    Print #1, "CONTINUOUS"  '1176
    Print #1, "0"       '1177
    Print #1, "LAYER"   '1178
    Print #1, "2"       '1179
    Print #1, "_0-5_CAM05"  '1180
    Print #1, "70"      '1181
    Print #1, "64"      '1182
    Print #1, "62"      '1183
    Print #1, "7"       '1184
    Print #1, "6"       '1185
    Print #1, "CONTINUOUS"  '1186
    Print #1, "0"       '1187
    Print #1, "LAYER"
    Print #1, "2"
    Print #1, "_0-6_CAM06"
    Print #1, "70"
    Print #1, "64"
    Print #1, "62"
    Print #1, "7"
    Print #1, "6"
    Print #1, "CONTINUOUS"
    Print #1, "0"
    Print #1, "LAYER"
    Print #1, "2"
    Print #1, "_0-7_CAM07"
    Print #1, "70"
    Print #1, "64"
    Print #1, "62"
    Print #1, "7"
    Print #1, "6"
    Print #1, "CONTINUOUS"
    Print #1, "0"
    Print #1, "LAYER"   '1188
    Print #1, "2"       '1189
    Print #1, "_1-0_"   '1190
    Print #1, "70"      '1191
    Print #1, "64"      '1192
    Print #1, "62"      '1193
    Print #1, "7"       '1194
    Print #1, "6"       '1195
    Print #1, "CONTINUOUS"  '1196
    Print #1, "0"       '1197
    Print #1, "LAYER"   '1198
    Print #1, "2"       '1199
    Print #1, "ADD_OBJECT"    '1200
    Print #1, "70"      '1201
    Print #1, "64"      '1202
    Print #1, "62"      '1203
    Print #1, "7"       '1204
    Print #1, "6"       '1205
    Print #1, "CONTINUOUS"  '1206
    Print #1, "0"           '1207
    Print #1, "ENDTAB"      '1208
    Print #1, "0"           '1209
    Print #1, "ENDSEC"      '1210
    Print #1, "0"           '1211
    Print #1, "SECTION"     '1212
    Print #1, "2"           '1213
    Print #1, "BLOCKS"      '1214
    Print #1, "0"           '1215
    Print #1, "ENDSEC"      '1216
    Print #1, "0"           '1217
    Print #1, "SECTION"     '1218
    Print #1, "2"           '1219
    Print #1, "ENTITIES"    '1220
    'Print #1, "0"
End Sub
Sub foot()
    Print #3, "0"
    Print #3, "ENDSEC"
    Print #3, "0"
    Print #3, "EOF"
End Sub
