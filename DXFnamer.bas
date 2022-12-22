Attribute VB_Name = "DXFnamer"
Sub main()
    ptime ("start time")
    Dim n, a, i As Integer
    Dim originpath, generatepath As String
    Dim getline, txtline() As String
    ChDir DesktopFilepath
    originpath = Application.GetOpenFilename(filefilter:="DXFdata(*.dxf;),", Title:="レイヤー名を修正するファイルを選択")
    
    If (originpath = False) Then Exit Sub
    
    generatepath = Application.GetSaveAsFilename("output", filefilter:="DXFdata(*.dxf;),", Title:="生成先のファイルを指定")
    
    If (generatepath = originpath) Then Exit Sub
    
    Open originpath For Input As #1
        Do While Not EOF(1)
            Line Input #1, getline
            If (getline = "ENTITIES") Then
                Do While Not EOF(1)
                    i = i + 1
                    Line Input #1, getline
                Loop
            End If
        Loop
    Close #1
    
    Open originpath For Input As #1
        Do While Not EOF(1)
            Line Input #1, getline
            If (getline = "ENTITIES") Then
                ReDim txtline(i)
                Do While Not EOF(1)
                    n = n + 1
                    Line Input #1, txtline(n)
                Loop
            End If
        Loop
    Close #1
    
    header (generatepath)
    
    Open generatepath For Append As #4
        Do While (1)
            a = a + 1
            txtline(a) = Left(txtline(a), 5)
            If (txtline(a) = "_0-0_") Then txtline(a) = txtline(a) & "ORIGIN"
            If (txtline(a) = "_0-1_") Then txtline(a) = txtline(a) & "CAM01"
            If (txtline(a) = "_0-2_") Then txtline(a) = txtline(a) & "CAM02"
            If (txtline(a) = "_0-3_") Then txtline(a) = txtline(a) & "CAM03"
            If (txtline(a) = "_0-4_") Then txtline(a) = txtline(a) & "CAM04"
            If (txtline(a) = "_0-5_") Then txtline(a) = txtline(a) & "CAM05"
            If (txtline(a) = "_0-6_") Then txtline(a) = txtline(a) & "CAM06"
            If (txtline(a) = "_0-7_") Then txtline(a) = txtline(a) & "CAM07"
            Print #4, txtline(a)
            If (txtline(a) = "EOF") Then Exit Do
        Loop
    Close #4
    
    ptime ("Program Finished")
End Sub
Private Function header(filepath As String)
    Open filepath For Output As #1
        Print #1, "0"
        Print #1, "SECTION"
        Print #1, "2"
        Print #1, "HEADER"
        Print #1, "9"
        Print #1, "$ACADVER"
        Print #1, "1"
        Print #1, "AC1009"
        Print #1, "9"
        Print #1, "$INSBASE"
        Print #1, "10"
        Print #1, "0"
        Print #1, "20"
        Print #1, "0"
        Print #1, "30"
        Print #1, "0"
        Print #1, "9"
        Print #1, "$EXTMIN"
        Print #1, "10"
        Print #1, "0"
        Print #1, "20"
        Print #1, "0"
        Print #1, "9"
        Print #1, "$EXTMAX"
        Print #1, "10"
        Print #1, "297"
        Print #1, "20"
        Print #1, "210"
        Print #1, "9"
        Print #1, "$LIMMIN"
        Print #1, "10"
        Print #1, "0"
        Print #1, "20"
        Print #1, "0"
        Print #1, "9"
        Print #1, "$LIMMAX"
        Print #1, "10"
        Print #1, "297"
        Print #1, "20"
        Print #1, "210"
        Print #1, "9"
        Print #1, "$LTSCALE"
        Print #1, "40"
        Print #1, "1"
        Print #1, "0"
        Print #1, "ENDSEC"
        Print #1, "0"
        Print #1, "SECTION"
        Print #1, "2"
        Print #1, "TABLES"
        Print #1, "0"
        Print #1, "TABLE"
        Print #1, "2"
        Print #1, "LTYPE"
        Print #1, "70"
        Print #1, "9"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "CONTINUOUS"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "????"
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "0"
        Print #1, "40"
        Print #1, "0"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "DASHED1"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "--  --  --  --  --  --  --  --  "
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "2"
        Print #1, "40"
        Print #1, "2.5"
        Print #1, "49"
        Print #1, "1.25"
        Print #1, "49"
        Print #1, "-1.25"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "DASHED2"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "----    ----    ----    ----    "
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "2"
        Print #1, "40"
        Print #1, "5"
        Print #1, "49"
        Print #1, "2.5"
        Print #1, "49"
        Print #1, "-2.5"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "DASHED3"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "------  ------  ------  ------  "
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "2"
        Print #1, "40"
        Print #1, "5"
        Print #1, "49"
        Print #1, "3.75"
        Print #1, "49"
        Print #1, "-1.25"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "CENTER1"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "----- - ----- - ----- - ----- - "
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "4"
        Print #1, "40"
        Print #1, "10"
        Print #1, "49"
        Print #1, "6.25"
        Print #1, "49"
        Print #1, "-1.25"
        Print #1, "49"
        Print #1, "1.25"
        Print #1, "49"
        Print #1, "-1.25"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "CENTER2"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "----------  --  ----------  --  "
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "4"
        Print #1, "40"
        Print #1, "20"
        Print #1, "49"
        Print #1, "12.5"
        Print #1, "49"
        Print #1, "-2.5"
        Print #1, "49"
        Print #1, "2.5"
        Print #1, "49"
        Print #1, "-2.5"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "PHANTOM1"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "--- - - --- - - --- - - --- - - "
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "6"
        Print #1, "40"
        Print #1, "10"
        Print #1, "49"
        Print #1, "3.25"
        Print #1, "49"
        Print #1, "-1.25"
        Print #1, "49"
        Print #1, "1.25"
        Print #1, "49"
        Print #1, "-1.25"
        Print #1, "49"
        Print #1, "1.25"
        Print #1, "49"
        Print #1, "-1.25"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "PHANTOM2"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "--------  -  -  --------  -  -  "
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "6"
        Print #1, "40"
        Print #1, "20"
        Print #1, "49"
        Print #1, "10"
        Print #1, "49"
        Print #1, "-2.5"
        Print #1, "49"
        Print #1, "1.25"
        Print #1, "49"
        Print #1, "-2.5"
        Print #1, "49"
        Print #1, "1.25"
        Print #1, "49"
        Print #1, "-2.5"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "DOT"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "-   -   -   -   -   -   -   -   "
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "2"
        Print #1, "40"
        Print #1, "2.5"
        Print #1, "49"
        Print #1, "0.625"
        Print #1, "49"
        Print #1, "-1.875"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "DUMMY"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "?_?~?["
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "0"
        Print #1, "40"
        Print #1, "0"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "RAND1"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "?????_????1"
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "4"
        Print #1, "40"
        Print #1, "10"
        Print #1, "49"
        Print #1, "1"
        Print #1, "49"
        Print #1, "-2"
        Print #1, "49"
        Print #1, "3"
        Print #1, "49"
        Print #1, "-4"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "RAND2"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "?????_????2"
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "4"
        Print #1, "40"
        Print #1, "14"
        Print #1, "49"
        Print #1, "2"
        Print #1, "49"
        Print #1, "-3"
        Print #1, "49"
        Print #1, "4"
        Print #1, "49"
        Print #1, "-5"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "RAND3"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "?????_????3"
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "4"
        Print #1, "40"
        Print #1, "18"
        Print #1, "49"
        Print #1, "3"
        Print #1, "49"
        Print #1, "-4"
        Print #1, "49"
        Print #1, "5"
        Print #1, "49"
        Print #1, "-6"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "RAND4"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "?????_????4"
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "4"
        Print #1, "40"
        Print #1, "22"
        Print #1, "49"
        Print #1, "4"
        Print #1, "49"
        Print #1, "-5"
        Print #1, "49"
        Print #1, "6"
        Print #1, "49"
        Print #1, "-7"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "RAND5"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "?????_????5"
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "4"
        Print #1, "40"
        Print #1, "26"
        Print #1, "49"
        Print #1, "5"
        Print #1, "49"
        Print #1, "-6"
        Print #1, "49"
        Print #1, "7"
        Print #1, "49"
        Print #1, "-8"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "LONG1"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "----------  --  ----------  --  "
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "4"
        Print #1, "40"
        Print #1, "40"
        Print #1, "49"
        Print #1, "32.5"
        Print #1, "49"
        Print #1, "-2.5"
        Print #1, "49"
        Print #1, "2.5"
        Print #1, "49"
        Print #1, "-2.5"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "LONG2"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "--------  -  -  --------  -  -  "
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "6"
        Print #1, "40"
        Print #1, "40"
        Print #1, "49"
        Print #1, "30"
        Print #1, "49"
        Print #1, "-2.5"
        Print #1, "49"
        Print #1, "1.25"
        Print #1, "49"
        Print #1, "-2.5"
        Print #1, "49"
        Print #1, "1.25"
        Print #1, "49"
        Print #1, "-2.5"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "LONG3"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "+++++++++++++++..+++++++++++++++"
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "2"
        Print #1, "40"
        Print #1, "40"
        Print #1, "49"
        Print #1, "37.5"
        Print #1, "49"
        Print #1, "-2.5"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "LONG4"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "+++++++++++++++..+++++++++++++++"
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "2"
        Print #1, "40"
        Print #1, "80"
        Print #1, "49"
        Print #1, "75"
        Print #1, "49"
        Print #1, "-5"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "LONG5"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "----------  --  ----------  --  "
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "4"
        Print #1, "40"
        Print #1, "60"
        Print #1, "49"
        Print #1, "37.5"
        Print #1, "49"
        Print #1, "-7.5"
        Print #1, "49"
        Print #1, "7.5"
        Print #1, "49"
        Print #1, "-7.5"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "dashed"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "-------   -------------   ------"
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "2"
        Print #1, "40"
        Print #1, "7.5"
        Print #1, "49"
        Print #1, "6"
        Print #1, "49"
        Print #1, "-1.5"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "dashed_spaced"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "----        --------        ----"
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "2"
        Print #1, "40"
        Print #1, "12"
        Print #1, "49"
        Print #1, "6"
        Print #1, "49"
        Print #1, "-6"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "long_dashed_dotted"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "------------- --- --------------"
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "4"
        Print #1, "40"
        Print #1, "15.25"
        Print #1, "49"
        Print #1, "12"
        Print #1, "49"
        Print #1, "-1.5"
        Print #1, "49"
        Print #1, "0.25"
        Print #1, "49"
        Print #1, "-1.5"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "long_dashed_double-dotted"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "----------- --- --- ------------"
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "6"
        Print #1, "40"
        Print #1, "17"
        Print #1, "49"
        Print #1, "12"
        Print #1, "49"
        Print #1, "-1.5"
        Print #1, "49"
        Print #1, "0.25"
        Print #1, "49"
        Print #1, "-1.5"
        Print #1, "49"
        Print #1, "0.25"
        Print #1, "49"
        Print #1, "-1.5"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "long_dashed_triplicate-dotted"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "--------- --- --- --- ----------"
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "8"
        Print #1, "40"
        Print #1, "18.75"
        Print #1, "49"
        Print #1, "12"
        Print #1, "49"
        Print #1, "-1.5"
        Print #1, "49"
        Print #1, "0.25"
        Print #1, "49"
        Print #1, "-1.5"
        Print #1, "49"
        Print #1, "0.25"
        Print #1, "49"
        Print #1, "-1.5"
        Print #1, "49"
        Print #1, "0.25"
        Print #1, "49"
        Print #1, "-1.5"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "dotted"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "   -       -       -       -    "
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "2"
        Print #1, "40"
        Print #1, "1.75"
        Print #1, "49"
        Print #1, "0.25"
        Print #1, "49"
        Print #1, "-1.5"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "chain"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "-------------- -- --------------"
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "4"
        Print #1, "40"
        Print #1, "18.5"
        Print #1, "49"
        Print #1, "12"
        Print #1, "49"
        Print #1, "-1.5"
        Print #1, "49"
        Print #1, "3.5"
        Print #1, "49"
        Print #1, "-1.5"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "chain_double_dash"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "------------ -- -- -------------"
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "6"
        Print #1, "40"
        Print #1, "23.5"
        Print #1, "49"
        Print #1, "12"
        Print #1, "49"
        Print #1, "-1.5"
        Print #1, "49"
        Print #1, "3.5"
        Print #1, "49"
        Print #1, "-1.5"
        Print #1, "49"
        Print #1, "3.5"
        Print #1, "49"
        Print #1, "-1.5"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "dashed_dotted"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "-----------     -     ----------"
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "4"
        Print #1, "40"
        Print #1, "9.25"
        Print #1, "49"
        Print #1, "6"
        Print #1, "49"
        Print #1, "-1.5"
        Print #1, "49"
        Print #1, "0.25"
        Print #1, "49"
        Print #1, "-1.5"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "double-dashed_dotted"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "------   -----------   -   -----"
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "6"
        Print #1, "40"
        Print #1, "16.75"
        Print #1, "49"
        Print #1, "6"
        Print #1, "49"
        Print #1, "-1.5"
        Print #1, "49"
        Print #1, "6"
        Print #1, "49"
        Print #1, "-1.5"
        Print #1, "49"
        Print #1, "0.25"
        Print #1, "49"
        Print #1, "-1.5"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "dashed_double-dotted"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "---------    -    -    ---------"
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "6"
        Print #1, "40"
        Print #1, "11"
        Print #1, "49"
        Print #1, "6"
        Print #1, "49"
        Print #1, "-1.5"
        Print #1, "49"
        Print #1, "0.25"
        Print #1, "49"
        Print #1, "-1.5"
        Print #1, "49"
        Print #1, "0.25"
        Print #1, "49"
        Print #1, "-1.5"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "double-dashed_double-dotted"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "-----   ---------   -   -   ----"
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "8"
        Print #1, "40"
        Print #1, "18.5"
        Print #1, "49"
        Print #1, "6"
        Print #1, "49"
        Print #1, "-1.5"
        Print #1, "49"
        Print #1, "6"
        Print #1, "49"
        Print #1, "-1.5"
        Print #1, "49"
        Print #1, "0.25"
        Print #1, "49"
        Print #1, "-1.5"
        Print #1, "49"
        Print #1, "0.25"
        Print #1, "49"
        Print #1, "-1.5"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "dashed_triplicate-dotted"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "-------    -    -    -    ------"
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "8"
        Print #1, "40"
        Print #1, "12.75"
        Print #1, "49"
        Print #1, "6"
        Print #1, "49"
        Print #1, "-1.5"
        Print #1, "49"
        Print #1, "0.25"
        Print #1, "49"
        Print #1, "-1.5"
        Print #1, "49"
        Print #1, "0.25"
        Print #1, "49"
        Print #1, "-1.5"
        Print #1, "49"
        Print #1, "0.25"
        Print #1, "49"
        Print #1, "-1.5"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "double-dashed_triplicate-dotted"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "-----  ---------  -  -  -  -----"
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "10"
        Print #1, "40"
        Print #1, "20.25"
        Print #1, "49"
        Print #1, "6"
        Print #1, "49"
        Print #1, "-1.5"
        Print #1, "49"
        Print #1, "6"
        Print #1, "49"
        Print #1, "-1.5"
        Print #1, "49"
        Print #1, "0.25"
        Print #1, "49"
        Print #1, "-1.5"
        Print #1, "49"
        Print #1, "0.25"
        Print #1, "49"
        Print #1, "-1.5"
        Print #1, "49"
        Print #1, "0.25"
        Print #1, "49"
        Print #1, "-1.5"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "undefined"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "--------------------------------"
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "0"
        Print #1, "40"
        Print #1, "0"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "_"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "                                "
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "0"
        Print #1, "40"
        Print #1, "0"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "__18"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "                                "
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "0"
        Print #1, "40"
        Print #1, "0"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "__19"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "                                "
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "0"
        Print #1, "40"
        Print #1, "0"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "__20"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "                                "
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "0"
        Print #1, "40"
        Print #1, "0"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "__21"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "                                "
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "0"
        Print #1, "40"
        Print #1, "0"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "__22"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "                                "
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "0"
        Print #1, "40"
        Print #1, "0"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "__23"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "                                "
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "0"
        Print #1, "40"
        Print #1, "0"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "__24"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "                                "
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "0"
        Print #1, "40"
        Print #1, "0"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "__25"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "                                "
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "0"
        Print #1, "40"
        Print #1, "0"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "__26"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "                                "
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "0"
        Print #1, "40"
        Print #1, "0"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "__27"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "                                "
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "0"
        Print #1, "40"
        Print #1, "0"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "__28"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "                                "
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "0"
        Print #1, "40"
        Print #1, "0"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "__29"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "                                "
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "0"
        Print #1, "40"
        Print #1, "0"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "__30"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "                                "
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "0"
        Print #1, "40"
        Print #1, "0"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "__31"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "                                "
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "0"
        Print #1, "40"
        Print #1, "0"
        Print #1, "0"
        Print #1, "LTYPE"
        Print #1, "2"
        Print #1, "__32"
        Print #1, "70"
        Print #1, "64"
        Print #1, "3"
        Print #1, "                                "
        Print #1, "72"
        Print #1, "65"
        Print #1, "73"
        Print #1, "0"
        Print #1, "40"
        Print #1, "0"
        Print #1, "0"
        Print #1, "ENDTAB"
        Print #1, "0"
        Print #1, "TABLE"
        Print #1, "2"
        Print #1, "STYLE"
        Print #1, "5"
        Print #1, "3"
        Print #1, "100"
        Print #1, "AcDbSymbolTable"
        Print #1, "70"
        Print #1, "1"
        Print #1, "0"
        Print #1, "STYLE"
        Print #1, "5"
        Print #1, "10"
        Print #1, "100"
        Print #1, "AcDbSymbolTableRecord"
        Print #1, "100"
        Print #1, "AcDbTextStyleTableRecord"
        Print #1, "2"
        Print #1, "STANDARD"
        Print #1, "70"
        Print #1, "0"
        Print #1, "40"
        Print #1, "0"
        Print #1, "41"
        Print #1, "1"
        Print #1, "50"
        Print #1, "0"
        Print #1, "71"
        Print #1, "0"
        Print #1, "42"
        Print #1, "0.2"
        Print #1, "3"
        Print #1, "txt"
        Print #1, "4"
        Print #1, "bigfont.shx"
        Print #1, "0"
        Print #1, "STYLE"
        Print #1, "5"
        Print #1, "26"
        Print #1, "100"
        Print #1, "AcDbSymbolTableRecord"
        Print #1, "100"
        Print #1, "AcDbTextStyleTableRecord"
        Print #1, "2"
        Print #1, "TATEGAKI"
        Print #1, "70"
        Print #1, "68"
        Print #1, "40"
        Print #1, "0"
        Print #1, "41"
        Print #1, "1"
        Print #1, "50"
        Print #1, "0"
        Print #1, "71"
        Print #1, "0"
        Print #1, "42"
        Print #1, "1"
        Print #1, "3"
        Print #1, "txt"
        Print #1, "4"
        Print #1, "bigfont.shx"
        Print #1, "0"
        Print #1, "ENDTAB"
        Print #1, "0"
        Print #1, "TABLE"
        Print #1, "2"
        Print #1, "LAYER"
        Print #1, "70"
        Print #1, "8"
        Print #1, "0"
        Print #1, "LAYER"
        Print #1, "2"
        Print #1, "_0-0_ORIGIN"
        Print #1, "70"
        Print #1, "64"
        Print #1, "62"
        Print #1, "7"
        Print #1, "6"
        Print #1, "CONTINUOUS"
        Print #1, "0"
        Print #1, "LAYER"
        Print #1, "2"
        Print #1, "_0-1_CAM01"
        Print #1, "70"
        Print #1, "64"
        Print #1, "62"
        Print #1, "7"
        Print #1, "6"
        Print #1, "CONTINUOUS"
        Print #1, "0"
        Print #1, "LAYER"
        Print #1, "2"
        Print #1, "_0-2_CAM02"
        Print #1, "70"
        Print #1, "64"
        Print #1, "62"
        Print #1, "7"
        Print #1, "6"
        Print #1, "CONTINUOUS"
        Print #1, "0"
        Print #1, "LAYER"
        Print #1, "2"
        Print #1, "_0-3_CAM03"
        Print #1, "70"
        Print #1, "64"
        Print #1, "62"
        Print #1, "7"
        Print #1, "6"
        Print #1, "CONTINUOUS"
        Print #1, "0"
        Print #1, "LAYER"
        Print #1, "2"
        Print #1, "_0-4_CAM04"
        Print #1, "70"
        Print #1, "64"
        Print #1, "62"
        Print #1, "7"
        Print #1, "6"
        Print #1, "CONTINUOUS"
        Print #1, "0"
        Print #1, "LAYER"
        Print #1, "2"
        Print #1, "_0-5_CAM05"
        Print #1, "70"
        Print #1, "64"
        Print #1, "62"
        Print #1, "7"
        Print #1, "6"
        Print #1, "CONTINUOUS"
        Print #1, "0"
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
        Print #1, "LAYER"
        Print #1, "2"
        Print #1, "_1-0_"
        Print #1, "70"
        Print #1, "64"
        Print #1, "62"
        Print #1, "7"
        Print #1, "6"
        Print #1, "CONTINUOUS"
        Print #1, "0"
        Print #1, "LAYER"
        Print #1, "2"
        Print #1, "ADD_OBJECT"
        Print #1, "70"
        Print #1, "64"
        Print #1, "62"
        Print #1, "7"
        Print #1, "6"
        Print #1, "CONTINUOUS"
        Print #1, "0"
        Print #1, "ENDTAB"
        Print #1, "0"
        Print #1, "ENDSEC"
        Print #1, "0"
        Print #1, "SECTION"
        Print #1, "2"
        Print #1, "BLOCKS"
        Print #1, "0"
        Print #1, "ENDSEC"
        Print #1, "0"
        Print #1, "SECTION"
        Print #1, "2"
        Print #1, "ENTITIES"
        'Print #1, "0"
    Close #1
End Function
Private Function ptime(mes As String) As String
    ptime = Format(Now, "yyyy-mm-dd-hh-nn-ss")
    Debug.Print mes & " : " & ptime
End Function
Private Function DesktopFilepath() As String
    DesktopFilepath = "C:\Users\" & Environ("Username") & "\Desktop\"
End Function
