Option Explicit
Sub MatchMyTables()
    Dim TableRange1 As Range, TableRange2 As Range, r As Range
    Dim ResultSht As Worksheet
    Dim Arr1(), Arr2(), ColumnArr1(), ColumnArr2(), TempArr(), TempArr_patt()
    Dim ColumnProperties1(), ColumnProperties2()
    Dim FullArray()
    Dim i As Long, j As Long, k As Long, ii As Long, jj As Long
    Dim Key1 As String, Key2 As String, keyArr() As String
    Dim ColumnsToWrite(), ColumnsToWrite_c As Long, ResultRow As Long, ind_1 As Long, ind_2 As Long
    Dim Left_s As Long  'left step
    Dim GoDown_ind As Long, GoUp_ind As Long
    Dim Table1Name As String, Table2Name As String
    Dim FoundPerc As Double, RazMah As Double
    Dim FindPairs_Algo As Boolean, ShowErrors As Boolean, FindPatterns_Algo As Boolean, PaintNotFound As Boolean
    Dim RowCountOptim As Boolean
    Dim K_a As Long, K_b As Long, K_c As Long, K_Tanimoto, K_Tanimoto_max, K_Tanimoto_Index
    Dim K_a_p As Long, K_b_p As Long, K_c_p As Long, K_Tanimoto_p, K_Tanimoto_max_p, K_Tanimoto_Index_p
    Dim UniqKeys As New Collection, UniqValues As Long
    Dim t As Date, dtFinish As Date, dtWork As Date
    Dim MaxColArr1 As Long, MaxColArr2 As Long
    
    TableMatch_frm.Show
    
     t = Time          ' Set up timer
    '**Set Table 1 and Table 2
    If TableMatch_frm.TextBox1 = vbNullString Then
        i = ThisWorkbook.Worksheets("Table1").UsedRange.Rows.count: j = ThisWorkbook.Worksheets("Table1").UsedRange.Columns.count
        Set TableRange1 = ThisWorkbook.Worksheets("Table1").Range(cells(1, 1).Address & ":" & cells(i, j).Address)
    Else
        Set TableRange1 = Range(TableMatch_frm.TextBox1)
    End If
    If TableMatch_frm.TextBox3 = vbNullString Then
        i = ThisWorkbook.Worksheets("Table2").UsedRange.Rows.count: j = ThisWorkbook.Worksheets("Table 2").cells(1, 1).End(xlToRight).Column 'ThisWorkbook.Worksheets("Table 2").UsedRange.Columns.Count
        Set TableRange2 = ThisWorkbook.Worksheets("Table2").Range(cells(1, 1).Address & ":" & cells(i, j).Address)
    Else
        Set TableRange2 = Range(TableMatch_frm.TextBox3)
    End If
        
    'Set key columns for our tables
    If TableMatch_frm.TextBox1 = vbNullString Then
        Key1 = "1 8 14"
    Else
        Key1 = TableMatch_frm.TextBox2
    End If
    If TableMatch_frm.TextBox4 = vbNullString Then
        Key2 = "1 7 8"
    Else
        Key2 = TableMatch_frm.TextBox4
    End If
    
    FindPairs_Algo = TableMatch_frm.Cb_Datas.Value
    FindPatterns_Algo = TableMatch_frm.Cb_DataPatterns
    ShowErrors = TableMatch_frm.Cb_ErrorsMarks
    PaintNotFound = TableMatch_frm.Cb_PaintNotFound
    RowCountOptim = TableMatch_frm.Cb_RowCountOptim
    
    Table1Name = TableMatch_frm.TextBox5
    Table2Name = TableMatch_frm.TextBox6
    
    Unload TableMatch_frm
    
    '**Create Result Worksheet if it not exist
    If Not (WorksheetIsExist("Result")) Then
        Set ResultSht = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.count))
        ResultSht.Name = "Result"
    Else
        Set ResultSht = ActiveWorkbook.Worksheets("Result")
        ResultSht.cells.Clear
        ResultSht.UsedRange.ClearOutline
    End If
    
    TurnCalculations_OFF            'Turn off screen updating e.t.c.
    ResultSht.cells.Clear
    
    On Error Resume Next
    '**
    Arr1 = TableRange1.Value
    ReDim Preserve Arr1(1 To TableRange1.Rows.count, 1 To TableRange1.Columns.count + 1)
    'Create table 1 key
    keyArr = Split(Key1)
    k = 0
    If RowCountOptim Then
        ReDim TempArr(1 To TableRange1.Rows.count + TableRange2.Rows.count - 2)
        ReDim TempArr_patt(1 To TableRange1.Rows.count + TableRange2.Rows.count - 2)
    End If
    MaxColArr1 = UBound(Arr1, 2)
    For i = 1 To UBound(Arr1, 1)
        For j = 0 To UBound(keyArr)
            Arr1(i, MaxColArr1) = Arr1(i, MaxColArr1) & IIf(j = 0, "", "_") & Arr1(i, keyArr(j))
        Next
        If i > 1 And RowCountOptim Then
            TempArr(i - 1) = Arr1(i, MaxColArr1)
            TempArr_patt(i - 1) = 1
            If i Mod 2000 = 0 Then
                Application.StatusBar = "Adding keys for table 1 into first Table"
                DoEvents
            End If
        End If
    Next

    k = i - 2
    jj = 0      'Var to calculate pairs
    'Create table 2 key
    Arr2 = TableRange2.Value
    ReDim Preserve Arr2(1 To TableRange2.Rows.count, 1 To TableRange2.Columns.count + 2)
    '+1 -- Key; +2 -- if it Was found in second Array
    keyArr = Split(Key2)
    MaxColArr2 = UBound(Arr2, 2)
    For i = 1 To UBound(Arr2, 1)
        For j = 0 To UBound(keyArr)
            Arr2(i, MaxColArr2 - 1) = Arr2(i, MaxColArr2 - 1) & IIf(j = 0, "", "_") & Arr2(i, keyArr(j))
        Next
        If i > 1 And RowCountOptim Then
            ii = 0
            For j = 1 To k
                If TempArr_patt(j) = 1 Then
                    If Arr2(i, MaxColArr2 - 1) = TempArr(j) Then
                        ii = 1
                        TempArr(j) = TempArr(k)
                        TempArr_patt(j) = TempArr_patt(k)
                        k = k - 1: jj = jj + 1
                        Exit For
                    End If
                End If
            Next
            If i Mod 1000 = 0 Then
                Application.StatusBar = "Searching keys of Table 2 in Table 1 keys and adding new ones into the table( row " & CStr(i) & "/" & CStr(UBound(Arr2, 1)) & " of table 2)"
                DoEvents
            End If
        End If
        If ii = 0 And i > 1 And RowCountOptim Then
            k = k + 1
            TempArr(k) = Arr2(i, MaxColArr2 - 1)
            TempArr_patt(k) = 2
        End If
        'UniqKeys.Add Arr2(i, UBound(Arr2, 2)), CStr(Arr2(i, UBound(Arr2, 2)))
    Next
    If RowCountOptim Then UniqValues = k + jj + 2
    On Error GoTo 0
    Erase TempArr
    Erase TempArr_patt
    
    ReDim ColumnProperties1(1 To TableRange1.Columns.count, 1 To 3) '1: Array Type | 2: Index of column in Second Array | 3: idn...
    ReDim ColumnProperties2(1 To TableRange2.Columns.count, 1 To 3)
    
    Application.StatusBar = "Determining numeric and symbolic columns"
    
    For i = 1 To TableRange1.Columns.count
        Set r = TableRange1.Columns(i).Offset(1, 0).Resize(TableRange1.Rows.count - 1)
        j = r.cells.count - Application.WorksheetFunction.count(r) - Application.WorksheetFunction.CountBlank(r)
        If j > 0 Or Application.WorksheetFunction.count(r) = 0 Then
            ColumnProperties1(i, 1) = "Symbolic"
            
        Else
            ColumnProperties1(i, 1) = "Numeric"
        End If
    Next
    For i = 1 To TableRange2.Columns.count
        Set r = TableRange2.Columns(i).Offset(1, 0).Resize(TableRange2.Rows.count - 1)
        j = r.cells.count - Application.WorksheetFunction.count(r) - Application.WorksheetFunction.CountBlank(r)
        If j > 0 Or Application.WorksheetFunction.count(r) = 0 Then
            ColumnProperties2(i, 1) = "Symbolic"
            
        Else
            ColumnProperties2(i, 1) = "Numeric"
        End If
    Next
    
    Application.StatusBar = "Check matching columns names"
    
    For i = 1 To TableRange1.Columns.count
        For j = 1 To TableRange2.Columns.count
            If ColumnProperties1(i, 2) = 0 Then
                If Trim(TableRange1(i)) = Trim(TableRange2(j)) Then
                    ColumnProperties1(i, 2) = j
                    ColumnProperties2(j, 2) = i
                    ColumnsToWrite_c = ColumnsToWrite_c + 1
                    ReDim Preserve ColumnsToWrite(1 To 2, 1 To ColumnsToWrite_c)
                    ColumnsToWrite(1, ColumnsToWrite_c) = i: ColumnsToWrite(2, ColumnsToWrite_c) = j
                    '              Column number in table 1                 Column number in table 2
                    Exit For
                End If
            End If
        Next
    Next
    
    If FindPairs_Algo Then  '**************************
    '**Values and patterns matching block
    Application.StatusBar = "Creating columns with uniques values in Table 1"
    
    If TableRange1.Rows.count >= 10000 Then j = 10000 Else j = TableRange1.Rows.count
    '**Analyze table 1 **
    ReDim ColumnArr1(1 To j, 1 To TableRange1.Columns.count, 3)
    For i = 1 To TableRange1.Columns.count
        If ColumnProperties1(i, 2) = 0 And ColumnProperties1(i, 1) = "Symbolic" Then    'Search only for symbolic data
            Set r = TableRange1.Columns(i).Offset(1, 0).Resize(TableRange1.Rows.count - 1)
            TempArr = r.Value
            For j = 1 To UBound(TempArr)
                If j = 10001 Then Exit For
                ColumnArr1(j, i, 0) = TempArr(j, 1)
            Next
            TempArr = Application.WorksheetFunction.Transpose(r)
            Uniq TempArr
            k = UBound(TempArr)
            
            If FindPatterns_Algo Then
                ReDim TempArr_patt(1 To UBound(TempArr))
                For j = 1 To UBound(TempArr_patt)
                    TempArr_patt(j) = ConvertDataToPattern(TempArr(j))
                Next
                Uniq TempArr_patt
                For j = 1 To UBound(TempArr_patt)
                    If j = 10001 Then Exit For
                    ColumnArr1(j, i, 3) = TempArr_patt(j)
                Next
            End If
            
            For j = 1 To UBound(TempArr)
                FoundPerc = Application.WorksheetFunction.CountIf(r, TempArr(j)) / r.Rows.count
                If j = 10001 Then Exit For
                ColumnArr1(j, i, 1) = TempArr(j) 'Uniq value
                ColumnArr1(j, i, 2) = FoundPerc
                TempArr(j) = FoundPerc
            Next
            
            If UBound(TempArr) / r.Rows.count < 0.5 Then    'Search outlines only if unique values below 50% of overall values count
                'Calculating Scope
                RazMah = Application.WorksheetFunction.Percentile(TempArr, 0.75) - Application.WorksheetFunction.Percentile(TempArr, 0.25)
                If RazMah = 0 Then RazMah = 0.00001
                'Calculate low border, all percentages under this value we'll count as outlines
                RazMah = Application.WorksheetFunction.Percentile(TempArr, 0.25) - RazMah * 1.5
                For j = 1 To k
                    If ColumnArr1(j, i, 2) <= RazMah Then
                        'ColumnArr1(j, i, 0) = ColumnArr1(k, i, 0)
                        ColumnArr1(j, i, 2) = ColumnArr1(k, i, 2)
                        ColumnArr1(j, i, 1) = ColumnArr1(k, i, 1)
                        'ColumnArr1(k, i, 0) = 0
                        ColumnArr1(k, i, 1) = 0 'Uniq Value
                        ColumnArr1(k, i, 2) = 0 'Percentage value
                        k = k - 1
                    End If
                Next
            End If
        End If
    Next
    DoEvents
    Application.StatusBar = "Creating columns with uniques values in Table 2"
    
    If TableRange2.Rows.count >= 10000 Then j = 10000 Else j = TableRange2.Rows.count
    '**Analyze table 2 **
    ReDim ColumnArr2(1 To j, 1 To TableRange2.Columns.count, 3)
    For i = 1 To TableRange2.Columns.count
        If ColumnProperties2(i, 2) = 0 And ColumnProperties2(i, 1) = "Symbolic" Then    'Search only for symbolic data
            Set r = TableRange2.Columns(i).Offset(1, 0).Resize(TableRange2.Rows.count - 1)
            TempArr = r.Value
            For j = 1 To UBound(TempArr)
                If j = 10001 Then Exit For
                ColumnArr2(j, i, 0) = TempArr(j, 1)
            Next
            TempArr = Application.WorksheetFunction.Transpose(r)
            Uniq TempArr
            k = UBound(TempArr)
            
            If FindPatterns_Algo Then
                'If i = 4 Then Stop
                ReDim TempArr_patt(1 To UBound(TempArr))
                For j = 1 To UBound(TempArr_patt)
                    TempArr_patt(j) = ConvertDataToPattern(TempArr(j))
                Next
                Uniq TempArr_patt
                For j = 1 To UBound(TempArr_patt)
                    If j = 10001 Then Exit For
                    ColumnArr2(j, i, 3) = TempArr_patt(j)
                Next
            End If
            
            For j = 1 To UBound(TempArr)
                FoundPerc = Application.WorksheetFunction.CountIf(r, TempArr(j)) / r.Rows.count
                If j = 10001 Then Exit For
                ColumnArr2(j, i, 1) = TempArr(j)  'Uniq value
                ColumnArr2(j, i, 2) = FoundPerc
                TempArr(j) = FoundPerc
            Next
            If UBound(TempArr) / r.Rows.count < 0.5 Then  'Search outlines only if unique values below 50% of overall values count
                'Calculating Scope
                RazMah = Application.WorksheetFunction.Percentile(TempArr, 0.75) - Application.WorksheetFunction.Percentile(TempArr, 0.25)
                If RazMah = 0 Then RazMah = 0.00001
                'Calculate low border, all percentages under this value we'll count as outlines
                RazMah = Application.WorksheetFunction.Percentile(TempArr, 0.25) - RazMah * 1.5
                For j = 1 To k
                    If ColumnArr2(j, i, 2) <= RazMah Then
                        'ColumnArr2(j, i, 0) = ColumnArr2(k, i, 0)
                        ColumnArr2(j, i, 2) = ColumnArr2(k, i, 2)
                        ColumnArr2(j, i, 1) = ColumnArr2(k, i, 1)
                        'ColumnArr2(k, i, 0) = 0
                        ColumnArr2(k, i, 1) = 0 'Unique value
                        ColumnArr2(k, i, 2) = 0 'Percent
                        k = k - 1
                    End If
                Next
            End If
        End If
    Next
    Erase TempArr_patt      'Erasing Array to free some memory
    Erase TempArr           'Erasing Array to free some memory
    
    
    For i = 1 To TableRange1.Columns.count
        DoEvents
        Application.StatusBar = "Searching most suitable column for Table 1 [" & CStr(Arr1(1, i)) & "]"
        
        If ColumnProperties1(i, 2) = 0 And ColumnProperties1(i, 1) = "Symbolic" Then    'Search only for symbolic data
            K_Tanimoto_max = 0
            K_Tanimoto_Index = 0
            K_Tanimoto_max_p = 0
            K_Tanimoto_Index_p = 0
            For j = 1 To TableRange2.Columns.count
                If ColumnProperties2(j, 2) = 0 And ColumnProperties2(j, 1) = "Symbolic" Then
                    K_a = 0: K_a_p = 0
                    K_b = 0: K_b_p = 0
                    K_c = 0: K_c_p = 0
                    For ii = 1 To UBound(ColumnArr1, 1)
                        If ColumnArr1(ii, i, 1) = 0 Then Exit For
                        K_a = K_a + 1
                        For jj = 1 To UBound(ColumnArr2, 1)
                            If ColumnArr1(ii, i, 1) = ColumnArr2(jj, j, 1) Then
                                K_c = K_c + 1: Exit For
                            End If
                            If ColumnArr2(jj, j, 1) = 0 Then K_b = jj - 1: Exit For
                        Next
                    Next
                    If FindPatterns_Algo Then                    '**Paterns check part***
                        For ii = 1 To UBound(ColumnArr1, 1)
                            If ColumnArr1(ii, i, 3) = 0 Then Exit For
                            K_a_p = K_a_p + 1
                            For jj = 1 To UBound(ColumnArr2, 1)
                                If ColumnArr1(ii, i, 3) = ColumnArr2(jj, j, 3) Then
                                    K_c_p = K_c_p + 1: Exit For
                                End If
                                If ColumnArr2(jj, j, 3) = 0 Then K_b_p = jj - 1: Exit For
                            Next
                        Next
                        If K_b_p = 0 Then
                            For jj = 1 To UBound(ColumnArr2, 1)
                                If ColumnArr2(jj, j, 3) = 0 Then Exit For
                                K_b_p = K_b_p + 1
                            Next
                        End If
                    End If
                    '*****
                    If K_b = 0 Then
                        For jj = 1 To UBound(ColumnArr2, 1)
                            If ColumnArr2(jj, j, 1) = 0 Then Exit For
                            K_b = K_b + 1
                        Next
                    End If
                    K_Tanimoto = K_c / (K_a + K_b - K_c)
                    If K_Tanimoto > K_Tanimoto_max Then K_Tanimoto_max = K_Tanimoto: K_Tanimoto_Index = j
                    If K_Tanimoto = 1 Then Exit For     '100% match
                    If FindPatterns_Algo Then
                        K_Tanimoto_p = K_c_p / (K_a_p + K_b_p - K_c_p)
                        If K_Tanimoto_p > K_Tanimoto_max_p Then K_Tanimoto_max_p = K_Tanimoto_p: K_Tanimoto_Index_p = j
                    End If
                    'k=c/(a+b-c)
                End If
            Next
            If K_Tanimoto_max > 0.15 Then
                ColumnProperties1(i, 2) = K_Tanimoto_Index
                ColumnProperties2(K_Tanimoto_Index, 2) = i
                ColumnsToWrite_c = ColumnsToWrite_c + 1
                ReDim Preserve ColumnsToWrite(1 To 2, 1 To ColumnsToWrite_c)
                ColumnsToWrite(1, ColumnsToWrite_c) = i: ColumnsToWrite(2, ColumnsToWrite_c) = K_Tanimoto_Index
                Arr1(1, i) = Arr1(1, i) & " (matches ~ " & Format(CStr(K_Tanimoto_max * 100), "##") & "%)"
                Arr2(1, K_Tanimoto_Index) = Arr2(1, K_Tanimoto_Index) & " (matches ~ " & Format(CStr(K_Tanimoto_max * 100), "##") & "%)"
            Else
                If FindPatterns_Algo And K_Tanimoto_max_p > 0.15 Then
                    ColumnProperties1(i, 2) = K_Tanimoto_Index_p
                    ColumnProperties2(K_Tanimoto_Index_p, 2) = i
                    ColumnsToWrite_c = ColumnsToWrite_c + 1
                    ReDim Preserve ColumnsToWrite(1 To 2, 1 To ColumnsToWrite_c)
                    ColumnsToWrite(1, ColumnsToWrite_c) = i: ColumnsToWrite(2, ColumnsToWrite_c) = K_Tanimoto_Index_p
                    Arr1(1, i) = Arr1(1, i) & " (matches -- " & Format(CStr(K_Tanimoto_max_p * 100), "##") & "%)"
                    Arr2(1, K_Tanimoto_Index_p) = Arr2(1, K_Tanimoto_Index_p) & " (matches -- " & Format(CStr(K_Tanimoto_max_p * 100), "##") & "%)"
                End If
            End If
        End If
    Next
    End If      '*********************************************
    
    '**End of values and patterns matching

    Application.StatusBar = "Determining columns without pair from both of tables"
    
    For i = 1 To TableRange1.Columns.count
        If ColumnProperties1(i, 2) = 0 Then
            ColumnsToWrite_c = ColumnsToWrite_c + 1
            ReDim Preserve ColumnsToWrite(1 To 2, 1 To ColumnsToWrite_c)
            ColumnsToWrite(1, ColumnsToWrite_c) = i
        End If
    Next
    For i = 1 To TableRange2.Columns.count
        If ColumnProperties2(i, 2) = 0 Then
            ColumnsToWrite_c = ColumnsToWrite_c + 1
            ReDim Preserve ColumnsToWrite(1 To 2, 1 To ColumnsToWrite_c)
            ColumnsToWrite(2, ColumnsToWrite_c) = i
        End If
    Next
    
    k = UBound(Arr2, 1) + UBound(Arr1, 1)
    If RowCountOptim Then k = UniqValues
    ReDim FullArray(1 To k, 1 To ColumnsToWrite_c * 2 + 2)
    k = 1
    
    Left_s = 2
    ResultSht.cells(1, 1) = "KEY"
    ResultSht.cells(1, 2) = "Found in tables:"
    
    FullArray(k, 1) = "KEY"
    FullArray(k, 2) = "Found in tables"
    
    Range(ResultSht.cells(1, 1), ResultSht.cells(1, 2)).Interior.Color = 65535
    For i = 1 To ColumnsToWrite_c
        If ColumnsToWrite(1, i) <> vbNullString Then
            ResultSht.cells(1, (i - 1) * 2 + 1 + Left_s) = Arr1(1, ColumnsToWrite(1, i)) & "[Table " & Table1Name & "]"
            FullArray(k, (i - 1) * 2 + 1 + Left_s) = Arr1(1, ColumnsToWrite(1, i)) & "[Table " & Table1Name & "]"
            
            If InStr(1, Arr1(1, ColumnsToWrite(1, i)), "matches ~") > 0 Then
                ResultSht.cells(1, (i - 1) * 2 + 1 + Left_s).Interior.Color = 11389944
            Else
                If InStr(1, Arr1(1, ColumnsToWrite(1, i)), "matches -") > 0 Then
                    ResultSht.cells(1, (i - 1) * 2 + 1 + Left_s).Interior.Color = 13434879
                Else
                    ResultSht.cells(1, (i - 1) * 2 + 1 + Left_s).Interior.Color = 6740479
                End If
            End If
            If ColumnProperties1(ColumnsToWrite(1, i), 1) = "Numeric" Then
                With ResultSht.cells(1, (i - 1) * 2 + 1 + Left_s).Font
                    .Color = vbWhite
                    .Italic = True
                    .Underline = True
                End With
            End If
        Else
            ResultSht.cells(1, (i - 1) * 2 + 1 + Left_s) = "Not found in [Table " & Table1Name & "]"
            FullArray(k, (i - 1) * 2 + 1 + Left_s) = "Not found in [Table " & Table1Name & "]"
            
            ResultSht.cells(2, (i - 1) * 2 + 1 + Left_s).cells.Interior.Color = 14408667
            ResultSht.cells(1, (i - 1) * 2 + 1 + Left_s).Interior.Color = 13224393
        End If
        If ColumnsToWrite(2, i) <> vbNullString Then
            ResultSht.cells(1, (i - 1) * 2 + 2 + Left_s) = Arr2(1, ColumnsToWrite(2, i)) & "[Table " & Table2Name & "]"
            FullArray(k, (i - 1) * 2 + 2 + Left_s) = Arr2(1, ColumnsToWrite(2, i)) & "[Table " & Table2Name & "]"
            
            If InStr(1, Arr2(1, ColumnsToWrite(2, i)), "matches ~") > 0 Then
                ResultSht.cells(1, (i - 1) * 2 + 2 + Left_s).Interior.Color = 14395790
            Else
                If InStr(1, Arr2(1, ColumnsToWrite(2, i)), "matches -") > 0 Then
                    ResultSht.cells(1, (i - 1) * 2 + 2 + Left_s).Interior.Color = vbMagenta
                Else
                    ResultSht.cells(1, (i - 1) * 2 + 2 + Left_s).Interior.Color = 15773696
                End If
            End If
            If ColumnProperties2(ColumnsToWrite(2, i), 1) = "Numeric" Then
                With ResultSht.cells(1, (i - 1) * 2 + 2 + Left_s).Font
                    .Color = vbWhite
                    .Italic = True
                    .Underline = True
                End With
            End If
        Else
            ResultSht.cells(1, (i - 1) * 2 + 2 + Left_s) = "Not found in [Table " & Table2Name & "]"
            FullArray(k, (i - 1) * 2 + 2 + Left_s) = "Not found in [Table " & Table2Name & "]"
            
            ResultSht.cells(2, (i - 1) * 2 + 2 + Left_s).cells.Interior.Color = 14408667
            ResultSht.cells(1, (i - 1) * 2 + 2 + Left_s).Interior.Color = 13224393
        End If
    Next
    With Range(ResultSht.cells(1, 1), ResultSht.cells(1, ColumnsToWrite_c * 2 + 2))
        .Font.Bold = True
        .Borders.Weight = xlThin
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    
    '*** Erase Arrays to free some memory ***
    Erase ColumnProperties1
    Erase ColumnProperties2
    
    ResultRow = 1
    For ind_1 = 2 To UBound(Arr1, 1)
        'ResultRow = ResultRow + 1
        k = k + 1
        ind_2 = 0
        GoDown_ind = ind_1
        GoUp_ind = ind_1
        Do While GoDown_ind > 1 Or GoUp_ind <= UBound(Arr2, 1)
            If GoDown_ind > 1 And GoDown_ind <= UBound(Arr2, 1) Then
                If Arr2(GoDown_ind, UBound(Arr2, 2) - 1) = Arr1(ind_1, UBound(Arr1, 2)) And Arr2(GoDown_ind, UBound(Arr2, 2)) = 0 Then
                    ind_2 = GoDown_ind: Arr2(GoDown_ind, UBound(Arr2, 2)) = 1
                    Exit Do
                End If
            End If
            If GoUp_ind <= UBound(Arr2, 1) Then
                If Arr2(GoUp_ind, UBound(Arr2, 2) - 1) = Arr1(ind_1, UBound(Arr1, 2)) And Arr2(GoUp_ind, UBound(Arr2, 2)) = 0 Then
                    ind_2 = GoUp_ind: Arr2(GoUp_ind, UBound(Arr2, 2)) = 1
                    Exit Do
                End If
            End If
            GoUp_ind = GoUp_ind + 1
            GoDown_ind = GoDown_ind - 1
        Loop
        'ResultSht.cells(ResultRow, 1) = Arr1(ind_1, UBound(Arr1, 2))
        FullArray(k, 1) = Arr1(ind_1, UBound(Arr1, 2))
        'ResultSht.cells(ResultRow, 2) = "'1" & IIf(ind_2 <> 0, " 2", "")
        FullArray(k, 2) = "'1" & IIf(ind_2 <> 0, " 2", "")
        For j = 1 To ColumnsToWrite_c
            If ColumnsToWrite(1, j) <> vbNullString Then
                'ResultSht.cells(ResultRow, (j - 1) * 2 + 1 + Left_s).Value = Arr1(ind_1, ColumnsToWrite(1, j))
                FullArray(k, (j - 1) * 2 + 1 + Left_s) = Arr1(ind_1, ColumnsToWrite(1, j))
            Else
                'ResultSht.Cells(ResultRow, (j - 1) * 2 + 1 + Left_s).Value = "..column not found.."
            End If
            If ColumnsToWrite(2, j) <> vbNullString Then
                If ind_2 <> 0 Then
                    'ResultSht.cells(ResultRow, (j - 1) * 2 + 2 + Left_s).Value = Arr2(ind_2, ColumnsToWrite(2, j))
                    FullArray(k, (j - 1) * 2 + 2 + Left_s) = Arr2(ind_2, ColumnsToWrite(2, j))
                Else
                    FullArray(k, (j - 1) * 2 + 2 + Left_s) = "..key not found.."
                    'With ResultSht.cells(ResultRow, (j - 1) * 2 + 2 + Left_s)
                     '   .Value = "..key not found.."
                     '   If PaintNotFound Then
                    '      .Font.Size = 8
                    '        .Font.Color = vbWhite
                     '   End If
                    'End With
                End If
            Else
                'ResultSht.Cells(ResultRow, (j - 1) * 2 + 2 + Left_s).Value = "..column not found.."
            End If
        Next
        If ind_1 Mod 2000 = 0 Then
            DoEvents
            Application.StatusBar = "Processing " & CStr(ind_1) & "/" & CStr(UBound(Arr1, 1)) & " row in " & Table1Name
        End If
    Next
    For ind_2 = 2 To UBound(Arr2, 1)
        If Arr2(ind_2, UBound(Arr2, 2)) = 0 Then
            k = k + 1
            'ResultRow = ResultRow + 1
            'ResultSht.cells(ResultRow, 1) = Arr2(ind_2, UBound(Arr2, 2) - 1)
            FullArray(k, 1) = Arr2(ind_2, UBound(Arr2, 2) - 1)
            'ResultSht.cells(ResultRow, 2) = "'2"
            FullArray(k, 2) = "'2"
            For j = 1 To ColumnsToWrite_c
                If ColumnsToWrite(1, j) <> vbNullString Then
                    FullArray(k, (j - 1) * 2 + 1 + Left_s) = "..key not found.."
'                    With ResultSht.cells(ResultRow, (j - 1) * 2 + 1 + Left_s)
'                        .Value = "..key not found.."
'                        If PaintNotFound Then
'                            .Font.Size = 8
'                            .Font.Color = vbWhite
'                        End If
'                    End With
                End If
                If ColumnsToWrite(2, j) <> vbNullString Then
                    'ResultSht.cells(ResultRow, (j - 1) * 2 + 2 + Left_s).Value = Arr2(ind_2, ColumnsToWrite(2, j))
                    FullArray(k, (j - 1) * 2 + 2 + Left_s) = Arr2(ind_2, ColumnsToWrite(2, j))
                End If
            Next
            If ind_2 Mod 2000 = 0 Then
                DoEvents
                Application.StatusBar = "Processing " & CStr(ind_2) & "/" & CStr(UBound(Arr1, 2)) & " row in " & Table2Name
            End If
        End If
    Next

    Range(ResultSht.cells(1, 1), ResultSht.cells(k, ColumnsToWrite_c * 2 + 2)).Value2 = FullArray
mark:
    
    Erase FullArray
    Set TableRange1 = Range(ResultSht.cells(1, 1), ResultSht.cells(k, ColumnsToWrite_c * 2 + 2))
    For Each TableRange2 In TableRange1.cells
        If TableRange2.Value2 = "..key not found.." Then
            With TableRange2
                If PaintNotFound Then
                    .Font.Size = 8
                    If TableRange2.Column Mod 2 = 0 Then .Font.Color = vbWhite
                End If
            End With
        End If
    Next
    ii = TableRange1.Rows.count
    i = 0 'for flag
    For j = 4 To ColumnsToWrite_c * 2 + 2
        Set r = Range(ResultSht.cells(2, j), ResultSht.cells(ii, j))
        If Not (ResultSht.cells(2, j).Interior.Color = 14408667) And (j Mod 2 = 0) And i = 0 Then
            CreateCondFormat r
            If ShowErrors Then
                ResultSht.cells(1, ColumnsToWrite_c * 2 + 3).FormulaArray = "=SUM((" & r.Address & "<>" & r.Offset(0, -1).Address & ")*1)"
                k = ResultSht.cells(1, ColumnsToWrite_c * 2 + 3).Value2 - Application.WorksheetFunction.CountIf(r, "..key not found..")
                k = k - Application.WorksheetFunction.CountIf(r.Offset(0, -1), "..key not found..")
                If k > 0 Then
                    With r.cells(1).Offset(-1, 0)
                        .Value = .Value & " " & CStr(k) & " errors."
                        .Characters(Len(r.cells(1).Offset(-1, 0)) - Len(CStr(k) & " errors."), Len(CStr(k) & " errors.")).Font.Color = vbRed '16711884
                        .Characters(Len(r.cells(1).Offset(-1, 0)) - Len(CStr(k) & " errors."), Len(CStr(k) & " errors.")).Font.Size = 12
                    End With
                End If
            End If
        Else
            If ResultSht.cells(2, j).Interior.Color = 14408667 Then r.Interior.Color = 14408667: i = 1
        End If
    Next
    ResultSht.cells(1, ColumnsToWrite_c * 2 + 3).Clear
    
    Range(ResultSht.cells(1, 1), ResultSht.cells(1, ColumnsToWrite_c * 2 + 2)).AutoFilter
    ResultSht.Activate
    TurnCalculations_ON
    Application.StatusBar = False
    
    dtFinish = Time
    dtWork = (dtFinish - t)
    If Minute(dtWork) = 0 Then
        MsgBox "All done in " & Second(dtWork) & " seconds "
    Else
        MsgBox "All done in " & Minute(dtWork) & " min. " & Second(dtWork) & " sec. "
    End If
End Sub
Public Sub TurnCalculations_ON()
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
End Sub
Public Sub TurnCalculations_OFF()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
End Sub
Sub CreateCondFormat(r As Range)
    Dim myFormula As String
    myFormula = "=" & r.cells(1).Address(0, 0) & "<>" & r.cells(1).Offset(0, -1).Address(0, 0)
    r.FormatConditions.Add Type:=xlExpression, Formula1:=myFormula
    With r.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13421823
        .TintAndShade = 0
    End With
    myFormula = "=" & r.cells(1).Address(0, 0) & "=" & r.cells(1).Offset(0, -1).Address(0, 0)
    r.FormatConditions.Add Type:=xlExpression, Formula1:=myFormula
    r.FormatConditions(r.FormatConditions.count).SetFirstPriority
    With r.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13434828
        .TintAndShade = 0
    End With
End Sub
 Private Sub Uniq(ByRef a)
    Dim Dict, b(), i As Long, k As Long
    ReDim b(1 To UBound(a))
    k = 1
    Set Dict = CreateObject("Scripting.Dictionary")
    With Dict
         For i = 1 To UBound(a)
            If Not .Exists(a(i)) Then .Add a(i), i: ReDim Preserve b(1 To k): b(k) = a(i): k = k + 1
         Next
    End With
    a = b
End Sub
Public Function WorksheetIsExist(iName$) As Boolean
   On Error Resume Next
   WorksheetIsExist = (TypeName(ActiveWorkbook.Worksheets(iName$)) = "Worksheet")
   On Error GoTo 0
End Function
Public Function ConvertDataToPattern(ByVal myVal As String) As String
    Dim i As Long, answ$, myChar$
    myVal = LCase(myVal)
    For i = 1 To Len(myVal)
        myChar$ = Mid(myVal, i, 1)
        If myChar$ Like "*[abcdefghijklmnopqrstuvwxyz]*" Or myChar$ Like "*[абвгджзеёиклмнопрстуфхчшщэюя]*" Then
            answ$ = answ$ & "c"
        Else
            If myChar$ Like "#" Then
                answ$ = answ$ & "n"
            Else
                answ$ = answ$ & "s"
            End If
        End If
    Next
    ConvertDataToPattern = answ$
End Function
