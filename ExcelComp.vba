' This tool need a "Setting" page to config the worksheets or pivot tables that you want to compare

' if you want to compare two worksheets, you can call ComparePivottables
' I use "PRD" and "UAT" as the names of the pivottables
' you can set the name as you will

Sub Main()
    'CopyWorksheet
    
    'SetUATConnection
    
    SetValueDate
    
    CompareSheet
    
End Sub

Sub CompareSheet()
    Dim Sht As Worksheet
    Dim pt As PivotTable
    Dim rng As Range
    Dim ws_prd As Worksheet
    Dim ws_uat As Worksheet
    
    Dim startTime As Date
    Dim endTime As Date
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Set Sht = ThisWorkbook.Worksheets("Settings")
    
    n = 12
    
    While Sht.Cells(n, 1).Value <> ""
        startTime = Now()
        Set ws_prd = ThisWorkbook.Worksheets(Sht.Cells(n, 1).Value)
        Set ws_uat = ThisWorkbook.Worksheets(Sht.Cells(n, 1).Value & "_UAT")
        
        With ws_prd.Cells.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        
        With ws_uat.Cells.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        
        For Each pt In ws_prd.PivotTables
            Set rng = pt.TableRange1
            For i = rng.Row To rng.Row + rng.Rows.Count - 1
                For j = rng.Column To rng.Column + rng.Columns.Count - 1
                    If ws_prd.Cells(i, j).Value <> ws_uat.Cells(i, j).Value Then
                        If TypeName(ws_prd.Cells(i, j).Value) = "String" Or TypeName(ws_prd.Cells(i, j).Value) = "Empty" Then
                            With ws_prd.Cells(i, j).Interior
                                .Pattern = xlSolid
                                .PatternColorIndex = xlAutomatic
                                .Color = 65535
                                .TintAndShade = 0
                                .PatternTintAndShade = 0
                            End With
                            
                            With ws_uat.Cells(i, j).Interior
                                .Pattern = xlSolid
                                .PatternColorIndex = xlAutomatic
                                .Color = 65535
                                .TintAndShade = 0
                                .PatternTintAndShade = 0
                            End With
                        Else
                            'prdcell = Application.Ceiling(ws_prd.Cells(i, j).Value, 1)
                            'uatcell = Application.Ceiling(ws_uat.Cells(i, j).Value, 1)
                            If Round(ws_prd.Cells(i, j).Value) = Round(ws_uat.Cells(i, j).Value) Then
                                With ws_prd.Cells(i, j).Interior
                                    .Pattern = xlNone
                                    .TintAndShade = 0
                                    .PatternTintAndShade = 0
                                End With
                                
                                With ws_uat.Cells(i, j).Interior
                                    .Pattern = xlNone
                                    .TintAndShade = 0
                                    .PatternTintAndShade = 0
                                End With
                            Else
                                With ws_prd.Cells(i, j).Interior
                                .Pattern = xlSolid
                                .PatternColorIndex = xlAutomatic
                                .Color = 65535
                                .TintAndShade = 0
                                .PatternTintAndShade = 0
                                End With
                                
                                With ws_uat.Cells(i, j).Interior
                                    .Pattern = xlSolid
                                    .PatternColorIndex = xlAutomatic
                                    .Color = 65535
                                    .TintAndShade = 0
                                    .PatternTintAndShade = 0
                                End With
                            End If
                           
                        End If
                    Else
                        With ws_prd.Cells(i, j).Interior
                            .Pattern = xlNone
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                        
                        With ws_uat.Cells(i, j).Interior
                            .Pattern = xlNone
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                    End If
                Next j
            Next i
            
        Next
        
        n = n + 1
    Wend
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
End Sub

Sub ComparePivottables()
    Dim startTime As Date
    Dim endTime As Date
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Set Sht = ThisWorkbook.Worksheets("Settings")
    
    n = 12
    
    While Sht.Cells(n, 1).Value <> ""
        startTime = Now()
        Set ws = ThisWorkbook.Worksheets(Sht.Cells(n, 1).Value)
        
        With ws.Cells.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        
        Set pvt_prd = ws.PivotTables("PRD")
        Set pvt_uat = ws.PivotTables("UAT")
        
        Set uat_rng = pvt_uat.TableRange1
        Set prd_rng = pvt_prd.TableRange1
        startCol = pvt_prd.TableRange1.Column
        RowCount = WorksheetFunction.Max(prd_rng.Rows.Count, uat_rng.Rows.Count)
        ColCount = WorksheetFunction.Max(prd_rng.Columns.Count, uat_rng.Columns.Count)
        
        For i = uat_rng.Row To uat_rng.Row + RowCount - 1
            For j = uat_rng.Column To uat_rng.Column + ColCount - 1
                Debug.Print ("i=" & i & ", " & "j=" & j & ", Prd Col:" & j + startCol & ", UAT Cell:" & ws.Cells(i, j).Value & ", Prd Cell:" & ws.Cells(i, j + startCol - 1).Value)
                If ws.Cells(i, j).Value <> ws.Cells(i, j + startCol - 1).Value Then
                            If TypeName(ws.Cells(i, j).Value) = "String" Or TypeName(ws.Cells(i, j).Value) = "Empty" Then
                                With ws.Cells(i, j).Interior
                                    .Pattern = xlSolid
                                    .PatternColorIndex = xlAutomatic
                                    .Color = 65535
                                    .TintAndShade = 0
                                    .PatternTintAndShade = 0
                                End With
                                
                                With ws.Cells(i, j + startCol - 1).Interior
                                    .Pattern = xlSolid
                                    .PatternColorIndex = xlAutomatic
                                    .Color = 65535
                                    .TintAndShade = 0
                                    .PatternTintAndShade = 0
                                End With
                            Else
                                'prdcell = Application.Ceiling(ws_prd.Cells(i, j).Value, 1)
                                'uatcell = Application.Ceiling(ws_uat.Cells(i, j).Value, 1)
                                If Round(ws.Cells(i, j).Value) = Round(ws.Cells(i, j + startCol - 1).Value) Then
                                    With ws.Cells(i, j).Interior
                                        .Pattern = xlNone
                                        .TintAndShade = 0
                                        .PatternTintAndShade = 0
                                    End With
                                    
                                    With ws.Cells(i, j + startCol).Interior
                                        .Pattern = xlNone
                                        .TintAndShade = 0
                                        .PatternTintAndShade = 0
                                    End With
                                Else
                                    With ws.Cells(i, j).Interior
                                    .Pattern = xlSolid
                                    .PatternColorIndex = xlAutomatic
                                    .Color = 65535
                                    .TintAndShade = 0
                                    .PatternTintAndShade = 0
                                    End With
                                    
                                    With ws.Cells(i, j + startCol - 1).Interior
                                        .Pattern = xlSolid
                                        .PatternColorIndex = xlAutomatic
                                        .Color = 65535
                                        .TintAndShade = 0
                                        .PatternTintAndShade = 0
                                    End With
                                End If
                               
                            End If
                        Else
                            With ws.Cells(i, j).Interior
                                .Pattern = xlNone
                                .TintAndShade = 0
                                .PatternTintAndShade = 0
                            End With
                            
                            With ws.Cells(i, j + startCol).Interior
                                .Pattern = xlNone
                                .TintAndShade = 0
                                .PatternTintAndShade = 0
                            End With
                        End If
            Next j
        Next i
        
        endTime = Now()
        
        dtDuration = DateDiff("s", startTime, endTime)
        Sht.Cells(n, 4).Value = dtDuration
        
        n = n + 1
    Wend
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub


Sub SetValueDate()
    Dim Sht As Worksheet
    Dim dateFilter As String
    Dim pt As PivotTable
    Dim dateInRow As Integer
    Dim dateFiledStr As String
    Dim dateCubField As CubeField
    Dim datePvtField As PivotField
    
    Dim startTime As Date
    Dim endTime As Date
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    MyName = ThisWorkbook.Name
    Set Sht = ThisWorkbook.Worksheets("Settings")
    'Sht.Calculate
    'Config table starting row
    n = 12
    
    dateFiledStr = "[Value Date].[Value Date].[Value Date]"
    
    szDate = Format(Sht.Cells(2, 4), "yyyymmdd")
    dateFilter = "[Value Date].[Value Date].&[" & szDate & "]"
    
    'Set effective date for all worksheet which have pivot table
    For Each ws In ThisWorkbook.Worksheets
        startTime = Now()
        ws.Application.Calculation = xlCalculationManual
        If ws.PivotTables.Count > 0 And ws.Visible = xlSheetVisible Then
            For Each pvt In ws.PivotTables
                dateInRow = 0
                
                If pvt.PivotCache.SourceType <> 1 Then
                    Set pt = pvt
                    
                    If PivotFieldsExists(pt, dateFiledStr) Then
                        Set datePvtField = pt.PivotFields(dateFiledStr)
                        
                        For Each pvtField In pt.RowFields
                            If pvtField.Name = dateFiledStr Then
                                dateInRow = 1
                            End If
                        Next pvtField
                        
                        If dateInRow <> 1 Then
                            With datePvtField
                                .ClearAllFilters
                                .CubeField.EnableMultiplePageItems = True
                                .VisibleItemsList = Array(dateFilter)
                                .CubeField.Orientation = xlPageField
                                .CubeField.Position = 1
                            End With
                        End If
                    End If
                    
                    ThisWorkbook.Save
                End If
            Next
        End If
        
        ws.Calculate
        
        endTime = Now()
        
        If ws.Name <> Sht.Name And ws.Visible = xlSheetVisible Then
            dtDuration = DateDiff("s", startTime, endTime)
            Sht.Cells(n, 3).Value = dtDuration
            n = n + 1
        End If
    Next
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Sub SetUATConnection()
    Dim Sht As Worksheet
    Dim pt As PivotTable
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    For Each Sht In ThisWorkbook.Worksheets
        If InStr(Sht.Name, "_UAT") <> 0 Then
            For Each pt In Sht.PivotTables
                If pt.PivotCache.SourceType <> 1 Then
                    'pt.PivotCache.Refresh
                    pt.ChangeConnection ThisWorkbook.Connections("UDW_UAT")
                    'Application.Wait (Now + TimeValue("0:00:25"))
                End If
            Next
        End If
    Next
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Public Function ConvertToInteger(Cell As Range) As Integer
   On Error GoTo NOT_AN_INTEGER
   ConvertToInteger = CInt(Cell.Value)
   Exit Function
NOT_AN_INTEGER:
   ConvertToInteger = 0
End Function

Function CellType(c)
    Application.Volatile
    Set c = c.Range("A1")
    Select Case True
        Case IsEmpty(c): CellType = "Blank"
        Case Application.IsText(c): CellType = "Text"
        Case Application.IsLogical(c): CellType = "Logical"
        Case Application.IsErr(c): CellType = "Error"
        Case IsDate(c): CellType = "Date"
        Case InStr(1, c.Text, ":") <> 0: CellType = "Time"
        Case IsNumeric(c): CellType = "Value"
    End Select
End Function


Sub CopyWorksheet()
    Dim Sht As Worksheet
    Dim ws As Worksheet
    
    Set Sht = ThisWorkbook.Worksheets("Settings")
    
    n = 12
    
    While Sht.Cells(n, 1).Value <> ""
        Set ws_prd = ThisWorkbook.Worksheets(Sht.Cells(n, 1).Value)
        old_name = ws_prd.Name & " (2)"
        new_name = ws_prd.Name & "_UAT"
        
        If Not CheckExists(new_name) Then
            ws_prd.Copy After:=ws_prd
            Sheets(old_name).Name = new_name
        End If
        n = n + 1
    Wend
End Sub

Function CheckExists(shtName As Variant) As Boolean
  CheckExists = False
  For Each ws In Worksheets
    If shtName = ws.Name Then
      CheckExists = True
      Exit Function
    End If
  Next ws
End Function


Function ElapsedTime(endTime As Date, startTime As Date)
    Dim strOutput As String
    Dim Interval As Date
     
    ' Calculate the time interval.
    Interval = endTime - startTime
  
    ' Format and print the time interval in seconds.
    strOutput = Int(CSng(Interval * 24 * 3600)) & " Seconds"
    Debug.Print strOutput
         
    ' Format and print the time interval in minutes and seconds.
    strOutput = Int(CSng(Interval * 24 * 60)) & ":" & Format(Interval, "ss") _
        & " Minutes:Seconds"
    Debug.Print strOutput
     
    ' Format and print the time interval in hours, minutes and seconds.
    strOutput = Int(CSng(Interval * 24)) & ":" & Format(Interval, "nn:ss") _
           & " Hours:Minutes:Seconds"
    Debug.Print strOutput
         
    ' Format and print the time interval in days, hours, minutes and seconds.
    strOutput = Int(CSng(Interval)) & " days " & Format(Interval, "hh") _
        & " Hours " & Format(Interval, "nn") & " Minutes " & _
        Format(Interval, "ss") & " Seconds"
    Debug.Print strOutput
 
End Function


Function PivotFieldsExists(pt As PivotTable, pfName As String) As Boolean
    Dim pf As PivotField
    PivotFieldsExists = False
    For Each pf In pt.PivotFields
        If pf.Name = pfName Then
            PivotFieldsExists = True
        End If
    Next pf
End Function












