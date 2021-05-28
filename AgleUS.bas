Attribute VB_Name = "AgleUS"
'AGILE USA

Sub Formatowanie_Agile()
'
' Format_Agile Macro
'

' Functions
'--------------------------------------------------
    'Set up path
    SetCurrentDirectoryA "some path to folder"

    Obsluga_pliku
    dodawanie_sheetu
'---------------------------------------------------------------------------------------------------------
    
    Set CopyFrom = ActiveWorkbook
    Sheets("Sheet1").Select

    Range("N16").Select
'Filter
    Selection.AutoFilter
    ActiveSheet.Range("$A$15").AutoFilter Field:=14, Criteria1:= _
        "<>0" _
        , Operator:= _
        xlFilterValues
        
'copy items to temp sheet
    Dim TotalValueColum As Variant
    TotalValueColum = Application.WorksheetFunction.Sum(Columns("N:N"))
    If (TotalValueColum <> 0) Then
        Range("A15").Activate
        Range(Selection, Selection.End(xlDown)).Select
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.Copy
    
        Sheets("Sheet2").Select
        Range("A1").Select
        ActiveSheet.Paste
    End If
    
    If (TotalValueColum = 0) Then
        Range("A15").Activate
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.Copy
        Sheets("Sheet2").Select
        Range("A1").Select
        ActiveSheet.Paste
    End If
    
    Sheets("Sheet1").Select
    Application.CutCopyMode = False
    ActiveSheet.AutoFilterMode = False

    Range("A15").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Sheets("Sheet3").Select
    Range("A1").Select
    ActiveSheet.Paste
   
    Application.CutCopyMode = False
  
  
     If ActiveSheet.AutoFilterMode Then
     ActiveSheet.AutoFilterMode = False
     End If
        
    Sheets("Sheet1").Select
    
    ' setup destription temp sheet2
    Sheets("Sheet2").Select
    
     If ActiveSheet.AutoFilterMode Then
     ActiveSheet.AutoFilterMode = False
     End If
    
    'desc
    Cells(1, 17).Value = "Description"
    'PN
    Cells(1, 18).Value = "Project"
    'TASK
    Cells(1, 19).Value = "TASK"
    'EXP type
    Cells(1, 20).Value = "EXP TYPE"
    'value
    Cells(1, 21).Value = "VALUE"
    
    i = 2
    Do While Cells(i, 1) <> 0
    If Cells(i, 1) <> 0 Then
    Cells(i, 17).Value = Cells(i, 2) & " " & Cells(i, 6) & " " & Cells(i, 7)
    'PN
    Cells(i, 18).Value = Cells(i, 7)
    'TASK
    Cells(i, 19).Value = Mid(Cells(i, 8), 2, 200)
        If Cells(i, 19) = 0 Then
            Cells(i, 19).Value = "brak"
        End If
    
    'EXP type
    Cells(i, 20).Value = "Tax-Sales"
    'value
    Cells(i, 21).Value = Cells(i, 14)
    End If
    i = i + 1
    Loop
    
    
' setup description and others temp sheet3
'--------------------------------------------------
    Sheets("Sheet3").Select
    
     If ActiveSheet.AutoFilterMode Then
     ActiveSheet.AutoFilterMode = False
     End If
    
    i = 2
    Do While Cells(i, 1) <> 0
    If Cells(i, 1) <> 0 Then
    'desc
    Cells(i, 17).Value = Cells(i, 2) & " " & Cells(i, 6) & " " & Cells(i, 7)
    'PN
    Cells(i, 18).Value = "BALANCE"
    'TASK
    Cells(i, 19).Value = "227004"
    'EXP type
    Cells(i, 20).Value = "Balance"
    'value
    Cells(i, 21).Value = Cells(i, 13)
    End If
    i = i + 1
    Loop
'---------------------------------------------------------------------------------------------------------
    
    
'items to sheet2 from sheet3
'--------------------------------------------------
      
    Sheets("Sheet3").Select
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Sheets("Sheet2").Select
    
' Paste Data Below the Last Used Row
    Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    
' delete sheet3
    Application.DisplayAlerts = False
    Sheets("Sheet3").Delete
    Application.DisplayAlerts = True
'
'    Sheets("Sheet1").Select
'---------------------------------------------------------------------------------------------------------

    
    
'open M4A
'--------------------------------------------------
    Dim folderPath As String
    folderPath = Application.ActiveWorkbook.Path
    Filename = "file name"
    'file path
    'Workbooks.Open ("some file path")
    Workbooks.Open ("some file path")
    
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs (PathName & Filename)
    Application.DisplayAlerts = True
'---------------------------------------------------------------------------------------------------------
    
    
'paste data to M4A
'--------------------------------------------------
    Set CopyTo = ActiveWorkbook
    
    CopyFrom.Activate
    
    Sheets("Sheet2").Select
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    
    CopyTo.Activate
    
    Sheets("Upload").Select
    Range("B1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
'---------------------------------------------------------------------------------------------------------
    
'close workbook
'--------------------------------------------------
    Application.DisplayAlerts = False
    CopyFrom.Activate
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
'---------------------------------------------------------------------------------------------------------
    
    
'M4A format
'--------------------------------------------------
    CopyTo.Activate
    
    Sheets("Upload").Select
  
'  description r2 bu12
    Call Kolumn_copy((Range("BU12").Address), (Range("R2").Address))
'  Amount
    Call Kolumn_copy((Range("BV12").Address), (Range("V2").Address))
'  PN
    Call Kolumn_copy((Range("DM12").Address), (Range("S2").Address))
'  TASK
    Call Kolumn_copy((Range("DN12").Address), (Range("T2").Address))
'    EXP type
    Call Kolumn_copy((Range("DP12").Address), (Range("U2").Address))
    

    i = 12
        Do While Cells(i, 73) <> 0
    If Cells(i, 73) <> 0 Then
    '   Item
        Cells(i, 72).Value = "Item"
    '  accouting Date
        Cells(i, 100).Value = Date
    '     Organization
        Cells(i, 121).Value = "xxx"
    '   Item Date
        Cells(i, 122).Value = Date
        End If
    i = i + 1
    Loop
    
'HEADER
    '   Operating Unit
    Cells(12, 12).Value = "US_OU"
    '   Legal Enitity
    Cells(12, 13).Value = "xxx"
    '   Customer Taxpayer
    Cells(12, 14).Value = "xxx"
    '   Supplier Number
    Cells(12, 16).Value = "xxx"
    '   Supplier Name
    Cells(12, 17).Value = "xxx"
    '   Supplier Site
    Cells(12, 18).Value = "xxx"
    '   Invoice Type
    Cells(12, 19).Value = "Standard"
    '   Invoice Number
    Cells(12, 21).Value = "xxx"
    '   Invoice Date
    Cells(12, 23).Value = Date
    '   GL Date
    Cells(12, 24).Value = Date
     '   Description
    Cells(12, 28).Value = "KRV UPLOAD Labor w/e month/day/year"
     '   Invoice Total
    Cells(12, 29).Value = "TO be ADD"
     '   Pay Group
    Cells(12, 37).Value = "INVOICES"
     '   Tax Country
    Cells(12, 40).Value = "United States"
     '   Currency Code
    Cells(12, 44).Value = "USD"
    
    
    
    


'---------------------------------------------------------------------------------------------------------
  
    
End Sub







