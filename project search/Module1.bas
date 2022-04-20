Attribute VB_Name = "Module1"
Sub delete_results()
'
' delete results
'

    With Application
    .StatusBar = "Wait"
    .ScreenUpdating = False
    End With
    
    Range("A18").Select
    Selection.CurrentRegion.Select
    Selection.Offset(rowOffset:=1).Select
    Selection.Delete
    Range("A1").Select
    

    With Application
        .ScreenUpdating = True
        .CutCopyMode = False
        .StatusBar = False
    End With


    MsgBox ("Done")

    

End Sub



Sub find_project_details()
'
' find_project_details
'

    With Application
    .StatusBar = "Wait"
    .ScreenUpdating = False
    End With
    
    
    iter_cell = 4
    szukany_projekt = Range("B" & iter_cell).Value
    
    Do While (szukany_projekt <> "")
    
    '---------find last row
        last_row = Application.ActiveSheet.UsedRange.Rows.Count
        last_row_plus_1 = last_row + 1
        
        
    '--------Szukanie danych z an data
        looking_value_row = -1
            
        On Error GoTo Leavex
            Sheets("an data").Select
            Cells.Find(What:=szukany_projekt, After:=ActiveCell, LookIn:=xlFormulas2, _
                LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                MatchCase:=True, SearchFormat:=False).Activate
           
            'wiersz znalezionego projektu
            looking_value_row = ActiveCell.Row
                
            Sheets("Project search").Select
            Range("C" & iter_cell).Value = "ok"
            
            Range("A" & last_row_plus_1).Value = szukany_projekt
Leavex:
    
        'If not found goto next iteration
        If looking_value_row = "-1" Then
            Sheets("Project search").Select
            Range("C" & iter_cell).Value = "not found"
            GoTo not_found
        End If
    
    
        'Sub one_cell_copy(to_zm As String, from_zm As String)
        Call one_cell_copy("C" & last_row_plus_1, "M" & looking_value_row)          'Project Accountant
        Call one_cell_copy("D" & last_row_plus_1, "L" & looking_value_row)          'Project Manager
        Call one_cell_copy("E" & last_row_plus_1, "S" & looking_value_row)          'Allow Cross Charge Flag
        Call one_cell_copy("F" & last_row_plus_1, "B" & looking_value_row)          'Project Status
        Call one_cell_copy("G" & last_row_plus_1, "R" & looking_value_row)          'Project Currency
        Call one_cell_copy("H" & last_row_plus_1, "Q" & looking_value_row)          'Project OU
        Call one_cell_copy("I" & last_row_plus_1, "F" & looking_value_row)          'Legal Entity
        Call one_cell_copy("J" & last_row_plus_1, "U" & looking_value_row)          'Project Creation Date
        Call one_cell_copy("K" & last_row_plus_1, "C" & looking_value_row)          'Project Closed Date
        Call one_cell_copy("L" & last_row_plus_1, "D" & looking_value_row)          'Class-Controlling PU
    
    
        
        '-----------------------Tasks-------------------
        
        If Range("C1").Value = "no" Or Range("C1").Value = "NO" Then
            GoTo leave_task
        End If
        
        Sheets("PROJECT_TASK_FILE").Select
        
        On Error GoTo Leave_task_2
        Cells.Find(What:=szukany_projekt, After:=ActiveCell, LookIn:=xlFormulas2, _
                LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                MatchCase:=True, SearchFormat:=False).Activate
            
        Range("B1").Select
        
        Selection.AutoFilter
        ActiveSheet.ListObjects("PROJECT_TASK_FILE").Range.AutoFilter Field:=1, _
            Criteria1:=szukany_projekt
    
        Selection.Offset(rowOffset:=1).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy
        Sheets("Project search").Select
        Range("B" & last_row_plus_1).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            
        
        Application.CutCopyMode = False
        Sheets("PROJECT_TASK_FILE").Select
        Range("A1").Select
        Selection.AutoFilter
        Sheets("Project search").Select
        Range("A1").Select
        
        '-----------------------Tasks-------------------^^
    
    
Leave_task_2:
leave_task:
        Sheets("Project search").Select

not_found:

        On Error GoTo -1 'clean error GoTO
        iter_cell = iter_cell + 1
        szukany_projekt = Range("B" & iter_cell).Value
    
    Loop
    
    
    With Application
        .ScreenUpdating = True
        .CutCopyMode = False
        .StatusBar = False
    End With


    MsgBox ("Done")

End Sub




Sub one_cell_copy(to_zm As String, from_zm As String)
    
    Sheets("an data").Select
    Range(from_zm).Select
    Selection.Copy
    Sheets("Project search").Select
    Range(to_zm).PasteSpecial (xlPasteAll)
    Application.CutCopyMode = False
    
    Sheets("Project search").Range("A1").Select


End Sub










