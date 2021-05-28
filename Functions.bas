Attribute VB_Name = "Functions"


Public Declare PtrSafe Function SetCurrentDirectoryA Lib "kernel32" (ByVal lpPathName As String) As Long


Sub Obsluga_pliku()
'
' set up file
'


'pop up to choose file
Dim my_FileName As Variant
    
my_FileName = Application.GetOpenFilename(FileFilter:="Excel Files,*.xl*;*.xm*")
If my_FileName <> False Then
    Workbooks.Open Filename:=my_FileName
End If
   
If my_FileName = False Then
    MsgBox ("wybierz ponownie plik")
    End
End If
    




End Sub


Sub dodawanie_sheetu()
'
' add worksheet
'

Dim sheet As Worksheet
Set sheet = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
Set sheet = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))


End Sub

Sub Kolumn_copy(to_zm As String, from_zm As String)

  Sheets("Upload").Select
  Range(from_zm).Select
  Range(Selection, Selection.End(xlDown)).Select
  Selection.Copy
  Sheets("Upload Form").Select
  Range(to_zm).PasteSpecial (xlPasteValues)
  Application.CutCopyMode = False



End Sub

Sub Kolumn_copy_ARC(to_zm As String, from_zm As String)

    Sheets("Upload").Select
  Range(from_zm).Select
  Range(Selection, Selection.End(xlDown)).Select
  Selection.Copy
  Sheets("UploadForm").Select
  Range(to_zm).PasteSpecial (xlPasteValues)
  Application.CutCopyMode = False



End Sub


Sub FileCopy()
'
' FileCopy
'


FileCopy ("some file path", "some file path")
    




End Sub




