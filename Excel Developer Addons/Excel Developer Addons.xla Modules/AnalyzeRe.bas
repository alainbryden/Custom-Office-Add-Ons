Attribute VB_Name = "AnalyzeRe"
Option Explicit

Public Sub MakeStatic()
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    
    Dim response As VbMsgBoxResult
    response = MsgBox("Warning! The current workbook will be saved before creating " & _
                      "a static copy. Would you like to continue?", vbYesNo)
    
    If response = vbNo Then Exit Sub
    wb.Save
    
    Dim path As String, extpos As Integer, original_path
    original_path = wb.FullName
    path = wb.FullName
    extpos = InStrRev(path, ".")
    path = Mid(path, 1, extpos - 1) & " (Static)." & Right(path, Len(path) - extpos)
    wb.SaveAs path
    
    Application.EnableEvents = False
    
    Application.Calculation = xlCalculationManual
    Dim ws As Worksheet, activeWs As Worksheet
    Set activeWs = wb.ActiveSheet
    For Each ws In wb.Worksheets
        Dim prevVisible
        prevVisible = ws.Visible
        ws.Visible = xlSheetVisible
        ws.Activate
        Application.ScreenUpdating = False
        
        Dim FoundCell As Range, FirstMatch As Range
        Set FoundCell = ws.Cells.Find(What:="ARe.", LookIn:=xlFormulas)
        Do Until FoundCell Is Nothing
            If FoundCell.HasFormula Then
                If Not FoundCell.Comment Is Nothing Then FoundCell.Comment.Delete
                If Not FoundCell.HasArray Then
                    FoundCell.AddComment "Cell formula is:" & vbNewLine & FoundCell.Formula
                    FoundCell.Value = FoundCell.value2
                Else
                    FoundCell.AddComment "Cells " & FoundCell.CurrentArray.Address & _
                        " array-formula is:" & vbNewLine & FoundCell.FormulaArray
                    FoundCell.CurrentArray.Formula = FoundCell.CurrentArray.value2
                End If
            Else
                'Some cells contain text that matches the find, remember the first of these
                If FirstMatch Is Nothing Then
                    Set FirstMatch = FoundCell
                Else
                    'If we've looped back to the first match, we're done
                    If FirstMatch = FoundCell Then Exit Do
                End If
            End If
            Set FoundCell = ws.Cells.FindNext(FoundCell)
        Loop
        
        ws.Visible = prevVisible
        Application.ScreenUpdating = True
    Next ws
    
    activeWs.Activate

    Application.Calculation = xlCalculationAutomatic
    wb.Application.DisplayAlerts = False
    wb.SaveAs path, ConflictResolution:=xlLocalSessionChanges
    wb.Application.DisplayAlerts = True
    Application.EnableEvents = True
    
    Workbooks.Open original_path
    wb.Close SaveChanges:=False
End Sub
