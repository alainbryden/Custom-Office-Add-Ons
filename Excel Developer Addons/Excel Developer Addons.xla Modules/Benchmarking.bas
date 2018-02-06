Attribute VB_Name = "Benchmarking"
Option Explicit
#If VBA7 Then
    Private Declare PtrSafe Function GetTickCount Lib "Kernel32" ()
#Else
    Private Declare Function GetTickCount Lib "Kernel32"
#End If

'Compare the speed of two functions
Function TestSpeed() As String()
    Dim V1#, V2#
    Dim i As Long
    Dim starting
    Dim delta1, delta2
    Application.ScreenUpdating = False
    starting = GetTickCount
    For i = 0 To 100
        ActiveSheet.Range("O4:O14").Calculate
        Application.Calculation = xlCalculationManual
    Next i
    delta1 = GetTickCount - starting
    starting = GetTickCount
    For i = 0 To 100
        ActiveSheet.Range("P4:P14").Calculate
        Application.Calculation = xlCalculationManual
    Next i
    delta2 = GetTickCount - starting
    Application.ScreenUpdating = True
    MsgBox "1. " & delta1 & "ms taken. Result " & V1 & vbNewLine _
         & "2. " & delta2 & "ms taken. Result " & V2 & vbNewLine
End Function
