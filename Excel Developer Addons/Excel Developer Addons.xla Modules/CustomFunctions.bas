Attribute VB_Name = "CustomFunctions"
Option Explicit

Public Function JoinRange(ByRef rng As Range, Optional delim As String = ",") As String
    On Error GoTo tryTranspose
    JoinRange = Join(WorksheetFunction.Transpose(WorksheetFunction.Transpose(rng)), delim)
    Exit Function
tryTranspose:
    JoinRange = Join(WorksheetFunction.Transpose(rng), delim)
End Function
