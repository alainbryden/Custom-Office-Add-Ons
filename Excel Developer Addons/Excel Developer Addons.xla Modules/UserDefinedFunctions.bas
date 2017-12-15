Attribute VB_Name = "UserDefinedFunctions"
Option Explicit

Const Timeout As Double = 0.0001

Public GoogleCache As Variant

' Joins a range of cells into a string using some delimiter.
Public Function JoinRange(ByRef rng As Range, Optional delim As String = ",") As String
    On Error GoTo tryTranspose
    JoinRange = Join(WorksheetFunction.Transpose(WorksheetFunction.Transpose(rng)), delim)
    Exit Function
tryTranspose:
    JoinRange = Join(WorksheetFunction.Transpose(rng), delim)
End Function

' Hooks Internet Explorer to perform a google search to some question.
Function Google(Optional ByVal search As String = "")
    On Error GoTo onError
    Google = ""
    
    If IsEmpty(GoogleCache) Then Set GoogleCache = CreateObject("Scripting.Dictionary")
    If GoogleCache.Exists(search) Then
        Google = GoogleCache(search)
        GoTo skipCache
    End If
    
    Dim startTime As Date
    startTime = Now()
    
    Dim url As String
    url = "https://www.google.ca/search?q=" & WorksheetFunction.EncodeURL(search)
    
    Dim IE As Variant
    Set IE = CreateObject("InternetExplorer.Application")
    
    With IE
        .Visible = False
        .stop
        .navigate url
        delay 2
        While .Busy
            DoEvents
            If (Now() - startTime > Timeout) Then
                Google = "(Timeout)"
                GoTo skipCache
            End If
        Wend
    
        'delay 1
        
        If IsNull(.Document.GetElementById("searchform")) Then
            Google = "(Offline)"
            GoTo skipCache
        End If
        
        ' A list of google class names that contain "instant answers"
        Dim arrClassses() As String, i As Integer
        arrClassses = Split("_XWk,wob_t,cwcot,vk_bk", ",")
        
        On Error Resume Next
        For i = LBound(arrClassses) To UBound(arrClassses)
            Google = .Document.GetElementsByClassName(arrClassses(i))(0).innerText
            If Google <> "" Then GoTo cache
        Next
        ' Unit conversion answers are hidden in the 'value' of an input box.
        Google = .Document.GetElementById("_Cif").GetElementsByClassName("_eif")(0).Value
        If Google <> "" Then GoTo cache
        
        
        Google = "No easy answer"
    End With
    
cache:
    GoogleCache.Add search, Google
skipCache:
    If Not IsEmpty(IE) Then
        IE.Visible = False
        IE.Quit
    End If
    Set IE = Nothing
    Exit Function
onError:
    Google = "Error: " & Err.Description
End Function

Private Sub delay(seconds As Long)
   Dim endTime As Date
   endTime = DateAdd("s", seconds, Now())
   Do While Now() < endTime
       DoEvents
   Loop
End Sub
