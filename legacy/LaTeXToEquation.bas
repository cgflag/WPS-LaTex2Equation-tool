Attribute VB_Name = "LaTeXToEquation"
Option Explicit

' LaTeX $...$ batch converter for Word / WPS
' ASCII-only strings for WPS VBA encoding compatibility

Private Const HIT_SEP As String = "||HIT||"

Public Sub ConvertAllLaTeXFormulas()
    Dim doc As Document
    Dim hits As Collection
    Dim i As Long
    Dim okCount As Long
    Dim failCount As Long
    Dim displayNum As Long
    Dim hitData As String
    Dim parts() As String
    
    If ActiveDocument Is Nothing Then
        MsgBox "Open a document first.", vbExclamation
        Exit Sub
    End If
    
    Set doc = ActiveDocument
    Set hits = CollectFormulaHits(doc)
    
    If hits.Count = 0 Then
        MsgBox "No LaTeX formulas found ($...$ or $$...$$).", vbInformation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    On Error GoTo CleanUp
    
    For i = hits.Count To 1 Step -1
        hitData = hits(i)
        parts = Split(hitData, HIT_SEP)
        
        If ConvertOneHit(doc, CLng(parts(0)), CLng(parts(1)), parts(2), CBool(parts(3)), displayNum) Then
            okCount = okCount + 1
        Else
            failCount = failCount + 1
        End If
    Next i
    
CleanUp:
    Application.ScreenUpdating = True
    
    If Err.Number <> 0 Then
        MsgBox "Runtime error: " & Err.Number & vbCrLf & Err.Description, vbCritical
    Else
        MsgBox "Done." & vbCrLf & "OK: " & okCount & vbCrLf & "Skipped: " & failCount, vbInformation
    End If
End Sub

Private Function CollectFormulaHits(ByVal doc As Document) As Collection
    Dim col As Collection
    Dim para As Paragraph
    Dim paraText As String
    Dim regex As Object
    Dim m As Object
    Dim latex As String
    Dim isDisplay As Boolean
    
    Set col = New Collection
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = False
    regex.Pattern = "\$\$([^$]+)\$\$|\$([^$]+)\$"
    
    For Each para In doc.Paragraphs
        If ShouldProcessParagraph(para) Then
            paraText = ParaTextWithoutMark(para)
            
            For Each m In regex.Execute(paraText)
                If Len(m.SubMatches(0)) > 0 Then
                    latex = CStr(m.SubMatches(0))
                    isDisplay = IsDisplayParagraph(paraText, m.Value, latex, True)
                Else
                    latex = CStr(m.SubMatches(1))
                    isDisplay = IsDisplayParagraph(paraText, m.Value, latex, False)
                End If
                
                col.Add CStr(para.Range.Start + m.FirstIndex) & HIT_SEP & _
                          CStr(para.Range.Start + m.FirstIndex + m.Length) & HIT_SEP & _
                          latex & HIT_SEP & CStr(isDisplay)
            Next m
        End If
    Next para
    
    Set CollectFormulaHits = col
End Function

Private Function ShouldProcessParagraph(ByVal para As Paragraph) As Boolean
    Dim rng As Range
    Dim inTable As Boolean
    Dim shapeCount As Long
    
    ShouldProcessParagraph = False
    
    On Error Resume Next
    inTable = para.Range.Information(wdWithInTable)
    On Error GoTo 0
    If inTable Then Exit Function
    
    Set rng = para.Range
    
    On Error Resume Next
    If rng.StoryType <> wdMainTextStory Then
        On Error GoTo 0
        Exit Function
    End If
    
    shapeCount = rng.ShapeRange.Count
    On Error GoTo 0
    If shapeCount > 0 Then Exit Function
    
    ShouldProcessParagraph = True
End Function

Private Function ParaTextWithoutMark(ByVal para As Paragraph) As String
    Dim t As String
    t = para.Range.Text
    If Len(t) > 0 Then
        If Right(t, 1) = vbCr Then t = Left(t, Len(t) - 1)
        If Right(t, 1) = vbLf Then t = Left(t, Len(t) - 1)
    End If
    ParaTextWithoutMark = t
End Function

Private Function IsDisplayParagraph(ByVal paraText As String, ByVal fullMatch As String, _
                                    ByVal latex As String, ByVal isDoubleDelim As Boolean) As Boolean
    Dim trimmed As String
    trimmed = Trim(paraText)
    
    If isDoubleDelim Then
        IsDisplayParagraph = (trimmed = fullMatch)
    Else
        IsDisplayParagraph = (trimmed = fullMatch) Or (trimmed = "$" & latex & "$")
    End If
End Function

Private Function ConvertOneHit(ByVal doc As Document, ByVal rangeStart As Long, ByVal rangeEnd As Long, _
                             ByVal latex As String, ByVal isDisplay As Boolean, _
                             ByRef displayNum As Long) As Boolean
    Dim rng As Range
    Dim para As Paragraph
    Dim savedText As String
    Dim oMath As Object
    Dim converted As Range
    
    Set rng = doc.Range(rangeStart, rangeEnd)
    savedText = rng.Text
    
    On Error GoTo Restore
    
    rng.Text = latex
    
    Set converted = rng.OMaths.Add(rng)
    Set oMath = converted.OMaths(1)
    oMath.BuildUp
    
    If isDisplay Then
        displayNum = displayNum + 1
        Set para = converted.Paragraphs(1)
        FormatDisplayEquation para, displayNum
    End If
    
    ConvertOneHit = True
    Exit Function
    
Restore:
    On Error Resume Next
    rng.Text = savedText
    On Error GoTo 0
    ConvertOneHit = False
End Function

Private Sub FormatDisplayEquation(ByVal para As Paragraph, ByVal eqNum As Long)
    Dim pageWidth As Single
    Dim rightMargin As Single
    Dim centerPos As Single
    Dim rightPos As Single
    Dim rng As Range
    
    pageWidth = para.Range.PageSetup.PageWidth
    rightMargin = para.Range.PageSetup.RightMargin
    centerPos = pageWidth / 2
    rightPos = pageWidth - rightMargin
    
    Set rng = para.Range
    If Len(rng.Text) > 0 Then
        If Right(rng.Text, 1) = vbCr Then rng.End = rng.End - 1
        If Right(rng.Text, 1) = vbLf Then rng.End = rng.End - 1
    End If
    
    rng.InsertBefore vbTab
    rng.Collapse wdCollapseEnd
    rng.InsertAfter vbTab & "(" & eqNum & ")"
    
    With para
        .Alignment = wdAlignParagraphLeft
        .SpaceBefore = 6
        .SpaceAfter = 6
        .TabStops.ClearAll
        .TabStops.Add Position:=centerPos, Alignment:=wdAlignTabCenter
        .TabStops.Add Position:=rightPos, Alignment:=wdAlignTabRight
    End With
End Sub

Public Sub ConvertLaTeXInSelection()
    Dim rng As Range
    Dim regex As Object
    Dim m As Object
    Dim hits As Collection
    Dim i As Long
    Dim displayNum As Long
    Dim latex As String
    Dim hitData As String
    Dim parts() As String
    
    Set hits = New Collection
    Set rng = Selection.Range
    If Len(rng.Text) = 0 Then
        MsgBox "Select some text first.", vbExclamation
        Exit Sub
    End If
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.Pattern = "\$\$([^$]+)\$\$|\$([^$]+)\$"
    
    For Each m In regex.Execute(rng.Text)
        If Len(m.SubMatches(0)) > 0 Then
            latex = CStr(m.SubMatches(0))
        Else
            latex = CStr(m.SubMatches(1))
        End If
        
        hitData = CStr(rng.Start + m.FirstIndex) & HIT_SEP & _
                  CStr(rng.Start + m.FirstIndex + m.Length) & HIT_SEP & _
                  latex & HIT_SEP & "False"
        hits.Add hitData
    Next m
    
    If hits.Count = 0 Then
        MsgBox "No LaTeX formulas in selection.", vbInformation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    For i = hits.Count To 1 Step -1
        hitData = hits(i)
        parts = Split(hitData, HIT_SEP)
        ConvertOneHit ActiveDocument, CLng(parts(0)), CLng(parts(1)), parts(2), CBool(parts(3)), displayNum
    Next i
    Application.ScreenUpdating = True
    
    MsgBox "Selection done.", vbInformation
End Sub

Public Sub CheckOMathSupport()
    Dim L1 As String
    Dim L2 As String
    Dim L3 As String
    Dim L4 As String
    Dim L5 As String
    Dim testRng As Range
    Dim testOMath As Object
    Dim oCount As Long
    Dim errNum As Long
    Dim errDesc As String
    Dim ver As Single
    
    If ActiveDocument Is Nothing Then
        MsgBox "Open a document first.", vbExclamation, "OMath Check"
        Exit Sub
    End If
    
    L1 = "App: " & Application.Name
    L2 = "Version: " & Application.Version
    L3 = ""
    L4 = ""
    L5 = ""
    
    ver = Val(Application.Version)
    If ver > 0 And ver < 15 Then
        L5 = "Note: Version < 15 (2013). No LaTeX support. Use convert_latex_docx.py instead."
    End If
    
    Err.Clear
    On Error Resume Next
    
    Set testRng = ActiveDocument.Range(0, 0)
    testRng.Text = "x^2"
    
    errNum = Err.Number
    errDesc = Err.Description
    Err.Clear
    
    If errNum <> 0 Then
        L3 = "OMaths: FAIL (write test range)"
        L4 = "Err " & errNum & ": " & errDesc
    Else
        Set testRng = testRng.OMaths.Add(testRng)
        errNum = Err.Number
        errDesc = Err.Description
        Err.Clear
        
        If errNum <> 0 Then
            L3 = "OMaths: NOT SUPPORTED"
            L4 = "Err " & errNum & ": " & errDesc
            testRng.Text = ""
        Else
            oCount = testRng.OMaths.Count
            
            If oCount < 1 Then
                L3 = "OMaths: STUB (Add OK but Count=0)"
                L4 = "WPS/old Word API stub. VBA macro cannot convert."
                L5 = "Use: python convert_latex_docx.py your.docx"
                testRng.Text = ""
            Else
                Set testOMath = testRng.OMaths(1)
                testOMath.BuildUp
                errNum = Err.Number
                errDesc = Err.Description
                
                If errNum <> 0 Then
                    L3 = "OMaths: PARTIAL (BuildUp failed)"
                    L4 = "Err " & errNum & ": " & errDesc & " Count=" & oCount
                Else
                    L3 = "OMaths: OK - run ConvertAllLaTeXFormulas"
                    L4 = "Test x^2 at doc start (delete manually)"
                End If
            End If
        End If
    End If
    
    On Error GoTo 0
    
    MsgBox L1 & vbCrLf & L2 & vbCrLf & vbCrLf & L3 & vbCrLf & L4 & vbCrLf & L5, vbInformation, "OMath Check"
End Sub
