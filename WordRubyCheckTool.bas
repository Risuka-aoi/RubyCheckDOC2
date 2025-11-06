Attribute VB_Name = "WordRubyCheckTool"
Option Explicit

Public Sub RunWordRubyCheck()
    Dim originalDoc As Document
    Dim tempDoc As Document
    Dim docxPath As String
    Dim records As Collection
    Dim messageLines As Collection

    Set originalDoc = ActiveDocument
    Set records = New Collection
    Set messageLines = New Collection

    docxPath = SaveAsTemporaryDocx(originalDoc)

    If docxPath <> "" Then
        Set tempDoc = Documents.Open(FileName:=docxPath, AddToRecentFiles:=False, Visible:=False)
    Else
        Set tempDoc = originalDoc
    End If

    CollectShapeText tempDoc, records
    CollectActiveXText tempDoc, records

    OutputResults records, messageLines

    If docxPath <> "" Then
        tempDoc.Close SaveChanges:=False
        On Error Resume Next
        Kill docxPath
        On Error GoTo 0
    End If
End Sub

Private Function SaveAsTemporaryDocx(ByVal sourceDoc As Document) As String
    Dim docName As String
    Dim tempPath As String
    Dim tempFileName As String

    docName = sourceDoc.Name

    If LCase$(Right$(docName, 4)) = ".doc" Then
        tempPath = Environ$("TEMP")
        If Right$(tempPath, 1) <> "\" Then
            tempPath = tempPath & "\"
        End If
        tempFileName = tempPath & Left$(docName, Len(docName) - 4) & "_rubycheck.docx"
        sourceDoc.SaveCopyAs FileName:=tempFileName, FileFormat:=wdFormatXMLDocument
        SaveAsTemporaryDocx = tempFileName
    Else
        SaveAsTemporaryDocx = ""
    End If
End Function

Private Sub CollectShapeText(ByVal targetDoc As Document, ByRef records As Collection)
    Dim shp As Shape
    Dim objectLabel As String

    For Each shp In targetDoc.Shapes
        If shp.TextFrame.HasText Then
            If shp.Type = msoTextBox Then
                objectLabel = "TextBox"
            Else
                objectLabel = "Shape"
            End If
            InspectTextRange shp.TextFrame.TextRange, objectLabel, records
        End If
    Next shp
End Sub

Private Sub CollectActiveXText(ByVal targetDoc As Document, ByRef records As Collection)
    Dim shp As Shape
    Dim inlineShp As InlineShape

    For Each shp In targetDoc.Shapes
        If shp.Type = msoOLEControlObject Then
            InspectActiveX shp, records
        End If
    Next shp

    For Each inlineShp In targetDoc.InlineShapes
        If inlineShp.Type = wdInlineShapeOLEControlObject Then
            InspectInlineActiveX inlineShp, records
        End If
    Next inlineShp
End Sub

Private Sub InspectActiveX(ByVal shp As Shape, ByRef records As Collection)
    Dim ctrl As Object
    Dim textValue As String
    Dim acquisitionStatus As String

    acquisitionStatus = "ActiveX"

    On Error Resume Next
    Set ctrl = shp.OLEFormat.Object
    If Err.Number <> 0 Then
        textValue = "未取得"
        Err.Clear
    Else
        textValue = ctrl.Text
    End If
    On Error GoTo 0

    If textValue = "未取得" Then
        AppendPlaceholderRecord records, acquisitionStatus, "未取得"
    Else
        InspectPlainText textValue, acquisitionStatus, records
    End If
End Sub

Private Sub InspectInlineActiveX(ByVal inlineShp As InlineShape, ByRef records As Collection)
    Dim ctrl As Object
    Dim textValue As String
    Dim acquisitionStatus As String

    acquisitionStatus = "ActiveX"

    On Error Resume Next
    Set ctrl = inlineShp.OLEFormat.Object
    If Err.Number <> 0 Then
        textValue = "未取得"
        Err.Clear
    Else
        textValue = ctrl.Text
    End If
    On Error GoTo 0

    If textValue = "未取得" Then
        AppendPlaceholderRecord records, acquisitionStatus, "未取得"
    Else
        InspectPlainText textValue, acquisitionStatus, records
    End If
End Sub

Private Sub InspectPlainText(ByVal textValue As String, ByVal objectLabel As String, ByRef records As Collection)
    Dim parts() As String
    Dim i As Long
    Dim recordItem As RubyCheckRecord

    If Len(textValue) = 0 Then Exit Sub

    parts = Split(textValue, vbCrLf)
    For i = LBound(parts) To UBound(parts)
        Set recordItem = New RubyCheckRecord
        recordItem.No = records.Count + 1
        recordItem.PageNumber = 0
        recordItem.TargetText = parts(i)
        recordItem.RubyText = ""
        recordItem.RubyPresence = "未取得"
        recordItem.RubyFontName = ""
        recordItem.RubyFontSize = ""
        recordItem.ObjectType = objectLabel
        recordItem.Notes = "未取得"
        records.Add recordItem
    Next i
End Sub

Private Sub AppendPlaceholderRecord(ByRef records As Collection, ByVal objectLabel As String, ByVal message As String)
    Dim recordItem As RubyCheckRecord

    Set recordItem = New RubyCheckRecord
    recordItem.No = records.Count + 1
    recordItem.PageNumber = 0
    recordItem.TargetText = message
    recordItem.RubyText = ""
    recordItem.RubyPresence = "未取得"
    recordItem.RubyFontName = ""
    recordItem.RubyFontSize = ""
    recordItem.ObjectType = objectLabel
    recordItem.Notes = message

    records.Add recordItem
End Sub

Private Sub InspectTextRange(ByVal targetRange As Range, ByVal objectLabel As String, ByRef records As Collection)
    Dim charIndex As Long
    Dim segmentStart As Long
    Dim segmentRange As Range
    Dim currentUnderline As WdUnderline

    If targetRange.Characters.Count = 0 Then Exit Sub

    segmentStart = 1
    currentUnderline = targetRange.Characters(segmentStart).Font.Underline

    For charIndex = 2 To targetRange.Characters.Count + 1
        If charIndex <= targetRange.Characters.Count Then
            If targetRange.Characters(charIndex).Font.Underline = currentUnderline Then
                GoTo ContinueLoop
            End If
        End If

        If currentUnderline <> wdUnderlineNone Then
            Set segmentRange = targetRange.Duplicate
            segmentRange.SetRange _
                Start:=targetRange.Characters(segmentStart).Start, _
                End:=targetRange.Characters(charIndex - 1).End
            RecordRuby segmentRange, objectLabel, records
        End If

        If charIndex <= targetRange.Characters.Count Then
            segmentStart = charIndex
            currentUnderline = targetRange.Characters(charIndex).Font.Underline
        End If
ContinueLoop:
    Next charIndex
End Sub

Private Sub RecordRuby(ByVal segmentRange As Range, ByVal objectLabel As String, ByRef records As Collection)
    Dim cleanedText As String
    Dim recordItem As RubyCheckRecord
    Dim expectedRuby As String
    Dim notes As String
    Dim rubyPresence As String
    Dim rubyText As String
    Dim rubyFontName As String
    Dim rubyFontSizeValue As Variant
    Dim rubyFontSizeText As String

    cleanedText = Replace(segmentRange.Text, vbCr, "")
    cleanedText = Replace(cleanedText, vbLf, "")
    cleanedText = Trim$(cleanedText)

    If Len(cleanedText) = 0 Then Exit Sub

    expectedRuby = ExpectedRubyFor(cleanedText)
    rubyText = ExtractRubyDetails(segmentRange, rubyFontName, rubyFontSizeValue)

    If rubyText <> "" Then
        rubyPresence = "OK"
    Else
        rubyPresence = "NG"
    End If

    If IsNumeric(rubyFontSizeValue) Then
        rubyFontSizeText = Format$(CDbl(rubyFontSizeValue), "0.##") & "pt"
    Else
        rubyFontSizeText = ""
    End If

    notes = BuildNotes(expectedRuby, rubyText, rubyFontName, rubyFontSizeValue)

    Set recordItem = New RubyCheckRecord
    recordItem.No = records.Count + 1
    recordItem.PageNumber = segmentRange.Information(wdActiveEndAdjustedPageNumber)
    recordItem.TargetText = cleanedText
    recordItem.RubyText = rubyText
    recordItem.RubyPresence = rubyPresence
    recordItem.RubyFontName = rubyFontName
    recordItem.RubyFontSize = rubyFontSizeText
    recordItem.ObjectType = objectLabel
    recordItem.Notes = notes

    records.Add recordItem
End Sub

Private Function ExpectedRubyFor(ByVal targetText As String) As String
    Select Case targetText
        Case "0"
            ExpectedRubyFor = "ｾﾞﾛ"
        Case "O"
            ExpectedRubyFor = "ｵｰ"
        Case "l"
            ExpectedRubyFor = "ｴﾙ"
        Case "I"
            ExpectedRubyFor = "ｱｲ"
        Case "|"
            ExpectedRubyFor = "ﾊﾟｲﾌﾟ"
        Case "1"
            ExpectedRubyFor = "ｲﾁ"
        Case Else
            ExpectedRubyFor = ""
    End Select
End Function

Private Function BuildNotes(ByVal expectedRuby As String, ByVal rubyText As String, ByVal rubyFontName As String, ByVal rubyFontSizeValue As Variant) As String
    Dim noteItems As Collection
    Dim currentNote As Variant
    Dim result As String
    Dim hasRubyFontSize As Boolean

    Set noteItems = New Collection

    If expectedRuby <> "" Then
        If rubyText = "" Then
            noteItems.Add "ルビなし"
        ElseIf rubyText <> expectedRuby Then
            noteItems.Add "期待値: " & expectedRuby
        End If
    Else
        noteItems.Add "対象外"
    End If

    If rubyText <> "" Then
        If Len(rubyFontName) = 0 Then
            noteItems.Add "フォント未設定"
        End If
        If IsNumeric(rubyFontSizeValue) Then
            hasRubyFontSize = (CDbl(rubyFontSizeValue) <> 0)
        Else
            hasRubyFontSize = False
        End If
        If Not hasRubyFontSize Then
            noteItems.Add "サイズ未設定"
        End If
    End If

    If noteItems.Count = 0 Then
        BuildNotes = "-"
    Else
        For Each currentNote In noteItems
            If result = "" Then
                result = CStr(currentNote)
            Else
                result = result & ", " & CStr(currentNote)
            End If
        Next currentNote
        BuildNotes = result
    End If
End Function

Private Function ExtractRubyDetails(ByVal segmentRange As Range, ByRef rubyFontName As String, ByRef rubyFontSizeValue As Variant) As String
    Dim rubyObject As Object
    Dim rubyText As String

    rubyFontName = ""
    rubyFontSizeValue = Null

    On Error Resume Next
    Set rubyObject = segmentRange.Ruby
    If Err.Number <> 0 Then
        Err.Clear
        Set rubyObject = Nothing
    End If
    On Error GoTo 0

    If rubyObject Is Nothing Then
        ExtractRubyDetails = ""
        Exit Function
    End If

    On Error Resume Next
    rubyText = rubyObject.Text
    If Err.Number <> 0 Then
        rubyText = ""
        Err.Clear
    End If

    rubyFontName = rubyObject.Font.Name
    If Err.Number <> 0 Then
        rubyFontName = ""
        Err.Clear
    End If

    rubyFontSizeValue = rubyObject.Font.Size
    If Err.Number <> 0 Then
        rubyFontSizeValue = Null
        Err.Clear
    End If
    On Error GoTo 0

    ExtractRubyDetails = rubyText
End Function

Private Sub OutputResults(ByVal records As Collection, ByVal messageLines As Collection)
    Dim recordItem As RubyCheckRecord
    Dim header As String
    Dim line As String

    header = "No." & vbTab & "Page" & vbTab & "対象文字列" & vbTab & "ルビ" & vbTab & _
             "ルビ有無" & vbTab & "フォント名" & vbTab & "ルビサイズ" & vbTab & "オブジェクト種別" & vbTab & "備考"

    Debug.Print header
    messageLines.Add header

    For Each recordItem In records
        line = recordItem.No & vbTab & _
               recordItem.PageNumber & vbTab & _
               recordItem.TargetText & vbTab & _
               recordItem.RubyText & vbTab & _
               recordItem.RubyPresence & vbTab & _
               recordItem.RubyFontName & vbTab & _
               recordItem.RubyFontSize & vbTab & _
               recordItem.ObjectType & vbTab & _
               recordItem.Notes
        Debug.Print line
        messageLines.Add line
    Next recordItem

    MsgBox JoinCollection(messageLines, vbCrLf), vbInformation, "Wordルビ振りチェックツール"
End Sub

Private Function JoinCollection(ByVal items As Collection, ByVal delimiter As String) As String
    Dim element As Variant

    For Each element In items
        If JoinCollection = "" Then
            JoinCollection = CStr(element)
        Else
            JoinCollection = JoinCollection & delimiter & CStr(element)
        End If
    Next element
End Function
