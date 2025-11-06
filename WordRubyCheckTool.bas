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

    CollectMainStoryText tempDoc, records
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

Private Sub CollectMainStoryText(ByVal targetDoc As Document, ByRef records As Collection)
    InspectTextRange targetDoc.Content, "Document", records
End Sub

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
        recordItem.RubyPresence = "NG"
        recordItem.RubyFontName = ""
        recordItem.RubyFontSize = ""
        recordItem.ObjectType = objectLabel
        recordItem.Notes = ""
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
    recordItem.RubyPresence = "NG"
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
    Dim characterIndex As Long
    Dim characterRange As Range
    Dim characterText As String
    Dim recordItem As RubyCheckRecord
    Dim rubyFontName As String
    Dim rubyFontSizeValue As Variant
    Dim rubyFontSizeText As String
    Dim rubyPresence As String
    Dim rubyText As String

    If segmentRange.Characters.Count = 0 Then Exit Sub

    For characterIndex = 1 To segmentRange.Characters.Count
        Set characterRange = segmentRange.Characters(characterIndex).Duplicate
        characterText = NormalizeCharacterText(characterRange.Text)

        If Len(characterText) = 0 Then GoTo ContinueNextCharacter

        rubyText = ExtractRubyDetails(characterRange, rubyFontName, rubyFontSizeValue, characterText)

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

        Set recordItem = New RubyCheckRecord
        recordItem.No = records.Count + 1
        recordItem.PageNumber = characterRange.Information(wdActiveEndAdjustedPageNumber)
        recordItem.TargetText = characterText
        recordItem.RubyText = rubyText
        recordItem.RubyPresence = rubyPresence
        recordItem.RubyFontName = rubyFontName
        recordItem.RubyFontSize = rubyFontSizeText
        recordItem.ObjectType = objectLabel
        recordItem.Notes = ""

        records.Add recordItem

ContinueNextCharacter:
    Next characterIndex
End Sub

Private Function NormalizeCharacterText(ByVal rawText As String) As String
    Dim cleaned As String

    cleaned = Replace(rawText, vbCr, "")
    cleaned = Replace(cleaned, vbLf, "")
    cleaned = Replace(cleaned, Chr$(0), "")
    cleaned = Replace(cleaned, Chr$(7), "")

    NormalizeCharacterText = cleaned
End Function

Private Function ExtractRubyDetails(ByVal segmentRange As Range, ByRef rubyFontName As String, ByRef rubyFontSizeValue As Variant, ByRef baseText As String) As String
    Dim rubyObject As Object
    Dim rubyText As String
    Dim baseRange As Range

    rubyFontName = ""
    rubyFontSizeValue = Null

    On Error Resume Next
    Set rubyObject = CallByName(segmentRange, "Ruby", VbGet)
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
    rubyText = rubyObject.RubyText
    If Err.Number <> 0 Then
        rubyText = ""
        Err.Clear
    End If

    If Len(baseText) = 0 Then
        Set baseRange = rubyObject.Base.Duplicate
        baseRange.SetRange Start:=segmentRange.Start, End:=segmentRange.End
        baseText = NormalizeCharacterText(baseRange.Text)
        If Err.Number <> 0 Then
            baseText = ""
            Err.Clear
        End If
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

    header = "No." & vbTab & "Page" & vbTab & "対象文字" & vbTab & _
             "ルビ有無" & vbTab & "フォント名" & vbTab & "ルビサイズ" & vbTab & "オブジェクト種別" & vbTab & "備考"

    Debug.Print header
    messageLines.Add header

    For Each recordItem In records
        line = recordItem.No & vbTab & _
               recordItem.PageNumber & vbTab & _
               recordItem.TargetText & vbTab & _
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
