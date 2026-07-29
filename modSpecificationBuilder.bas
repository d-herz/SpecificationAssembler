Option Explicit

Public Sub BuildSpecificationBook( _
    ByVal outputDocument As Document, _
    ByVal specificationFiles As Collection)

    Dim specificationFile As Variant

    Dim isFirstSpecification As Boolean

    isFirstSpecification = True

    For Each specificationFile In specificationFiles

        InsertSpecification _
            outputDocument:=outputDocument, _
            specificationPath:=CStr(specificationFile), _
            insertSectionBreak:=Not isFirstSpecification

        isFirstSpecification = False

    Next specificationFile

End Sub

Public Sub InsertSpecification( _
    ByVal outputDocument As Document, _
    ByVal specificationPath As String, _
    ByVal insertSectionBreak As Boolean)

    Dim sourceDocument As Document
    Dim sourceRange As Range
    Dim destinationRange As Range
    
    Dim insertionStart As Long
    Dim insertionEnd As Long
    

    Set sourceDocument = Documents.Open( _
        FileName:=specificationPath, _
        ReadOnly:=True, _
        AddToRecentFiles:=False)

    Set sourceRange = sourceDocument.Range

    sourceRange.Copy

    Set destinationRange = outputDocument.Range
    destinationRange.Collapse wdCollapseEnd
    
  
    
    destinationRange.InsertBreak _
    Type:=wdSectionBreakNextPage
    
    insertionStart = destinationRange.Start
    
    destinationRange.PasteAndFormat wdFormatOriginalFormatting
    
    insertionEnd = outputDocument.Range.End
    
    ApplyHeadingStyles _
        outputDocument:=outputDocument, _
        rangeStart:=insertionStart, _
        rangeEnd:=insertionEnd
    
    '
    'Debug.Print "Inserted specification:"
    'Debug.Print "Start = " & insertionStart
    'Debug.Print "End   = " & insertionEnd
    '

    sourceDocument.Close SaveChanges:=False

End Sub


Public Sub ApplyHeadingStyles( _
    ByVal outputDocument As Document, _
    ByVal rangeStart As Long, _
    ByVal rangeEnd As Long)

    Dim specificationRange As Range
    Dim paragraph As paragraph

    Set specificationRange = outputDocument.Range( _
        Start:=rangeStart, _
        End:=rangeEnd)

    For Each paragraph In specificationRange.Paragraphs
        If paragraph.Style = outputDocument.Styles(wdStyleNormal) Then

            If paragraph.Range.Font.Bold = True _
            And paragraph.Range.Font.Underline = wdUnderlineSingle Then

                paragraph.Style = outputDocument.Styles(wdStyleHeading1)

            End If

        End If

        'Debug.Print _
            "[" & paragraph.Style & "] " & _
            "Bold=" & paragraph.Range.Font.Bold & _
            ", Underline=" & paragraph.Range.Font.Underline & _
            ", Size=" & paragraph.Range.Font.Size & _
            ", Text=""" & Left$(Replace(paragraph.Range.Text, vbCr, ""), 50) & """"

    Next paragraph

End Sub


