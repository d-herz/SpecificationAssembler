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
    
    destinationRange.PasteAndFormat wdFormatOriginalFormatting

    sourceDocument.Close SaveChanges:=False

End Sub

