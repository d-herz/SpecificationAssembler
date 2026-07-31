Option Explicit

Public Sub ImportDocumentStyles( _
    ByVal outputDocument As Document, _
    ByVal firstSpecificationPath As String)

    Application.OrganizerCopy _
        Source:=firstSpecificationPath, _
        Destination:=outputDocument.FullName, _
        Name:="Heading 1", _
        Object:=wdOrganizerObjectStyles

End Sub

