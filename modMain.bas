Option Explicit


'==============================================================================
' Project: Specification Assembler
' Author: Daniel Hyres
' Date: 08/2027
'
' Description:
'   Assembles multiple CTDOT specification word documents into a single master specification
'   book while preserving formatting, headers, footers, and page numbering.
'
'
' Version History
'
' 0.1.0
'   - Project initialized
'   - Folder picker implemented
'   - Specification discovery implemented

'==============================================================================


Public Sub Main()

    Dim specificationFolder As String
    Dim specificationFiles As Collection
    Dim specificationFile As Variant
    Dim outputDocument As Document

    specificationFolder = GetSpecificationFolder()
    

    If Len(specificationFolder) = 0 Then

        MsgBox "Operation cancelled."

        Exit Sub

    End If

    Set specificationFiles = GetSpecificationFiles(specificationFolder)
    
    Dim templatePath As String

    templatePath = GetDocumentTemplatePath(CStr(specificationFiles(1)))

    Debug.Print templatePath

    Debug.Print "Found " & specificationFiles.Count & " specification(s)."

    For Each specificationFile In specificationFiles

        Debug.Print specificationFile

    Next specificationFile
    
    Set outputDocument = CreateOutputDocument( _
    templatePath:=templatePath)

    BuildSpecificationBook _
        outputDocument:=outputDocument, _
        specificationFiles:=specificationFiles
        
        
    SaveOutputDocument _
        outputDocument:=outputDocument, _
        specificationFolder:=specificationFolder

    MsgBox "Master specification created successfully."

End Sub




