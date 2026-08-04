Option Explicit


'==============================================================================
' Project: Specification Assembler
' Author: Daniel Hyres
' Date: 08/2027
'
' Description:
'   Assembles multiple CTDOT specification word documents into a single master specification
'   book while preserving formatting, headers, footers, adding page numbering and a TOC.
'
'
' Version History
'
' 0.1.0
'   - Project initialized
'   - Folder picker implemented
'   - Specification discovery implemented
'
'
' Known Issues
' - Heading 1 style definition not yet imported from source document
' - Page numbering not yet implemented
' - Automatic PDF export not yet implemented
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

    'Debug.Print templatePath

    Debug.Print "Found " & specificationFiles.Count & " specification(s)."

    For Each specificationFile In specificationFiles

        Debug.Print specificationFile

    Next specificationFile
    
    Set outputDocument = CreateOutputDocument( _
    templatePath:=templatePath)
    
    Dim temporaryOutputPath As String

    temporaryOutputPath = GetTemporaryOutputPath()

    outputDocument.SaveAs2 _
        FileName:=temporaryOutputPath, _
        FileFormat:=wdFormatXMLDocument
    
    '
    ImportDocumentStyles _
        outputDocument:=outputDocument, _
        firstSpecificationPath:=specificationFiles(1)
    '
    
    InitializeFrontMatter _
    outputDocument:=outputDocument

    BuildSpecificationBook _
        outputDocument:=outputDocument, _
        specificationFiles:=specificationFiles
        
        
    UpdateTableOfContents _
        outputDocument:=outputDocument
        
        
    SaveOutputDocument _
        outputDocument:=outputDocument, _
        specificationFolder:=specificationFolder
        
    DeleteTemporaryOutputFile
        

    MsgBox "Master specification created successfully."

End Sub




