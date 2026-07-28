Option Explicit


Public Function CreateOutputDocument( _
    ByVal templatePath As String) As Document

    Dim outputDocument As Document

   Set outputDocument = Documents.Add

    Set CreateOutputDocument = outputDocument

End Function


Public Sub SaveOutputDocument( _
    ByVal outputDocument As Document, _
    ByVal specificationFolder As String)

    Dim saveFileDialog As FileDialog
    Dim outputFilePath As String
    Dim defaultFileName As String
    Dim initialSavePath As String
    
    defaultFileName = _
    "Master Specifications_" & Format(Date, "yyyy-mm-dd") & ".docx"
    
    initialSavePath = _
    GetParentFolder(specificationFolder) & "\" & defaultFileName

    Set saveFileDialog = Application.FileDialog(msoFileDialogSaveAs)

    With saveFileDialog

        .Title = "Save Master Specification"

        .InitialFileName = initialSavePath

        If .Show <> -1 Then

            MsgBox "Save cancelled."

            Exit Sub

        End If

        outputFilePath = .SelectedItems(1)

    End With

    If LCase$(Right$(outputFilePath, 5)) <> ".docx" Then

        outputFilePath = outputFilePath & ".docx"

    End If

    outputDocument.SaveAs2 _
        FileName:=outputFilePath, _
        FileFormat:=wdFormatXMLDocument

End Sub









'=========================================================================
' Legacy prototype function.
' Used during template investigation.
' Currently not used by the assembler.
'=========================================================================
Public Function GetDocumentTemplatePath( _
    ByVal specificationPath As String) As String

    Dim sourceDocument As Document

    Debug.Print "Opening:"
    Debug.Print specificationPath

    Set sourceDocument = Documents.Open( _
        FileName:=specificationPath, _
        AddToRecentFiles:=False)

    Debug.Print "Opened successfully."

    Debug.Print "Template:"
    Debug.Print sourceDocument.AttachedTemplate.FullName

    GetDocumentTemplatePath = sourceDocument.AttachedTemplate.FullName

    sourceDocument.Close SaveChanges:=False

End Function

