Option Explicit


Public Function GetSpecificationFolder() As String

    Dim dlg As FileDialog

    Set dlg = Application.FileDialog(msoFileDialogFolderPicker)

    With dlg

        .Title = "Select Folder Containing Specification Documents"

        .AllowMultiSelect = False

        If .Show = -1 Then
            GetSpecificationFolder = .SelectedItems(1)
        Else
            GetSpecificationFolder = vbNullString
        End If

    End With

End Function



Public Function GetParentFolder( _
    ByVal folderPath As String) As String
    
    Dim fileSystem As Object
    Dim folder As Object
    
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set folder = fileSystem.GetFolder(folderPath)
    
    GetParentFolder = folder.ParentFolder.Path
    
End Function


