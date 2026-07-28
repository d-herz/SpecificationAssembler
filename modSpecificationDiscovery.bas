Option Explicit


Public Function GetSpecificationFiles(ByVal specificationFolder As String) As Collection

    Dim specificationFiles As Collection
    Dim fileSystem As Object
    Dim folder As Object
    Dim file As Object
    Dim fileExtension As String

    Set specificationFiles = New Collection

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set folder = fileSystem.GetFolder(specificationFolder)

    For Each file In folder.Files

        fileExtension = LCase(fileSystem.GetExtensionName(file.Name))

        Select Case fileExtension

            Case "doc", "docx"
                specificationFiles.Add Item:=file.Path

        End Select

    Next file

    Set GetSpecificationFiles = specificationFiles

End Function
