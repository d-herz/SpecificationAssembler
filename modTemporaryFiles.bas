Option Explicit


Public Function GetTemporaryOutputPath() As String

    Dim temporaryFolder As String

    temporaryFolder = Environ$("TEMP")

    If Right$(temporaryFolder, 1) <> "\" Then

        temporaryFolder = temporaryFolder & "\"

    End If

    GetTemporaryOutputPath = _
        temporaryFolder & "SpecificationAssembler_Temp.docx"

End Function

