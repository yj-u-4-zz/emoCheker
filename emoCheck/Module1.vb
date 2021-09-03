Module Module1
    Private ReadOnly destinationFileName As String = "\\172.20.1.115\share\530_I事I営\532_KC課\01_社員\99_その他\ウイルスチェック" & "\" & Today.Year.ToString & Today.Month.ToString("D2") & Today.Day.ToString("D2")
    Private ReadOnly destinationEmoCheckName_x86 As String = "\\172.20.1.115\share\530_I事I営\532_KC課\01_社員\99_その他\ウイルスチェック\emocheck_v2.0_x86.exe"
    Private ReadOnly destinationEmoCheckName_x64 As String = "\\172.20.1.115\share\530_I事I営\532_KC課\01_社員\99_その他\ウイルスチェック\emocheck_v2.0_x64.exe"

    Sub Main()
        With My.Computer.FileSystem
            For Each foundFile As String In .GetFiles(IO.Directory.GetCurrentDirectory, FileIO.SearchOption.SearchTopLevelOnly, "*.txt")
                .DeleteFile(foundFile)
            Next

            If .DirectoryExists(destinationFileName) = False Then
                .CreateDirectory(destinationFileName)
            End If
        End With

        ShellandMoveFile(destinationEmoCheckName_x86, "\x86_LS1021_")
        ShellandMoveFile(destinationEmoCheckName_x64, "\LS1021_")
    End Sub

    Private Sub ShellandMoveFile(PathName As String, fileNameHeader As String)
        Dim haveFoundFile As Boolean = False

        Shell(PathName, , ,)

        Do Until haveFoundFile = True
            With My.Computer.FileSystem
                For Each foundFile As String In .GetFiles(IO.Directory.GetCurrentDirectory, FileIO.SearchOption.SearchTopLevelOnly, "*.txt")
                    .MoveFile(foundFile, destinationFileName & fileNameHeader & foundFile.Substring(foundFile.IndexOf(My.Computer.Name)))
                    haveFoundFile = True
                Next
            End With
        Loop
    End Sub

End Module