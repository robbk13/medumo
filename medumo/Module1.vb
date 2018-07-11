Imports WinSCP
Module Module1

    Sub Main()




        'Dim CurrentSessionOptions As New SessionOptions With {.Protocol = Protocol.Sftp, .HostName = "medumo.brickftp.com", .UserName = "urology@ucsf", .Password = "Smspilot9", .SshHostKeyFingerprint = "ssh-rsa 4096 JvS7SrgY9QfsC2otdG0TGo0aWcvvieGg1R2Vx8/5VSw=", .TimeoutInMilliseconds = 10000}
        'CurrentSessionOptions.AddRawSettings("ProxyPort", "1")
        'Dim Currentsession As New Session With {
        '    .SessionLogPath = "D:\medumofiles\log\logSftp.log",
        '    .DebugLogPath = "D:\medumofiles\debug\logSftp.log"
        '}

        'Dim p As String = Currentsession.SessionLogPath
        'Dim d As String = Currentsession.DebugLogPath
        'Dim q As String = Currentsession.ExecutablePath



        'Currentsession.Open(CurrentSessionOptions)


        Dim DidItWork As Boolean = SFTP.UploadCSV()


        If DidItWork = False Then
            SFTP.NotfiySuccess(String.Format("ERROR: File {0} NOT to FTP Site at {1:D} - Maintenance Required.", IO.Path.GetFileName("D:\medumofiles\SMS_biopsy_2018-06-27.csv"), Now), DidItWork)

        Else
            'SFTP.NotfiySuccess(String.Format("File {0} Posted to FTP Site at {1:D}", IO.Path.GetFileName("D:\medumofiles\SMS_biopsy_2018-06-27.csv"), Now))

        End If


    End Sub

End Module
