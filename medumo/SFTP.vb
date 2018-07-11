Imports WinSCP
Imports System.Data.SqlClient
Imports System.Text
Imports System.IO
Imports System.Net

Public Class SFTP
    Public Const fPath = "D:\medumofiles\"

    Public Shared Function UploadCSV() As Boolean
        Dim DidWorkCorrectly As Boolean


        Try
            Dim TPlate As XElement = <q><![CDATA[UPDATE q
set [HasBeenUploaded] = {0}, [HasBeenUploadedTS] = getdate()
FROM [dbo].[MEDUMO_Current_Appts] q INNER JOIN [dbo].[vMEDUMO_Current_Appts_For_Upload] z
ON z.id = q.id]]></q>
            Dim sql As String = Nothing
            Dim DB As New medumoDataContext
            Dim SourceUploadFilePath As String = SFTP.WriteFile(SFTP.CreateCSVText)
            Dim CurrentSessionOptions As New SessionOptions With {.Protocol = Protocol.Sftp, .HostName = "medumo.brickftp.com", .UserName = "urology@ucsf", .Password = "Smspilot9", .SshHostKeyFingerprint = "ssh-rsa 4096 JvS7SrgY9QfsC2otdG0TGo0aWcvvieGg1R2Vx8/5VSw="}
            CurrentSessionOptions.AddRawSettings("ProxyPort", "1")
            Dim Currentsession As New Session
            Currentsession.Open(CurrentSessionOptions)
            Try
                Dim Settings_Upload As New TransferOptions With {.TransferMode = TransferMode.Ascii}
                Dim tr As TransferOperationResult
                tr = Currentsession.PutFiles(SourceUploadFilePath, "/", False, Settings_Upload)
                tr.Check()
                For Each trans In tr.Transfers
                    Dim log As New DataTransfer With {.ActionName = "Uploade SMS Appt List", .ActivityType = "SFTP File Upload", .SucceedFail = True, .ResultDesc = SourceUploadFilePath}
                    DB.DataTransfers.InsertOnSubmit(log)
                    sql = String.Format(TPlate.Value, 1)
                Next
                NotfiySuccess(String.Format("File {0} Posted to FTP Site at {1:D}", IO.Path.GetFileName(SourceUploadFilePath), Now), True)
            Catch ex As Exception
                Dim log As New DataTransfer With {.ActionName = "Uploade SMS Appt List", .ActivityType = "SFTP File Upload", .SucceedFail = False, .ResultDesc = "!!!!!!!  FAIL !!!!!!!!!!! " & SourceUploadFilePath}
                DB.DataTransfers.InsertOnSubmit(log)
                sql = String.Format(TPlate.Value, 1)

            End Try

            DB.SubmitChanges()
            Using cnt As New SqlConnection("Data Source=ccwp10;Initial Catalog=DataInventory;Integrated Security=True")
                cnt.Open()
                Using cmd As New SqlCommand(sql, cnt)
                    cmd.ExecuteNonQuery()
                End Using
            End Using
            Currentsession.Close()
            DidWorkCorrectly = True
        Catch ex As Exception
            DidWorkCorrectly = False
        End Try
        Return DidWorkCorrectly

    End Function
    Public Shared Function CreateCSVText() As String
        Dim SB As New StringBuilder
        Dim lb As New StringBuilder


        Dim lx As String = "MRN,FirstName,LastName,DOB,Sex,Zip,LanguagePrimary,LanguageCare,LanguageWritten,InterpreterNeeded,EnglishFluency,PrefCommunicationMethod,SendSMS,PhoneHome,PhoneWork,PhoneMobile,Email,StartDateTime,DurationMinutes,EndDateTime,ApptType,Provider,ApptStatus"

        SB.AppendLine(lx)
        Dim CHeads() As String = lx.Split(",")

        Dim sql As XElement = <q><![CDATA[SELECT [ID]
      ,[MRN]
      ,[FirstName]
      ,[LastName]
      ,[DOB]
      ,[Sex]
      ,[Zip]
      ,[LanguagePrimary]
      ,[LanguageCare]
      ,[LanguageWritten]
      ,[InterpreterNeeded]
      ,[EnglishFluency]
      ,[PrefCommunicationMethod]
      ,[SendSMS]
      ,[PhoneHome]
      ,[PhoneWork]
      ,[PhoneMobile]
      ,[Email]
      ,[StartDateTime]
      ,[DurationMinutes]
      ,[EndDateTime]
      ,[ApptType]
      ,[Provider]
      ,[ApptStatus]
      ,[HasBeenUploaded]
      ,[HasBeenUploadedTS]
  FROM [dbo].[vMEDUMO_Current_Appts_For_Upload]]]></q>
        Using cnt As New SqlConnection("Data Source=ccwp10;Initial Catalog=DataInventory;Integrated Security=True")
            cnt.Open()
            Using cmd As New SqlCommand(sql.Value, cnt)
                Dim R As IDataReader = cmd.ExecuteReader
                While R.Read
                    lb.Clear()
                    Dim IsFirst As Boolean = True
                    For i As Integer = 1 To R.FieldCount - 3
                        If IsFirst Then
                            IsFirst = False
                            lb.Append(R(i))
                        Else
                            lb.Append(",")
                            If Not String.IsNullOrEmpty(R(i).ToString) Then
                                If CType(R(i), String).Contains(",") Then
                                    lb.Append(Chr(34))
                                    lb.Append(R(i).ToString.Trim)
                                    lb.Append(Chr(34))
                                Else
                                    lb.Append(R(i).ToString.Trim)
                                End If
                            End If
                        End If
                    Next
                    Dim Q() As String = lb.ToString.Split(",")
                    SB.AppendLine(lb.ToString)
                End While
            End Using
        End Using
        Return SB.ToString
    End Function

    Public Shared Function WriteFile(vCSV As String) As String
        Dim FileNameTemplate As String = String.Format("SMS_biopsy_{0:yyyy-MM-dd}.csv", Today)
        Dim Fname As String = String.Format("{0}{1}", fPath, FileNameTemplate)
        My.Computer.FileSystem.WriteAllText(Fname, vCSV, False)
        Return Fname
    End Function


    '    // Set up session options
    'SessionOptions sessionOptions = New SessionOptions
    '{
    '    Protocol = Protocol.Sftp,
    '    HostName = "medumo.brickftp.com",
    '    UserName = "urology@ucsf",
    '    Password = "Smspilot9",
    '    SshHostKeyFingerprint = "ssh-rsa 4096 JvS7SrgY9QfsC2otdG0TGo0aWcvvieGg1R2Vx8/5VSw=",
    '};

    'sessionOptions.AddRawSettings("ProxyPort", "1");

    'Using (Session session = New Session())
    '{
    '    // Connect
    '    session.Open(sessionOptions);

    '    // Your code
    '}


    Public Shared Sub NotfiySuccess(vMessage As String, DiditWork As Boolean)

        Dim URL As String = "https://ucsfurology.slack.com/messages/D9DMETY0N/files/FASFN5551/"
        URL = "https://hooks.slack.com/services/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"


        Dim webrequest As WebRequest = WebRequest.Create(URL)

        webrequest.Method = "POST"
        webrequest.ContentType = "application/json"
        Dim mydata As String
        If DiditWork Then
            mydata = "{'text':'" & vMessage & "', 'icon_emoji': ':bell:','username':'Notification'}"

        Else
            mydata = "{'text':'" & vMessage & "', 'icon_emoji': ':no_entry_sign:','username':'Notification'}"

        End If
        'Dim mydata As String = "{'text':'" & vMessage & "', 'icon_emoji': ':bell:','username':'Notification'}"


        webrequest.GetRequestStream.Write(System.Text.Encoding.UTF8.GetBytes(mydata), 0, System.Text.Encoding.UTF8.GetBytes(mydata).Count)

        Dim webresponse As HttpWebResponse = webrequest.GetResponse()

        Dim stream As System.IO.Stream = webresponse.GetResponseStream()
        Dim reader As New StreamReader(stream, Encoding.UTF8)
        Dim contents As String = reader.ReadToEnd()
        Debug.Print(contents)

    End Sub


End Class
