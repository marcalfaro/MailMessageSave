Imports System.Net.Mail

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Using _testMail As MailMessage = New MailMessage
            With _testMail
                .Body = "This is a test email"
                .To.Add(New MailAddress("email@domain.com"))
                .From = New MailAddress("sender@domain.com")
                .Subject = "Test email"
                '            .Save($"{IO.Path.Combine(Application.StartupPath, "test.eml")}")

                Dim tmpFolder As String = IO.Path.Combine(Application.StartupPath, System.Guid.NewGuid.ToString)
                If Not IO.Directory.Exists(tmpFolder) Then
                    IO.Directory.CreateDirectory(tmpFolder)
                End If

                Using client As SmtpClient = New SmtpClient("mysmtphost")
                    'client.UseDefaultCredentials = true;
                    client.DeliveryMethod = SmtpDeliveryMethod.SpecifiedPickupDirectory
                    client.PickupDirectoryLocation = tmpFolder
                    client.Send(_testMail)
                End Using

                Dim filePath As String = IO.Directory.GetFiles(tmpFolder).Single
                MsgBox(filePath)
            End With
        End Using



    End Sub
End Class
