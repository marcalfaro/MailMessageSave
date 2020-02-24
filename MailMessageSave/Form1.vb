Option Explicit On
Imports System.Net.Mail

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Using _testMail As MailMessage = New MailMessage
            With _testMail
                'mark as draft
                .Headers.Add("X-Unsent", "1")

                .From = New MailAddress("sender@domain.com")
                .To.Add(New MailAddress("email@domain.com"))
                .Subject = "Test email"
                .Body = "This is a test email"


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
