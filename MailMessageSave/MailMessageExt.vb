Imports System.IO
Imports System.Net.Mail
Imports System.Reflection
Imports System.Runtime.CompilerServices

Module MailMessageExt
    <Extension()>
    Sub Save(ByVal Message As MailMessage, ByVal FileName As String)
        Dim assembly As Assembly = GetType(SmtpClient).Assembly
        Dim _mailWriterType As Type = assembly.[GetType]("System.Net.Mail.MailWriter")

        Using _fileStream As FileStream = New FileStream(FileName, FileMode.Create)
            Dim _mailWriterContructor As ConstructorInfo = _mailWriterType.GetConstructor(BindingFlags.Instance Or BindingFlags.NonPublic, Nothing, New Type() {GetType(Stream)}, Nothing)
            Dim _mailWriter As Object = _mailWriterContructor.Invoke(New Object() {_fileStream})
            Dim _sendMethod As MethodInfo = GetType(MailMessage).GetMethod("Send", BindingFlags.Instance Or BindingFlags.NonPublic)
            _sendMethod.Invoke(Message, BindingFlags.Instance Or BindingFlags.NonPublic, Nothing, New Object() {_mailWriter, True}, Nothing)
            Dim _closeMethod As MethodInfo = _mailWriter.[GetType]().GetMethod("Close", BindingFlags.Instance Or BindingFlags.NonPublic)
            _closeMethod.Invoke(_mailWriter, BindingFlags.Instance Or BindingFlags.NonPublic, Nothing, New Object() {}, Nothing)
        End Using

    End Sub
End Module
