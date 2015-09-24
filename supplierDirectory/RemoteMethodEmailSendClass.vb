
Imports Contensive.BaseClasses
Imports Contensive.Addons.SupplierDirectory.Constants

Namespace Contensive.Addons.SupplierDirectory
    Public Class RemoteMethodEmailSendClass
        Inherits AddonBaseClass
        '
        ' Entry Point for this Addon
        '
        Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
            Dim toAddress As String = ""
            Dim fromAddress As String = ""
            Dim subject As String = ""
            Dim body As String = ""
            Dim orgid As Integer = 0
            Dim cs As CPCSBaseClass
            Dim testToAddress As String
            '
            Try
                orgid = CP.Utils.EncodeInteger(CP.Request.GetText("organizationid"))
                If orgid = 0 Then
                    '
                    ' Bad organzation ID
                    '
                    Call CP.Utils.AppendLogFile("RemoteMethodEmailSendClass.Execute, OrganizationID=0")
                Else
                    '
                    ' to address from organzation record
                    '
                    cs = CP.CSNew
                    cs.Open("organizations", "id=" & orgid)
                    If cs.OK Then
                        toAddress = cs.GetText("directoryemail")
                    End If
                    Call cs.Close()
                    '
                    ' from address & body from the form
                    '
                    fromAddress = CP.Site.GetProperty("EMAILFROMADDRESS")
                    If fromAddress = "" Then
                        fromAddress = "unknown@unknown.com"
                    End If
                    subject = "Supplier Directory Contact Email"
                    body = "" _
                        & vbCrLf & "This email was created by a visitor to the contact form on " & CP.Site.Domain _
                        & vbCrLf & "" _
                        & vbCrLf & "subject: " & CP.Request.GetText("Subject") _
                        & vbCrLf & "to: " & toAddress _
                        & vbCrLf & "from: " & CP.Request.GetText("email") _
                        & vbCrLf & "" _
                        & vbCrLf & CP.Request.GetText("message")
                    '
                    ' check for test mode - email sent to a custom address
                    '
                    testToAddress = CP.Site.GetProperty("Supplier Directory Test Email Address", "support@contensive.com")
                    If testToAddress <> "" Then
                        '
                        ' test mode
                        '
                        body = "" _
                            & "test mode: email toAddress was " & toAddress _
                            & vbCrLf & body
                        toAddress = testToAddress
                    End If
                    '
                    Call CP.Email.Send(toAddress, fromAddress, subject, body, , False)
                End If
            Catch ex As Exception
                Call CP.Site.ErrorReport("RemoteMethodEmailSendClass.Execute eception, " & ex.Message)
            End Try
            '
            Execute = ""
        End Function

    End Class
End Namespace

