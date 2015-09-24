
Imports Contensive.BaseClasses
Imports Contensive.Addons.SupplierDirectory.Constants
'
Namespace Contensive.Addons.SupplierDirectory
    '
    '
    '
    Public Class EmailNotificationClass
        Inherits AddonBaseClass
        '
        ' Entry Point for this Addon
        '
        Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
            Execute = New String("")
            Try
                Dim hint As String = "init"
                Dim emailCnt As Integer
                Dim adId As Integer
                Dim adName As String
                Dim cs As CPCSBaseClass
                Dim csEmail As CPCSBaseClass
                Dim emailBody As String
                Dim emailToId As Integer
                Dim emailFromAddress As String
                Dim emailSubject As String
                Dim emailTextMemberId As Integer
                Dim adOrganizationId As Integer
                Dim csOrg As CPCSBaseClass
                Dim errorMessage As String = ""
                Dim emailModifiedById As Integer
                Dim emailId As Integer = 0
                Dim emailFromAddressDefault As String = CP.Site.GetProperty("EMAILFROMADDRESS")
                Dim emailTestAddress As String = CP.Site.GetProperty("Supplier Directory Test Email Address", "support@contensive.com")
                Dim TestMode As Boolean = CP.Utils.EncodeBoolean(CP.Site.GetProperty("Supplier Directory Test Mode", "true"))
                '
                emailCnt = CP.Utils.EncodeInteger(CP.Site.GetProperty("reviewemail-cnt"))
                Call CP.Site.SetProperty("reviewemail-cnt", (emailCnt + 1))
                cs = CP.CSNew
                Call cs.Open("Directory Banner Ads", "(SendReviewRequest<>0)")
                hint &= ",100"
                Do While cs.OK
                    hint &= ",200"
                    emailId = cs.GetInteger("ReviewEmailID")
                    If emailId <> 0 Then
                        hint &= ",210"
                        adId = cs.GetInteger("id")
                        adName = cs.GetText("name")
                        emailModifiedById = cs.GetInteger("ModifiedBy")
                        adOrganizationId = cs.GetInteger("OrganizationId")
                        emailToId = 0
                        '
                        csOrg = CP.CSNew
                        Call csOrg.Open("organizations", "id=" & adOrganizationId)
                        If Not csOrg.OK Then
                            errorMessage &= vbCrLf & "Banner Ad #" & adId & ", '" & adName & "' can not send notification email because it does not have a valid organization."
                        Else
                            emailToId = csOrg.GetInteger("directoryAccountContactID")
                        End If
                        Call csOrg.Close()
                        '
                        csEmail = CP.CSNew
                        Call csEmail.Open("system Email", "id=" & emailId)
                        If Not csEmail.OK Then
                            hint &= ",220"
                            errorMessage &= vbCrLf & "Banner Ad #" & adId & ", '" & adName & "' can not send notification email because it does not have a valid system email."
                        Else
                            hint &= ",230"
                            emailBody = csEmail.GetText("CopyFilename")
                            emailBody = Replace(emailBody, "##name##", adName)
                            emailTextMemberId = csEmail.GetInteger("TestMemberID")
                            emailSubject = csEmail.GetText("subject")
                            emailFromAddress = csEmail.GetText("fromAddress")
                            If TestMode Then
                                hint &= ",240, emailTestAddress=[" & emailTestAddress & "]"
                                Call CP.Email.send(emailTestAddress, emailFromAddress, "[Test mode true] " & emailSubject, "[Test Mode, toAddressMemberID=" & emailToId & "]<br>" & emailBody, True, True)
                            Else
                                hint &= ",250"
                                Call CP.Email.sendUser(emailToId, emailFromAddress, emailSubject, emailBody, True, True)
                            End If
                        End If
                        Call csEmail.Close()
                        Call CP.Site.SetProperty("reviewemail-id", emailId)
                        hint &= ",260"
                        If errorMessage <> "" Then
                            hint &= ",270"
                            If TestMode Then
                                hint &= ",280"
                                Call CP.Email.send(emailTestAddress, emailFromAddressDefault, "[Test mode true] Could not notify account contact for banner ad changes", "[Test Mode, toAddressMemberID=" & emailModifiedById & "]" & vbCrLf & errorMessage, True, False)
                            ElseIf (emailModifiedById <> 0) Then
                                hint &= ",290"
                                Call CP.Email.sendUser(emailModifiedById, emailFromAddressDefault, "Could not notify account contact for banner ad changes", errorMessage, True, False)
                            End If

                        End If
                    End If
                    hint &= ",299"
                    Call cs.SetField("sendReviewRequest", 0)
                    Call cs.GoNext()
                Loop
                Call cs.Close()
                hint &= ",300"
                Call CP.Utils.AppendLogFile("exiting EmailNotificationClass.execute, hint=" & hint)
            Catch ex As Exception
                Call CP.Site.ErrorReport("EmailNotificationClass.execute, " & ex.Message)
            End Try
        End Function
    End Class
End Namespace