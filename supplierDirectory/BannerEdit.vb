
Imports Contensive.BaseClasses
Imports Contensive.Addons.SupplierDirectory.Constants

Namespace Contensive.Addons.SupplierDirectory
    Friend Class BannerEditClass
        '
        '============================================================================
        ' process form, return next formId
        '============================================================================
        '
        Friend Function processForm(ByVal cp As CPBaseClass, ByVal common As CommonClass) As Integer
            Dim actionSave As Boolean = False
            Dim cs As Contensive.BaseClasses.CPCSBaseClass = cp.CSNew
            Dim csPeople As CPCSBaseClass = cp.CSNew()
            '
            processForm = formIdBannerEdit
            Try
                Select Case cp.Doc.GetText("button")
                    Case buttonSave
                        actionSave = True
                        processForm = formIdBannerEdit
                    Case buttonOK
                        actionSave = True
                        processForm = formIdMyAccount
                    Case buttonCancel
                        processForm = formIdMyAccount
                    Case Else
                End Select
                '
                If actionSave Then
                    Call cs.Open("Directory Banner Ads", "id=" & cp.Doc.GetText(rnBannerAdID))
                    If cs.OK Then
                        Call cs.SetFormInput("name")
                        Call cs.SetFormInput("approvedByAccount")
                        Call cs.SetFormInput("typeid")
                        Call cs.SetFormInput("link")
                        Call cs.SetFormInput("title")
                        '
                        If cp.Request.GetBoolean("imageFilename.delete") Then
                            Call cs.SetField("Imagefilename", "")
                        End If
                        Call cs.SetFormInput("ImageFilename")
                    End If
                    Call cs.Close()
                End If
            Catch ex As Exception
                Call cp.Site.ErrorReport("BannerEditClass.getForm, " & ex.Message)
            End Try
            '
        End Function
        '
        '============================================================================
        '   get Home Form
        '============================================================================
        '
        Friend Function getForm(ByVal cp As CPBaseClass, ByVal common As CommonClass, ByVal orgId As Integer) As String
            getForm = ""
            Try
                '
                Dim pos As Integer = 0
                Dim virtualFilename As String = ""
                Dim ImageFilename As String = ""
                Dim copy As String = ""
                Dim button As String = ""
                Dim content As String = ""
                Dim qs As String = ""
                Dim rqs As String = cp.Doc.RefreshQueryString()
                Dim cs As CPCSBaseClass = cp.CSNew()
                Dim csPeople As CPCSBaseClass = cp.CSNew()
                Dim bannerAdId As Integer = 0
                '
                cp.Site.TestPoint("BannerEditClass.getForm")
                '
                ' build form
                '
                content = common.getLayout(cp, "Supplier Directory Banner Edit")
                '
                ' Populate the replacement fields
                '
                bannerAdId = cp.Utils.EncodeInteger(cp.Doc.GetText(rnBannerAdID))
                content = content.Replace("##bannerAdId##", bannerAdId)
                Call cs.Open("directory banner ads", "id=" & bannerAdId.ToString)
                content = content.Replace("##name##", common.getHtmlInputText(cp, cs, "name"))
                content = content.Replace("##approvedByAccount##", common.getHtmlCheckbox(cp, cs, "approvedByAccount"))
                content = content.Replace("##type##", common.getHtmlSelectList(cp, cs, "typeid", "Vertical Banner-120x600,Horizontal Banner-1000x72"))
                content = content.Replace("##link##", common.getHtmlInputText(cp, cs, "link"))
                content = content.Replace("##imageFilename##", common.getHtmlInputFile(cp, cs, "imageFilename"))
                content = content.Replace("##title##", common.getHtmlInputText(cp, cs, "title"))
                content &= cp.Html.Hidden(rnOrganizationID, orgId.ToString)
                Call cs.Close()
                '
                qs = cp.Doc.RefreshQueryString
                qs = cp.Utils.ModifyQueryString(qs, rnformId, formIdBannerEdit)
                getForm = cp.Html.Form(content, "BannerEdit", "BannerEditForm", , qs, "post")
            Catch ex As Exception
                Call cp.Site.ErrorReport("BannerEditClass.getForm, " & ex.Message)
            End Try
            '
        End Function
        '
        '
        '
        Private Function bufferCopy(ByVal copy As String, ByVal defaultCopy As String) As String
            bufferCopy = copy
            If bufferCopy = "" Then
                bufferCopy = defaultCopy
            End If
        End Function
        '
        '
        '
    End Class
    '
    '
    '
End Namespace
