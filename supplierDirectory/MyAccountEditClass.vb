
Imports Contensive.BaseClasses
Imports Contensive.Addons.SupplierDirectory.Constants

Namespace Contensive.Addons.SupplierDirectory
    Friend Class MyAccountEditClass
        '
        '============================================================================
        ' process form, return next formId
        '============================================================================
        '
        Friend Function processForm(ByVal cp As CPBaseClass, ByVal common As CommonClass, ByVal orgid As Integer) As Integer
            Dim actionSave As Boolean = False
            Dim cs As Contensive.BaseClasses.CPCSBaseClass = cp.CSNew
            Dim csPeople As CPCSBaseClass = cp.CSNew()
            Dim accountContactID As Integer
            Dim isEnhancedListing As Boolean
            Dim isvalid As Boolean = False
            Dim csOrg As CPCSBaseClass = cp.CSNew
            Dim stateName As String = ""
            Dim stateAbbr As String = ""
            Dim csState As CPCSBaseClass
            '
            Try
                '
                If True Then
                    processForm = formIdMyAccountEdit
                    Select Case cp.doc.getText("button")
                        Case buttonSave
                            actionSave = True
                            processForm = formIdMyAccountEdit
                        Case buttonOK
                            actionSave = True
                            processForm = formIdMyAccount
                        Case buttonCancel
                            processForm = formIdMyAccount
                        Case Else
                    End Select
                    '
                    If actionSave Then
                        Call cs.Open("organizations", "id=" & cp.User.OrganizationID)
                        If cs.OK Then
                            isEnhancedListing = cs.GetBoolean("directoryisEnhancedListing")
                            accountContactID = cs.GetInteger("directoryaccountContactID")
                            cp.Site.TestPoint("myAccountEditClass.processForm, 200, directoryaccountContactID=" & accountContactID)
                            Call csPeople.Open("people", "id=" & accountContactID)
                            'If Not csPeople.OK Then
                            '    cp.Site.TestPoint("myAccountEditClass.processForm, 210")
                            '    Call csPeople.Close()
                            '    Call csPeople.Insert("people")
                            '    If csPeople.OK Then
                            '        cp.Site.TestPoint("myAccountEditClass.processForm, 220")
                            '        accountContactID = csPeople.GetInteger("id")
                            '        Call cs.SetField("directoryaccountContactID", accountContactID)
                            '    End If
                            '    Call cp.Group.AddUser("Account Contacts", accountContactID)
                            'End If
                            stateAbbr = ""
                            stateName = cp.Request.GetText("directoryState")
                            If stateName <> "" Then
                                csState = cp.CSNew
                                Call csState.Open("states", "name=" & cp.Db.EncodeSQLText(stateName), "id", , "abbreviation")
                                If csState.OK() Then
                                    stateAbbr = csState.GetText("abbreviation")
                                End If
                                Call csState.Close()
                            End If
                            Call cs.SetField("directoryState", stateAbbr)
                            '
                            Call cs.SetFormInput("directoryname")
                            Call cs.SetFormInput("directoryAddress1")
                            Call cs.SetFormInput("directoryAddress2")
                            Call cs.SetFormInput("directoryCity")
                            'Call cs.SetFormInput("directoryState")
                            Call cs.SetFormInput("directoryZip")
                            Call cs.SetFormInput("directoryCountry")
                            Call cs.SetFormInput("directoryProfileContact")
                            Call cs.SetFormInput("directoryPhone")
                            Call cs.SetFormInput("directoryFax")
                            Call cs.SetFormInput("directoryWeb")
                            Call cs.SetFormInput("directoryEmail")
                            If cp.Request.GetBoolean("directorylistingimagefilename.delete") Then
                                Call cs.SetField("directorylistingimagefilename", "")
                            End If
                            Call cs.SetFormInput("directorylistingimagefilename")
                            If isEnhancedListing Then
                                Call cs.SetFormInput("directoryProfileApprovedByAccount")
                                Call cs.SetFormInput("directoryProfileDescription")
                                If cp.Request.GetBoolean("directoryProfileImageFilename.delete") Then
                                    Call cs.SetField("directoryProfileImageFilename", "")
                                End If
                                Call cs.SetFormInput("directoryProfileImageFilename")
                            End If
                            If csPeople.OK Then
                                Call csPeople.SetFormInput("name", "contactName")
                                Call csPeople.SetFormInput("phone", "contactPhone")
                                Call csPeople.SetFormInput("email", "contactEmail")
                            End If
                            Call csPeople.Close()
                        End If
                        Call cs.Close()
                    End If
                End If
            Catch ex As Exception
                Call cp.Site.ErrorReport("MyAccountEditClass.getForm, " & ex.Message)
            End Try
            '
        End Function
        '
        '============================================================================
        '   get Home Form
        '============================================================================
        '
        Friend Function getForm(ByVal cp As CPBaseClass, ByVal common As CommonClass, ByVal orgid As Integer) As String
            getForm = ""
            Try
                '
                Dim isEnhancedListing As Boolean
                Dim pos As Integer = 0
                Dim virtualFilename As String = ""
                Dim ImageFilename As String = ""
                Dim listingImageFilename As String = ""
                Dim copy As String = ""
                Dim button As String = ""
                Dim content As String = ""
                Dim qs As String = ""
                Dim rqs As String = cp.Doc.RefreshQueryString()
                Dim cs As CPCSBaseClass = cp.CSNew()
                Dim csPeople As CPCSBaseClass = cp.CSNew()
                Dim accountContactID As Integer
                Dim directoryState As String = ""
                Dim OrganizationId As Long = 0
                Dim csState As CPCSBaseClass
                Dim stateAbbr As String = ""
                Dim stateName As String = ""
                '
                cp.Site.TestPoint("MyAccountEditClass.getForm")
                '
                ' populate the organization info
                '
                cp.Site.TestPoint("MyAccountEditClass.getForm, 100")
                OrganizationId = cp.doc.getText(rnOrganizationID)

                Call cs.Open("organizations", "(id=" & OrganizationId & ")and(directoryaccountcontactid=" & cp.Db.EncodeSQLNumber(cp.User.Id) & ")")
                If Not cs.OK Then
                    '
                    ' bad org id
                    '
                    content = cp.Html.p("There was a problem editing this organization. Please use your back button to return to the previous screen.")
                Else
                    '
                    ' build form
                    '
                    content = common.getLayout(cp, "Supplier Directory My Account Edit")
                    isEnhancedListing = cs.GetBoolean("directoryisEnhancedListing")
                    cp.Site.TestPoint("MyAccountEditClass.getForm, 200")
                    accountContactID = cs.GetInteger("directoryaccountContactID")
                    '
                    stateAbbr = cs.GetText("directoryState")
                    stateName = ""
                    If stateAbbr <> "" Then
                        csState = cp.CSNew
                        Call csState.Open("states", "abbreviation=" & cp.Db.EncodeSQLText(stateAbbr), "id", , "name")
                        If csState.OK() Then
                            stateName = csState.GetText("name")
                        End If
                        Call csState.Close()
                        If (stateAbbr <> "") And (stateName = "") Then
                            stateName = stateAbbr
                        End If
                    End If
                    directoryState = stateName
                    '
                    cp.Site.TestPoint("MyAccountEditClass.getForm, 300")
                    content = content.Replace("##orgInputName##", common.getHtmlInputText(cp, cs, "directoryName"))
                    cp.Site.TestPoint("MyAccountEditClass.getForm, 310")
                    content = content.Replace("##orgInputAddress1##", common.getHtmlInputText(cp, cs, "directoryAddress1"))
                    content = content.Replace("##orgInputAddress2##", common.getHtmlInputText(cp, cs, "directoryAddress2"))
                    content = content.Replace("##orgInputCity##", common.getHtmlInputText(cp, cs, "directoryCity"))
                    content = content.Replace("##orgInputState##", common.getStateSelect(cp, "directoryState", directoryState, "directoryState"))
                    content = content.Replace("##orgInputZip##", common.getHtmlInputText(cp, cs, "directoryZip"))
                    content = content.Replace("##orgInputCountry##", common.getHtmlInputText(cp, cs, "directoryCountry"))
                    '
                    content = content.Replace("##orgInputProfileContact##", common.getHtmlInputText(cp, cs, "directoryProfileContact"))
                    content = content.Replace("##orgInputPhone##", common.getHtmlInputText(cp, cs, "directoryPhone"))
                    content = content.Replace("##orgInputFax##", common.getHtmlInputText(cp, cs, "directoryFax"))
                    content = content.Replace("##orgInputWeb##", common.getHtmlInputText(cp, cs, "directoryWeb"))
                    content = content.Replace("##orgInputEmail##", common.getHtmlInputText(cp, cs, "directoryEmail"))
                    '
                    cp.Site.TestPoint("MyAccountEditClass.getForm, 400")
                    content = content.Replace("##orgInputListingImage##", common.getHtmlInputFile(cp, cs, "directorylistingImageFilename", "listingImageFilename"))
                    If Not isEnhancedListing Then
                        content = content.Replace("##orgInputProfileApprovedByAccount##", "(available only for enhanced listings)")
                        content = content.Replace("##orgInputProfileDescription##", "(available only for enhanced listings)")
                        content = content.Replace("##orgInputProfileImage##", "(available only for enhanced listings)")
                    Else
                        content = content.Replace("##orgInputProfileApprovedByAccount##", common.getHtmlCheckbox(cp, cs, "directoryprofileApprovedByAccount", "profileApprovedByAccount"))
                        content = contentReplaceTextArea(cp, content, cs, "directoryProfileDescription", "directoryprofileDescription", "##orgInputProfileDescription##")
                        content = content.Replace("##orgInputProfileImage##", common.getHtmlInputFile(cp, cs, "directoryprofileImageFilename", "profileImageFilename"))
                    End If
                    '
                    cp.Site.TestPoint("MyAccountEditClass.getForm, 500")
                    Call csPeople.Open("people", "id=" & cp.User.Id)
                    content = content.Replace("##orgInputContactName##", common.getHtmlInputText(cp, csPeople, "name", , , "contactName"))
                    content = content.Replace("##orgInputContactPhone##", common.getHtmlInputText(cp, csPeople, "phone", , , "contactPhone"))
                    content = content.Replace("##orgInputContactEmail##", common.getHtmlInputText(cp, csPeople, "Email", , , "contactEmail"))
                    Call csPeople.Close()
                End If
                Call cs.Close()
                '
                qs = cp.Doc.RefreshQueryString
                qs = cp.Utils.ModifyQueryString(qs, rnformId, formIdMyAccountEdit)
                qs = cp.Utils.ModifyQueryString(qs, rnOrganizationID, orgid)
                getForm = cp.Html.Form(content, "myAccountEdit", "myAccountEditForm", , qs, "post")
            Catch ex As Exception
                Call cp.Site.ErrorReport("MyAccountEditClass.getForm, " & ex.Message)
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
        Private Function contentReplaceInput(ByVal cp As Contensive.BaseClasses.CPBaseClass, ByVal content As String, ByVal cs As Contensive.BaseClasses.CPCSBaseClass, ByVal fieldName As String, Optional ByVal HtmlClass As String = "", Optional ByVal replaceDst As String = "", Optional ByVal inputName As String = "") As String
            Dim copy As String = ""
            '
            If replaceDst = "" Then
                replaceDst = "##orgInput" & fieldName & "##"
            End If
            If inputName = "" Then
                inputName = fieldName
            End If
            If HtmlClass = "" Then
                HtmlClass = fieldName
            End If
            copy = cp.Html.InputText(inputName, cs.GetText(fieldName), , , , HtmlClass)
            If HtmlClass <> "" Then
                copy = copy.Replace(">", " class=""" & HtmlClass & """>")
            End If
            contentReplaceInput = content.Replace(replaceDst, copy)
        End Function
        '
        '
        '
        Private Function contentReplaceTextArea(ByVal cp As Contensive.BaseClasses.CPBaseClass, ByVal content As String, ByVal cs As Contensive.BaseClasses.CPCSBaseClass, ByVal fieldName As String, Optional ByVal HtmlClass As String = "", Optional ByVal replaceDst As String = "", Optional ByVal inputName As String = "") As String
            Dim copy As String = ""
            '
            If replaceDst = "" Then
                replaceDst = "##orgInput" & fieldName & "##"
            End If
            If inputName = "" Then
                inputName = fieldName
            End If
            If HtmlClass = "" Then
                HtmlClass = fieldName
            End If
            copy = cp.Html.InputText(inputName, cs.GetText(fieldName), 2, , , HtmlClass)
            If HtmlClass <> "" Then
                copy = copy.Replace(">", " class=""" & HtmlClass & """>")
            End If
            contentReplaceTextArea = content.Replace(replaceDst, copy)
        End Function
        '
        '
        '
    End Class
    '
    '
    '
    '
    '
End Namespace
