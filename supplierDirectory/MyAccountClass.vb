
Imports Contensive.BaseClasses
Imports Contensive.Addons.SupplierDirectory.Constants

Namespace Contensive.Addons.SupplierDirectory
    Friend Class MyAccountClass
        '
        '============================================================================
        '   get Home Form
        '============================================================================
        '
        Friend Function getForm(ByVal cp As CPBaseClass, ByVal common As CommonClass, ByVal orgId As Integer) As String
            getForm = ""
            Try
                '
                'Dim orgId As Long = 0
                Dim categoryName As String = ""
                Dim subcategoryName As String = ""
                Dim sql As String = ""
                Dim subcategorylist As String = ""
                Dim placement As String = ""
                Dim typeId As Integer = 0
                Dim typeName As String = ""
                Dim active As String
                Dim recordId As Integer = 0
                Dim recordName As String = ""
                Dim w As String = ""
                Dim copy As String = ""
                Dim adListRow As String
                Dim adListRows As String = ""
                Dim content As String = ""
                Dim qs As String = ""
                Dim rqs As String = cp.Doc.RefreshQueryString()
                Dim cs As CPCSBaseClass = cp.CSNew()
                Dim csPeople As Contensive.BaseClasses.CPCSBaseClass = cp.CSNew()
                Dim accountContactID As Integer
                Dim line3 As String = ""
                Dim city As String = ""
                Dim state As String = ""
                Dim zip As String = ""
                Dim country As String = ""
                Dim directoryIsEnhancedListing As Boolean = False
                Dim directoryProfileApprovedByAccount As Boolean = False
                Dim directoryProfileDescription As String = ""
                Dim directoryProfileImageFilename As String = ""
                Dim directoryListingImageFilename As String = ""
                Dim qsShowcase As String = ""
                Dim csOrg As CPCSBaseClass = cp.CSNew
                '
                cp.Site.TestPoint("MyAccountClass.getForm")
                '
                ' populate the organization info
                '
                Call csOrg.Open("organizations", "(id=" & orgId & ")and(directoryAccountContactId=" & cp.User.Id & ")")
                If Not csOrg.OK Then
                    content = common.getPermissionError()
                Else
                    content = common.getLayout(cp, "Supplier Directory My Account Home")
                    qs = cp.Utils.ModifyQueryString(rqs, rnformId, formIdMyAccountEdit)
                    qs = cp.Utils.ModifyQueryString(qs, rnOrganizationID, orgId)
                    content = content.Replace("##myAccountEditLink##", "?" & qs)
                    '
                    cp.Site.TestPoint("MyAccountClass.getForm, 100")
                    '
                    accountContactID = csOrg.GetInteger("directoryaccountContactID")
                    Call csPeople.Open("people", "id=" & accountContactID)
                    If Not csPeople.OK Then
                        cp.Site.TestPoint("MyAccountClass.getForm, 110")
                        Call csPeople.Close()
                        Call csPeople.Insert("people")
                        If csPeople.OK Then
                            cp.Site.TestPoint("MyAccountClass.getForm, 120")
                            accountContactID = csPeople.GetInteger("id")
                            Call csOrg.SetField("directoryaccountContactID", accountContactID)
                        End If
                    End If
                    '
                    cp.Site.TestPoint("MyAccountClass.getForm, 130")
                    copy = csOrg.GetText("directoryname")
                    If copy = "" Then
                        copy = "(no name)"
                    End If
                    cp.Site.TestPoint("MyAccountClass.getForm, 140")
                    content = content.Replace("##orgName##", bufferCopy(copy, "(no name)"))
                    '
                    content = content.Replace("##orgAddress1##", bufferCopy(csOrg.GetText("directoryaddress1"), "&nbsp;"))
                    content = content.Replace("##orgAddress2##", bufferCopy(csOrg.GetText("directoryaddress2"), "&nbsp;"))
                    city = csOrg.GetText("directorycity")
                    state = csOrg.GetText("directorystate")
                    zip = csOrg.GetText("directoryzip")
                    country = csOrg.GetText("directorycountry")
                    If (city <> "") And (state <> "") Then
                        line3 = city & ",&nbsp;" & state
                    Else
                        line3 = city & state
                    End If
                    If (line3 <> "") And (zip <> "") Then
                        line3 = line3 & "&nbsp;" & zip
                    Else
                        line3 = city & zip
                    End If
                    If (line3 <> "") And (country <> "") Then
                        line3 = line3 & ",&nbsp;" & country
                    Else
                        line3 = city & country
                    End If
                    content = content.Replace("##orgAddress3##", bufferCopy(line3, "&nbsp;"))
                    '
                    content = content.Replace("##orgProfileContact##", bufferCopy(csOrg.GetText("directoryProfileContact"), "&nbsp;"))
                    content = content.Replace("##orgPhone##", bufferCopy(common.formatPhone(cp, csOrg.GetText("directoryphone")), "&nbsp;"))
                    content = content.Replace("##orgFax##", bufferCopy(common.formatPhone(cp, csOrg.GetText("directoryfax")), "&nbsp;"))
                    content = content.Replace("##orgWeb##", bufferCopy(csOrg.GetText("directoryweb"), "&nbsp;"))
                    content = content.Replace("##orgEmail##", bufferCopy(csOrg.GetText("directoryemail"), "&nbsp;"))
                    '
                    directoryIsEnhancedListing = csOrg.GetBoolean("directoryIsEnhancedListing")
                    directoryProfileApprovedByAccount = csOrg.GetBoolean("directoryProfileApprovedByAccount")
                    directoryProfileDescription = csOrg.GetText("directoryProfileDescription")
                    directoryProfileImageFilename = csOrg.GetText("directoryProfileImageFilename")
                    directoryListingImageFilename = csOrg.GetText("directoryListingImageFilename")
                    '
                    copy = ""
                    copy &= "<div class=""listingImageCaption"">Image:</div>"
                    If directoryListingImageFilename <> "" Then
                        copy &= "<div class=""listingImage""><img src=""" & cp.Site.FilePath & directoryListingImageFilename & """></div>"
                    Else
                        copy &= "<div class=""listingImageMessage"">No image available.</div>"
                    End If
                    content = content.Replace("##listingImage##", copy)
                    '
                    '
                    ' populate the people record
                    '
                    cp.Site.TestPoint("MyAccountClass.getForm, 200")
                    If Not csPeople.OK Then
                        content = content.Replace("##orgContactName##", "&nbsp;")
                        content = content.Replace("##orgContactPhone##", "&nbsp;")
                        content = content.Replace("##orgContactEmail##", "&nbsp;")
                    Else
                        content = content.Replace("##orgContactName##", bufferCopy(csPeople.GetText("name"), "&nbsp;"))
                        content = content.Replace("##orgContactPhone##", bufferCopy(common.formatPhone(cp, csPeople.GetText("phone")), "&nbsp;"))
                        content = content.Replace("##orgContactEmail##", bufferCopy(csPeople.GetText("email"), "&nbsp;"))
                    End If
                    Call csPeople.Close()
                    '
                    ' Enhanced Listing Section
                    '
                    '"##orgEnhancedListing##"
                    copy = ""
                    If Not (directoryIsEnhancedListing) Then
                        copy &= "<div class=""profileView"">Enhanced listing has not been added.</div>"
                    Else
                        If Not (directoryProfileApprovedByAccount) Then
                            copy &= "<div class=""profileView"">Enhanced listing is available but has not been approved by you.</div>"
                        End If
                        qs = rqs
                        qs = cp.Utils.ModifyQueryString(qs, rnformId, formIdProfile)
                        qs = cp.Utils.ModifyQueryString(qs, rnOrganizationID, orgId)
                        copy &= "<div class=""profileView""><a href=""?" & qs & """>View Company Profile</a></div>"
                        copy &= "<div class=""profileDescriptionCaption"">Description:</div>"
                        copy &= "<div class=""profileDescription"">" & directoryProfileDescription & "</div>"
                        copy &= "<div class=""profileImageCaption"">Image:</div>"
                        If directoryProfileImageFilename <> "" Then
                            copy &= "<div class=""profileImage""><img src=""" & cp.Site.FilePath & directoryProfileImageFilename & """></div>"
                        Else
                            copy &= "<div class=""profileImageMessage"">No image available. Listing image will be used.</div>"
                        End If
                    End If
                    content = content.Replace("##orgEnhancedListing##", copy)
                    '
                    ' Populate the subcategory List
                    '
                    subcategorylist = ""
                    sql = "select h.id, h.category, h.name as subcategoryName, p.name as placementName" _
                        & " from (( organizationheadingrules r" _
                        & " left join tempheadings h on r.headingid=h.id )" _
                        & " left join directoryplacements p on r.placementid=p.id )" _
                        & " where(r.organizationid=" & orgId & ")" _
                        & " order by h.category, h.name, p.name desc"
                    cs.OpenSQL("", sql)
                    Do While cs.OK
                        placement = cs.GetText("placementName")
                        If placement <> "" Then
                            placement = "&nbsp;(&nbsp;" & placement & "&nbsp;)"
                        End If
                        categoryName = cs.GetText("category")
                        qs = rqs
                        qs = cp.Utils.ModifyQueryString(qs, rnformId, formIdCategoryHeading)
                        qs = cp.Utils.ModifyQueryString(qs, rnCategory, categoryName)
                        categoryName = "<a href=""?" & qs & """>" & categoryName & "</a>"
                        '
                        qs = rqs
                        qs = cp.Utils.ModifyQueryString(qs, rnformId, formIdHeadingSearchREsults)
                        qs = cp.Utils.ModifyQueryString(qs, rnHeadingID, cs.GetText("id"))
                        subcategoryName = "<a href=""?" & qs & """>" & cs.GetText("subcategoryName") & "</a>"
                        subcategorylist &= cp.Html.div(categoryName & "&nbsp;&gt;&nbsp;" & subcategoryName & placement, , "subcategoryListRow")
                        cs.GoNext()
                    Loop
                    cs.Close()
                    If subcategorylist = "" Then
                        subcategorylist = "<div class=""emptyRow"">Your account is not currently assigned to any subcategories.</div>"
                    End If
                    content = content.Replace("##subcategoryListRows##", subcategorylist)
                    '
                    ' List Banner Ads
                    '
                    cp.Site.TestPoint("MyAccountClass.getForm, 300")
                    adListRow = common.getLayout(cp, "Supplier Directory My Account Home adListRow")
                    cs.Open("Directory Banner Ads", "(organizationid=" & orgId & ")and(approved<>0)")
                    Do While cs.OK
                        recordId = cs.GetInteger("id")
                        recordName = cs.GetText("name")
                        If recordName = "" Then
                            recordName = "banner ad " & recordId.ToString
                        End If
                        If cs.GetBoolean("ApprovedByAccount") Then
                            active = "Yes"
                        Else
                            active = "No"
                        End If
                        typeName = ""
                        Select Case cs.GetInteger("TypeID")
                            Case 1
                                typeName = "Vertical Banner-120x600"
                            Case 2
                                typeName = "Horizontal Banner-1000x72"
                        End Select

                        w = adListRow
                        w = w.Replace("##name##", recordName)
                        w = w.Replace("##active##", active)
                        w = w.Replace("##type##", bufferCopy(typeName, "&nbsp;"))
                        w = w.Replace("##expires##", bufferCopy(cs.GetText("DateExpires"), "(No Expiration)"))
                        qs = cp.Utils.ModifyQueryString(rqs, rnformId, formIdBannerEdit)
                        qs = cp.Utils.ModifyQueryString(qs, rnBannerAdID, recordId.ToString)
                        qs = cp.Utils.ModifyQueryString(qs, rnOrganizationID, orgId)
                        w = w.Replace("##edit##", "<a href=""?" & qs & """>Edit<a/>")
                        adListRows &= w
                        cs.GoNext()
                    Loop
                    Call cs.Close()
                    '
                    ' List Showcase Ads
                    '
                    cp.Site.TestPoint("MyAccountClass.getForm, 400")
                    cs.Open("Directory Showcase Ads", "(organizationid=" & orgId & ")and(approved<>0)")
                    Do While cs.OK
                        recordId = cs.GetInteger("id")
                        recordName = cs.GetText("name")
                        If recordName = "" Then
                            recordName = "banner ad " & recordId.ToString
                        End If
                        If cs.GetBoolean("ApprovedByAccount") Then
                            active = "Yes"
                        Else
                            active = "No"
                        End If
                        typeName = "Showcase Ad"
                        w = adListRow
                        w = w.Replace("##name##", recordName)
                        w = w.Replace("##active##", active)
                        w = w.Replace("##type##", bufferCopy(typeName, "&nbsp;"))
                        w = w.Replace("##expires##", bufferCopy(cs.GetText("DateExpires"), "(No Expiration)"))
                        qs = rqs
                        qs = cp.Utils.ModifyQueryString(qs, rnformId, formIdShowcaseEdit)
                        qs = cp.Utils.ModifyQueryString(qs, rnShowcaseAdID, recordId.ToString)
                        qs = cp.Utils.ModifyQueryString(qs, rnOrganizationID, orgId)
                        qsShowcase = rqs
                        qsShowcase = cp.Utils.ModifyQueryString(qsShowcase, rnformId, formIdShowcaseAd)
                        qsShowcase = cp.Utils.ModifyQueryString(qsShowcase, rnShowcaseAdID, recordId.ToString)
                        w = w.Replace("##edit##", "<a href=""?" & qs & """>Edit<a/>&nbsp-&nbsp<a href=""?" & qsShowcase & """>Public&nbsp;View</a>")
                        adListRows &= w
                        cs.GoNext()
                    Loop
                    Call cs.Close()
                    '
                    '
                    '
                    cp.Site.TestPoint("MyAccountClass.getForm, 500")
                    If adListRows = "" Then
                        adListRows = "<tr><td class=""emptyRow"">Your account currently has no active advertising items.</td></tr>"
                    End If
                    content = content.Replace("##adListRows##", adListRows)
                End If
                Call csOrg.Close()
                '
                '
                '
                getForm = content
            Catch ex As Exception
                Call cp.Site.ErrorReport("MyAccountClass.getForm, " & ex.Message)
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
        Private Function GetRecordName(ByVal CP As Contensive.BaseClasses.CPBaseClass, ByVal ContentName As String, ByVal recordId As Integer) As String
            GetRecordName = ""
            Try
                Dim cs As Contensive.BaseClasses.CPCSBaseClass = CP.CSNew
                '
                Call cs.Open(ContentName, "(id=" & recordId & ")", , , "name")
                If cs.OK Then
                    GetRecordName = cs.GetText("name")
                End If
                Call cs.Close()

            Catch ex As Exception
                Call CP.Site.ErrorReport("MyAccountClass.GetRecordName, " & ex.Message)
            End Try
            '


        End Function



    End Class
End Namespace
