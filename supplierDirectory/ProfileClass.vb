
Imports Contensive.BaseClasses
Imports Contensive.Addons.SupplierDirectory.Constants

Namespace Contensive.Addons.SupplierDirectory
    Friend Class ProfileClass
        '
        '============================================================================
        '   get Home Form
        '============================================================================
        '
        Friend Function getForm(ByVal cp As CPBaseClass, ByVal common As CommonClass) As String
            '
            Dim csShowcase As CPCSBaseClass
            Dim showcaseCaption As String = ""
            Dim csHeadings As CPCSBaseClass
            Dim emailForm As String = ""
            Dim address1Line As String = ""
            Dim address2Line As String = ""
            Dim address3Line As String = ""
            Dim categoryList As String = ""
            Dim profileDescription As String = ""
            Dim profileDescriptionLine As String = ""
            Dim profileImageLine As String = ""
            Dim organizationID As Integer = 0
            Dim pagePtr As Integer = 0
            Dim pageCnt As Integer = 0
            Dim layout As String = ""
            Dim cell As String = ""
            Dim memberIcon As String = ""
            'Dim addressLine As String = ""
            'Dim descriptionLine As String = ""
            Dim webLine As String = ""
            Dim web As String
            Dim weblink As String = ""
            Dim imageFilename As String = ""
            Dim memberIconSrc As String = ""
            Dim memberCompanyLink As String = ""
            Dim IsAssociationMember As Boolean = False
            Dim copy As String = ""
            'Dim pos As Integer
            Dim companyNameLine As String = ""
            Dim faxLine As String = ""
            Dim phoneLine As String = ""
            Dim profileButtonLine As String = ""
            Dim cs As CPCSBaseClass
            'Dim pageNumber As Integer = 0
            'Dim pageSize As Integer = 0
            'Dim categoryLast As String = ""
            Dim rqs As String = ""
            Dim qs As String = ""
            Dim cellLeft As String = ""
            Dim cellRight As String = ""
            Dim cellCenter As String = ""
            Dim cellRowMax As Integer = 0
            Dim cellRowPtr As Integer = 0
            Dim Hint As String = ""
            Dim rowTotal As Integer = 0
            Dim cacheName As String = ""
            Dim headingID As Integer = 0
            Dim headingName As String = ""
            Dim categoryName As String = ""
            Dim recordList As String = ""
            Dim recordCnt As Long = 0
            Dim pageSizeOptions As String = ""
            Dim sql As String = ""
            Dim viewOptions As String = ""
            Dim recordCntMessage As String = ""
            Dim image As String = ""
            Dim companyName As String = ""
            Dim ListingNormalSrc As String = ""
            Dim memberCompanyCaption As String = ""
            Dim featureLine As String = ""
            Dim IsExhibitor As Boolean = False
            Dim exhibitorCaption As String = ""
            Dim exhibitorLink As String = ""
            Dim webLinkLine As String = ""
            Dim email As String = ""
            Dim recordPtr As Integer = 0
            Dim pageNavigation As String = ""
            Dim profileContact As String
            Dim profileContactLine As String
            Dim emailLinkLine As String = ""
            'Dim hint As String = ""
            '
            cp.Site.TestPoint("ProfileClass.getForm")
            '
            getForm = ""
            Try
                organizationID = cp.Request.GetInteger(rnOrganizationID)
                If organizationID <> 0 Then
                    hint = "100"
                    '
                    ' attempt cache read
                    '
                    cacheName = "Supplier Directory Profile, " & organizationID
                    getForm = cp.Cache.Read(cacheName)
                    If getForm = "" Then
                        hint = "200"
                        '
                        ' build form
                        '
                        cp.Site.TestPoint("ProfileClass.getForm, building form")
                        cp.Doc.AddRefreshQueryString(rnOrganizationID, organizationID)
                        rqs = cp.Doc.RefreshQueryString
                        '
                        cs = cp.CSNew
                        Call cs.Open("organizations", "id=" & organizationID)

                        'sql = "select o.*,h.name as headingName,h.category as categoryName" _
                        '    & " from ((organizations o " _
                        '    & " left join OrganizationSubcategoryRules r on r.organizationid=o.id)" _
                        '    & " left join directorySubcategories h on r.headingid=h.id)" _
                        '    & " where (o.id=" & organizationID & ")" _
                        '    & " and ((r.approved<>0)or(r.id is null))" _
                        '    & " order by h.category,h.name"
                        'cs = cp.CSNew
                        'Call cs.OpenSQL("", sql)
                        If Not cs.OK Then
                            '
                            ' Organization not found
                            '
                            getForm = "<p>The organzation you requested could not be displayed. Please use the back button to return.</p>"
                            getForm = cp.Content.GetCopy("Supplier Directory Profile Page Organzation Not Found", getForm)
                        Else
                            hint = "300"
                            memberCompanyCaption = cp.Site.GetProperty("Supplier Directory Member Company Caption", "Member Company")
                            memberCompanyLink = cp.Site.GetProperty("Supplier Directory Member Company Link", "/")
                            exhibitorCaption = cp.Site.GetProperty("Supplier Directory Exhibitor Caption", "Exhibitor")
                            exhibitorLink = cp.Site.GetProperty("Supplier Directory Exhibitor Link", "/")
                            showcaseCaption = cp.Site.GetProperty("Supplier Directory Showcase Caption", "Product Showcase")
                            '
                            companyName = cs.GetText("directoryname").Replace(" ", "&nbsp;")
                            weblink = cs.GetText("directorylink")
                            web = cs.GetText("directoryweb")
                            imageFilename = cs.GetText("directoryProfileImageFilename")
                            If imageFilename = "" Then
                                imageFilename = cs.GetText("directoryListingImageFilename")
                            End If
                            IsAssociationMember = cs.GetBoolean("directoryIsAssociationMember")
                            IsExhibitor = cs.GetBoolean("directoryIsExhibitor")
                            profileDescription = cs.GetText("directoryprofileDescription")
                            profileContact = cs.GetText("directoryprofileContact")
                            '
                            If web = "" Then
                                web = weblink
                            End If
                            If weblink = "" Then
                                weblink = web
                            End If
                            If (web <> "") Then
                                If weblink.Substring(0, 4).ToLower <> "http" Then
                                    weblink = "http://" & weblink
                                End If
                            End If
                            '
                            companyNameLine = cp.Html.div(companyName, , "companyNameLine")
                            '
                            If imageFilename <> "" Then
                                image = "<img alt=""" & cp.Utils.EncodeHTML(companyName) & """ src=""" & cp.Site.FilePath & imageFilename & """>"
                                If weblink <> "" Then
                                    image = "<a href=""" & weblink & """ target=""_blank"">" & image & "</a>"
                                End If
                                profileImageLine = cp.Html.div(image, , "profileImageLine")
                            End If
                            '
                            profileDescriptionLine = cs.GetText("directoryProfiledescription")
                            'If (profileDescriptionLine <> "") And (weblink <> "") Then
                            'profileDescriptionLine = "<a style=""description"" href=""" & weblink & """>" & profileDescriptionLine & "</a>"
                            'End If
                            If profileDescriptionLine <> "" Then
                                profileDescriptionLine = profileDescriptionLine.Replace(vbCrLf, "<br>")
                                profileDescriptionLine = profileDescriptionLine.Replace(vbLf, "<br>")
                                profileDescriptionLine = profileDescriptionLine.Replace(vbCr, "<br>")
                                profileDescriptionLine = cp.Html.div(profileDescriptionLine, , "descriptionLine")
                            End If
                            '
                            hint = "400"
                            featureLine = ""
                            If IsAssociationMember Then
                                featureLine &= "<div><a href=""" & memberCompanyLink & """ target=""_blank"">" & memberCompanyCaption & "</a></div>"
                            End If
                            If IsExhibitor Then
                                featureLine &= "<div><a href=""" & exhibitorLink & """ target=""_blank"">" & exhibitorCaption & "</a></div>"
                            End If
                            sql = "select top 1 id from DirectoryShowcaseAds where (organizationid=" & organizationID & ")and(approved<>0)and(ApprovedByAccount<>0)"
                            csShowcase = cp.CSNew
                            Call csShowcase.OpenSQL("", sql)
                            If csShowcase.OK Then
                                qs = rqs
                                qs = cp.Utils.ModifyQueryString(qs, rnShowcaseAdID, csShowcase.GetInteger("ID").ToString)
                                qs = cp.Utils.ModifyQueryString(qs, rnformId, formIdShowcaseAd)
                                featureLine &= "<div><a href=""?" & qs & """>" & showcaseCaption & "</a></div>"
                            End If
                            Call csShowcase.Close()
                            If featureLine <> "" Then
                                featureLine = cp.Html.div(featureLine, , "featureLine")
                            End If
                            '
                            phoneLine = cs.GetText("directoryPhone")
                            If phoneLine <> "" Then
                                phoneLine = common.formatPhone(cp, phoneLine)
                                phoneLine = cp.Html.div(phoneLine.Replace(" ", "&nbsp;"), , "phoneLine")
                            End If
                            '
                            hint = "500"
                            faxLine = cs.GetText("directoryfax")
                            If faxLine <> "" Then
                                faxLine = common.formatPhone(cp, faxLine)
                                faxLine = cp.Html.div(faxLine.Replace(" ", "&nbsp;") & "&nbsp;fax", , "faxLine")
                            End If
                            '
                            If (web <> "") Then
                                webLine = cp.Html.div("<a href=""" & weblink & """ target=""_blank"">" & web & "</a>", , "webLine")
                            End If
                            '
                            email = cs.GetText("directoryemail")
                            If email <> "" Then
                                emailForm = common.getLayout(cp, "Supplier Directory Email Popup Form")
                                emailForm = emailForm.Replace("##companyName##", companyName)
                                emailLinkLine = "" _
                                    & vbCrLf & "<a id=""supplierDirectoryEmailFormLink##organizationID##"" href=""#supplierDirectoryEmailPopup##organizationID##"">E-mail</a>" _
                                    & vbCrLf & emailForm _
                                    & vbCrLf & "<script>" _
                                        & "$(document).ready(function() {" _
                                            & "$('#supplierDirectoryEmailFormLink##organizationID##').fancybox({" _
                                                & "'titleShow':false," _
                                                & "'transitionIn':'fade'," _
                                                & "'transitionOut':'fade'," _
                                                & "'overlayOpacity':'.6'," _
                                                & "'overlayColor':'#000000'" _
                                            & "});" _
                                        & "});" _
                                    & "</script>" _
                                    & ""
                                emailLinkLine = emailLinkLine.Replace("##organizationID##", organizationID)
                                emailLinkLine = cp.Html.div(emailLinkLine, , "emailLinkLine")
                            End If
                            '
                            hint = "600"
                            address1Line = ""
                            copy = cs.GetText("directoryaddress1")
                            If copy <> "" Then
                                address1Line = cp.Html.div(copy.Replace(" ", "&nbsp;"), , "address1Line")
                            End If
                            '
                            address2Line = ""
                            copy = cs.GetText("directoryaddress2")
                            If copy <> "" Then
                                address2Line = cp.Html.div(copy.Replace(" ", "&nbsp;"), , "address2Line")
                            End If
                            '
                            address3Line = ""
                            copy = cs.GetText("directorycity")
                            If copy <> "" Then
                                address3Line &= ",&nbsp;" & copy.Replace(" ", "&nbsp;")
                            End If
                            copy = cs.GetText("directorystate")
                            If copy <> "" Then
                                If address3Line <> "" Then
                                    address3Line &= ",&nbsp;"
                                End If
                                address3Line &= copy.Replace(" ", "&nbsp;")
                            End If
                            copy = cs.GetText("directoryzip")
                            If copy <> "" Then
                                If address3Line <> "" Then
                                    address3Line &= "&nbsp;"
                                End If
                                address3Line &= copy.Replace(" ", "&nbsp;")
                            End If
                            copy = cs.GetText("directorycountry")
                            If copy <> "" Then
                                Select Case copy.ToLower
                                    Case "", "united states of america", "united states", "us", "usa", "u.s.a", "u.s."
                                    Case Else
                                        If address3Line <> "" Then
                                            address3Line &= "<br>"
                                        End If
                                        address3Line &= copy.Replace(" ", "&nbsp;")
                                End Select
                            End If
                            If address3Line <> "" Then
                                address3Line = cp.Html.div(address3Line.Substring(7), , "address3Line")
                            End If
                            '
                            emailForm = ""
                            '
                            hint = "700"
                            categoryList = ""
                            csHeadings = cp.CSNew
                            sql = "select h.category as categoryname,h.name as headingName,h.id as headingId" _
                                & " from directorySubcategories h left join OrganizationSubcategoryRules r on r.headingid=h.id" _
                                & " where r.organizationid=" & organizationID _
                                & " and (r.approved<>0)" _
                                & " order by h.category, h.name"
                            cp.Site.TestPoint("heading sql=" & sql)
                            Call csHeadings.OpenSQL("", sql)
                            Do While csHeadings.OK
                                '
                                Dim categoryLink As String
                                Dim headingLink As String
                                '
                                categoryLink = csHeadings.GetText("categoryname")
                                If categoryLink <> "" Then
                                    qs = rqs
                                    qs = cp.Utils.ModifyQueryString(qs, rnOrganizationID, "", False)
                                    qs = cp.Utils.ModifyQueryString(qs, rnOrganizationID, "", False)
                                    qs = cp.Utils.ModifyQueryString(qs, rnformId, formIdCategoryHeading)
                                    qs = cp.Utils.ModifyQueryString(qs, rnCategory, categoryLink)
                                    categoryLink = "<a href=""?" & qs & """>" & categoryLink & "</a>"
                                    qs = rqs
                                    qs = cp.Utils.ModifyQueryString(qs, rnOrganizationID, "", False)
                                    qs = cp.Utils.ModifyQueryString(qs, rnformId, formIdHeadingSearchREsults)
                                    qs = cp.Utils.ModifyQueryString(qs, rnHeadingID, csHeadings.GetInteger("headingId"))
                                    headingLink = "<a href=""?" & qs & """>" & csHeadings.GetText("headingname") & "</a>"
                                    categoryList &= cp.Html.div(categoryLink & "&nbsp;&gt;&nbsp;" & headingLink)
                                End If
                                csHeadings.GoNext()
                            Loop
                            Call csHeadings.Close()
                            If categoryList <> "" Then
                                categoryList = cp.Html.div(categoryList, , "categoryList")
                            End If
                            profileContactLine = ""
                            If profileContact <> "" Then
                                profileContactLine = cp.Html.div(profileContact, , "profileContactLine")
                            End If
                            '
                            ' stuff layout and save bake
                            '
                            hint = "800"
                            getForm = common.getLayout(cp, "Supplier Directory Profile")
                            If getForm = "" Then
                                cp.Site.TestPoint("ProfileClass.getForm, profile layout not found")
                            Else
                                qs = rqs
                                qs = cp.Utils.ModifyQueryString(qs, "searchButton", "", False)
                                qs = cp.Utils.ModifyQueryString(qs, "searchText", "", False)
                                qs = cp.Utils.ModifyQueryString(qs, "searchPhrase", "", False)
                                qs = cp.Utils.ModifyQueryString(qs, "searchLocation", "", False)
                                qs = cp.Utils.ModifyQueryString(qs, rnHeadingID, "", False)
                                qs = cp.Utils.ModifyQueryString(qs, rnCategory, "", False)
                                qs = cp.Utils.ModifyQueryString(qs, rnformId, formIdHome)
                                getForm = getForm.Replace("##home##", "<a href=""?" & qs & """>All Categories</a>")
                                getForm = getForm.Replace("##companyNameLine##", companyNameLine)
                                getForm = getForm.Replace("##profileImageLine##", profileImageLine)
                                getForm = getForm.Replace("##profileDescriptionLine##", profileDescriptionLine)
                                getForm = getForm.Replace("##featureLine##", featureLine)
                                getForm = getForm.Replace("##phoneLine##", phoneLine)
                                getForm = getForm.Replace("##faxLine##", faxLine)
                                getForm = getForm.Replace("##webLine##", faxLine)
                                getForm = getForm.Replace("##emailLinkLine##", emailLinkLine)
                                getForm = getForm.Replace("##address1Line##", address1Line)
                                getForm = getForm.Replace("##address2Line##", address2Line)
                                getForm = getForm.Replace("##address3Line##", address3Line)
                                getForm = getForm.Replace("##emailLinkLine##", emailLinkLine)
                                getForm = getForm.Replace("##profileContactLine##", profileContactLine)
                                getForm = getForm.Replace("##categoryList##", categoryList)
                                Call cp.Cache.Save(cacheName, getForm, "Organizations, Organization Subcategories, Directory Subcategories")
                            End If
                        End If
                        Call cs.Close()
                    End If
                End If
            Catch ex As Exception
                Call cp.Site.ErrorReport("ProfileClass.getForm, hint=[" & hint & "],  " & ex.Message)
            End Try
            '
        End Function

    End Class
End Namespace
