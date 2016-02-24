
Imports Contensive.BaseClasses
Imports Contensive.Addons.SupplierDirectory.Constants

Namespace Contensive.Addons.SupplierDirectory
    Friend Class ShowcaseAdClass
        '
        '============================================================================
        '   get Home Form
        '============================================================================
        '
        Friend Function getForm(ByVal cp As CPBaseClass, ByVal common As CommonClass) As String
            '
            Dim csHeadings As CPCSBaseClass
            Dim emailForm = ""
            Dim address1Line As String = ""
            Dim address2Line As String = ""
            Dim address3Line As String = ""
            Dim categoryList As String = ""
            Dim showcaseAdDescription As String = ""
            Dim showcaseAdDescriptionLine As String = ""
            Dim showcaseAdImageLine As String = ""
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
            Dim showcaseAdButtonLine As String = ""
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
            Dim emailLine As String = ""
            Dim recordPtr As Integer = 0
            Dim pageNavigation As String = ""
            Dim profileContact As String
            'Dim showcaseAdContactLine As String
            Dim showcaseAdID As Integer
            Dim profileContactLine As String = ""
            Dim showcaseCaption As String = ""
            Dim listingLink As String = ""
            '
            cp.Site.TestPoint("showcaseAdClass.getForm")
            '
            getForm = ""
            Try
                showcaseAdID = cp.Request.GetInteger(rnshowcaseAdID)
                If showcaseAdID <> 0 Then
                    '
                    ' attempt cache read
                    '
                    cacheName = "Supplier Directory Showcase Ad, " & showcaseAdID
                    getForm = common.cacheRead(cp, cacheName)
                    If getForm = "" Then
                        '
                        ' build form
                        '
                        cp.Site.TestPoint("showcaseAdClass.getForm, building form")
                        cp.Doc.AddRefreshQueryString(rnShowcaseAdID, showcaseAdID)
                        rqs = cp.Doc.RefreshQueryString
                        '
                        cs = cp.CSNew
                        sql = "select o.*" _
                            & ",s.imageFilename as showcaseAdImageFilename" _
                            & ",s.caption as showcaseAdCaption" _
                            & ",s.copy as showcaseAdDescription" _
                            & ",s.link as showcaseAdLink" _
                            & ",s.listingImageFilename as showcaseAdListingImageFilename" _
                            & " from DirectoryShowcaseAds s left join Organizations o on o.id=s.organizationid" _
                            & " where s.id=" & showcaseAdID
                        Call cp.Site.TestPoint("showcaseAdClass.getForm, sql=" & sql)
                        Call cs.OpenSQL("", sql)
                        If Not cs.OK Then
                            '
                            ' showcaseAd not found
                            '
                            getForm = "<p>The organzation you requested could not be displayed. Please use the back button to return.</p>"
                            getForm = cp.Content.GetCopy("Supplier Directory Showcase Ad Not Found", getForm)
                        Else
                            memberCompanyCaption = cp.Site.GetProperty("Supplier Directory Member Company Caption", "Member Company")
                            memberCompanyLink = cp.Site.GetProperty("Supplier Directory Member Company Link", "/")
                            exhibitorCaption = cp.Site.GetProperty("Supplier Directory Exhibitor Caption", "Exhibitor")
                            exhibitorLink = cp.Site.GetProperty("Supplier Directory Exhibitor Link", "/")
                            '
                            organizationID = cs.GetInteger("id")
                            companyName = cs.GetText("directoryname")
                            weblink = cs.GetText("directorylink")
                            web = cs.GetText("directoryweb")
                            imageFilename = cs.GetText("showcaseAdImageFilename")
                            If imageFilename = "" Then
                                imageFilename = cs.GetText("showcaseAdListingImageFilename")
                            End If
                            IsAssociationMember = cs.GetBoolean("directoryIsAssociationMember")
                            IsExhibitor = cs.GetBoolean("directoryIsExhibitor")
                            showcaseAdDescription = cs.GetText("showcaseAdDescription")
                            profileContact = cs.GetText("directoryprofileContact")
                            showcaseCaption = cs.GetText("showcaseAdCaption")
                            '

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
                                showcaseAdImageLine = cp.Html.div(image, , "showcaseAdImageLine")
                            End If
                            '
                            showcaseAdDescriptionLine = cs.GetText("showcaseAddescription")

                            If showcaseAdDescriptionLine <> "" Then
                                showcaseAdDescriptionLine = showcaseAdDescriptionLine.Replace(vbCrLf, "<br>")
                                showcaseAdDescriptionLine = showcaseAdDescriptionLine.Replace(vbLf, "<br>")
                                showcaseAdDescriptionLine = showcaseAdDescriptionLine.Replace(vbCr, "<br>")
                                showcaseAdDescriptionLine = cp.Html.div(showcaseAdDescriptionLine, , "descriptionLine")
                            End If
                            '
                            featureLine = ""
                            If IsAssociationMember Then
                                featureLine &= "<div><a href=""" & memberCompanyLink & """>" & memberCompanyCaption & "</a></div>"
                            End If
                            If IsExhibitor Then
                                featureLine &= "<div><a href=""" & exhibitorLink & """>" & exhibitorCaption & "</a></div>"
                            End If
                            If featureLine <> "" Then
                                featureLine = cp.Html.div(featureLine, , "featureLine")
                            End If
                            '
                            phoneLine = cs.GetText("directoryPhone")
                            If phoneLine <> "" Then
                                phoneLine = common.formatPhone(cp, phoneLine)
                                phoneLine = cp.Html.div(phoneLine, , "phoneLine")
                            End If
                            '
                            faxLine = cs.GetText("directoryfax")
                            If faxLine <> "" Then
                                faxLine = common.formatPhone(cp, faxLine)
                                faxLine = cp.Html.div(faxLine & "&nbsp;fax", , "faxLine")
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
                                emailLine = "" _
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
                                emailLine = emailLine.Replace("##organizationID##", organizationID)
                                emailLine = cp.Html.div(emailLine, , "emailLine")
                            End If
                            '
                            address1Line = ""
                            copy = cs.GetText("directoryaddress1")
                            If copy <> "" Then
                                address1Line = cp.Html.div(copy, , "address1Line")
                            End If
                            '
                            address2Line = ""
                            copy = cs.GetText("directoryaddress2")
                            If copy <> "" Then
                                address2Line = cp.Html.div(copy, , "address2Line")
                            End If
                            '
                            copy = cs.GetText("directorycity")
                            If copy <> "" Then
                                address3Line = copy
                            End If
                            copy = cs.GetText("directorystate")
                            If copy <> "" Then
                                If address3Line <> "" Then
                                    address3Line &= ", "
                                End If
                                address3Line &= copy
                            End If
                            copy = cs.GetText("directorycountry")
                            If copy <> "" Then
                                Select Case copy.ToLower
                                    Case "", "united states of america", "united states", "us", "usa", "u.s.a", "u.s."
                                    Case Else
                                        If address3Line <> "" Then
                                            address3Line &= " "
                                        End If
                                        address3Line &= copy
                                End Select
                            End If
                            copy = cs.GetText("directoryzip")
                            If copy <> "" Then
                                If address3Line <> "" Then
                                    address3Line &= " "
                                End If
                                address3Line &= copy
                            End If
                            If address3Line <> "" Then
                                address3Line = cp.Html.div(address3Line, , "address3Line")
                            End If
                            '
                            emailForm = ""
                            '
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
                            ''
                            ''
                            'categoryList = ""
                            'csHeadings = cp.CSNew
                            'sql = "select h.category as categoryname,h.name as headingName" _
                            '    & " from directorySubcategories h left join OrganizationSubcategoryRules r on r.headingid=h.id" _
                            '    & " where r.organizationid=" & organizationID _
                            '    & " and (r.approved<>0)" _
                            '    & " order by h.category, h.name"
                            'cp.Site.TestPoint("heading sql=" & sql)
                            'Call csHeadings.OpenSQL("", sql)
                            'Do While csHeadings.OK
                            '    copy = csHeadings.GetText("categoryname")
                            '    If copy <> "" Then
                            '        categoryList &= cp.Html.div(copy & ", " & csHeadings.GetText("headingname"))
                            '    End If
                            '    csHeadings.GoNext()
                            'Loop
                            'Call csHeadings.Close()
                            'If categoryList <> "" Then
                            '    categoryList = cp.Html.div(categoryList, , "categoryList")
                            'End If
                            ''
                            'profileContactLine = ""
                            If profileContact <> "" Then
                                profileContactLine = cp.Html.div(profileContact, , "profileContactLine")
                            End If
                            qs = rqs
                            qs = cp.Utils.ModifyQueryString(qs, "searchButton", "", False)
                            qs = cp.Utils.ModifyQueryString(qs, "searchText", companyName, True)
                            qs = cp.Utils.ModifyQueryString(qs, "searchPhrase", "", False)
                            qs = cp.Utils.ModifyQueryString(qs, "searchLocation", "", False)
                            qs = cp.Utils.ModifyQueryString(qs, rnHeadingID, "", False)
                            qs = cp.Utils.ModifyQueryString(qs, rnCategory, "", False)
                            qs = cp.Utils.ModifyQueryString(qs, rnformId, FormTextSearchResults)
                            listingLink = "" _
                                & "<div class=""publicListingLink""><a href=""?" & qs & """>Directory Listing</a></div>"
                            '
                            ' stuff layout and save bake
                            '
                            getForm = common.getLayout(cp, "Supplier Directory Showcase Ad")
                            If getForm = "" Then
                                cp.Site.TestPoint("showcaseAdClass.getForm, Showcase Ad layout not found")
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
                                getForm = getForm.Replace("##showcaseAdImageLine##", showcaseAdImageLine)
                                getForm = getForm.Replace("##showcaseAdDescriptionLine##", showcaseAdDescriptionLine)
                                getForm = getForm.Replace("##featureLine##", featureLine)
                                getForm = getForm.Replace("##phoneLine##", phoneLine)
                                getForm = getForm.Replace("##faxLine##", faxLine)
                                getForm = getForm.Replace("##webLine##", faxLine)
                                getForm = getForm.Replace("##emailLine##", emailLine)
                                getForm = getForm.Replace("##address1Line##", address1Line)
                                getForm = getForm.Replace("##address2Line##", address2Line)
                                getForm = getForm.Replace("##address3Line##", address3Line)
                                getForm = getForm.Replace("##emailForm##", emailForm)
                                getForm = getForm.Replace("##profileContactLine##", profileContactLine)
                                getForm = getForm.Replace("##categoryList##", categoryList)
                                getForm = getForm.Replace("##showcaseSubhead##", showcaseCaption)
                                getForm = getForm.Replace("##listingLink##", listingLink)
                                Call cp.Cache.Save(cacheName, getForm, "Organizations, Organization Subcategories, Directory Subcategories")
                            End If
                        End If
                        Call cs.Close()
                    End If
                End If
            Catch ex As Exception
                Call cp.Site.ErrorReport("showcaseAdClass.getForm, " & ex.Message)
            End Try
            '
        End Function

    End Class
End Namespace
