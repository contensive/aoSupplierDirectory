
Imports Contensive.BaseClasses
Imports Contensive.Addons.SupplierDirectory.Constants

Namespace Contensive.Addons.SupplierDirectory
    Friend Class HeadingSearchResultsClass

        '
        '============================================================================
        '   get Home Form
        '============================================================================
        '
        Friend Function getForm(ByVal cp As CPBaseClass, ByVal common As CommonClass, ByVal mapMode As Boolean) As String
            '
            '
            Dim returnHtml As String = ""
            '
            '
            Dim sqlCriteria As String = ""
            Dim emailFormPopupLink As String = ""
            Dim showcaseCaption As String = ""
            Dim organizationID As Integer = 0
            Dim csShowcase As CPCSBaseClass
            Dim pagePtr As Integer = 0
            Dim pageCnt As Integer = 0
            Dim ListingEnhancedSrc As String = ""
            Dim cell As String = ""
            Dim memberIcon As String = ""
            Dim addressLine As String = ""
            Dim descriptionLine As String = ""
            Dim fullDescriptionLine As String = ""
            Dim webLine As String = ""
            Dim web As String
            Dim weblink As String = ""
            Dim imageFilename As String = ""
            Dim memberIconSrc As String = ""
            Dim memberCompanyLink As String = ""
            Dim IsAssociationMember As Boolean = False
            Dim copy As String = ""
            Dim fullCopy As String = ""
            Dim pos As Integer
            Dim companyLine As String = ""
            Dim faxLine As String = ""
            Dim phoneLine As String = ""
            Dim profileButtonLine As String = ""
            Dim cs As CPCSBaseClass
            Dim pageNumber As Integer = 0
            Dim pageSize As Integer = 0
            Dim categoryLast As String = ""
            Dim rqs As String = ""
            Dim qs As String = ""
            Dim qs10 As String = ""
            Dim qs50 As String = ""
            Dim qs100 As String = ""
            Dim cellLeft As String = ""
            Dim cellRight As String = ""
            Dim cellCenter As String = ""
            Dim cellRowMax As Integer = 0
            Dim cellRowPtr As Integer = 0
            Dim Hint As String = ""
            Dim rowTotal As Integer = 0
            Dim cacheName As String = ""
            Dim headingID As Integer = 0
            Dim subCategoryID As Integer = 0
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
            Dim emailBody As String = ""
            Dim recordPtr As Integer = 0
            Dim pageNavigation As String = ""
            Dim emailForm As String = ""
            'Dim searchTermsMessage As String = ""
            Dim isEnhancedListing As Boolean = False
            Dim contactLine As String = ""
            '
            Dim aliasName As String = ""
            Dim aliasPrefix As String = cp.Site.GetProperty("Supplier Directory Vanity URL Prefix")
            Dim csA As BaseClasses.CPCSBaseClass = cp.CSNew()
            Dim form As CPBlockBaseClass = cp.BlockNew()
            Dim myEmailForm As CPBlockBaseClass = cp.BlockNew()
            Dim buyerResultslSection As CPBlockBaseClass = cp.BlockNew
            Dim resultsGuideListingSection As CPBlockBaseClass = cp.BlockNew
            Dim blockLayout As CPBlockBaseClass = cp.BlockNew
            Dim exitLoop As Boolean = False
            Dim buyerGuideBoxHtml As String
            '
            cp.Site.TestPoint("HeadingSearchResultsClass.getForm")
            '
            returnHtml = ""
            Try
                subCategoryID = cp.Request.GetInteger(rnSubcategoryID)
                If subCategoryID <> 0 Then
                    pageNumber = cp.Request.GetInteger(rnPageNumber)
                    If pageNumber = 0 Then
                        pageNumber = 1
                    End If
                    pageSize = cp.Request.GetInteger(rnPageSize)
                    If pageSize = 0 Then
                        pageSize = cp.Utils.EncodeInteger(cp.Visit.GetProperty("Supplier Directory PageSize", "10"))
                    Else
                        Call cp.Visit.SetProperty("Supplier Directory PageSize", CStr(pageSize))
                        pageNumber = 1
                    End If
                    cp.Site.TestPoint("pageNumber=" & pageNumber & ",pageSize=" & pageSize & "")
                    '
                    Call cp.Doc.AddRefreshQueryString(rnCategoryID, subCategoryID)
                    Call cp.Doc.AddRefreshQueryString(rnPageNumber, pageNumber)
                    '
                    ' attempt cache read
                    '
                    cacheName = "Supplier Directory Subcategory Search Results, " & subCategoryID & "-" & pageNumber & "-" & pageSize
                    returnHtml = common.cacheRead(cp, cacheName)
                    If returnHtml = "" Then
                        '
                        ' no cache, verify heading
                        '
                        cs = cp.CSNew
                        Dim sdSql As String = ""
                        sdSql = "select ds.name,ds.id as subId,dc.name as categoryName, ds.categoryId, dc.id from directorySubcategories ds inner Join directoryCategories dc " _
                                     & "on ds.categoryID = dc.id where ds.id=" & subCategoryID
                        cs.OpenSQL("", sdSql)
                        If cs.OK Then
                            headingName = cs.GetText("name")
                            categoryName = cs.GetText("categoryName")
                        End If
                        Call cs.Close()
                        If headingName <> "" Then
                            '
                            ' valid heading, build form
                            '
                            cp.Doc.AddRefreshQueryString(rnCategoryID, subCategoryID)
                            cp.Site.TestPoint("HeadingSearchResultsClass.getForm, build form for heading [" & headingName & "]")
                            rqs = cp.Doc.RefreshQueryString
                            '
                            ' pageSizeOptions
                            '
                            pageSizeOptions = vbCrLf & vbTab & "<span class=""pageSizeOptions"">"
                            Select Case pageSize
                                Case 10
                                    pageSizeOptions &= vbCrLf & vbTab & vbTab & "<span class=""select10"">"
                                Case 50
                                    pageSizeOptions &= vbCrLf & vbTab & vbTab & "<span class=""select50"">"
                                Case 100
                                    pageSizeOptions &= vbCrLf & vbTab & vbTab & "<span class=""select100"">"
                                Case Else
                                    pageSizeOptions &= vbCrLf & vbTab & vbTab & "<span>"
                            End Select
                            qs10 = cp.Utils.ModifyQueryString(rqs, rnPageSize, "10")
                            qs50 = cp.Utils.ModifyQueryString(rqs, rnPageSize, "50")
                            qs100 = cp.Utils.ModifyQueryString(rqs, rnPageSize, "100")
                            pageSizeOptions &= "" _
                                & vbCrLf & vbTab & vbTab & vbTab & "<a class=""size10"" href=""?" & qs10 & """>10</a>" _
                                & vbCrLf & vbTab & vbTab & vbTab & "<a class=""size50"" href=""?" & qs50 & """>50</a>" _
                                & vbCrLf & vbTab & vbTab & vbTab & "<a class=""size100"" href=""?" & qs100 & """>100</a>" _
                                & vbCrLf & vbTab & vbTab & "</span>" _
                                & vbCrLf & vbTab & "</span>" _
                                & ""
                            ''
                            '' view options
                            ''
                            'qs = rqs
                            'If mapMode Then
                            '    qs = cp.Utils.ModifyQueryString(qs, rnMapMode, "0")
                            '    viewOptions = "<a class=""viewOption"" href=""?" & qs & """>List View</a>"
                            'Else
                            '    qs = cp.Utils.ModifyQueryString(qs, rnMapMode, "1")
                            '    viewOptions = "<a class=""viewOption"" href=""?" & qs & """>Map View</a>"
                            'End If
                            '
                            ' record count
                            '
                            sqlCriteria = "" _
                                & " (r.headingid=" & subCategoryID & ")" _
                                & " and ((r.dateexpires is null)or(r.dateexpires>" & cp.Db.EncodeSQLDate(Now) & "))" _
                                & " and (r.approved<>0)" _
                                & " and (o.includeInDirectory=1)"

                            sql = "select count(*) as cnt" _
                                & " from organizations o left join OrganizationSubcategoryRules r on r.organizationid=o.id" _
                                & " where " & sqlCriteria

                            cp.Utils.AppendLog("sql Statement.log", sql)

                            Call cs.OpenSQL(sql, "")
                            If cs.OK Then
                                recordCnt = cs.GetInteger("cnt")
                            End If
                            cs.Close()
                            If recordCnt = 0 Then
                                recordCntMessage = "No listings found"
                            ElseIf recordCnt = 1 Then
                                recordCntMessage = "1 listing found"
                            Else
                                recordCntMessage = recordCnt & " listings found"
                            End If
                            '
                            ' replace with new tempaltes
                            '

                            'ListingEnhancedSrc = common.getLayout(cp, "Supplier Directory Listing Enhanced")
                            'ListingNormalSrc = common.getLayout(cp, "Supplier Directory Listing Normal")
                            'memberCompanyCaption = cp.Site.GetProperty("Supplier Directory Member Company Caption", "Member Company")
                            'memberCompanyLink = cp.Site.GetProperty("Supplier Directory Member Company Link", "/")
                            'exhibitorCaption = cp.Site.GetProperty("Supplier Directory Exhibitor Caption", "Exhibitor")
                            'exhibitorLink = cp.Site.GetProperty("Supplier Directory Exhibitor Link", "/")
                            'showcaseCaption = cp.Site.GetProperty("Supplier Directory Showcase Caption", "Product Showcase")
                            '
                            'returnHtml = common.getLayout(cp, "Supplier Directory Subcategory Search Results")

                            Dim formSrc As String
                            Dim listingEnhanded As CPBlockBaseClass = cp.BlockNew()
                            Dim listingNormal As CPBlockBaseClass = cp.BlockNew()

                            form.OpenLayout("bgSearchResultsPage.html")
                            formSrc = form.GetHtml()

                            ListingEnhancedSrc = form.GetOuter(".enhancedList")
                            listingEnhanded.Load(ListingEnhancedSrc)
                            '

                            ListingNormalSrc = form.GetOuter(".normalList")
                            listingNormal.Load(ListingNormalSrc)
                            '
                            memberCompanyCaption = cp.Site.GetProperty("Supplier Directory Member Company Caption", "Member Company")
                            memberCompanyLink = cp.Site.GetProperty("Supplier Directory Member Company Link", "/")
                            exhibitorCaption = cp.Site.GetProperty("Supplier Directory Exhibitor Caption", "Exhibitor")
                            exhibitorLink = cp.Site.GetProperty("Supplier Directory Exhibitor Link", "/")
                            showcaseCaption = cp.Site.GetProperty("Supplier Directory Showcase Caption", "Product Showcase")
                            '
                            ' the new stuff
                            '
                            '
                           


                            ' save cache
                            'Dim layout As CPBlockBaseClass = cp.BlockNew
                            'buyerResultslSection.OpenLayout("04buyers-guide.html")
                            'Dim resultContents As String = buyerResultslSection.GetOuter(".tlpaTierSection")
                            'buyerResultslSection.Load(resultContents)
                            'Dim buyersResultBox As CPBlockBaseClass = cp.BlockNew()
                            'Dim buyersResultSponserHbox As CPBlockBaseClass = cp.BlockNew()
                            'buyersResultBox.Load(buyerResultslSection.GetOuter(".buyerResults"))
                            'buyersResultSponserHbox.Load(buyersGuideSection.GetOuter(".sponsHBox"))

                            Dim layout As CPBlockBaseClass = cp.BlockNew
                            buyerResultslSection.OpenLayout("bgSearchResultsPage.html")
                            Dim resultContents As String = buyerResultslSection.GetOuter(".buyerGuide-results")
                            buyerResultslSection.Load(resultContents)

                            Dim rebuyersResultBox As CPBlockBaseClass = cp.BlockNew()


                            resultsGuideListingSection.OpenLayout("BuyersGuide Listings Enhanced")
                            Dim ResultslistSection As String = resultsGuideListingSection.GetOuter(".buyerResultsInner")
                            resultsGuideListingSection.Load(ResultslistSection)

                            'Dim buyersResultSponserHbox As CPBlockBaseClass = cp.BlockNew()
                            'buyersResultBox.Load(buyerResultslSection.GetOuter(".buyerResults"))
                            'buyersResultSponserHbox.Load(buyersGuideSection.GetOuter(".sponsHBox"))



                            'blockLayout.Load(layout.GetOuter("#js-buyerResults"))
                            'Dim buyerResultsBox As CPBlockBaseClass = cp.BlockNew()
                            'Dim resultULInner As String = buyerResultsBox.GetInner(".buyerResultsUL")
                            'Dim tmpBlock As CPBlockBaseClass = cp.BlockNew()
                            'tmpBlock.Load(vbCrLf & "<div>xxx</div>" & resultULInner)
                            '
                            '
                            'blockLayout.Load(layout.GetOuter("#js-buyerResults"))
                            Dim buyerResultsBox As CPBlockBaseClass = cp.BlockNew()
                            Dim resultULInner As String = buyerResultsBox.GetInner(".buyerResultsInner")
                            buyerResultsBox.Load(resultULInner)
                            Dim tmpBlock As CPBlockBaseClass = cp.BlockNew()
                            'tmpBlock.Load(vbCrLf & "<div>xxx</div>" & resultULInner)
                            Dim emailLink As String = "##emailLinkLine##"
                            Dim webLines As String = "##webLine##"
                            Dim resultULliTemplate As String = tmpBlock.GetInner(".buyerResultsInner")

                            form.SetInner(".bgHome", "##home##")
                            form.SetInner(".bgCategory", "##category##")
                            form.SetInner(".bgHeading", "##heading##")
                            form.SetInner(".buyerResultsUL", "##list##")
                            returnHtml = form.GetHtml()
                            '
                            listingNormal.SetInner(".company", "##companyLine##")
                            listingNormal.SetInner(".location", "##addressLine##")
                            listingNormal.SetInner(".sales-phone", "##phoneLine##")
                            listingNormal.SetInner(".fax", "##faxLine##")
                            listingNormal.SetInner(".description", "##descriptionLine##")
                            listingNormal.SetInner(".longDescription", "##longdescriptionLine##")
                            listingNormal.SetInner(".emailBtn", "##emailLinkLine##")
                            listingNormal.SetOuter(".buyerBtn", cp.Site.EncodeAppRootPath(webLines))
                            ListingNormalSrc = listingNormal.GetHtml()
                            '
                            listingEnhanded.SetInner(".buy-logo-inner", "##image##")
                            listingEnhanded.SetInner(".company", "##companyLine##")
                            listingEnhanded.SetInner(".location", "##addressLine##")
                            listingEnhanded.SetInner(".sales-phone", "##phoneLine##")
                            listingEnhanded.SetInner(".fax", "##faxLine##")
                            listingEnhanded.SetInner(".description", "##descriptionLine##")
                            listingEnhanded.SetInner(".longDescription", "##longdescriptionLine##")
                            listingEnhanded.SetOuter(".emailBtn", "##emailLinkLine##")
                            listingEnhanded.SetOuter(".bgLearnMore", "##profileButtonLine##")
                            'bgLearnMore
                            listingEnhanded.SetOuter(".buyerBtn", cp.Site.EncodeAppRootPath(webLines))
                            ListingEnhancedSrc = listingEnhanded.GetHtml()
                            '
                            'oneBuyer
                            'buyerResultslSection.SetInner(".buyerGuide-results", buyerResultslSection.GetHtml())
                            'resultsGuideListingSection.SetInner(".buyerResultsInner", resultsGuideListingSection.GetHtml())
                            '
                            '
                            ' Dim formSrc As String
                            Dim listingBannerAdSrc As String = ""
                            ' Dim form As CPBlockBaseClass = CP.BlockNew()
                            Dim listingBannerAd As CPBlockBaseClass = cp.BlockNew()
                            listingBannerAdSrc = listingBannerAd.GetInner("#spbHorizontal")
                            listingBannerAd.Load(listingBannerAdSrc)
                            'listingBannerAd.OpenLayout("bgSearchResultsPage.html")
                            'listingBannerAdSrc = listingBannerAd.GetHtml()
                            ' put the #names# into the layout
                            '    '.sponsHBoxInner = bottom banner
                            '
                            listingBannerAd.SetInner(".sponsHBoxInner", "##bannerBottom##")


                            listingBannerAdSrc = listingBannerAd.GetHtml()
                            ' run the old code
                            '
                            '
                            '
                            'returnHtml = buyerResultslSection.GetHtml()
                            '
                            ' ----- now continue on the old way, with the new layouts
                            '
                            If recordCnt = 0 Then
                                recordList = cp.Content.GetCopy("Heading Search Results, no listings found", "<p class=""emptyList"">No companies were found under this listing.</p>")
                            Else
                                '
                                ' /replace with new tempaltes
                                '
                                sql = "select o.*" _
                                    & " from (( OrganizationSubcategoryRules r" _
                                    & " left join organizations o on r.organizationid=o.id )" _
                                    & " left join directoryplacements s on r.placementid=s.id )" _
                                    & " where " & sqlCriteria _
                                    & " order by s.priority desc,( o.directoryIsEnhancedListing * o.directoryprofileApprovedByAccount ) desc, o.directoryname"
                                'sql = "select o.*" _
                                '    & " from (((organizations o " _
                                '    & " left join OrganizationSubcategoryRules r on r.organizationid=o.id)" _
                                '    & " left join (select * from directoryheadingplacement where headingid=" & headingID & ") p on p.organizationid=o.id)" _
                                '    & " left join directoryplacements s on p.placementid=s.id)" _
                                '    & " where(r.headingid = " & headingID & ")" _
                                '    & " and ((p.approved<>0)or(p.id is null))" _
                                '    & " order by s.priority desc,o.directoryisEnhancedListing desc, o.name"
                                Call cs.OpenSQL(sql, "", pageSize, pageNumber)
                                recordPtr = 0
                                Do While cs.OK And (recordPtr < pageSize)
                                    organizationID = cs.GetInteger("id")
                                    companyName = cs.GetText("directoryName")
                                    weblink = cs.GetText("directorylink")
                                    web = cs.GetText("directoryweb")
                                    imageFilename = cs.GetText("directoryListingImageFilename")
                                    IsAssociationMember = cs.GetBoolean("directoryIsAssociationMember")
                                    IsExhibitor = cs.GetBoolean("directoryIsExhibitor")
                                    profileButtonLine = ""
                                    descriptionLine = ""
                                    webLinkLine = ""
                                    emailBody = ""
                                    isEnhancedListing = cs.GetBoolean("directoryProfileApprovedByAccount") And cs.GetBoolean("directoryisEnhancedListing")
                                    '
                                    If web = "" Then
                                        web = weblink
                                    End If
                                    If weblink = "" Then
                                        weblink = web
                                    End If
                                    If (weblink <> "") Then
                                        If weblink.IndexOf("://") < 0 Then
                                            weblink = "http://" & weblink
                                        End If
                                    End If
                                    '
                                    featureLine = ""
                                    If IsAssociationMember And (memberCompanyCaption <> "") Then
                                        featureLine &= ",&nbsp;<a href=""" & memberCompanyLink & """ target=""_blank"">" & memberCompanyCaption & "</a>"
                                    End If
                                    If IsExhibitor And (exhibitorCaption <> "") Then
                                        featureLine &= ",&nbsp;<a href=""" & exhibitorLink & """ target=""_blank"">" & exhibitorCaption & "</a>"
                                    End If
                                    sql = "select top 1 id from DirectoryShowcaseAds where (organizationid=" & organizationID & ")and(approved<>0)and(ApprovedByAccount<>0)"
                                    csShowcase = cp.CSNew
                                    Call csShowcase.OpenSQL("", sql)
                                    If csShowcase.OK Then
                                        qs = rqs
                                        qs = cp.Utils.ModifyQueryString(qs, rnShowcaseAdID, csShowcase.GetInteger("ID").ToString)
                                        qs = cp.Utils.ModifyQueryString(qs, rnformId, formIdShowcaseAd)
                                        featureLine &= ",&nbsp;<a href=""?" & qs & """>" & showcaseCaption & "</a>"
                                    End If
                                    Call csShowcase.Close()
                                    If featureLine <> "" Then
                                        featureLine = cp.Html.div(featureLine.Substring(7), , "featureLine")
                                    End If
                                    '
                                    contactLine = cs.GetText("directoryprofileContact")
                                    If contactLine <> "" Then
                                        contactLine = cp.Html.div(contactLine.Replace(" ", "&nbsp;"), , "contactLine")
                                    End If
                                    '
                                    phoneLine = cs.GetText("directoryphone")
                                    If phoneLine <> "" Then
                                        phoneLine = common.formatPhone(cp, phoneLine)
                                        phoneLine = cp.Html.div(phoneLine.Replace(" ", "&nbsp;"), , "phoneLine")
                                    End If
                                    '
                                    faxLine = cs.GetText("directoryfax")
                                    If faxLine <> "" Then
                                        faxLine = common.formatPhone(cp, faxLine)
                                        faxLine = cp.Html.div(faxLine.Replace(" ", "&nbsp;") & "&nbsp;Fax", , "faxLine")
                                    End If
                                    '
                                    addressLine = ""
                                    copy = cs.GetText("directoryaddress1")
                                    If copy <> "" Then
                                        addressLine &= ",&nbsp;" & copy.Replace(" ", "&nbsp;")
                                    End If
                                    copy = cs.GetText("directoryaddress2")
                                    If copy <> "" Then
                                        addressLine &= ",&nbsp;" & copy.Replace(" ", "&nbsp;")
                                    End If
                                    copy = cs.GetText("directorycity")
                                    If copy <> "" Then
                                        addressLine &= ", " & copy.Replace(" ", "&nbsp;")
                                    End If
                                    copy = cs.GetText("directorystate")
                                    If copy <> "" Then
                                        addressLine &= ",&nbsp;" & copy.Replace(" ", "&nbsp;")
                                    End If
                                    copy = cs.GetText("directoryzip")
                                    If copy <> "" Then
                                        addressLine &= "&nbsp;" & copy.Replace(" ", "&nbsp;")
                                    End If
                                    copy = cs.GetText("directorycountry")
                                    Select Case copy.ToLower
                                        Case "", "united states of america", "united states", "us", "usa", "u.s.a", "u.s."
                                        Case Else
                                            addressLine &= "<br>" & copy.Replace(" ", "&nbsp;")
                                    End Select
                                    If addressLine <> "" Then
                                        addressLine = cp.Html.div(addressLine.Substring(7), , "addressLine")
                                    End If
                                    If Not isEnhancedListing Then
                                        '
                                        ' Normal Listing
                                        '
                                        cell = ListingNormalSrc
                                        '
                                        companyLine = ""
                                        If companyName <> "" Then
                                            companyLine = cp.Html.div(companyName, , "companyLine")
                                        End If
                                    Else
                                        '
                                        ' Enhanced listing
                                        '
                                        cell = ListingEnhancedSrc
                                        '
                                        companyLine = ""
                                        If companyName <> "" Then
                                            If (weblink <> "") Then
                                                ' companyLine = cp.Html.div("<a href=""" & weblink & """ target=""_blank"">" & companyName & "</a>", , "companyLine")
                                                companyLine = cp.Html.div(companyName, , "companyLine")
                                            Else
                                                companyLine = cp.Html.div(companyName, , "companyLine")
                                            End If
                                        End If
                                        '
                                        If (web <> "") Then
                                            webLine = cp.Html.div("<a href=""" & weblink & """ target=""_blank"">" & web & "</a>", , "webLine")
                                            webLinkLine = cp.Html.div("<a href=""" & weblink & """ target=""_blank"">web</a>", , "webLinkLine")
                                        End If
                                        '
                                        '
                                        myEmailForm.OpenLayout("bgOldEmailForm.html")
                                        emailForm = myEmailForm.GetOuter(".emailForm")
                                        emailForm = emailForm.Replace("##companyName##", companyName)
                                        emailForm = emailForm.Replace("##organizationID##", organizationID.ToString())

                                        email = cs.GetText("directoryemail")
                                        If email <> "" Then
                                            emailBody = "" _
                                                & vbCrLf & "<a class=""emailBtn"" id=""supplierDirectoryEmailFormLink" & organizationID & """ href=""#supplierDirectoryEmailPopup" & organizationID & """>E-mail</a>" _
                                                & vbCrLf & emailForm _
                                                & vbCrLf & "<script>" _
                                                    & "$(document).ready(function() {" _
                                                        & "$('#supplierDirectoryEmailFormLink" & organizationID & "').fancybox({" _
                                                            & "'titleShow':false," _
                                                            & "'transitionIn':'fade'," _
                                                            & "'transitionOut':'fade'," _
                                                            & "'overlayOpacity':'.6'," _
                                                            & "'overlayColor':'#000000'" _
                                                        & "});" _
                                                    & "});" _
                                                & "</script>" _
                                                & ""
                                            cp.Utils.AppendLog("emailLink.log", emailBody)
                                        End If
                                        '
                                        If imageFilename = "" Then
                                            image = "<img width=""120"" height=""60"" alt=""" & cp.Utils.EncodeHTML(companyName) & """ src=""/cclib/images/spacer.gif"">"
                                        Else
                                            image = "<img width=""120"" height=""60"" alt=""" & cp.Utils.EncodeHTML(companyName) & """ src=""" & cp.Site.FilePath & imageFilename & """>"
                                            If weblink <> "" Then
                                                image = "<a href=""" & weblink & """ target=""_blank"">" & image & "</a>"
                                            End If
                                        End If
                                        '
                                        fullCopy = cs.GetText("directoryprofiledescription")
                                        copy = cs.GetText("directoryprofiledescription")
                                        If copy <> "" Then
                                            If copy.Length > 300 Then
                                                pos = copy.IndexOf(".", 300)
                                                If (pos < 300) Or (pos > 350) Then
                                                    pos = copy.IndexOf(" ", 300)
                                                End If
                                                If (pos < 300) Or (pos > 350) Then
                                                    copy = copy.Substring(0, 300) & "..."
                                                Else
                                                    copy = copy.Substring(0, pos) & "..."
                                                End If
                                            End If
                                        End If
                                        descriptionLine = copy
                                        'If (descriptionLine <> "") And (weblink <> "") Then
                                        'descriptionLine = "<a style=""description"" href=""" & weblink & """ target=""_blank"">" & descriptionLine & "</a>"
                                        'End If
                                        If descriptionLine <> "" Then
                                            descriptionLine = cp.Html.div(descriptionLine, , "descriptionLine")
                                        End If
                                        fullDescriptionLine = fullCopy
                                        '
                                        qs = rqs
                                        qs = cp.Utils.ModifyQueryString(qs, "searchButton", "")
                                        qs = cp.Utils.ModifyQueryString(qs, "searchText", "")
                                        qs = cp.Utils.ModifyQueryString(qs, "searchPhrase", "")
                                        qs = cp.Utils.ModifyQueryString(qs, "searchLocation", "")
                                        qs = cp.Utils.ModifyQueryString(qs, rnformId, formIdProfile)
                                        qs = cp.Utils.ModifyQueryString(qs, rnOrganizationID, cs.GetInteger("id"))
                                        '
                                        aliasName = aliasPrefix & "/" & cs.GetText("name")
                                        aliasName = Replace(aliasName, "'", "-")
                                        aliasName = Replace(aliasName, """", "-")
                                        aliasName = Replace(aliasName, " ", "-")
                                        '
                                        If csA.Open("Link Aliases", "name=" & cp.Db.EncodeSQLText(aliasName)) Then
                                            profileButtonLine = "<a href=""?" & qs & """>Profile</a>"
                                            '& "##profileButtonLine##"
                                            'profileButtonLine = "<a href=""" & aliasName & """>Profile</a>"
                                        Else
                                            profileButtonLine = "<a href=""?" & qs & """>Profile</a>"
                                        End If
                                        csA.Close()
                                        '
                                        If profileButtonLine <> "" Then
                                            profileButtonLine = cp.Html.div(profileButtonLine, , "profileButtonLine")
                                        End If
                                    End If
                                    cell = cell.Replace("##image##", image)
                                    cell = cell.Replace("##companyLine##", companyLine)
                                    cell = cell.Replace("##featureLine##", featureLine)
                                    cell = cell.Replace("##addressLine##", addressLine)
                                    Dim value As String = ""
                                    Dim length As Integer = value.Length
                                    value += descriptionLine
                                    length = value.Length
                                    cp.Utils.AppendLog("discription.log", length)
                                   
                                    cell = cell.Replace("##descriptionLine##", fullDescriptionLine & "##profileButtonLine##")
                                   
                                    'cell = cell.Replace("##descriptionLine##", descriptionLine)
                                    cell = cell.Replace("##webLine##", webLine)
                                    cell = cell.Replace("##profileButtonLine##", profileButtonLine)
                                    cell = cell.Replace("##phoneLine##", phoneLine)
                                    cell = cell.Replace("##faxLine##", faxLine)
                                    cell = cell.Replace("##webLine##", webLine)
                                    'cell = cell.Replace("##webLinkLine##", webLinkLine)
                                    cell = cell.Replace("##emailLinkLine##", emailBody)
                                    cell = cell.Replace("##contactLine##", contactLine)
                                    recordList &= cell
                                    recordPtr += 1
                                    Call cs.GoNext()
                                Loop
                                Call cs.Close()
                            End If
                            '
                            ' page Navigation
                            '
                            pageCnt = Int((recordCnt + pageSize - 1) / pageSize)
                            cp.Site.TestPoint("recordCnt=" & recordCnt & ",pageSize=" & pageSize & ",pageCnt=" & pageCnt & "")
                            If pageCnt > 1 Then
                                For pagePtr = 1 To pageCnt
                                    qs = rqs
                                    qs = cp.Utils.ModifyQueryString(qs, rnPageNumber, pagePtr)
                                    pageNavigation &= cp.Html.li(cp.Html.div("<a href=""?" & qs & """>" & pagePtr & "</a>"))
                                Next
                                pageNavigation = cp.Html.ul(pageNavigation)
                            End If
                            '
                            qs = rqs
                            qs = cp.Utils.ModifyQueryString(qs, "searchButton", "")
                            qs = cp.Utils.ModifyQueryString(qs, "searchText", "")
                            qs = cp.Utils.ModifyQueryString(qs, "searchPhrase", "")
                            qs = cp.Utils.ModifyQueryString(qs, "searchLocation", "")
                            qs = cp.Utils.ModifyQueryString(qs, rnHeadingID, "", False)
                            qs = cp.Utils.ModifyQueryString(qs, rnCategory, "", False)
                            qs = cp.Utils.ModifyQueryString(qs, rnformId, formIdHome)
                            returnHtml = returnHtml.Replace("##home##", "<a href=""?" & qs & """>All Categories</a>")
                            qs = cp.Utils.ModifyQueryString(qs, rnCategory, categoryName)
                            qs = cp.Utils.ModifyQueryString(qs, rnformId, formIdCategoryHeading)
                            returnHtml = returnHtml.Replace("##category##", "<a href=""?" & qs & """>" & categoryName & "</a>")
                            returnHtml = returnHtml.Replace("##heading##", headingName)
                            returnHtml = returnHtml.Replace("##viewOptions##", viewOptions)
                            returnHtml = returnHtml.Replace("##pageSizeOptions##", pageSizeOptions)
                            returnHtml = returnHtml.Replace("##recordCntMessage##", recordCntMessage)
                            returnHtml = returnHtml.Replace("##list##", recordList)
                            returnHtml = returnHtml.Replace("##pageNavigation##", pageNavigation)
                            returnHtml = returnHtml.Replace("##showcaseList##", listingBannerAdSrc)

                            '
                            Call cp.Cache.Save(cacheName, returnHtml, "Organizations, Organization Subcategories, Directory Subcategories")
                        End If
                    End If
                End If
            Catch ex As Exception
                Call cp.Site.ErrorReport("HeadingSearchResultsClass, " & ex.Message)
            End Try
            Return returnHtml
        End Function

    End Class
End Namespace
