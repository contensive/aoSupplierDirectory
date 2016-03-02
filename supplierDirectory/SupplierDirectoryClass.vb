
Imports Contensive.BaseClasses
Imports Contensive.Addons.SupplierDirectory.Constants
'
Namespace Contensive.Addons.SupplierDirectory

    Public Class SupplierDirectoryClass
        Inherits AddonBaseClass
        '
        Dim Global_HorizontalBannerList As String = ""
        Dim Global_VerticalBannerList As String = ""
        '
        ' Entry Point for this Addon
        '
        Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
            '
            Dim cs As Contensive.BaseClasses.CPCSBaseClass = CP.CSNew
            Dim blockWithLogin As Boolean
            Dim showcaseEdit As ShowcaseEditClass
            Dim bannerEdit As BannerEditClass
            Dim categoryHeading As CategoryHeadingClass
            Dim MyAccountEdit As MyAccountEditClass
            Dim MyAccount As MyAccountClass
            Dim footerNav As String = ""
            Dim AllowBannerAds As Boolean = False
            Dim copy As String = ""
            Dim searchBanner As String = ""
            Dim home As HomeClass
            Dim banners() As String
            Dim bannerTopPtr As Integer
            Dim bannerBottomPtr As Integer
            Dim bannerLeftPtr As Integer
            Dim bannerRightPtr As Integer
            Dim bannerCnt As Integer
            Dim formId As Integer = 0
            Dim layout As String = ""
            Dim s As String = ""
            Dim content As String = ""
            'Dim CS As CPCSBaseClass
            Dim bannerLinkPrefix As String = ""
            Dim bannerLeft As String = ""
            Dim bannerRight As String = ""
            Dim bannerTop As String = ""
            Dim bannerBottom As String = ""
            Dim bannerWrapper As String = ""
            Dim bannerList As String = ""
            Dim bannerRow As String = ""
            Dim bannerAttr() As String
            Dim bannerAlt As String = ""
            Dim bannerLink As String = ""
            Dim bannerImageFilename As String = ""
            Dim hint As String = ""
            Dim common As New CommonClass
            Dim headingSearchResults As HeadingSearchResultsClass
            Dim profile As ProfileClass
            Dim showcaseAd As ShowcaseAdClass
            Dim textSearchResults As TextSearchResultsClass
            Dim qs As String = ""
            Dim rqs As String = CP.Doc.RefreshQueryString
            Dim searchText As String = ""
            Dim searchButton As String = ""
            Dim searchPhrase As Boolean = False
            Dim searchLocation As String = ""
            Dim searchAdvanced As Boolean = False
            Dim mapMode = False
            Dim orgId As Integer = 0
            Dim js As String = ""
            '
            Execute = ""
            Try
                '
                ' --------------------------------------------------------------------
                ' common requests
                ' --------------------------------------------------------------------
                '
                copy = CP.Request.GetText(rnMapMode)
                If copy <> "" Then
                    If CP.Utils.EncodeBoolean(copy) Then
                        mapMode = True
                    End If
                    Call CP.Site.SetProperty("Supplier Directory Map Mode", mapMode.ToString)
                Else
                    mapMode = CP.Site.GetProperty("Supplier Directory Map Mode", "0")
                End If
                searchButton = CP.Request.GetText("searchButton")
                formId = CP.Request.GetInteger(rnformId)
                If formId = 0 Then
                    formId = formIdHome
                End If
                orgId = CP.Utils.EncodeInteger(CP.Doc.Var(rnOrganizationID))
                '
                ' --------------------------------------------------------------------
                ' authenticated check (only allow public pages to non-auth)
                ' --------------------------------------------------------------------
                '
                blockWithLogin = False
                If Not CP.User.IsAuthenticated Then
                    Select Case formId
                        Case formIdHome, formIdHeadingSearchREsults, formIdCategoryHeading, formIdProfile, FormSearchMapResults, FormTextSearchResultsMore, FormTextSearchResults, formIdShowcaseAd
                            '
                            ' Allow Anon
                            '
                        Case Else
                            '
                            ' require authentication
                            '
                            blockWithLogin = True
                    End Select
                End If
                If blockWithLogin Then
                    '
                    ' block with login
                    '
                    Call CP.Doc.AddRefreshQueryString(rnformId, formId.ToString)
                    content = common.getLayout(CP, "Supplier Directory Login")
                    content = content.Replace("##login##", CP.Utils.ExecuteAddon("Login Form"))
                    qs = rqs
                    qs = CP.Utils.ModifyQueryString(qs, rnformId, formIdHome)
                    content = content.Replace("##homeLink##", "?" & qs)
                Else
                    '
                    ' continue without login form
                    '
                    searchText = CP.Request.GetText("searchText")
                    searchLocation = CP.Request.GetText("searchLocation")
                    If ((searchText <> "") Or (searchLocation <> "")) And (CP.Visit.CookieSupport) Then
                        '
                        ' --------------------------------------------------------------------
                        ' search button hit, skip processesing and go to results page
                        ' --------------------------------------------------------------------
                        '
                        'searchText = CP.Request.GetText("searchText")
                        'searchLocation = CP.Request.GetText("searchLocation")
                        If searchPhrase Or (searchLocation <> "") Then
                            searchAdvanced = True
                        End If
                        Call CP.Doc.AddRefreshQueryString("searchButton", CP.Request.GetText("searchButton"))
                        Call CP.Doc.AddRefreshQueryString("searchText", searchText)
                        Call CP.Doc.AddRefreshQueryString("searchPhrase", searchPhrase)
                        Call CP.Doc.AddRefreshQueryString("searchLocation", searchLocation)
                        formId = FormTextSearchResults
                    Else
                        '
                        ' --------------------------------------------------------------------
                        ' Process Previous Form
                        ' --------------------------------------------------------------------
                        '
                        If formId > 0 Then
                            Select Case formId
                                Case formIdShowcaseEdit
                                    '
                                    '
                                    '
                                    If common.isAccountContact(CP, orgId) Then
                                        showcaseEdit = New ShowcaseEditClass
                                        formId = showcaseEdit.processForm(CP, common)
                                    End If
                                Case formIdBannerEdit
                                    '
                                    '
                                    '
                                    If common.isAccountContact(CP, orgId) Then
                                        bannerEdit = New BannerEditClass
                                        formId = bannerEdit.processForm(CP, common)
                                    End If
                                Case formIdMyAccountEdit
                                    '
                                    '
                                    '
                                    '
                                    If common.isAccountContact(CP, orgId) Then
                                        MyAccountEdit = New MyAccountEditClass
                                        formId = MyAccountEdit.processForm(CP, common, orgId)
                                    End If
                                Case formIdMyAccount
                                    '
                                    '
                                    '
                                Case formIdCategoryHeading
                                    '
                                    '
                                    '
                                Case formIdHeadingSearchREsults
                                    '
                                    '
                                    '
                                Case formIdProfile
                                    '
                                    '
                                    '
                                Case formIdShowcaseAd
                                    '
                                    '
                                    '
                                Case FormTextSearchResults
                                    '
                                    '
                                    '
                                Case FormTextSearchResultsMore
                                    '
                                    '
                                    '
                                Case FormSearchMapResults
                                    '
                                    '
                                    '
                                Case Else
                                    '
                                    ' Home Form - Nothing to process, stay on Home Form
                                    '
                                    formId = formIdHome
                            End Select
                        End If
                    End If
                    '
                    ' --------------------------------------------------------------------
                    ' Generate next form
                    ' --------------------------------------------------------------------
                    '
                    CP.Doc.AddRefreshQueryString(rnformId, formId)
                    Select Case formId
                        Case formIdShowcaseEdit
                            '
                            ' showcase Edit
                            '
                            If Not common.isAccountContact(CP, orgId) Then
                                content = common.getPermissionError()
                            Else
                                showcaseEdit = New ShowcaseEditClass
                                content = showcaseEdit.getForm(CP, common, orgId)
                                AllowBannerAds = False
                            End If
                        Case formIdBannerEdit
                            '
                            ' Banner Edit
                            '
                            If Not common.isAccountContact(CP, orgId) Then
                                content = common.getPermissionError()
                            Else
                                bannerEdit = New BannerEditClass
                                content = bannerEdit.getForm(CP, common, orgId)
                                AllowBannerAds = False
                            End If
                        Case formIdMyAccountEdit
                            '
                            ' My Account Edit
                            '
                            If Not common.isAccountContact(CP, orgId) Then
                                content = common.getPermissionError()
                            Else
                                MyAccountEdit = New MyAccountEditClass
                                content = MyAccountEdit.getForm(CP, common, orgId)
                                AllowBannerAds = False
                            End If
                        Case formIdMyAccount
                            '
                            ' My Account
                            '
                            If Not common.isAccountContact(CP, orgId) Then
                                content = common.getPermissionError()
                            Else
                                MyAccount = New MyAccountClass
                                content = MyAccount.getForm(CP, common, orgId)
                                AllowBannerAds = False
                            End If
                        Case formIdCategoryHeading
                            '
                            ' Heading Category List
                            '
                            categoryHeading = New CategoryHeadingClass
                            content = categoryHeading.getForm(CP, common)
                            AllowBannerAds = True
                        Case formIdHeadingSearchREsults
                            '
                            ' Heading Search Results
                            '
                            headingSearchResults = New HeadingSearchResultsClass
                            content = headingSearchResults.getForm(CP, common, mapMode)
                            AllowBannerAds = True
                        Case formIdProfile
                            '
                            ' Profile
                            '
                            profile = New ProfileClass
                            content = profile.getForm(CP, common)
                            AllowBannerAds = True
                        Case formIdShowcaseAd
                            '
                            ' Showcase Ad
                            '
                            showcaseAd = New ShowcaseAdClass
                            content = showcaseAd.getForm(CP, common)
                            AllowBannerAds = True
                        Case FormTextSearchResults
                            '
                            ' text Search Results
                            '
                            textSearchResults = New TextSearchResultsClass
                            content = textSearchResults.getForm(CP, common, searchText, searchPhrase, searchLocation, mapMode)
                            AllowBannerAds = True
                        Case FormTextSearchResultsMore
                            '
                            '
                            '
                        Case FormSearchMapResults
                            '
                            '
                            '
                    End Select
                End If
                '
                ' --------------------------------------------------------------------
                ' If blank, do the home page
                ' --------------------------------------------------------------------
                '
                If content = "" Then
                    '
                    ' If no content returned, go to home page
                    '

                    home = New HomeClass
                    content = home.getForm(CP, common)
                    AllowBannerAds = True
                End If
                '
                ' --------------------------------------------------------------------
                ' Banner Wrapper
                ' --------------------------------------------------------------------
                '
                If AllowBannerAds Then
                    bannerLinkPrefix = CP.Request.Protocol & CP.Site.Domain & CP.Site.FilePath
                    bannerWrapper = common.getLayout(CP, "Supplier Directory Banner Wrapper")
                    bannerWrapper = bannerWrapper.Replace("##content##", content)
                    '
                    ' Horizontal Banners
                    '
                    Randomize()
                    bannerTop = ""
                    bannerBottom = ""
                    bannerList = getHorizontalBannerList(CP, common)
                    If bannerList <> "" Then
                        banners = bannerList.Split(vbCrLf)
                        bannerCnt = banners.Length
                        If bannerCnt > 0 Then
                            '
                            ' Banner Top
                            '
                            bannerTopPtr = Int(bannerCnt * Rnd())
                            If bannerTopPtr >= bannerCnt Then
                                bannerTopPtr = 0
                            End If
                            bannerRow = banners(bannerTopPtr).ToString
                            bannerAttr = bannerRow.Split(vbTab)
                            bannerAlt = ""
                            bannerLink = ""
                            If bannerAttr.Length > 0 Then
                                bannerImageFilename = bannerAttr(0)
                                If bannerAttr.Length > 1 Then
                                    bannerLink = bannerAttr(1)
                                    If bannerAttr.Length > 2 Then
                                        bannerAlt = bannerAttr(2)
                                    End If
                                End If
                                bannerTop = "" _
                                    & vbCrLf & vbTab & "<a href=""" & bannerLink & """ target=""_blank"">" _
                                    & vbCrLf & vbTab & vbTab & "<img src=""" & bannerLinkPrefix & bannerImageFilename & """ alt=""" & bannerAlt & """ border=""0"">" _
                                    & vbCrLf & vbTab & "</a>"
                            End If
                            '
                            ' Banner Bottom
                            '
                            bannerBottomPtr = Int(bannerCnt * Rnd())
                            If bannerBottomPtr >= bannerCnt Then
                                bannerBottomPtr = 0
                            End If
                            If (bannerCnt > 1) And (bannerTopPtr = bannerBottomPtr) Then
                                bannerBottomPtr += 1
                                If bannerBottomPtr >= bannerCnt Then
                                    bannerBottomPtr = 0
                                End If
                            End If
                            bannerRow = banners(bannerBottomPtr).ToString
                            bannerAttr = bannerRow.Split(vbTab)
                            bannerAlt = ""
                            bannerLink = ""
                            If bannerAttr.Length > 0 Then
                                bannerImageFilename = bannerAttr(0)
                                If bannerAttr.Length > 1 Then
                                    bannerLink = bannerAttr(1)
                                    If bannerAttr.Length > 2 Then
                                        bannerAlt = bannerAttr(2)
                                    End If
                                End If
                                bannerBottom = "" _
                                    & vbCrLf & vbTab & "<a href=""" & bannerLink & """ target=""_blank"">" _
                                    & vbCrLf & vbTab & vbTab & "<img src=""" & bannerLinkPrefix & bannerImageFilename & """ alt=""" & bannerAlt & """ border=""0"">" _
                                    & vbCrLf & vbTab & "</a>"
                            End If
                        End If
                    End If
                    bannerWrapper = bannerWrapper.Replace("##bannerTop##", bannerTop)
                    bannerWrapper = bannerWrapper.Replace("##bannerBottom##", bannerBottom)
                    '
                    ' Vertical Banners
                    '
                    bannerTop = ""
                    bannerBottom = ""
                    bannerList = getVerticalBannerList(CP, common)
                    If bannerList <> "" Then
                        banners = bannerList.Split(vbCrLf)
                        bannerCnt = banners.Length
                        If bannerCnt > 0 Then
                            '
                            ' Banner Left
                            '
                            bannerLeftPtr = Int(bannerCnt * Rnd())
                            If bannerLeftPtr >= bannerCnt Then
                                bannerLeftPtr = 0
                            End If
                            bannerRow = banners(bannerLeftPtr).ToString
                            bannerAttr = bannerRow.Split(vbTab)
                            bannerAlt = ""
                            bannerLink = ""
                            If bannerAttr.Length > 0 Then
                                bannerImageFilename = bannerAttr(0)
                                If bannerAttr.Length > 1 Then
                                    bannerLink = bannerAttr(1)
                                    If bannerAttr.Length > 2 Then
                                        bannerAlt = bannerAttr(2)
                                    End If
                                End If
                                bannerLeft = "" _
                                    & vbCrLf & vbTab & vbTab & "<a href=""" & bannerLink & """ target=""_blank"">" _
                                    & vbCrLf & vbTab & vbTab & vbTab & "<img src=""" & bannerLinkPrefix & bannerImageFilename & """ alt=""" & bannerAlt & """ border=""0"">" _
                                    & vbCrLf & vbTab & vbTab & "</a>"
                            End If
                            '
                            ' Banner Right
                            '
                            bannerRightPtr = Int(bannerCnt * Rnd())
                            If bannerRightPtr >= bannerCnt Then
                                bannerRightPtr = 0
                            End If
                            If (bannerCnt > 1) And (bannerLeftPtr = bannerRightPtr) Then
                                bannerRightPtr += 1
                                If bannerRightPtr >= bannerCnt Then
                                    bannerRightPtr = 0
                                End If
                            End If
                            bannerRow = banners(bannerRightPtr).ToString
                            bannerAttr = bannerRow.Split(vbTab)
                            bannerAlt = ""
                            bannerLink = ""
                            If bannerAttr.Length > 0 Then
                                bannerImageFilename = bannerAttr(0)
                                If bannerAttr.Length > 1 Then
                                    bannerLink = bannerAttr(1)
                                    If bannerAttr.Length > 2 Then
                                        bannerAlt = bannerAttr(2)
                                    End If
                                End If
                                bannerRight = "" _
                                    & vbCrLf & vbTab & vbTab & "<a href=""" & bannerLink & """ target=""_blank"">" _
                                    & vbCrLf & vbTab & vbTab & vbTab & "<img src=""" & bannerLinkPrefix & bannerImageFilename & """ alt=""" & bannerAlt & """ border=""0"">" _
                                    & vbCrLf & vbTab & vbTab & "</a>"
                            End If
                        End If
                    End If
                    bannerWrapper = bannerWrapper.Replace("##bannerLeft##", bannerLeft)
                    bannerWrapper = bannerWrapper.Replace("##bannerRight##", bannerRight)
                    '
                    ' Search form
                    '
                    qs = rqs
                    qs = CP.Utils.ModifyQueryString(qs, rnformId, FormTextSearchResults)
                    bannerWrapper = bannerWrapper.Replace("##stateSelect##", common.getStateSelect(CP, "searchLocation", "", "searchLocation"))
                    bannerWrapper = bannerWrapper.Replace("##action##", "?" & qs)
                    bannerWrapper = bannerWrapper.Replace("##searchText##", searchText)
                    bannerWrapper = bannerWrapper.Replace("##searchLocation##", searchLocation)
                    If searchPhrase Then
                        bannerWrapper = bannerWrapper.Replace("##searchPhraseChecked##", "checked")
                    Else
                        bannerWrapper = bannerWrapper.Replace("##searchPhraseChecked##", "")
                    End If
                    bannerWrapper = bannerWrapper.Replace("##bannerLogo##", CP.Content.GetCopy("Supplier Directory Banner Logo", "<div style=""text-align:center"">supplier directory logo</div>"))
                    If searchAdvanced Then
                        bannerWrapper &= vbCrLf & vbTab & "<script>cj.hide('searchNormal');cj.show('searchAdvanced');</script>"
                    End If
                    content = bannerWrapper
                End If
                '
                ' --------------------------------------------------------------------
                ' home return link Navigation
                ' --------------------------------------------------------------------
                '
                qs = rqs
                qs = CP.Utils.ModifyQueryString(qs, rnformId, formIdHome)
                content = content.Replace("##homeLink##", "?" & qs)
                '
                ' --------------------------------------------------------------------
                ' footer Navigation
                ' --------------------------------------------------------------------
                '
                If CP.User.IsAuthenticated() Then
                    cs = CP.CSNew
                    If cs.Open("organizations", "directoryaccountcontactid=" & CP.Db.EncodeSQLNumber(CP.User.Id)) Then
                        copy = ""
                        Do While cs.OK()
                            qs = rqs
                            qs = CP.Utils.ModifyQueryString(qs, rnformId, formIdMyAccount)
                            qs = CP.Utils.ModifyQueryString(qs, rnOrganizationID, cs.GetInteger("id"))
                            copy &= vbCrLf & "<span class=""footerNavItem""><a href=""?" & qs & """>Manage " & cs.GetText("name") & "</a></span>"
                            Call cs.GoNext()
                        Loop
                        qs = rqs
                        qs = CP.Utils.ModifyQueryString(qs, "method", "logout")
                        copy &= "|<span class=""footerNavItem"">You are logged in as " & CP.User.Name & ". <a href=""?" & qs & """>Click here to log out</a>.</span>"
                    End If
                    Call cs.Close()
                End If
                If copy = "" Then
                    js &= "jQuery('#footerNav').hide();"
                Else
                    content = content.Replace("##footerNavContent##", copy)
                End If
                '
                ' --------------------------------------------------------------------
                ' Done
                ' --------------------------------------------------------------------
                '
                If Not String.IsNullOrEmpty(js) Then
                    CP.Doc.AddHeadJavascript(js)
                End If
                Execute = content
            Catch ex As Exception
                Call CP.Site.ErrorReport("Execute ex.message=" & ex.Message)
            End Try
            '
            ' done
            '
        End Function
        '
        '============================================================================
        '   get Vertical Banner List
        '   crlf delimited list of all banners, image tab link tab title
        '============================================================================
        '
        Function getVerticalBannerList(ByVal cp As CPBaseClass, common As CommonClass) As String
            '
            getVerticalBannerList = ""
            Try
                getVerticalBannerList = common.cacheRead(cp, "VerticalBannerList")
                If getVerticalBannerList = "" Then
                    Call setBannerLists(cp)
                    getVerticalBannerList = Global_VerticalBannerList
                End If
            Catch ex As Exception
                Call cp.Site.ErrorReport("getVerticalBannerList, " & ex.Message)
            End Try
            '
        End Function
        '
        '============================================================================
        '   get Horizontal Banner List
        '   crlf delimited list of all banners, image tab link tab title
        '============================================================================
        '
        Function getHorizontalBannerList(ByVal cp As CPBaseClass, common As CommonClass) As String
            '
            getHorizontalBannerList = ""
            Try
                getHorizontalBannerList = Common.cacheRead(cp, "HorizontalBannerList")
                If getHorizontalBannerList = "" Then
                    Call setBannerLists(cp)
                    getHorizontalBannerList = Global_HorizontalBannerList
                End If
            Catch ex As Exception
                Call cp.Site.ErrorReport("getHorizontalBannerList, " & ex.Message)
            End Try
            '
        End Function
        '
        '============================================================================
        '   setup Banner Lists
        '============================================================================
        '
        Sub setBannerLists(ByVal cp As CPBaseClass)
            Try
                Dim Layout As String = ""
                Dim filename As String = ""
                Dim cs As BaseClasses.CPCSBaseClass
                Dim title As String = ""
                Dim imageFilename As String = ""
                Dim link As String = ""
                Dim sqlCriteria As String = ""
                Dim nowDate As Date = Now.Date
                '
                Global_HorizontalBannerList = ""
                Global_VerticalBannerList = ""
                cs = cp.CSNew
                sqlCriteria = "" _
                    & " (approved<>0)" _
                    & " and(ApprovedByAccount<>0)" _
                    & " and((DateExpires is null)or(DateExpires>" & cp.Db.EncodeSQLDate(nowDate) & "))" _
                    & ""
                cs.Open("Directory Banner Ads", sqlCriteria)
                Do While cs.OK
                    imageFilename = cs.GetText("imagefilename")
                    link = cs.GetText("link").Replace(vbCrLf, " ").Replace(vbTab, " ")
                    title = cs.GetText("title").Replace(vbCrLf, " ").Replace(vbTab, " ")
                    If (imageFilename <> "") And (link <> "") Then
                        If cs.GetInteger("typeid") = AdTypeIDHorizonatal Then
                            Global_HorizontalBannerList = Global_HorizontalBannerList _
                                & vbCrLf & imageFilename _
                                & vbTab & link _
                                & vbTab & title _
                                & ""
                        Else
                            Global_VerticalBannerList = Global_VerticalBannerList _
                                & vbCrLf & imageFilename _
                                & vbTab & link _
                                & vbTab & title _
                                & ""
                        End If
                    End If
                    Call cs.GoNext()
                Loop
                Call cs.Close()
                If Global_VerticalBannerList <> "" Then
                    Global_VerticalBannerList = Global_VerticalBannerList.Substring(2)
                End If
                If Global_HorizontalBannerList <> "" Then
                    Global_HorizontalBannerList = Global_HorizontalBannerList.Substring(2)
                End If
                '
                Call cp.Cache.Save("VerticalBannerList", Global_VerticalBannerList, "Directory Banner Ads")
                Call cp.Cache.Save("HorizontalBannerList", Global_HorizontalBannerList, "Directory Banner Ads")
            Catch ex As Exception
                Call cp.Site.ErrorReport("setBannerLists, " & ex.Message)
            End Try
        End Sub
        '
        '
        '
    End Class
End Namespace

