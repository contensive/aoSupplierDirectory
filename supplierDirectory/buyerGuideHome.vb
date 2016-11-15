
Imports Contensive.BaseClasses
Imports Contensive.Addons.SupplierDirectory.Constants

Namespace Contensive.Addons.SupplierDirectory
    Friend Class buyerGuideHomeClass
        '
        '============================================================================
        '   get Home Form
        '============================================================================
        '
        Function getForm(ByVal cp As CPBaseClass, ByVal common As CommonClass) As String
            Dim returnHtml As String = ""
            Try
                '
                Dim cs As CPCSBaseClass
                Dim category As String = ""
                Dim categoryLast As String = ""
                Dim heading As String = ""
                Dim rqs As String = ""
                Dim qs As String = ""
                Dim cellLeft As String = ""
                Dim cellRight As String = ""
                Dim cellRowMax As Integer = 0
                Dim cellRowPtr As Integer = 0
                Dim Hint As String = ""
                Dim rowTotal As Integer = 0
                Dim cacheName As String = ""
                Dim form As CPBlockBaseClass = cp.BlockNew()
                Dim tlpaTierSection As CPBlockBaseClass = cp.BlockNew
                Dim buyersGuideSection As CPBlockBaseClass = cp.BlockNew
                Dim blockLayout As CPBlockBaseClass = cp.BlockNew
                Dim exitLoop As Boolean = False
                Dim buyerGuideBoxHtml As String
                Dim rowPtr As Integer = 1
                Dim unUsedCatagory As String = ""
                Dim name As String = ""
                Dim categoryImage As String = ""
                Dim tmpHtml As String = ""
                Dim dirSQL As String = ""
                '
                ' read in the cache
                '
                cacheName = "Supplier Directory Home"
                returnHtml = common.cacheRead(cp, cacheName)
                If returnHtml = "" Then
                    '
                    ' cache is empty - recreate page
                    '
                    rqs = cp.Doc.RefreshQueryString
                    cs = cp.CSNew
                    'dirSQL = "select dc.Name as category, dc.categoryImage, ds.name,ds.id from directorySubcategories ds inner Join directoryCategories dc " _
                    '                 & "on ds.categoryID = dc.id where dc.name is not null and ds.active=1 order by dc.name,ds.name, ds.id"
                    ''cs.Open("Directory Subcategories", "(category is not null)and(name is not null)", "Category,Name", , "Category,Name,ID")
                    'cs.OpenSQL("", dirSQL)
                    'categoryImage = cs.GetText("categoryImage")
                    'If Not cs.OK Then
                    '    cellLeft = "<div>There are no Directory Subcategories to display.</div>"
                    '    cellLeft &= cp.Content.GetAddLink("Directory Subcategories", "", False, True)
                    '    cellRight = ""
                    'Else
                    '    '
                    '    ' count the total display rows needed so they can be divided up right and left
                    '    '
                    '    categoryLast = "start"
                    '    Do While cs.OK
                    '        category = cs.GetText("category")
                    '        If category <> categoryLast Then
                    '            rowTotal += 1
                    '        End If
                    '        rowTotal += 1
                    '        categoryLast = category
                    '        cs.GoNext()
                    '    Loop
                    '    '
                    '    ' left column
                    '    '
                    '    Call cs.GoFirst()
                    '    cellRowMax = rowTotal / 2
                    '    categoryLast = "start"
                    '    category = cs.GetText("category")
                    '    Do While cs.OK And ((cellRowPtr < cellRowMax) Or (category = categoryLast))
                    '        If category <> categoryLast Then
                    '            If categoryLast <> "start" Then
                    '                cellLeft += "" _
                    '                    & vbCrLf & vbTab & vbTab & vbTab & "</ul>" _
                    '                    & vbCrLf & vbTab & vbTab & "</li>" _
                    '                    & ""
                    '            End If
                    '            qs = rqs
                    '            qs = cp.Utils.ModifyQueryString(qs, rnCategory, category)
                    '            qs = cp.Utils.ModifyQueryString(qs, rnformId, formIdCategoryHeading)
                    '            cellLeft += "" _
                    '                & vbCrLf & vbTab & vbTab & "<li><a href=""?" & qs & """>" & category & "</a>" _
                    '                & vbCrLf & vbTab & vbTab & vbTab & "<ul class=""headings"">" _
                    '                & ""
                    '        End If
                    '        qs = rqs
                    '        qs = cp.Utils.ModifyQueryString(qs, rnHeadingID, cs.GetText("id"))
                    '        qs = cp.Utils.ModifyQueryString(qs, rnformId, formIdHeadingSearchREsults)
                    '        cellLeft += vbCrLf & vbTab & vbTab & vbTab & "<li><a href=""?" & qs & """>" & cs.GetText("name") & "</a></li>"
                    '        cellRowPtr += 1
                    '        cs.GoNext()
                    '        categoryLast = category
                    '        If cs.OK() Then
                    '            category = cs.GetText("category")
                    '        End If
                    '    Loop
                    '    If cellLeft <> "" Then
                    '        cellLeft = "" _
                    '            & vbCrLf & vbTab & "<ul class=""categories"">" _
                    '            & cellLeft _
                    '            & vbCrLf & vbTab & vbTab & vbTab & "</ul>" _
                    '            & vbCrLf & vbTab & vbTab & "</li>" _
                    '            & vbCrLf & vbTab & "</ul>" _
                    '            & ""
                    '    End If
                    'End If
                    'If cs.OK Then
                    '    '
                    '    ' right column
                    '    '
                    '    categoryLast = "start"
                    '    Do While cs.OK
                    '        qs = cp.Utils.ModifyQueryString(rqs, rnHeadingID, cs.GetText("id"))
                    '        If category <> categoryLast Then
                    '            If categoryLast <> "start" Then
                    '                cellRight += "" _
                    '                    & vbCrLf & vbTab & vbTab & vbTab & "</ul>" _
                    '                    & vbCrLf & vbTab & vbTab & "</li>" _
                    '                    & ""
                    '            End If
                    '            qs = rqs
                    '            qs = cp.Utils.ModifyQueryString(qs, rnCategory, category)
                    '            qs = cp.Utils.ModifyQueryString(qs, rnformId, formIdCategoryHeading)
                    '            cellRight += "" _
                    '                & vbCrLf & vbTab & vbTab & "<li><a href=""?" & qs & """>" & category & "</a>" _
                    '                & vbCrLf & vbTab & vbTab & vbTab & "<ul class=""headings"">" _
                    '                & ""
                    '            cellRowPtr -= 1
                    '        End If
                    '        qs = rqs
                    '        qs = cp.Utils.ModifyQueryString(qs, rnHeadingID, cs.GetText("id"))
                    '        qs = cp.Utils.ModifyQueryString(qs, rnformId, formIdHeadingSearchREsults)
                    '        cellRight += vbCrLf & vbTab & vbTab & vbTab & "<li><a href=""?" & qs & """>" & cs.GetText("name") & "</a></li>"
                    '        cellRowPtr += 1
                    '        cs.GoNext()
                    '        categoryLast = category
                    '        If cs.OK() Then
                    '            category = cs.GetText("category")
                    '        End If
                    '    Loop
                    '    If cellRight <> "" Then
                    '        cellRight = "" _
                    '            & vbCrLf & vbTab & "<ul class=""categoriesUL"">" _
                    '            & cellRight _
                    '            & vbCrLf & vbTab & vbTab & vbTab & "</ul>" _
                    '            & vbCrLf & vbTab & vbTab & "</li>" _
                    '            & vbCrLf & vbTab & "</ul>" _
                    '            & ""
                    '    End If
                    'End If
                    'Call cs.Close()
                    '
                    ' add wrapper
                    '
                    tlpaTierSection.OpenLayout("bgHome.html")
                    Dim tierContents As String = tlpaTierSection.GetOuter(".tlpaTierSection")
                    tlpaTierSection.Load(tierContents)
                    Dim buyersGuideBox As CPBlockBaseClass = cp.BlockNew()
                    Dim buyersGuildSponserHbox As CPBlockBaseClass = cp.BlockNew()
                    buyersGuideBox.Load(tlpaTierSection.GetOuter(".buyerGuideBox"))
                    buyersGuildSponserHbox.Load(buyersGuideSection.GetOuter(".sponsHBox"))
                    '
                    Dim mySql As String = ""
                    exitLoop = False
                    mySql = "select distinct ds.categoryId, dc.name from directorySubcategories ds inner join directoryCategories dc " _
                           & "on ds.categoryID = dc.id order by dc.name "
                    cs.OpenSQL("", mySql)
                    categoryLast = "start"
                    '
                    ' building li list of categories / subcategories - create one li for each category, and populate the inner ul with all the subcategories
                    '
                    Dim categoryULInner As String = buyersGuideBox.GetInner(".categoryUL")
                    Dim tmpBlock As CPBlockBaseClass = cp.BlockNew()
                    tmpBlock.Load(vbCrLf & "<div>xxx</div>" & categoryULInner)
                    ''
                    ''
                    'Dim imgBlock As CPBlockBaseClass = cp.BlockNew()
                    'imgBlock.Load(vbCrLf & "<div>xxx</div>" & categoryULInner)
                    'Dim imageULliTemplate As String = imgBlock.GetOuter(".tlpaTierSection")
                    'Dim imageULli As String = imageULliTemplate

                    tlpaTierSection.SetOuter(".s728", "<a href=""##link##"">##caption##</a>")
                    tlpaTierSection.SetOuter(".sponsHBoxInner", "<div>##showcaseList##</div>")
                    '
                    Dim categoryULliTemplate As String = tmpBlock.GetOuter(".categoryULItem")


                    Dim categoryULliList As String = ""
                    Dim csCat As CPCSBaseClass = cp.CSNew()
                    Dim catPtr As Integer = 0
                    Dim sql As String = ""
                    sql = "select distinct ds.categoryId,dc.name,dc.categoryImage " _
                                     & "from directorySubcategories ds inner Join directoryCategories dc " _
                                     & "on ds.categoryID = dc.id where (ds.active=1) order by dc.name"
                    If csCat.OpenSQL("", sql) Then
                        Do
                            '
                            ' build the subcategory menu for each catagorey
                            '
                            Dim categoryULli As String = categoryULliTemplate

                            Dim categoryULliBlock As CPBlockBaseClass = cp.BlockNew()
                            categoryULliBlock.Load(categoryULli)

                            '
                            category = csCat.GetText("name")
                            categoryImage = csCat.GetText("categoryImage")
                            '
                            '
                            categoryULliBlock.SetInner(".categoryName", category)
                            categoryULliBlock.SetOuter(".imageName", "<img class=""imageName"" src=""" & cp.Site.FilePath & categoryImage & """ alt=""Fleet Communication Systems and Equipment"" />")
                            ' categoryULliBlock.SetOuter(".sponsHBoxInner", "<div>##showcaseList##</div>")
                            ' categoryULliBlock.SetOuter(".sponsHBoxInner", "")
                            's728
                            '
                            Dim csSub As CPCSBaseClass = cp.CSNew()
                            Dim subList As String = ""
                            Dim subItem As String = ""
                            Dim subItemTEmplate As String = ""
                            Dim subSql As String = ""
                            Dim categoryId As Integer

                            subSql = "select distinct ds.name,ds.categoryID,dc.name as categoryName,ds.id as subcategoryId " _
                                     & "from directorySubcategories ds inner Join directoryCategories dc " _
                                     & "on ds.categoryID = dc.id where dc.name=" & cp.Db.EncodeSQLText(category)

                            If csSub.OpenSQL("", subSql) Then
                                Do
                                    '
                                    categoryId = csSub.GetInteger("categoryID")
                                    ' build li list of subcategories for this category
                                    '
                                    qs = rqs
                                    qs = cp.Utils.ModifyQueryString(qs, rnSubcategoryID, csSub.GetInteger("subcategoryId").ToString())
                                    qs = cp.Utils.ModifyQueryString(qs, rnformId, formIdHeadingSearchREsults)

                                    subList &= "<li><a href=""?" & qs & """>" & csSub.GetText("name") & "</a></li>"
                                    csSub.GoNext()
                                Loop While csSub.OK
                            End If
                            csSub.Close()
                            categoryULliBlock.SetInner(".categoryUL2", subList)
                            categoryULliList &= categoryULliBlock.GetHtml().Replace("show-cat-01", "show-cat-" & catPtr.ToString)
                            '
                            Call csSub.Close()
                            Call csCat.GoNext()
                            catPtr += 1
                        Loop While csCat.OK

                    End If
                    Call csCat.Close()
                    Call cs.Close()

                    Dim showcaseULliTemplate As String = tmpBlock.GetOuter(".sponsHBoxInner")
                    Dim orgid As Integer
                    Dim ImageFilename As String = ""
                    Dim showcaseULli As String = showcaseULliTemplate

                    Dim showcaseULliBlock As CPBlockBaseClass = cp.BlockNew()
                    '
                    'If cs.Open("Directory Showcase Ads") Then
                    '    Do
                    '        orgid = cs.GetInteger("OrganizationID")
                    '        ImageFilename = cs.GetText("ImageFilename")
                    '        showcaseULliBlock.SetInner(".s728", cp.Site.FilePath(ImageFilename))
                    '        cs.GoNext()
                    '    Loop While cs.OK

                    'End If

                    ''
                    'Call cs.Close()
                    buyersGuideBox.SetInner(".categoryUL", categoryULliList)
                    tlpaTierSection.SetInner(".tierFull", buyersGuideBox.GetHtml())
                    'returnHtml = returnHtml.Replace("##showcaseList##", getShowcaseAds(cp, common, rqs))
                    ' 
                    returnHtml = tlpaTierSection.GetHtml()
                    ''
                    Call cp.Cache.Save(cacheName, returnHtml, "Layouts,Directory Subcategories")
                End If
            Catch ex As Exception
                Call cp.Site.ErrorReport("HomeClass.getForm, " & ex.Message)
            End Try
            Return returnHtml
        End Function
        '
        '
        '
        Function getShowcaseAd(ByVal cp As CPBaseClass, ByVal Layout As String, ByVal OrganizationID As String, ByVal ImageSrc As String, ByVal Caption As String, ByVal Copy As String, ByVal Link As String) As String
            '
            '
            getShowcaseAd = ""
            Try
                getShowcaseAd = Layout
                getShowcaseAd = getShowcaseAd.Replace("##organizationID##", OrganizationID)
                getShowcaseAd = getShowcaseAd.Replace("##imageSrc##", ImageSrc)
                getShowcaseAd = getShowcaseAd.Replace("##caption##", Caption)
                getShowcaseAd = getShowcaseAd.Replace("##copy##", Copy)
                getShowcaseAd = getShowcaseAd.Replace("##link##", Link)
            Catch ex As Exception
                Call cp.Site.ErrorReport("getShowcaseAd, " & ex.Message)
            End Try
            '
        End Function
        '
        '============================================================================
        '   get the showcase banner list
        '============================================================================
        '
        Function getShowcaseAds(ByVal cp As CPBaseClass, ByVal common As CommonClass, ByVal rqs As String) As String
            '
            Dim cs As BaseClasses.CPCSBaseClass
            Dim Caption As String = ""
            Dim Copy As String = ""
            Dim Link As String = ""
            Dim ShowcaseAdList As String = ""
            Dim ptr As Integer
            Dim cnt As Integer
            Dim current As Integer
            Dim ShowcaseAds() As String
            Dim adElements() As String
            Dim Layout As String
            Dim ImageSrcPrefix As String
            Dim cacheName As String = ""
            Dim qs As String = ""
            '
            cp.Site.TestPoint("getShowcaseAds")
            getShowcaseAds = ""
            Try
                cacheName = "ShowcaseAdList"
                ShowcaseAdList = common.cacheRead(cp, cacheName)
                If ShowcaseAdList = "" Then
                    cs = cp.CSNew
                    cs.Open("Directory Showcase Ads", "(approved<>0)and(ApprovedByAccount<>0)")
                    Do While cs.OK
                        Caption = cs.GetText("caption").Replace(vbCrLf, " ").Replace(vbTab, " ")
                        Copy = cs.GetText("copy").Replace(vbCrLf, "<br>").Replace(vbTab, " ")
                        qs = cp.Utils.ModifyQueryString(rqs, rnShowcaseAdID, cs.GetInteger("ID").ToString)
                        qs = cp.Utils.ModifyQueryString(qs, rnformId, formIdShowcaseAd)
                        Link = "?" & qs
                        ShowcaseAdList += "" _
                            & vbCrLf & cs.GetInteger("organizationid") _
                            & vbTab & cs.GetText("imagefilename") _
                            & vbTab & Caption _
                            & vbTab & Copy _
                            & vbTab & Link _
                            & ""
                        cs.GoNext()
                    Loop
                    Call cs.Close()
                    If ShowcaseAdList <> "" Then
                        ShowcaseAdList.Substring(2)
                    End If
                    Call cp.Cache.Save(cacheName, ShowcaseAdList, "Directory Showcase Ads, Organizations")
                End If
                If ShowcaseAdList <> "" Then
                    'cp.Site.TestPoint("getShowcaseAds, showcaseAdList=" & ShowcaseAdList.Replace(vbCrLf, "<br>"))
                    ImageSrcPrefix = cp.Request.Protocol & cp.Site.Domain & cp.Site.FilePath
                    ShowcaseAds = Split(ShowcaseAdList, vbCrLf)
                    cnt = ShowcaseAds.Length
                    current = cp.Utils.EncodeInteger(cp.Site.GetProperty("ShowcaseAdCurrent", "0")) + 1
                    If current >= cnt Then
                        current = 0
                    End If
                    Call cp.Site.SetProperty("ShowcaseAdCurrent", CStr(current))
                    Layout = common.getLayout(cp, "Supplier Directory showcaseAd")
                    If current < cnt Then
                        For ptr = current To cnt - 1
                            cp.Site.TestPoint("ptr=" & ptr)
                            adElements = Split(ShowcaseAds(ptr), vbTab)
                            If adElements.Length > 4 Then
                                getShowcaseAds += getShowcaseAd(cp, Layout, adElements(0), ImageSrcPrefix & adElements(1), adElements(2), adElements(3), adElements(4))
                            End If
                        Next
                    End If
                    If current > 0 Then
                        For ptr = 0 To current - 1
                            cp.Site.TestPoint("ptr=" & ptr)
                            adElements = Split(ShowcaseAds(ptr), vbTab)
                            If adElements.Length > 4 Then
                                getShowcaseAds += getShowcaseAd(cp, Layout, adElements(0), ImageSrcPrefix & adElements(1), adElements(2), adElements(3), adElements(4))
                            End If
                        Next
                    End If
                End If
            Catch ex As Exception
                Call cp.Site.ErrorReport("getShowcaseAds, " & ex.Message)
            End Try
            '
        End Function

    End Class
End Namespace