
Imports Contensive.BaseClasses
Imports Contensive.Addons.SupplierDirectory.Constants

Namespace Contensive.Addons.SupplierDirectory
    Friend Class HomeClass
        '
        '============================================================================
        '   get Home Form
        '============================================================================
        '
        Function getForm(ByVal cp As CPBaseClass, ByVal common As CommonClass) As String
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
            '
            getForm = ""
            Try
                cacheName = "Supplier Directory Home"
                getForm = cp.Cache.Read(cacheName)
                If getForm = "" Then
                    rqs = cp.Doc.RefreshQueryString
                    cs = cp.CSNew
                    cs.Open("Directory Subcategories", "(category is not null)and(name is not null)", "Category,Name", , "Category,Name,ID")
                    If Not cs.OK Then
                        cellLeft = "<div>There are no Directory Subcategories to display.</div>"
                        cellLeft &= cp.Content.GetAddLink("Directory Subcategories", "", False, True)
                        cellRight = ""
                    Else
                        '
                        ' count the total display rows needed so they can be divided up right and left
                        '
                        categoryLast = "start"
                        Do While cs.OK
                            category = cs.GetText("category")
                            If category <> categoryLast Then
                                rowTotal += 1
                            End If
                            rowTotal += 1
                            categoryLast = category
                            cs.GoNext()
                        Loop
                        '
                        ' left column
                        '
                        Call cs.GoFirst()
                        cellRowMax = rowTotal / 2
                        categoryLast = "start"
                        category = cs.GetText("category")
                        Do While cs.OK And ((cellRowPtr < cellRowMax) Or (category = categoryLast))
                            If category <> categoryLast Then
                                If categoryLast <> "start" Then
                                    cellLeft += "" _
                                        & vbCrLf & vbTab & vbTab & vbTab & "</ul>" _
                                        & vbCrLf & vbTab & vbTab & "</li>" _
                                        & ""
                                End If
                                qs = rqs
                                qs = cp.Utils.ModifyQueryString(qs, rnCategory, category)
                                qs = cp.Utils.ModifyQueryString(qs, rnformId, formIdCategoryHeading)
                                cellLeft += "" _
                                    & vbCrLf & vbTab & vbTab & "<li><a href=""?" & qs & """>" & category & "</a>" _
                                    & vbCrLf & vbTab & vbTab & vbTab & "<ul class=""headings"">" _
                                    & ""
                            End If
                            qs = rqs
                            qs = cp.Utils.ModifyQueryString(qs, rnHeadingID, cs.GetText("id"))
                            qs = cp.Utils.ModifyQueryString(qs, rnformId, formIdHeadingSearchREsults)
                            cellLeft += vbCrLf & vbTab & vbTab & vbTab & "<li><a href=""?" & qs & """>" & cs.GetText("name") & "</a></li>"
                            cellRowPtr += 1
                            cs.GoNext()
                            categoryLast = category
                            If cs.OK() Then
                                category = cs.GetText("category")
                            End If
                        Loop
                        If cellLeft <> "" Then
                            cellLeft = "" _
                                & vbCrLf & vbTab & "<ul class=""categories"">" _
                                & cellLeft _
                                & vbCrLf & vbTab & vbTab & vbTab & "</ul>" _
                                & vbCrLf & vbTab & vbTab & "</li>" _
                                & vbCrLf & vbTab & "</ul>" _
                                & ""
                        End If
                    End If
                    If cs.OK Then
                        '
                        ' right column
                        '
                        categoryLast = "start"
                        Do While cs.OK
                            qs = cp.Utils.ModifyQueryString(rqs, rnHeadingID, cs.GetText("id"))
                            If category <> categoryLast Then
                                If categoryLast <> "start" Then
                                    cellRight += "" _
                                        & vbCrLf & vbTab & vbTab & vbTab & "</ul>" _
                                        & vbCrLf & vbTab & vbTab & "</li>" _
                                        & ""
                                End If
                                qs = rqs
                                qs = cp.Utils.ModifyQueryString(qs, rnCategory, category)
                                qs = cp.Utils.ModifyQueryString(qs, rnformId, formIdCategoryHeading)
                                cellRight += "" _
                                    & vbCrLf & vbTab & vbTab & "<li><a href=""?" & qs & """>" & category & "</a>" _
                                    & vbCrLf & vbTab & vbTab & vbTab & "<ul class=""headings"">" _
                                    & ""
                                cellRowPtr -= 1
                            End If
                            qs = rqs
                            qs = cp.Utils.ModifyQueryString(qs, rnHeadingID, cs.GetText("id"))
                            qs = cp.Utils.ModifyQueryString(qs, rnformId, formIdHeadingSearchREsults)
                            cellRight += vbCrLf & vbTab & vbTab & vbTab & "<li><a href=""?" & qs & """>" & cs.GetText("name") & "</a></li>"
                            cellRowPtr += 1
                            cs.GoNext()
                            categoryLast = category
                            If cs.OK() Then
                                category = cs.GetText("category")
                            End If
                        Loop
                        If cellRight <> "" Then
                            cellRight = "" _
                                & vbCrLf & vbTab & "<ul class=""categories"">" _
                                & cellRight _
                                & vbCrLf & vbTab & vbTab & vbTab & "</ul>" _
                                & vbCrLf & vbTab & vbTab & "</li>" _
                                & vbCrLf & vbTab & "</ul>" _
                                & ""
                        End If
                    End If
                    '
                    ' add wrapper
                    '
                    getForm = common.getLayout(cp, "Supplier Directory Home")
                    getForm = getForm.Replace("##leftColumn##", cellLeft)
                    getForm = getForm.Replace("##rightColumn##", cellRight)
                    getForm = getForm.Replace("##showcaseList##", getShowcaseAds(cp, common, rqs))
                    '
                    Call cp.Cache.Save(cacheName, getForm, "Layouts,Directory Subcategories")
                End If
            Catch ex As Exception
                Call cp.Site.ErrorReport("HomeClass.getForm, " & ex.Message)
            End Try
            '
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
                ShowcaseAdList = cp.Cache.Read(cacheName)
                If ShowcaseAdList = "" Then
                    cs = cp.CSNew
                    cs.Open("Directory Showcase Ads", "(approved<>0)and(ApprovedByAccount<>0)")
                    Do While cs.OK
                        Caption = cs.GetText("caption").Replace(vbCrLf, " ").Replace(vbTab, " ")
                        Copy = cs.GetText("copy").Replace(vbCrLf, "<br>").Replace(vbTab, " ")
                        qs = cp.Utils.ModifyQueryString(rqs, rnshowcaseAdID, cs.GetInteger("ID").ToString)
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