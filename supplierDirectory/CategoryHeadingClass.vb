
Imports Contensive.BaseClasses
Imports Contensive.Addons.SupplierDirectory.Constants

Namespace Contensive.Addons.SupplierDirectory
    Friend Class CategoryHeadingClass
        '
        '============================================================================
        '   get Home Form
        '============================================================================
        '
        Friend Function getForm(ByVal cp As CPBaseClass, ByVal common As CommonClass) As String
            '
            Dim cs As CPCSBaseClass
            Dim category As String = ""
            Dim categoryLast As String = ""
            Dim heading As String = ""
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
            '
            cp.Site.TestPoint("CategoryHeadingClass.getForm")
            '
            getForm = ""
            Try
                category = cp.Request.GetText(rnCategory)
                If category <> "" Then
                    cacheName = "Supplier Directory Subcategory List, " & category
                    getForm = common.cacheRead(cp, cacheName)
                    If getForm = "" Then
                        cs = cp.CSNew
                        cs.Open("Directory Subcategories", "(category=" & cp.Db.EncodeSQLText(category) & ")", "name")
                        If cs.OK Then
                            '
                            ' all is valid, build page
                            '
                            cp.Site.TestPoint("CategoryHeadingClass.getForm, build form for category [" & category & "]")
                            Call cp.Doc.AddRefreshQueryString(rnCategory, category)
                            rqs = cp.Doc.RefreshQueryString
                            '
                            ' count the total display rows needed so they can be divided up right and left
                            '
                            rowTotal = cs.GetRowCount
                            cellRowMax = rowTotal / 3
                            '
                            ' left column
                            '
                            cellRowPtr = 0
                            cellLeft = ""
                            Do While cs.OK And (cellRowPtr < cellRowMax)
                                qs = rqs
                                qs = cp.Utils.ModifyQueryString(qs, rnHeadingID, cs.GetText("id"))
                                qs = cp.Utils.ModifyQueryString(qs, rnformId, formIdHeadingSearchREsults)
                                cellLeft += vbCrLf & vbTab & vbTab & "<li><a href=""?" & qs & """>" & cs.GetText("name") & "</a></li>"
                                cellRowPtr += 1
                                cs.GoNext()
                            Loop
                            If cellLeft <> "" Then
                                cellLeft = "" _
                                    & vbCrLf & vbTab & "<ul class=""headings"">" _
                                    & cellLeft _
                                    & vbCrLf & vbTab & "</ul>" _
                                    & ""
                            End If
                            '
                            ' center column
                            '
                            cellRowPtr = 0
                            cellCenter = ""
                            Do While cs.OK And (cellRowPtr < cellRowMax)
                                qs = rqs
                                qs = cp.Utils.ModifyQueryString(qs, rnHeadingID, cs.GetText("id"))
                                qs = cp.Utils.ModifyQueryString(qs, rnformId, formIdHeadingSearchREsults)
                                cellCenter += vbCrLf & vbTab & vbTab & "<li><a href=""?" & qs & """>" & cs.GetText("name") & "</a></li>"
                                cellRowPtr += 1
                                cs.GoNext()
                            Loop
                            If cellCenter <> "" Then
                                cellCenter = "" _
                                    & vbCrLf & vbTab & "<ul class=""headings"">" _
                                    & cellCenter _
                                    & vbCrLf & vbTab & "</ul>" _
                                    & ""
                            End If
                            '
                            ' right column
                            '
                            cellRowPtr = 0
                            cellRight = ""
                            Do While cs.OK
                                qs = rqs
                                qs = cp.Utils.ModifyQueryString(qs, rnHeadingID, cs.GetText("id"))
                                qs = cp.Utils.ModifyQueryString(qs, rnformId, formIdHeadingSearchREsults)
                                cellRight += vbCrLf & vbTab & vbTab & "<li><a href=""?" & qs & """>" & cs.GetText("name") & "</a></li>"
                                cellRowPtr += 1
                                cs.GoNext()
                            Loop
                            If cellRight <> "" Then
                                cellRight = "" _
                                    & vbCrLf & vbTab & "<ul class=""headings"">" _
                                    & cellRight _
                                    & vbCrLf & vbTab & "</ul>" _
                                    & ""
                            End If
                            '
                            ' add wrapper
                            '
                            getForm = common.getLayout(cp, "Supplier Directory Subcategory List")
                            qs = rqs
                            qs = cp.Utils.ModifyQueryString(qs, "searchButton", "", False)
                            qs = cp.Utils.ModifyQueryString(qs, "searchText", "", False)
                            qs = cp.Utils.ModifyQueryString(qs, "searchPhrase", "", False)
                            qs = cp.Utils.ModifyQueryString(qs, "searchLocation", "", False)
                            qs = cp.Utils.ModifyQueryString(qs, rnHeadingID, "", False)
                            qs = cp.Utils.ModifyQueryString(qs, rnCategory, "", False)
                            qs = cp.Utils.ModifyQueryString(qs, rnformId, formIdHome)
                            getForm = getForm.Replace("##home##", "<a href=""?" & qs & """>All Categories</a>")
                            getForm = getForm.Replace("##categoryName##", category)
                            getForm = getForm.Replace("##leftColumn##", cellLeft)
                            getForm = getForm.Replace("##centerColumn##", cellCenter)
                            getForm = getForm.Replace("##rightColumn##", cellRight)
                            '
                            Call cp.Cache.Save(cacheName, getForm, "Layouts, Directory Subcategories")
                        End If
                    End If
                End If
            Catch ex As Exception
                Call cp.Site.ErrorReport("CategoryHeadingClass.getForm, " & ex.Message)
            End Try
            '
        End Function


    End Class
End Namespace
