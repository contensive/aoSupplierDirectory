
Imports Contensive.BaseClasses

Namespace Contensive.Addons.SupplierDirectory

    Public Class CommonClass
        '
        ' Privates
        '
        Private isAccountContact_Local As Boolean
        Private isAccountContact_Local_Loaded As Boolean
        '
        '============================================================================
        '   get Layout
        '============================================================================
        '
        Friend Function getLayout(ByVal cp As CPBaseClass, ByVal LayoutName As String) As String
            Dim Layout As String = ""
            Dim filename As String = ""
            Dim CSLayout As BaseClasses.CPCSBaseClass
            Dim isFound As Boolean = False
            Dim cacheName As String
            '
            getLayout = ""
            Try
                cacheName = "Layout, " & LayoutName
                getLayout = cacheRead(cp, cacheName)
                If getLayout = "" Then
                    CSLayout = cp.CSNew
                    CSLayout.Open("layouts", "name=" & cp.Db.EncodeSQLText(LayoutName))
                    If CSLayout.OK Then
                        isFound = True
                        getLayout = CSLayout.GetTextFile("layout")
                    End If
                    Call CSLayout.Close()
                    If Not isFound Then
                        CSLayout.Insert("layouts")
                        If CSLayout.OK Then
                            Call CSLayout.SetField("name", LayoutName)
                            Call CSLayout.SetField("layout", "<!-- blank layout created automatically -->")
                        End If
                    End If
                    Call cp.Cache.Save(cacheName, getLayout, "Layouts")
                End If
            Catch ex As Exception
                Call cp.Site.ErrorReport(ex)
                'Call cp.Site.ErrorReport("CommonClass.getLayout, " & ex.Message)
            End Try
            '
        End Function
        '
        '============================================================================
        ' format a phone number (703) 406-3655
        '============================================================================
        '
        Friend Function formatPhone(ByVal cp As Contensive.BaseClasses.CPBaseClass, ByVal source As String) As String
            Dim w As String = ""
            Dim ext As String = ""
            Dim o As String = ""
            Dim pos As Integer = 0
            '
            o = source
            Try
                w = source
                '
                'normalize to 7034063655x15
                '
                w = w.ToLower
                w = w.Replace("extension", "x")
                w = w.Replace("ext.", "x")
                w = w.Replace("ext", "x")
                w = w.Replace("ex", "x")
                w = w.Replace(" ", "")
                w = w.Replace(".", "")
                w = w.Replace("-", "")
                w = w.Replace("(", "")
                w = w.Replace(")", "")
                w = w.Replace(")", "")
                '
                ' break off extension
                '
                pos = w.IndexOf("x")
                If pos >= 0 Then
                    ext = w.Substring(pos + 1)
                    If pos > 0 Then
                        w = w.Substring(0, pos)
                    Else
                        w = ""
                    End If
                End If
                '
                ' format output
                '
                If (w.Length <> 10) Then
                    o = source
                Else
                    o = w.Substring((w.Length - 4), 4)
                    If w.Length > 7 Then
                        o = w.Substring((w.Length - 7), 3) & "-" & o
                        If w.Length >= 10 Then
                            o = "(" & w.Substring((w.Length - 10), 3) & ")&nbsp;" & o
                        End If
                        If w.Length >= 11 Then
                            o = w.Substring(1, (w.Length - 11)) & ")&nbsp;" & o
                        End If
                    End If
                    If ext <> "" Then
                        o &= " x" & ext
                    End If
                End If
            Catch ex As Exception
                Call cp.Site.ErrorReport(ex)
                'Call cp.Site.ErrorReport("CommonClass.formatPhone, " & ex.Message)
            End Try
            formatPhone = o
        End Function
        '
        '
        '
        Friend Function getHtmlInputText(ByVal cp As CPBaseClass, ByVal cs As CPCSBaseClass, ByVal fieldName As String, Optional ByVal textArea As Boolean = False, Optional ByVal HtmlClass As String = "", Optional ByVal inputName As String = "") As String
            Dim s As String = ""
            Dim lines As String = "1"
            '
            Try
                If inputName = "" Then
                    inputName = fieldName
                End If
                If HtmlClass = "" Then
                    HtmlClass = fieldName
                End If
                If cs.OK Then
                    s = cs.GetText(fieldName)
                End If
                If textArea Then
                    lines = "10"
                End If
                s = cp.Html.InputText(inputName, s, lines, , , HtmlClass)
                If HtmlClass <> "" Then
                    s = s.Replace(">", " class=""" & HtmlClass & """>")
                End If
            Catch ex As Exception
                Call cp.Site.ErrorReport(ex)
                'Call cp.Site.ErrorReport("CommonClass.getHtmlInputText, " & ex.Message)
            End Try
            getHtmlInputText = s
        End Function
        '
        '
        '
        Friend Function getHtmlCheckbox(ByVal cp As Contensive.BaseClasses.CPBaseClass, ByVal cs As Contensive.BaseClasses.CPCSBaseClass, ByVal fieldName As String, Optional ByVal HtmlClass As String = "", Optional ByVal inputName As String = "") As String
            Dim s As String = ""
            Dim lines As Integer = 1
            Dim currentValue As Boolean = False
            '
            Try
                If inputName = "" Then
                    inputName = fieldName
                End If
                If HtmlClass = "" Then
                    HtmlClass = fieldName
                End If
                If cs.OK Then
                    currentValue = cs.GetBoolean(fieldName)
                End If
                s = cp.Html.CheckBox(inputName, currentValue, HtmlClass)
                'If HtmlClass <> "" Then
                's = s.Replace(">", " class=""" & HtmlClass & """>")
                'End If
            Catch ex As Exception
                Call cp.Site.ErrorReport(ex)
                'Call cp.Site.ErrorReport("CommonClass.getHtmlCheckbox, " & ex.Message)
            End Try
            getHtmlCheckbox = s
        End Function
        '
        '
        '
        Friend Function getHtmlSelectList(ByVal cp As Contensive.BaseClasses.CPBaseClass, ByVal cs As Contensive.BaseClasses.CPCSBaseClass, ByVal fieldName As String, ByVal optionList As String, Optional ByVal HtmlClass As String = "", Optional ByVal inputName As String = "") As String
            Dim s As String = ""
            Dim lines As Integer = 1
            Dim currentValue As Integer = 0
            '
            Try
                If inputName = "" Then
                    inputName = fieldName
                End If
                If HtmlClass = "" Then
                    HtmlClass = fieldName
                End If
                If cs.OK Then
                    currentValue = cs.GetInteger(fieldName)
                End If
                s = cp.Html.SelectList(inputName, currentValue.ToString, optionList, , HtmlClass)
            Catch ex As Exception
                Call cp.Site.ErrorReport(ex)
                'Call cp.Site.ErrorReport("CommonClass.getHtmlSelect, " & ex.Message)
            End Try
            getHtmlSelectList = s
        End Function
        '
        '
        '
        Friend Function getHtmlInputFile(ByVal cp As Contensive.BaseClasses.CPBaseClass, ByVal cs As Contensive.BaseClasses.CPCSBaseClass, ByVal fieldName As String, Optional ByVal HtmlClass As String = "", Optional ByVal inputName As String = "") As String
            Dim s As String = ""
            Dim virtualFilename As String = ""
            Dim imageFilename As String = ""
            Dim pos As Integer = 0
            '
            Try
                If inputName = "" Then
                    inputName = fieldName
                End If
                If HtmlClass = "" Then
                    HtmlClass = fieldName
                End If
                If cs.OK Then
                    virtualFilename = cs.GetText(fieldName)
                End If
                If virtualFilename <> "" Then
                    virtualFilename = virtualFilename.Replace("\", "/")
                    imageFilename = virtualFilename
                    pos = InStrRev(imageFilename, "/")
                    If pos >= 0 Then
                        imageFilename = imageFilename.Substring(pos)
                    End If
                    s = "<a href=""" & cp.Site.FilePath & virtualFilename & """ target=""_blank"">" & imageFilename & "</a>"
                    s &= "&nbsp;" & cp.Html.CheckBox(inputName & ".delete") & "&nbsp;delete&nbsp;&nbsp;"
                    s &= "&nbsp;" & cp.Html.InputFile(inputName, HtmlClass)
                Else
                    s = cp.Html.InputFile(inputName, HtmlClass)
                End If
                If HtmlClass <> "" Then
                    s = s.Replace(">", " class=""" & HtmlClass & """>")
                End If
                getHtmlInputFile = s
            Catch ex As Exception
                Call cp.Site.ErrorReport(ex)
            End Try
            getHtmlInputFile = s
        End Function
        '
        '
        '
        Friend Function getStateSelect(ByRef cp As Contensive.BaseClasses.CPBaseClass, ByVal htmlName As String, ByVal currentValue As String, ByVal HtmlClass As String)
            Dim s As String = ""
            Try
                s = getLayout(cp, "Supplier Directory State Select")
                If htmlName <> "" Then
                    s = s.Replace("name=""State""", "name=""" & htmlName & """")
                End If
                If currentValue <> "" Then
                    s = Replace(s, "<option>" & currentValue & "</option>", "<option selected=""selected"">" & currentValue & "</option>", , , CompareMethod.Text)
                End If
                If HtmlClass <> "" Then
                    s = Replace(s, "class=""stateSelect""", "class=""" & HtmlClass & """", , , CompareMethod.Text)
                End If
            Catch ex As Exception
                Call cp.Site.ErrorReport(ex)
            End Try
            Return s
        End Function
        '
        '
        '
        Friend Function getPermissionError() As String
            getPermissionError = "" _
                & vbCrLf & vbTab & "<div class=""sdPermissionError"">" _
                & vbCrLf & vbTab & vbTab & "There was a problem with the page you requested. The account you are logged into may not have the necessary permissions to view this page. Please use your back button and try again." _
                & vbCrLf & vbTab & "</div>"
        End Function
        '
        '
        '
        Friend Function isAccountContact(ByVal cp As CPBaseClass, ByVal orgId As Integer) As Boolean
            Dim csOrg As CPCSBaseClass
            '
            Try
                If Not isAccountContact_Local_Loaded Then
                    csOrg = cp.CSNew
                    Call csOrg.Open("organizations", "(id=" & orgId & ")and(directoryAccountContactId=" & cp.User.Id & ")", , , "id")
                    isAccountContact_Local = csOrg.OK
                    Call csOrg.Close()
                    isAccountContact_Local_Loaded = True
                End If
                Return isAccountContact_Local
            Catch ex As Exception

            End Try

        End Function
        '====================================================================================================
        ''' <summary>
        ''' Read from the cache if not admin site
        ''' </summary>
        ''' <param name="cp"></param>
        ''' <param name="name"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Friend Function cacheRead(cp As CPBaseClass, name As String) As String
            Dim returnResults As String = ""
            If Not cp.Doc.IsAdminSite Then
                returnResults = cp.Cache.Read(name)
            End If
            Return returnResults
        End Function
        '====================================================================================================
        ''' <summary>
        ''' save to the cache if not admin site
        ''' </summary>
        ''' <param name="cp"></param>
        ''' <param name="name"></param>
        ''' <param name="value"></param>
        ''' <param name="clearOnContentList"></param>
        ''' <remarks></remarks>
        Friend Sub cacheSave(cp As CPBaseClass, name As String, value As String, clearOnContentList As String)
            If Not cp.Doc.IsAdminSite Then
                cp.Cache.Save(name, value, clearOnContentList)
            End If
        End Sub
        '
    End Class
End Namespace
