Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.BaseClasses

Namespace addonCollectionName
    '
    ' Sample Vb addon
    '
    Public Class HelloWorldClass
        Inherits AddonBaseClass
        '====================================================================================================
        '
        Private cp As CPBaseClass
        '
        '====================================================================================================
        ''' <summary>
        ''' Addon Interface
        ''' </summary>
        ''' <param name="CP"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
            Dim returnHtml As String
            Try

                Dim contentId As Integer = CP.Doc.GetInteger("contentId")
                Dim recordId As Integer = CP.Doc.GetInteger("recordId")
                If contentId > 0 And recordId > 0 Then
                    Select Case CP.Content.GetRecordName("content", contentId).ToLower
                        Case "organizations"
                            Call getContent(recordId)
                    End Select
                End If
                returnHtml = ""
            Catch ex As Exception
                errorReport(CP, ex, "execute")
                returnHtml = ""
            End Try
            Return returnHtml
        End Function
        '====================================================================================================
        ''' <summary>
        ''' Organization edit handler
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub getContent(organizationId As Integer)

            Dim cs As CPCSBaseClass = cp.CSNew()
            Dim csA As CPCSBaseClass = cp.CSNew()
            Dim aliasName As String
            Dim aliasPrefix As String
            Dim pageID As Integer
            '
            cs = cp.csNew()
            csA = cp.csNew()
            aliasPrefix = cp.site.GetProperty("Supplier Directory Vanity URL Prefix")
            pageID = cp.site.GetProperty("Supplier Directory Location")
            '
            If cs.Open("Organizations", "(id=" & organizationId & ")and(includeInDirectory=1) and (directoryIsEnhancedListing=1)", , , "id,name") Then
                Do While cs.OK()
                    aliasName = aliasPrefix & "/" & cs.GetText("name")
                    aliasName = Replace(aliasName, "'", "-")
                    aliasName = Replace(aliasName, """", "-")
                    aliasName = Replace(aliasName, " ", "-")
                    aliasName = Replace(aliasName, "&", "-")
                    aliasName = Replace(aliasName, "=", "-")
                    aliasName = Replace(aliasName, "?", "-")
                    aliasName = Replace(aliasName, "+", "-")
                    aliasName = Replace(aliasName, ",", "-")
                    aliasName = cp.Site.addLinkAlias(aliasName, pageID, qs)
                    '
                    Call csA.Open("Link Aliases", "name=" & cp.Db.EncodeSQLText(aliasName))
                    If Not csA.OK() Then
                        csA.Close()
                        Call csA.Insert("Link Aliases")
                    End If
                    If csA.OK() Then
                        Call csA.SetField("name", aliasName)
                        Call csA.SetField("PageID", pageID)
                        Call csA.SetField("QueryStringSuffix", "formId=4&organizationid=" & cs.GetInteger("ID"))
                    End If
                    csA.Close()
                    '
                    Call cs.GoNext()
                Loop
            End If
            Call cs.Close()

        End Sub        '
        '=====================================================================================
        ' common report for this class
        '=====================================================================================
        '
        Private Sub errorReport(ByVal cp As CPBaseClass, ByVal ex As Exception, ByVal method As String)
            Try
                CP.Site.ErrorReport(ex, "Unexpected error in sampleClass." & method)
            Catch exLost As Exception
                '
                ' stop anything thrown from cp errorReport
                '
            End Try
        End Sub
    End Class
End Namespace
