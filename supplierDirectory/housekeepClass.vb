
Option Strict On
Option Explicit On

Imports System.Xml
Imports System.Text
Imports System
Imports System.Collections.Generic
Imports Contensive.BaseClasses

Namespace Contensive.Addons.SupplierDirectory
    '
    ' Sample Vb addon
    '
    Public Class housekeepClass
        Inherits AddonBaseClass
        '
        Public Const spHousekeepStartTime As String = "supplierDirHousekeepStartTime"
        Public Const spHousekeepLastExecution As String = "supplierDirHousekeepLastExecution"
        Public Const aoVanityURLProcessor As String = "{ACCF60C6-4D62-4AB8-A10A-DDAC68968A77}"
        '
        '====================================================================================================
        ''' <summary>
        ''' housekeep Addon
        ''' </summary>
        ''' <param name="CP"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overrides Function Execute(ByVal CP As Contensive.BaseClasses.CPBaseClass) As Object
            Execute = ""
            Try
                '
                Dim rightNow As Date = Now()
                Dim cs As CPCSBaseClass = CP.CSNew()
                '
                ' housekeep executed hourly, but only runs at the predetermined hour.
                '
                appendLog(CP, "housekeep attempt")
                Dim housekeepStartTime As Integer = CP.Site.GetInteger(spHousekeepStartTime)
                Dim housekeepLastExecution As Date = CP.Site.GetDate(spHousekeepLastExecution, Date.MinValue.ToString)
                '
                If (housekeepLastExecution.Date < rightNow.Date) And (rightNow.Hour >= housekeepStartTime) Then
                    Call CP.Site.SetProperty(spHousekeepLastExecution, rightNow.ToString)
                    appendLog(CP, "housekeep needed, start")
                    '
                    ' make sure all yesterday's email notifications were sent
                    '
                    Dim email As New EmailNotificationClass
                    Call email.Execute(CP)
                    '
                    ' verify all the vanity URLs
                    '
                    Call CP.Utils.ExecuteAddon(aoVanityURLProcessor)
                    '
                    '
                    '
                    appendLog(CP, "housekeep, exit")
                End If
            Catch ex As Exception
                CP.Site.ErrorReport(ex, "Exception in Contensive.Addons.SupplierDirectory.housekeepClass")
            End Try
            '
            Return Execute
        End Function
        '
        Private Sub appendLog(cp As CPBaseClass, MsgBox As String)
            Dim rightNow As Date = Now()
            cp.Utils.AppendLog("housekeep\SupplierDir" & rightNow.Year & rightNow.Month.ToString().PadLeft(2, "0"c) & rightNow.Day.ToString().PadLeft(2, "0"c) & ".log", MsgBox)
        End Sub
    End Class
End Namespace
