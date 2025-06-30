Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic

Partial Class GentaniSheet
    Inherits System.Web.UI.Page

   
    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then

            Dim uDataBase As New Utility.DataBase
            Dim uCommon As New Utility.Common
            uDataBase.DefaultConnection = New SqlConnection(uCommon.GetConnectionString("SFD_Con"))
            Dim strSql As String = "SELECT * FROM [V_Gentani_01] WHERE (([RNO] = '" & Request.QueryString("DRNO") & "') AND ([DEVNO] = '" & Request.QueryString("DDEVNO") & "') AND ([CODENO] = '" & Request.QueryString("DCODENO") & "'))"
            Dim dt As DataTable = uDataBase.GetDataTable(strSql)

            Dim s As New List(Of String)
            s.Add("TNLITEM")
            s.Add("TNRITEM")
            s.Add("ECOL")
            s.Add("EITEM")

            For Each dc As DataColumn In dt.Columns
                Dim l As Object = Me.form1.FindControl("D" & dc.ColumnName)

                
                Dim v As String = uCommon.ReplaceDBnull(dt.Rows(0)(dc.ColumnName), "")
                If l IsNot Nothing Then
                    'If dc.ColumnName = "OTHER1" Or dc.ColumnName = "OTHER2" Then
                    '    If v.IndexOf(" ") <> -1 Then
                    '        v = v.Replace(" ", "<br>")
                    '        l.Style("top") = l.Style("top").Replace("px", "") - 10 & "px"
                    '    End If
                    'End If
                    l.Text = v
                Else
                    If s.Contains(dc.ColumnName) Then

                        l = Me.form1.FindControl("D" & dc.ColumnName & "1")
                        l.Text = v
                        l = Me.form1.FindControl("D" & dc.ColumnName & "2")
                        l.Text = v

                    End If
                End If
            Next
        End If
    End Sub
End Class
