Imports System.Data
Imports System.Drawing
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Collections.Generic

Partial Class ReferenceListCL
    Inherits System.Web.UI.Page

    Dim fpObj As New ForProject     '操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript

    Dim InputData As String

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '*****************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Me.IsPostBack Then   '不是PostBack
            GetData()
        End If
    End Sub

    Sub GetData()
        Dim SQL As String

        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        SQL = " Select  top 100  ExpCat+'--'+ExpItem as ExpItem,convert(char(10),Adate,111) Adate,TaxType,NetAmt,TaxAmt,Amt,Content,Remark,ACID from F_FundingSheetCLdt where 1=1 "
        If InputData <> "" Then
            SQL = SQL + " and ( ExpItem like '%" + InputData + "%')"
        End If
        SQL = SQL + " order by unique_id desc "
        OleDbConnection1.Open()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "Getata")
        GridView1.DataSource = DBDataSet1
        GridView1.DataBind()

        OleDbConnection1.Close()
    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click, DData.TextChanged
        InputData = DData.Text
        GetData()
    End Sub


    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        GridView1.PageIndex = e.NewPageIndex
    End Sub

    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged
        Dim fpObj As New ForProject
        Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
        '
        Dim js As String = ""
        js &= "window.opener.document.all.DExpItem1.value = '" & GridView1.SelectedRow.Cells(1).Text & "';"
        js &= "window.opener.document.all.DADate.value = '" & GridView1.SelectedRow.Cells(2).Text & "';"
        js &= "window.opener.document.all.DTaxType.value = '" & GridView1.SelectedRow.Cells(3).Text & "';"

        js &= "window.opener.document.all.DNetAmt.value = '" & GridView1.SelectedRow.Cells(4).Text & "';"
        js &= "window.opener.document.all.DTaxAmt.value = '" & GridView1.SelectedRow.Cells(5).Text & "';"
        js &= "window.opener.document.all.DAmt.value = '" & GridView1.SelectedRow.Cells(6).Text & "';"

        js &= "window.opener.document.all.DContent1.value = '" & GridView1.SelectedRow.Cells(7).Text & "';"
        If Trim(GridView1.SelectedRow.Cells(8).Text) = "&nbsp;" Then
            js &= "window.opener.document.all.DRemark1.value = '" & "" & "';"
        Else
            js &= "window.opener.document.all.DRemark1.value = '" & Trim(GridView1.SelectedRow.Cells(8).Text) & "';"
        End If
        js &= "window.opener.document.all.DACID.value = '" & GridView1.SelectedRow.Cells(9).Text & "';"


        js &= "window.opener.document.forms[0].submit();"
        js &= "window.close();"
        '
        Me.ClientScript.RegisterClientScriptBlock(GetType(String), "back", js, True)
    End Sub


    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged
        GetData()
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        If GridView1.Rows.Count > 0 Then
            e.Row.Cells(9).Visible = False
        End If

    End Sub
End Class
