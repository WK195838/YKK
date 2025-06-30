Imports System.Data
Imports System.Data.OleDb

Partial Class EMPListCL
    Inherits System.Web.UI.Page

    Dim InputData As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Me.IsPostBack Then   '不是PostBack
            If Request.QueryString("pKey") <> "" Then
                InputData = Request.QueryString("pKey")
                GetData()
            End If
        End If
    End Sub

    Sub GetData()
        Dim SQL As String
        'Dim UserID As String = Request.QueryString("userid")


        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        SQL = " select *  from M_wfsempCL "
        SQL = SQL + "where 1=1 "
        'Com_Code in ('10','11','12') "
        'SQL = SQL + "and substring(Dep_Code,1,1) = '1' "
        ' SQL = SQL + "and leav_cd <> 'Y' "
        If InputData <> "" Then
            SQL = SQL + " and ( DepName like '%" + InputData + "%' or empName like '%" + InputData + "%' or DepID like '%" + InputData + "%' or EMPID like '%" + InputData + "%')"
        End If

        SQL = SQL + " order by DepName, EMPID, empName "
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
        'js &= "window.opener.document.all.DDivision2.value = '" & GridView1.SelectedRow.Cells(1).Text & "';"
        'js &= "window.opener.document.all.DEmpName2.value = '" & GridView1.SelectedRow.Cells(2).Text & "';"
        'js &= "window.opener.document.all.DDivisionCode2.value = '" & GridView1.SelectedRow.Cells(3).Text & "';"
        'js &= "window.opener.document.all.DEmpID2.value = '" & GridView1.SelectedRow.Cells(4).Text & "';"
        'js &= "window.opener.document.all.DJobTitle2.value = '" & GridView1.SelectedRow.Cells(5).Text & "';"
        '  js &= "window.opener.document.all.CDivision2.checked = false;"
        'js &= "window.opener.document.forms[0].submit();"
        'js &= "window.close();"
        'Me.ClientScript.RegisterClientScriptBlock(GetType(String), "back", js, True)


        js &= "window.opener.document.all.DDivision2.value = '" & GridView1.SelectedRow.Cells(1).Text.Replace("&amp;", "&") & "';"
        js &= "window.opener.document.all.DEmpName2.value = '" & GridView1.SelectedRow.Cells(2).Text & "';"
        js &= "window.opener.document.all.DDivisionCode2.value = '" & GridView1.SelectedRow.Cells(3).Text & "';"
        js &= "window.opener.document.all.DEmpID2.value = '" & GridView1.SelectedRow.Cells(4).Text & "';"
        js &= "window.opener.document.all.D1.value = '" & GridView1.SelectedRow.Cells(4).Text & "';"
        js &= "window.opener.document.all.DJobTitle2.value = '" & GridView1.SelectedRow.Cells(5).Text & "';"
        js &= "window.opener.document.all.CDivision2.checked = false;"
        '   js &= "window.opener.document.forms[0].submit();"
        js &= "window.close();"


        Me.ClientScript.RegisterClientScriptBlock(GetType(String), "back", js, True)



    End Sub


    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged
        GetData()
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        '  e.Row.Cells(6).Visible = False
    End Sub
End Class
