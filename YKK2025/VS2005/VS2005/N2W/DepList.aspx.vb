Imports System.Data
Imports System.Data.OleDb
Partial Class DepList
    Inherits System.Web.UI.Page
    Dim InputData As String



    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Me.IsPostBack Then   '不是PostBack
            GetData()
        End If
    End Sub

    Sub GetData()
        Dim SQL As String
        'Dim UserID As String = Request.QueryString("userid")

        InputData = DData.Text
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        SQL = " select substring(data,1,7)DepID,substring(data,9,len(data)-1)DepName from M_referp"
        SQL = SQL + " where  cat = 3105  and dkey ='dep'"


        If InputData <> "" Then
            SQL = SQL + " and ( Data like '%" + InputData + "%' or Data like '%" + InputData + "%')"
        End If

        SQL = SQL + "order by data  "
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


        Dim DepID As String = Trim(GridView1.SelectedRow.Cells(1).Text)
        Dim DepName As String = Trim(GridView1.SelectedRow.Cells(2).Text)


        Dim js As String = ""
        js &= "var obj = window.opener.document.all.DDepID;"
        js &= "obj.value = '" & DepID & "';"
        js &= "var obj = window.opener.document.all.DDepName;"
        js &= "obj.value = '" & DepName & "';"
        js &= "window.opener.document.forms[0].submit();"
        js &= "window.close();"


        Me.ClientScript.RegisterClientScriptBlock(GetType(String), "back", js, True)

    End Sub


    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged
        GetData()
    End Sub
End Class
