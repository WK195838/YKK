Imports System.Data
Imports System.Data.OleDb
Partial Class NewColorAgainList
    Inherits System.Web.UI.Page
    Dim InputData As String



    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Me.IsPostBack Then   '不是PostBack
            GetData()
        End If
    End Sub

    Sub GetData()
        Dim SQL As String
        Dim No As String = Request.QueryString("pNo")


        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        SQL = " select no,convert(char(10),againSdate,111) as AgainSdate,convert(char(10),againedate,111) as Againedate,againdays,dyetimes, "
        SQL = SQL + " 'DTMW_NewColorAgain_02.aspx?&pFormNo='+Formno+'&pFormSno='+convert(char(1),FormSno) as URL"
        SQL = SQL + " from f_newcolorAgain where no='" + No + "'"



        OleDbConnection1.Open()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "Getata")
        GridView1.DataSource = DBDataSet1
        GridView1.DataBind()

        OleDbConnection1.Close()
    End Sub

    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        GridView1.PageIndex = e.NewPageIndex
    End Sub

    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged
        GetData()
    End Sub
End Class
