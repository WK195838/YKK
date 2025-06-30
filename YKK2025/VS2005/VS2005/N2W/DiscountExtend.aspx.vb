Imports System.Data
Imports System.Data.OleDb
Partial Class DiscountExtend
    Inherits System.Web.UI.Page
    Dim InputData As String
    Dim First As String
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()



    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Me.IsPostBack Then   '不是PostBack
            '   First = " Top 20 "
            GetData()

        End If
    End Sub

    Sub GetData()
        Dim SQL As String
        Dim formno As String = Request.QueryString("formno")
        Dim pNo As String
        pNo = Request.QueryString("pNo")

        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定


        SQL = " select  case when sts =0 then '開發中'   when sts = 1 then '完成(OK)'     when sts = 2 then  '取消/中止'     else '抽單' end as sts,convert(char(10),date,111) Date,no,appname as name,convert(char(10),ASdate,111)ASDATE,convert(char(10),AEdate,111)AEdate,FormNo,FormSno,"
        SQL = SQL + " 'http://10.245.1.10/N2W/DiscountSheet_02.aspx?pFormNo='+formno+'&pFormSno='+convert(char,formsno) as ViewURL "
        SQL = SQL & " from  F_DiscountSheet where Oformsno='" & pNo & "'"
        SQL = SQL + " order by no desc"

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
