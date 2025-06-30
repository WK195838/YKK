Imports System.Data
Imports System.Data.OleDb
Partial Class CopyQA
    Inherits System.Web.UI.Page
    Dim InputData As String
    Dim First As String
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()



    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Me.IsPostBack Then   '不是PostBack
            GetData()

        End If
    End Sub

    Sub GetData()
        Dim SQL As String
        Dim formno As String = Request.QueryString("formno")

        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        SQL = " select case when sts =0 then '開發中'  "
        SQL = SQL + " when sts = 1 then '完成(OK)'    "
        SQL = SQL + " when sts = 2 then  '取消/中止'    "
        SQL = SQL + " else '抽單' end as sts,F_QASheet.no,name,FormNo,FormSno, "
        SQL = SQL + " 'http://10.245.1.6/ISOSQC/QASheet_02.aspx?pFormNo=008002&pFormSno='+convert(char,formsno) as ViewURL,"
        SQL = SQL + " '('+Supplier+') '+Size+' '+Family+' '+Body+' '+Puller+' '+Color+' '+Finish as 'Type' "
        SQL = SQL + " from F_QASheet left JOIN F_QASheetDT ON F_QASheetDT.NO = F_QASheet.NO "
        SQL = SQL + " where formno ='008002' AND F_QASheet.NO NOT IN (SELECT DISTINCT ModNo FROM F_QAModSheet WHERE F_QAModSheet.Sts = 0 )  "
        SQL = SQL + " AND F_QASheet.Sts <> 2  "


        If InputData <> "" Then
            SQL = SQL + " and ( F_QASheet.no like '%" + InputData + "%'"
            SQL = SQL + " or Size+Family+Body+Puller+Color+Finish Like '%" + InputData.Replace(" ", "") + "%'"
            SQL = SQL + " )"

        End If
        SQL = SQL + " order by no desc"

        OleDbConnection1.Open()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "Getata")
        GridView1.DataSource = DBDataSet1
        GridView1.DataBind()

        OleDbConnection1.Close()
    End Sub


    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click, DData.TextChanged
        First = ""
        InputData = DData.Text
        GetData()
    End Sub

    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        GridView1.PageIndex = e.NewPageIndex
    End Sub

    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged

        Dim No As String = GridView1.SelectedRow.Cells(2).Text
        Dim js As String = ""
        js &= "var obj = window.opener.document.all.DNO1;"
        js &= "obj.value = '" & No & "';"
        js &= "window.opener.document.forms[0].submit();"
        js &= "window.close();"

        Me.ClientScript.RegisterClientScriptBlock(GetType(String), "back", js, True)

    End Sub


    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged
        GetData()
    End Sub
End Class
