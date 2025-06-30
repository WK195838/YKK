Imports System.Data
Imports System.Data.OleDb
Partial Class DTMW_NewColorCopy
    Inherits System.Web.UI.Page
    Dim InputData As String
    Dim First As String
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()



    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Me.IsPostBack Then   '不是PostBack
            First = " Top 20 "
            GetData()
         
        End If
    End Sub

    Sub GetData()
        Dim SQL As String
        Dim formno As String = Request.QueryString("formno")

        If formno = "005012" Then
            formno = "005011"
        End If
        'formno = "005001"

        SQL = " SELECT REPLACE(DATA,'01','02')URL  FROM m_REFERP"
        SQL = SQL + " WHERE  CAT  = '5001'"
        SQL = SQL + " AND  DKEY = 'URL-" + formno + "'"
        Dim DBAdapter2 As DataTable = uDataBase.GetDataTable(SQL)
        Dim URL As String = DBAdapter2.Rows(0).Item("URL")

        SQL = " SELECT Data  FROM m_REFERP"
        SQL = SQL + " WHERE  CAT  = '5001'"
        SQL = SQL + " AND  DKEY = 'Sheet-" + formno + "'"
        Dim DBAdapter3 As DataTable = uDataBase.GetDataTable(SQL)
        Dim Sheet As String = DBAdapter3.Rows(0).Item("Data")


        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        SQL = " select  " + First + "  case when sts =0 then '開發中'  "
        SQL = SQL + " when sts = 1 then '完成(OK)'    "
        SQL = SQL + " when sts = 2 then  '取消/中止'    "
        SQL = SQL + " else '抽單' end as sts, '" + Sheet + "' as Sheet,no,name,FormNo,FormSno, "
        SQL = SQL + " 'http://10.245.1.10/DTMW/'+'" + URL + "?pFormNo='+formno+'&pFormSno='+convert(char,formsno) as ViewURL"
        SQL = SQL + " from v_Newcolor where formno ='" + formno + "'"

        If InputData <> "" Then
            SQL = SQL + " and ( no like '%" + InputData + "%'"
            SQL = SQL + " or  name like '%" + InputData + "%'"
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

   
        Dim No As String = GridView1.SelectedRow.Cells(3).Text
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
