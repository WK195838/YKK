Imports System.Data
Imports System.Data.OleDb
Partial Class DiscountCopy
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

        formno = "003102"


        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定


        SQL = "  select" + First + "  case when a.sts =0 then '開發中'   when a.sts = 1 then '完成(OK)'     when a.sts = 2 then  '取消/中止' "
        SQL = SQL + "  else '抽單' end as sts,a.no,a.appname as name,a.FormNo,a.FormSno,itemcode,"
        SQL = SQL + " 'http://10.245.1.10/N2W/DiscountSheet_02.aspx?pFormNo='+a.formno+'&pFormSno='+convert(char,a.formsno) as ViewURL "
        SQL = SQL + " from  F_DiscountSheet a,F_DiscountSheetdt b  where a.formsno=b.formsno and a.formno ='003102' and a.sts=1 "

        If InputData <> "" Then
            SQL = SQL + " and ( a.no like '%" + InputData + "%'"
            SQL = SQL + " or  a.appname like '%" + InputData + "%'"
            SQL = SQL + " or  b.itemcode like '%" + InputData + "%'"
            SQL = SQL + " )"

        End If
        SQL = SQL + " order by a.no desc"

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
