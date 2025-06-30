Imports System.Data
Imports System.Data.OleDb
Partial Class YKKColorCodeListUAKIPLING
    Inherits System.Web.UI.Page
    Dim InputData As String
    Dim InputData1 As String



    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Me.IsPostBack Then   '不是PostBack
            '  GetData()

        End If
    End Sub

    Sub GetData()
        Dim SQL As String
        'Dim UserID As String = Request.QueryString("userid")


        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定
        '查詢Wave Color Code 
        SQL = " select   paccxx,cta1xx,ccp1xx,ccp2xx,cta2xx from  MST_color_structure  where 1=1 "
        If InputData <> "" Then
            SQL = SQL + " and  ctbnxx='" + InputData + "'"
        End If

        If InputData1 <> "" Then
            SQL = SQL + " and  paccxx like '" + InputData1 + "%'"
        Else
            SQL = " select  paccxx,cta1xx,ccp1xx,ccp2xx,cta2xx  from  MST_color_structure  where 1=1 "
            SQL = SQL + " and  ctbnxx='" + InputData + "'"
        End If
        '查詢待處理 Color Code 
        SQL = SQL + " union all "
        SQL = SQL + " select distinct paccxx,cta1xx,ccp1xx,ccp2xx,cta2xx from ("
        SQL = SQL + " select   ykkcolorcode as paccxx,'' as cta1xx , '' as  ccp1xx,'' as ccp2xx,'' as cta2xx from V_NewColor"
        SQL = SQL + " where  sts not in (2,3)  and  YKKColorCode like '" + InputData1 + "%'"
        '查詢待處理 Color Code 
        SQL = SQL + " union all "
        SQL = SQL + " select   ykkcolorcode as paccxx,'' as cta1xx , '' as  ccp1xx,'' as ccp2xx,'' as cta2xx from F_CombiSheet"
        SQL = SQL + " where sts not in (2,3) and YKKColorCode like '" + InputData1 + "%'"
        SQL = SQL + " )a where paccxx not in ( select   paccxx   from  MST_color_structure where  ctbnxx='" + InputData + "' )"

        SQL = "select * from (" + SQL + ")a  order by paccxx"


        OleDbConnection1.Open()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "Getata")
        GridView1.DataSource = DBDataSet1
        GridView1.DataBind()

        OleDbConnection1.Close()
    End Sub


    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click, DData.TextChanged
        InputData = DData.Text
        InputData1 = DData1.Text

        GetData()
    End Sub

    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        GridView1.PageIndex = e.NewPageIndex
    End Sub

    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged

        Dim fpObj As New ForProject
        Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()


        Dim ColorCode As String = GridView1.SelectedRow.Cells(1).Text


        Dim js As String = ""
        '  js &= "var obj = window.opener.document.all.DYKKColorCode;"

        js &= "var obj = window.opener.document.all." + Request.QueryString("field") + ";"


        js &= "obj.value = '" & ColorCode & "';"

        js &= "window.opener.document.forms[0].submit();"
        js &= "window.close();"


        Me.ClientScript.RegisterClientScriptBlock(GetType(String), "back", js, True)

    End Sub


    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged
        GetData()
    End Sub
End Class
