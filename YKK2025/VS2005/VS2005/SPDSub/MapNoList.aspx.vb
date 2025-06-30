Imports System.Data
Imports System.Data.OleDb
Partial Class MapNoList
    Inherits System.Web.UI.Page
    Dim InputData As String



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


        SQL = " select * from ( select '內製' Title,formno+'-'+ltrim(str(Formsno)) as SNo,No,Mapno,Buyer,SliderCode  from F_ManufInSheet"
        SQL = SQL + " where sts = 1"
        SQL = SQL + " and mapno <> ''"
        SQL = SQL + " union  all "
        SQL = SQL + " select  '外注' Title,formno+'-'+ltrim(str(Formsno)) as SNo,No,Mapno,Buyer,SliderCode  from F_ManufOutSheet"
        SQL = SQL + " where sts =1"
        SQL = SQL + "  and mapno <> ''"
        SQL = SQL + " )a where 1=1 and MapNo like '%" & InputData & "%'"

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


        Dim MapNo As String = GridView1.SelectedRow.Cells(3).Text
        Dim Buyer As String = GridView1.SelectedRow.Cells(4).Text
        Dim OFormNo As String = Mid(GridView1.SelectedRow.Cells(5).Text, 1, 6)
        Dim OFormSNo As String = Mid(GridView1.SelectedRow.Cells(5).Text, 8, Len(GridView1.SelectedRow.Cells(5).Text) - 7)
        Dim SliderCode As String = GridView1.SelectedRow.Cells(6).Text
     


        Dim js As String = ""
        js &= "var obj = window.opener.document.all.DOFormNo;"
        js &= "obj.value = '" & OFormNo & "';"
        js &= "var obj1 = window.opener.document.all.DOFormSno;"
        js &= "obj1.value = '" & OFormSNo & "';"
        js &= "var obj3 = window.opener.document.all.DMapNo;"
        js &= "obj3.value = '" & MapNo & "';"
        js &= "var obj4 = window.opener.document.all.DBuyer;"
        js &= "obj4.value = '" & Buyer & "';"
        '  js &= "var obj5 = window.opener.document.all.DSliderCode;"
        ' js &= "obj5.value = '" & SliderCode & "';"
        js &= "window.opener.document.forms[0].submit();"
        js &= "window.close();"


        Me.ClientScript.RegisterClientScriptBlock(GetType(String), "back", js, True)

    End Sub


    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged
        GetData()
    End Sub
End Class
