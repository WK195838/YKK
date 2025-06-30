Imports System.Data
Imports System.Data.OleDb
Partial Class OrderList
    Inherits System.Web.UI.Page
    Dim InputData As String
    Dim uCommon As New Utility.Common



    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Me.IsPostBack Then   '不是PostBack
            '  GetData()
            ' DData.Text = "OR04071102"
        End If
    End Sub

    Sub GetData()

        Dim SQL As String
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        Dim ConnectString As String = uCommon.GetAppSetting("WAVESOLEDB")
        '
        cn.ConnectionString = ConnectString
        '
        If Mid(InputData, 1, 2) <> "OR" Then
            InputData = "OR" + InputData
        End If
        SQL = "SELECT ORDN5C,BYRC5C,DK1C5C FROM WAVEDLIB.S5C00 "
        SQL = SQL + " WHERE ORDN5C ='" & InputData & "' "
      
        '
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, cn)
        DBAdapter1.Fill(ds, "S5C00")

        GridView1.DataSource = ds
        GridView1.DataBind()


    End Sub


    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click, DData.TextChanged
        InputData = Trim(DData.Text)
        GetData()
    End Sub

    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        GridView1.PageIndex = e.NewPageIndex
    End Sub

    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged

        Dim fpObj As New ForProject
        Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()

        Dim ORNO As String = Mid(GridView1.SelectedRow.Cells(1).Text, 3, Len(GridView1.SelectedRow.Cells(1).Text) - 1)
        Dim CUST As String = GridView1.SelectedRow.Cells(2).Text
        Dim Buyer As String = GridView1.SelectedRow.Cells(3).Text

       

        Dim js As String = ""
        js &= "var obj = window.opener.document.all.D1;"
        js &= "obj.value = '" & CUST & "';"
        js &= "var obj = window.opener.document.all.D2;"
        js &= "obj.value = '" & Buyer & "';"
        js &= "var obj = window.opener.document.all.DORNO;"
        js &= "obj.value = '" & ORNO & "';"
        js &= "window.opener.document.forms[0].submit();"
        js &= "window.close();"


        Me.ClientScript.RegisterClientScriptBlock(GetType(String), "back", js, True)

    End Sub


    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged
        GetData()
    End Sub
End Class
