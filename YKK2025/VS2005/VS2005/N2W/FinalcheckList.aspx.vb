Imports System.Data
Imports System.Data.OleDb
Partial Class FinalcheckList
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
        Dim formsno As String = Request.QueryString("formsno")
        InputData = DData.Text

        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        SQL = "select  no,accdepName,FORMSNO, "
        SQL = SQL + " 'http://10.245.1.10/N2W/FinalcheckModSheet_01.aspx?pOFormNo='+rtrim(convert(char,formsno))+'&pFormNo=003107&pFormSno=0&pStep=1&pSeqNo=1&pApplyID='+createuser+'&pUserID='+createuser "
        SQL = SQL + "  as ViewURL"
        SQL = SQL + " from F_Finalchecksheet "

        If InputData <> "" Then
            SQL = SQL + " where  no like '%" + InputData + "%' "

        End If
        SQL = SQL + " order by no "

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


        'Dim No As String = GridView1.SelectedRow.Cells(1).Text
        'Dim AppName As String = GridView1.SelectedRow.Cells(2).Text
        'Dim APPDepo As String = GridView1.SelectedRow.Cells(3).Text
        'Dim SNo As String = GridView1.SelectedRow.Cells(4).Text

        'Dim js As String = ""
        'js &= "var obj = window.opener.document.all.DNO1;"
        'js &= "obj.value = '" & No & "';"
        'js &= "var obj = window.opener.document.all.DAppName;"
        'js &= "obj.value = '" & AppName & "';"
        'js &= "var obj = window.opener.document.all.DAPPDepo;"
        'js &= "obj.value = '" & APPDepo & "';"
        'js &= "var obj = window.opener.document.all.DSNO;"
        'js &= "obj.value = '" & SNo & "';"
        'js &= "window.opener.document.forms[0].submit();"
        'js &= "window.close();"


        'Me.ClientScript.RegisterClientScriptBlock(GetType(String), "back", js, True)

    End Sub


    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged
        GetData()
    End Sub


    Protected Sub GridView1_DataBound(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        '  e.Row.Cells(4).Visible = False
    End Sub
End Class
