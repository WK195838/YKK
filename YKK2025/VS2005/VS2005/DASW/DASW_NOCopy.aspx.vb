Imports System.Data
Imports System.Data.OleDb
Partial Class DASW_NOCopy
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

        

        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        SQL = "select  no,appname,deponame,FORMSNO, "
        SQL = SQL + " 'http://10.245.1.6/DASW/DASW_DISPOSALSheet03.aspx?pFormNo='+formno+'&pFormSno='+rtrim(convert(char,formsno))+'&pStep=600&pSeqNo=1&pApplyID='+createuser+'&pUserID='+createuser "
        SQL = SQL + "  as ViewURL,"
        SQL = SQL + " 'http://10.245.1.6/DASW/DASW_DISPOSALSheet02.aspx?pFormNo='+formno+'&pFormSno='+rtrim(convert(char,formsno))+'&pStep=600&pSeqNo=1&pApplyID='+createuser+'&pUserID='+createuser "
        SQL = SQL + "  as ViewURLNO "
        SQL = SQL + " from f_disposalsheet"
        SQL = SQL + " where disposalym = LEFT(convert(char(10),getdate(),111),7)  "
        'SQL = SQL + "  where disposalym = '2022/02'  "

        'If Request.QueryString("pUserID") = "it013" Or Request.QueryString("pUserID") = "it003" Or Request.QueryString("pUserID") = "it004" Then
        '    SQL = SQL + " "
        'Else
        '    SQL = SQL + "  and createuser in('it013','it004','it003','" + Request.QueryString("pUserID") + "')"
        'End If
       


        If InputData <> "" Then
            SQL = SQL + " and ( no like '%" + InputData + "%'"

            SQL = SQL + " )"

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


        Dim No As String = GridView1.SelectedRow.Cells(1).Text
        Dim AppName As String = GridView1.SelectedRow.Cells(2).Text
        Dim APPDepo As String = GridView1.SelectedRow.Cells(3).Text
        Dim SNo As String = GridView1.SelectedRow.Cells(4).Text

        Dim js As String = ""
        js &= "var obj = window.opener.document.all.DNO1;"
        js &= "obj.value = '" & No & "';"
        js &= "var obj = window.opener.document.all.DAppName;"
        js &= "obj.value = '" & AppName & "';"
        js &= "var obj = window.opener.document.all.DAPPDepo;"
        js &= "obj.value = '" & APPDepo & "';"
        js &= "var obj = window.opener.document.all.DSNO;"
        js &= "obj.value = '" & SNo & "';"
        js &= "window.opener.document.forms[0].submit();"
        js &= "window.close();"


        Me.ClientScript.RegisterClientScriptBlock(GetType(String), "back", js, True)

    End Sub


    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged
        GetData()
    End Sub

  
    Protected Sub GridView1_DataBound(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        ' e.Row.Cells(3).Visible = False
    End Sub
End Class
