Imports System.Data
Imports System.Data.OleDb
Partial Class QCList
    Inherits System.Web.UI.Page
    Dim InputData As String
    Dim Count As Integer




    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Me.IsPostBack Then   '不是PostBack
            '  GetData()
            Count = 0
        End If
    End Sub

    Sub GetData()
        Dim SQL As String
        'Dim UserID As String = Request.QueryString("userid")


        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        SQL = " select  no,appname,qcno,substring(cpcontent,1,30)cpcontent,formno,formsno from F_ComplaintOutSheet  where sts in (0,1) "
        If InputData <> "" Then
            SQL = SQL + " and ( no like '%" + InputData + "%'" + " or appname like '%" + InputData + "%'" + " or qcno like '%" + InputData + "%'" + ")"
        End If

        SQL = SQL + " order by  no "
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



        Dim js As String = ""
        js &= "window.opener.document.all.DComplainNo.value = '" & GridView1.SelectedRow.Cells(1).Text.Replace("&amp;", "&") & "';"
        js &= "window.opener.document.all.DQC.value = '" & GridView1.SelectedRow.Cells(1).Text.Replace("&amp;", "&") & "';"
        js &= "window.opener.document.forms[0].submit();"
        js &= "window.close();"


        Me.ClientScript.RegisterClientScriptBlock(GetType(String), "back", js, True)

    End Sub


    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged
        GetData()
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound



        If (e.Row.RowType = DataControlRowType.DataRow) <> (e.Row.RowType = DataControlRowType.Footer) Then
            Dim h1 As New HyperLink
            Dim formno As String
            Dim formsno As String
            h1.Target = "_blank"
            h1.Text = e.Row.Cells(1).Text
            formno = e.Row.Cells(5).Text
            formsno = e.Row.Cells(6).Text
            ' 連結到客訴表單   
            h1.NavigateUrl = "http://10.245.1.10/N2W/ComplaintOutSheet_02.aspx?&pFormno=003109&" + formno + "&pFormSno=" + formsno
            ' e.Row.Cells(3).Text = ""
            e.Row.Cells(1).Controls.Add(h1)

            Count = Count + 0
        End If

        If Count <> 0 Then
            e.Row.Cells(5).Visible = False
            e.Row.Cells(6).Visible = False
        End If




    End Sub
End Class
