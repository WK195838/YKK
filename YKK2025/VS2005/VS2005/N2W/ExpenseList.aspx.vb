Imports System.Data
Imports System.Data.OleDb
Partial Class ExpenseList
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

        SQL = " select no,appname,freason,formno,formsno from F_ExpenseSheet  where sts =1 "
        If InputData <> "" Then
            SQL = SQL + " and ( no like '%" + InputData + "%'" + " or appname like '%" + InputData + "%'" + " )"
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
        js &= "window.opener.document.all.DENo.value = '" & GridView1.SelectedRow.Cells(1).Text.Replace("&amp;", "&") & "';"
        js &= "window.opener.document.all.DExpenseNo.value = '" & GridView1.SelectedRow.Cells(1).Text.Replace("&amp;", "&") & "';"
        '  js &= "window.opener.document.forms[0].submit();"
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
            formno = e.Row.Cells(4).Text
            formsno = e.Row.Cells(5).Text
            ' 連結到客訴表單   
            h1.NavigateUrl = "ExpenseSheet_02.aspx?pFormNo=" + formno + "&pFormSno=" + formsno
            ' e.Row.Cells(3).Text = ""
            e.Row.Cells(1).Controls.Add(h1)
            Count = Count + 1
        End If

        If Count <> 0 Then
            e.Row.Cells(4).Visible = False
            e.Row.Cells(5).Visible = False
        End If
    

    End Sub
End Class
