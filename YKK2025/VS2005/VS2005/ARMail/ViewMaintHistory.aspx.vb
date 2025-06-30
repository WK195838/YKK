Imports System.Data

Partial Class ViewMaintHistory
    Inherits PageBase

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            GV_Bind()
        End If
    End Sub

    Sub GV_Bind()

        Dim strSql As String = ""
        strSql &= "SELECT top 500 ProcProgram,ProcTime,ProcUserName + '(' + ProcUser + ')' as ProcUser ,[Function],Before,After from T_MaintHistory"

        If Not String.IsNullOrEmpty(TextBox1.Text.Trim) Then
            strSql &= " where ProcProgram like '%" & TextBox1.Text.Trim & "%' or ProcUser like '%" & TextBox1.Text.Trim & "%' or ProcUserName like '%" & TextBox1.Text.Trim & "%'"

        End If
        strSql &= " order by ProcTime DESC"

        Dim uDataBase As New Utility.DataBase
        uDataBase.DefaultConnection = New SqlClient.SqlConnection(ForProject.GetConnectionString())

        GridView1.DataSource = uDataBase.GetDataTable(strSql)
        GridView1.DataBind()

    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        GV_Bind()

    End Sub

    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged
        GV_Bind()
    End Sub

    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        GridView1.PageIndex = e.NewPageIndex
    End Sub

    
End Class
