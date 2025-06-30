Imports System.Data

Partial Class CategoriPicker
    Inherits PageBase

    Protected Sub Page_Load1(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If Not IsPostBack Then

            '預設CAT = 999

            ViewState("Condition") = " and [cat] = '999'"

            GV_DataBind()

        End If



    End Sub

    '資料繫結
    Sub GV_DataBind()

        If Not ForProject.CheckAdmin(Request.QueryString("pUserID")) Then
            ViewState("Condition") &= " and [Dkey] <> '998' and [Dkey] <> '999'"

        End If
        Dim sql As String = "Select [Cat],[Data] , [DKey] from M_Referp where 1=1 " & ViewState("Condition") & " order by Dkey, Data"


        Dim uDataBase As New Utility.DataBase
        uDataBase.DefaultConnection = New SqlClient.SqlConnection(ForProject.GetConnectionString())

        GridView1.DataSource = uDataBase.GetDataTable(sql)
        GridView1.DataBind()

    End Sub

    '傳回父視窗
    Sub SelectReturn(ByVal cat As String)

        Dim uJScript As New Utility.JScript
        Dim sControlId As String() = Request.QueryString("pControlID").Split(",")
        Dim sAttribute As String() = {"value", "value"}
        Dim sValue As String() = {cat, cat}
        uJScript.RegJavaScript(Me, "picker", uJScript.ReturnValue(sControlId, sAttribute, sValue))

    End Sub

    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged
        GV_DataBind()
    End Sub

    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        GridView1.PageIndex = e.NewPageIndex
    End Sub

    '查詢
    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        ViewState("Condition") = " and [cat] = '999' and ([Dkey] like '%" & DDKey.Text.Trim & "%' or [Data] like '%" & DDKey.Text.Trim & "')"
        GV_DataBind()
    End Sub

    Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        If e.CommandName = "Pick" Then
            SelectReturn(e.CommandArgument)
        End If

    End Sub
End Class
