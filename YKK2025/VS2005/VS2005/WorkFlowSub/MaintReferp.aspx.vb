Imports System.Data.SqlClient
Imports System.Data
Imports System.Web.Security

Partial Class MaintReferp
    Inherits System.Web.UI.Page

    Dim UserID As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        UserID = Request.QueryString("pUserID")
        If Not IsPostBack Then
            DCat.Text = Request.QueryString("pCat")
            GV_BIND()
        End If
        BSelect.Attributes.Add("onclick", "window.open('CategoriPicker.aspx?pUserID=" & UserID & "','','scrollbars=no,status=no,width=400,height=600'　);")
    End Sub

    Protected Sub BSerch_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BSerch.Click
        GV_BIND()
    End Sub

    Sub GV_BIND()
        If DCat.Text <> "" Then
            GridView1.Visible = True
            Dim sqlStr As String = ""
            sqlStr = "Select * from m_referp "
            sqlStr = sqlStr + "where cat ='" & DCat.Text & "' "
            If DDKey.Text <> "" And DDKey.Text <> "DKey" Then
                sqlStr = sqlStr + "  and dkey like '%" & DDKey.Text & "%' "
            End If
            sqlStr = sqlStr + "order by dkey, data"

            Dim dbcon As New SqlConnection(System.Web.Configuration.WebConfigurationManager.ConnectionStrings("WorkFlow_Con").ToString)

            Dim dbcmd As New SqlCommand(sqlStr, dbcon)
            dbcmd.Connection.Open()
            Dim dbAdpter As New SqlDataAdapter(dbcmd)
            Dim ds As New Data.DataSet
            dbAdpter.Fill(ds)
            GridView1.DataSource = ds
            GridView1.DataBind()
            dbcmd.Connection.Close()
        Else
            GridView1.Visible = False
        End If
    End Sub


    Protected Sub Page_PreInit(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreInit

    End Sub

    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged
        GV_BIND()
    End Sub

    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        GridView1.PageIndex = e.NewPageIndex
    End Sub

    Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        Dim jsStr As String


        Select Case e.CommandName


            Case "Upd"
                Response.Redirect("MaintItem.aspx?pFun=MOD&pID=" & e.CommandArgument & "&pUserID=" & UserID & "&pcat=" & DCat.Text)
            Case "Del"
                '刪除語法
                Me.ClientScript.RegisterClientScriptBlock(GetType(String), "abc", jsStr, True)
                Dim sqlStr As String = "Delete from M_Referp where Unique_ID = '" & e.CommandArgument & "'"
                Dim dbcon As New SqlConnection(System.Web.Configuration.WebConfigurationManager.ConnectionStrings("WorkFlow_Con").ToString)
                Dim dbcmd As New SqlCommand(sqlStr, dbcon)

                dbcmd.Connection.Open()
                dbcmd.ExecuteNonQuery()
                dbcmd.Connection.Close()
                '刪除後重新 binding
                GV_BIND()
        End Select
    End Sub


    Protected Sub Page_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRender
        DCat.Attributes.Add("readonly", "true")

    End Sub


    Protected Sub BADD_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BADD.Click
        Response.Redirect("MaintItem.aspx?pfun=ADD&pUserid=" & UserID & "&pcat=" & DCat.Text & "&pID=")
    End Sub

End Class
