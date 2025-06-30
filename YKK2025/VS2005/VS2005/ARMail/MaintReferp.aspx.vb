Imports System.Data


Partial Class MaintReferp
    Inherits PageBase

    Protected Sub Page_Load1(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim sql As String = "window.open('categoripicker.aspx?pUserID=" & Request.QueryString("pUserID") & "&pControlID=TextBox1,HiddenField1','','scrollbars=no,status=no,width=400,height=600,top=0,left=0');"

        Button1.Attributes.Add("onclick", sql)

        If IsPostBack Then
            If Not String.IsNullOrEmpty(HiddenField1.Value) Then
                TextBox1.Text = HiddenField1.Value
            End If
        End If
        
    End Sub
    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged
        GV_DataBind()
    End Sub

    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        GridView1.PageIndex = e.NewPageIndex
    End Sub

    Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        GV_DataBind()
        If e.CommandName <> "Page" Then
            Dim bolAdmin As Boolean = ForProject.CheckAdmin(Request.QueryString("pUserID"))
            Dim sUnique_ID As String = e.CommandArgument.ToString.Split(",")(0)
            Dim sLevel As String = e.CommandArgument.ToString.Split(",")(1)
            If e.CommandName = "Upt" Then
                'level 0 or1 才能修改
                If bolAdmin Or (Int(sLevel) <= 1) Then
                    Response.Redirect("MaintItem.aspx?pFun=MOD&pUserID=" & Request.QueryString("pUserID") & "&pCat=" & TextBox1.Text.Trim & "&pID=" & sUnique_ID)
                Else
                    Dim uJavaScript As New Utility.JScript
                    uJavaScript.PopMsg(Me, ForProject.strCannotMod)
                End If
            Else
                'level 0 or 2 才能刪除
                If bolAdmin Or (sLevel = "0" Or sLevel = "2") Then
                    Dim uDataBase As New Utility.DataBase
                    uDataBase.DefaultConnection = New SqlClient.SqlConnection(ForProject.GetConnectionString())
                    Dim rowindex As Integer = CType(CType(e.CommandSource, Button).NamingContainer, GridViewRow).RowIndex

                    Dim before As String = sLevel & "-" & TextBox1.Text.Trim & "-" & GridView1.Rows(rowindex).Cells(2).Text & "-" & GridView1.Rows(rowindex).Cells(3).Text
                    ForProject.InsertMaintHistory(Request.QueryString("pUserID"), "MaintReferp", "DEL", before, "")
                    Dim sql As String = "Delete from m_referp where Unique_ID = " & sUnique_ID
                    uDataBase.ExecuteNonQuery(sql)
                    GV_DataBind()
                Else
                    Dim uJavaScript As New Utility.JScript
                    uJavaScript.PopMsg(Me, ForProject.strCannotDel)
                End If
            End If
        End If

    End Sub

    Sub GV_DataBind()
        If Not String.IsNullOrEmpty(TextBox1.Text.Trim) Then
            Dim uDataBase As New Utility.DataBase
            uDataBase.DefaultConnection = New SqlClient.SqlConnection(ForProject.GetConnectionString())

            Label1.Text = "類別:" & TextBox1.Text.Trim & " (" & uDataBase.SelectQuery("Select [Data] from m_referp where cat='999' and dkey ='" & TextBox1.Text.Trim & "'") & ")"
            Dim sql As String = ""
            If Not String.IsNullOrEmpty(TextBox2.Text.Trim) Then
                sql = "Select [Level],[Unique_ID] , [Cat] , [Data] , [DKey] from M_Referp where [Cat] = '" & HiddenField1.Value & "' and dkey='" & TextBox2.Text.Trim & "' order by Cat,Dkey, Data"

            Else
                sql = "Select [Level],[Unique_ID] , [Cat] , [Data] , [DKey] from M_Referp where [Cat] = '" & HiddenField1.Value & "' order by Cat,Dkey, Data"

            End If

            GridView1.DataSource = uDataBase.GetDataTable(sql)
            GridView1.DataBind()
            GridView1.Visible = True
            Label1.Visible = True
        End If
    End Sub

    

    '新增
    Protected Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.Click
        If Not String.IsNullOrEmpty(TextBox1.Text.Trim) Then
            '一般使用者要level = 0 or 1才能新增,此level看cat=999的header
            Dim bolAdmin As Boolean = ForProject.CheckAdmin(Request.QueryString("pUserID"))
            If Not bolAdmin Then
                Dim uDataBase As New Utility.DataBase
                uDataBase.DefaultConnection = New SqlClient.SqlConnection(ForProject.GetConnectionString())
                Dim strLevel As String = uDataBase.SelectQuery("Select [Level] from m_referp where cat='999' and Dkey='" & TextBox1.Text.Trim & "'")
                If Int(strLevel) <= 1 Then
                    Response.Redirect("MaintItem.aspx?pFun=ADD&pUserID=" & Request.QueryString("pUserID") & "&pCat=" & TextBox1.Text.Trim & "&pID=")
                Else
                    Dim uJavaScript As New Utility.JScript
                    uJavaScript.PopMsg(Me, ForProject.strCannotIns)
                End If
            Else
                Response.Redirect("MaintItem.aspx?pFun=ADD&pUserID=" & Request.QueryString("pUserID") & "&pCat=" & TextBox1.Text.Trim & "&pID=")
            End If
        Else
            Dim uJavaScript As New Utility.JScript
            uJavaScript.PopMsg(Me, ForProject.strChoiceMsg)
        End If
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            CType(e.Row.Cells(0).FindControl("Button4"), Button).Attributes.Add("onclick", "return confirm('" & ForProject.strDeleteAlertMessage & "');")
        End If
    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        GridView1.Visible = False
        Label1.Visible = False
        TextBox2.Text = ""
    End Sub

    '類別+項目
    Protected Sub Button6_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button6.Click
        GV_DataBind()
    End Sub
End Class
