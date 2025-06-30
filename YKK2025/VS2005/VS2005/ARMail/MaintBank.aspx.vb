Imports System.data
Partial Class MaintBank
    Inherits PageBase

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click

        If Not String.IsNullOrEmpty(TextBox1.Text.Trim) Then
            Dim sql As String = " where 1=1 and CustCode like '%" & TextBox1.Text.Trim & "%' or CustName like '%" & TextBox1.Text.Trim & "%' "
            SqlDataSource1.SelectCommand = SqlDataSource1.SelectCommand.Replace("where 1=1", sql)
            
        Else
            SqlDataSource1.SelectCommand = "Select * from M_Bank where 1=1 Order by CustCode,CustName"
        End If
        GridView1.PageIndex = 0
        GridView1.DataBind()
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            CType(e.Row.Cells(0).FindControl("Button2"), Button).Attributes.Add("onclick", "return confirm('" & ForProject.strDeleteAlertMessage & "');")
        End If
    End Sub

    Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        Dim sUnique_ID As String = e.CommandArgument.ToString
        Dim uDataBase As New Utility.DataBase
        uDataBase.DefaultConnection = New SqlClient.SqlConnection(ForProject.GetConnectionString())

        Select Case e.CommandName

            Case "Upt"
                If ForProject.CheckUserAdmin(Request.QueryString("pUserID")) Then
                    Response.Redirect("MaintBankData.aspx?pFun=MOD&pUserID=" & Request.QueryString("pUserID") & "&pID=" & sUnique_ID)

                Else
                    Dim uJavaScript As New Utility.JScript
                    uJavaScript.PopMsg(Me, ForProject.strCannotMod)

                End If

            Case "Del"
                If ForProject.CheckUserAdmin(Request.QueryString("pUserID")) Then
                    Dim sql As String = "Select CustCode + '-' + CustName + '-' + SalesCode + '-' + BankName + '-' + BankAddress" & _
                                            " + '-' + BankACNo + '-' + BankACName + '-' + Swift  from m_bank where CustCode='" & sUnique_ID & "'"
                    Dim before As String = uDataBase.SelectQuery(sql)

                    ForProject.InsertMaintHistory(Request.QueryString("pUserID"), "MaintBank", "DEL", before, "")

                    sql = "Delete from m_Bank where CustCode = '" & sUnique_ID & "'"
                    uDataBase.ExecuteNonQuery(sql)
                    GridView1.PageIndex = 0
                    GridView1.DataBind()

                Else
                    Dim uJavaScript As New Utility.JScript
                    uJavaScript.PopMsg(Me, ForProject.strCannotDel)
                End If

           
        End Select

    End Sub

    '新增
    Protected Sub Button4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button4.Click
        If ForProject.CheckUserAdmin(Request.QueryString("pUserID")) Then
            Response.Redirect("MaintBankData.aspx?pFun=ADD&pUserID=" & Request.QueryString("pUserID") & "&pID=")
        Else
            Dim uJavaScript As New Utility.JScript
            uJavaScript.PopMsg(Me, ForProject.strCannotIns)
        End If
    End Sub
End Class
