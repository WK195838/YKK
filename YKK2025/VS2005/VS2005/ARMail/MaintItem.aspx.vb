Imports System.Data

Partial Class MaintItem
    Inherits PageBase
    Protected Sub Page_Load1(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Button2.Attributes.Add("onclick", "history.go(-1);return false;")
        Button1.Attributes.Add("onclick", "return confirm('" & ForProject.strSaveAlertMessage & "');")
        If Not IsPostBack Then
            Dim uDataBase As New Utility.DataBase
            uDataBase.DefaultConnection = New SqlClient.SqlConnection(ForProject.GetConnectionString())

            TextBox4.Text = Request.QueryString("pCat")
            If Request.QueryString("pFun") = "ADD" Then
                If TextBox4.Text <> "999" Then
                    '到DB找CAT
                    Dim strLevel As String = uDataBase.SelectQuery("Select Level from m_referp where cat='999' and dkey='" & TextBox4.Text & "'")
                    DropDownList1.SelectedIndex = strLevel
                    DropDownList1.Enabled = False
                End If

            Else

                'MOD
                Dim result As String() = uDataBase.SelectQuery("Select level , dkey , [data] from m_referp where unique_id=" & Request.QueryString("pID"), Utility.DataBase.SelectType.Columns).Split(",")
                TextBox1.Text = Request.QueryString("pID")
                TextBox4.Text = Request.QueryString("pCat")
                DropDownList1.SelectedIndex = result(0)
                TextBox2.Text = result(1)
                TextBox2.ReadOnly = True
                TextBox3.Text = result(2)
                If TextBox4.Text <> "999" Then

                    DropDownList1.Enabled = False
                End If
            End If
        End If
    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim uDataBase As New Utility.DataBase
        uDataBase.DefaultConnection = New SqlClient.SqlConnection(ForProject.GetConnectionString())

        Dim uJScript As New Utility.JScript
       

        Dim strValue As String = DropDownList1.SelectedValue & "-" & TextBox4.Text.Trim & "-" & TextBox2.Text.Trim & "-" & TextBox3.Text.Trim

        If Request.QueryString("pFun") = "ADD" Then

            '檢查是否已存在
            Dim strSql As String = "Select count(*) from m_referp where cat='" & TextBox4.Text.Trim & "' and dkey='" & TextBox2.Text.Trim & "' and data='" & TextBox3.Text & "'"
            Dim strResult As String = uDataBase.SelectQuery(strSql)
            If strResult <> "0" Then
                uJScript.PopMsg(Me, ForProject.strRecordExist)
            Else
                ForProject.InsertMaintHistory(Request.QueryString("pUserID"), "MaintItem", "ADD", "", strValue)

                'insert
                strSql = "insert into m_referp (Level,Cat,Dkey,Data,CreateUser,CreateTime,ModifyUser,ModifyTime) values (@Level,@Cat,@Dkey,@Data,@CreateUser,@CreateTime,@ModifyUser,@ModifyTime)"
                Dim cmdSql As New SqlClient.SqlCommand(strSql)
                cmdSql.Parameters.AddWithValue("@Level", DropDownList1.SelectedValue)
                cmdSql.Parameters.AddWithValue("@Cat", TextBox4.Text.Trim)
                cmdSql.Parameters.AddWithValue("@Dkey", TextBox2.Text.Trim)
                cmdSql.Parameters.AddWithValue("@Data", TextBox3.Text.Trim)
                cmdSql.Parameters.AddWithValue("@CreateUser", Request.QueryString("pUserID"))
                cmdSql.Parameters.AddWithValue("@CreateTime", Now)
                cmdSql.Parameters.AddWithValue("@ModifyUser", Request.QueryString("pUserID"))
                cmdSql.Parameters.AddWithValue("@ModifyTime", Now)
                uDataBase.ExecuteNonQuery(cmdSql)

                Response.Redirect("MaintReferp.aspx?pUserID=" & Request.QueryString("pUserID"))
            End If
        Else

            '檢查
            Dim strSql As String = "Select count(*) from m_referp where cat='" & TextBox4.Text.Trim & "' and dkey='" & TextBox2.Text.Trim & "'"
            Dim strResult As String = uDataBase.SelectQuery(strSql)
            If strResult = "0" Then
                uJScript.PopMsg(Me, ForProject.strRecordNotExist)
            Else
                '
                Dim result As String() = uDataBase.SelectQuery("Select level , dkey , [data] from m_referp where unique_id=" & Request.QueryString("pID"), Utility.DataBase.SelectType.Columns).Split(",")
                TextBox1.Text = Request.QueryString("pID")
                TextBox4.Text = Request.QueryString("pCat")

                Dim before As String = result(0) & "-" & TextBox4.Text & "-" & result(1) & "-" & result(2)
                ForProject.InsertMaintHistory(Request.QueryString("pUserID"), "MaintItem", "MOD", before, strValue)

                'MOD
                strSql = "Update m_referp SET Level=@Level,Cat=@Cat,Dkey=@Dkey,Data=@Data,ModifyUser=@ModifyUser,ModifyTime=@ModifyTime where Unique_ID = @Unique_ID"
                Dim cmdSql As New SqlClient.SqlCommand(strSql)
                cmdSql.Parameters.AddWithValue("@Level", DropDownList1.SelectedValue)
                cmdSql.Parameters.AddWithValue("@Cat", TextBox4.Text.Trim)
                cmdSql.Parameters.AddWithValue("@Dkey", TextBox2.Text.Trim)
                cmdSql.Parameters.AddWithValue("@Data", TextBox3.Text.Trim)
                cmdSql.Parameters.AddWithValue("@ModifyUser", Request.QueryString("pUserID"))
                cmdSql.Parameters.AddWithValue("@ModifyTime", Now)
                cmdSql.Parameters.AddWithValue("@Unique_ID", TextBox1.Text)
                uDataBase.ExecuteNonQuery(cmdSql)

                Response.Redirect("MaintReferp.aspx?pUserID=" & Request.QueryString("pUserID"))
            End If

        End If
    End Sub
End Class
