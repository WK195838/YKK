Imports System.Data
Partial Class MaintCustControl
    Inherits PageBase
   
    Protected Sub Page_Load1(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If Not IsPostBack Then
            GV_Bind()
        End If
    End Sub


    Sub GV_Bind()

        Dim strSql As String = ""
        strSql &= "SELECT  Unique_id,CustCode,case Active when 1 then '啟用' else '停用' end as active," & _
               "CustCode + '-' + CustName as CustName," & _
               "MailToName1 + case when MailToName2 <> '' then ',' + MailToName2 " & _
               "else '' end as MailToName ,salesman,case PDF_Create when 1 then '是' " & _
               "else '否' end as PDF_Create,case PDF_Recreate when 1 then '是' " & _
               "else '否' end as PDF_Recreate,PDF_Period,case SMTP_Send when 1 then '是' " & _
               "else '否' end as SMTP_Send,case SMTP_ReSend when 1 then '是' " & _
               "else '否' end as SMTP_ReSend, SMTP_Period FROM [M_CustControl] where 1=1 "

        If Not String.IsNullOrEmpty(DCust.Text.Trim) Then
            strSql &= " and (custcode like '%" & DCust.Text.Trim & "%' or custname like '%" & DCust.Text.Trim & "%') "
        End If
        If Not String.IsNullOrEmpty(DSalesMan.Text.Trim) Then
            strSql &= " and salesman like '%" & DSalesMan.Text.Trim & "%'"
        End If
        If Not String.IsNullOrEmpty(DSalesCode.Text.Trim) Then
            strSql &= " and salesCode like '%" & DSalesCode.Text.Trim & "%'"
        End If

        If DropDownList1.SelectedValue <> "ALL" Then
            strSql &= " and IntCust = " & DropDownList1.SelectedValue
        End If
        If DropDownList2.SelectedValue <> "ALL" Then
            strSql &= " and SMTP_Send = " & DropDownList2.SelectedValue
        End If
        strSql &= " order by custcode,custname,salesman"

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

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            CType(e.Row.Cells(0).FindControl("Button3"), Button).Attributes.Add("onclick", "return confirm('" & ForProject.strDeleteAlertMessage & "');")
        End If
    End Sub

    Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        Dim sUnique_ID As String = e.CommandArgument.ToString
        Dim uDataBase As New Utility.DataBase
        uDataBase.DefaultConnection = New SqlClient.SqlConnection(ForProject.GetConnectionString())

        Select Case e.CommandName

            Case "Upt"
                If ForProject.CheckUserAdmin(Request.QueryString("pUserID")) Then
                    Response.Redirect("MaintCustomer.aspx?pFun=MOD&pUserID=" & Request.QueryString("pUserID") & "&pID=" & sUnique_ID)

                Else
                    Dim uJavaScript As New Utility.JScript
                    uJavaScript.PopMsg(Me, ForProject.strCannotMod)

                End If

            Case "Del"
                If ForProject.CheckUserAdmin(Request.QueryString("pUserID")) Then
                    Dim sql As String = "Select CustCode + '-' + CustName + '-' + MailToName1 + '-' + MailToAddress1 + '-' + MailToPosition1" & _
                                            " + '-' + MailToName2 + '-' + MailToAddress2 + '-' + MailToPosition2 + '-' + Cast(MailCCList as char(1)) + '-' + SalesCode + '-' + SalesMan + '-'" & _
                                            "+ SalesMailAddress + '-' + Cast(PDF_Create  as char(1)) + '-' + Cast(SMTP_Send as char(1)) from m_custcontrol where unique_id=" & sUnique_ID
                    Dim before As String = uDataBase.SelectQuery(sql)

                    ForProject.InsertMaintHistory(Request.QueryString("pUserID"), "MaintCustControl", "DEL", before, "")

                    sql = "Delete from m_custcontrol where Unique_ID = " & sUnique_ID
                    uDataBase.ExecuteNonQuery(sql)
                    GV_Bind()

                Else
                    Dim uJavaScript As New Utility.JScript
                    uJavaScript.PopMsg(Me, ForProject.strCannotDel)
                End If

            Case "Link"
                Response.Redirect("ViewProcHistory.aspx?pUserID=" & Request.QueryString("pUserID") & "&pCust=" & sUnique_ID)
        End Select

    End Sub

    '新增
    Protected Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
        If ForProject.CheckUserAdmin(Request.QueryString("pUserID")) Then
            Response.Redirect("MaintCustomer.aspx?pFun=ADD&pUserID=" & Request.QueryString("pUserID") & "&pID=")
        Else
            Dim uJavaScript As New Utility.JScript
            uJavaScript.PopMsg(Me, ForProject.strCannotIns)
        End If
    End Sub
End Class
