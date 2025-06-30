
Partial Class _Default
    Inherits System.Web.UI.Page
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     主程式
    '**
    '*****************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    End Sub
    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("SBDCommissionSheet_01.aspx?pFormNo=" & DropDownList1.SelectedValue & "&pFormSno=" & TextBox2.Text.Trim & "&pStep=" & TextBox3.Text.Trim & "&pSeqNo=" & TextBox4.Text.Trim & "&pApplyID=" & TextBox5.Text.Trim)
        '  Response.Redirect("HRWYM_TimeOffSheet_02.aspx?pFormNo=001204 " & "&pFormSno=" & TextBox2.Text.Trim & "&pStep=" & TextBox3.Text.Trim & "&pSeqNo=" & TextBox4.Text.Trim & "&pApplyID=" & TextBox5.Text.Trim & "&pAgentID=")

    End Sub

  
    Protected Sub TextBox2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged

    End Sub

    Protected Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("SBDCommissionSheet_02.aspx?pFormNo=" & DropDownList1.SelectedValue & "&pFormSno=" & TextBox2.Text.Trim & "&pStep=" & TextBox3.Text.Trim & "&pSeqNo=" & TextBox4.Text.Trim & "&pApplyID=" & TextBox5.Text.Trim)
        '  Response.Redirect("HRWYM_TimeOffSheet_02.aspx?pFormNo=001204 " & "&pFormSno=" & TextBox2.Text.Trim & "&pStep=" & TextBox3.Text.Trim & "&pSeqNo=" & TextBox4.Text.Trim & "&pApplyID=" & TextBox5.Text.Trim & "&pAgentID=")

    End Sub

    Protected Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("SBDSurfaceSheet_01.aspx?pFormNo=003002" & "&pFormSno=" & TextBox2.Text.Trim & "&pStep=" & TextBox3.Text.Trim & "&pSeqNo=" & TextBox4.Text.Trim & "&pApplyID=" & TextBox5.Text.Trim)
        '  Response.Redirect("HRWYM_TimeOffSheet_02.aspx?pFormNo=001204 " & "&pFormSno=" & TextBox2.Text.Trim & "&pStep=" & TextBox3.Text.Trim & "&pSeqNo=" & TextBox4.Text.Trim & "&pApplyID=" & TextBox5.Text.Trim & "&pAgentID=")

    End Sub

    Protected Sub TextBox6_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox6.TextChanged

    End Sub

    Protected Sub Button4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button4.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("SBDSurfaceSheet_02.aspx?pFormNo=003002" & "&pFormSno=" & TextBox2.Text.Trim & "&pStep=" & TextBox3.Text.Trim & "&pSeqNo=" & TextBox4.Text.Trim & "&pApplyID=" & TextBox5.Text.Trim)
        '  Response.Redirect("HRWYM_TimeOffSheet_02.aspx?pFormNo=001204 " & "&pFormSno=" & TextBox2.Text.Trim & "&pStep=" & TextBox3.Text.Trim & "&pSeqNo=" & TextBox4.Text.Trim & "&pApplyID=" & TextBox5.Text.Trim & "&pAgentID=")

    End Sub
End Class
