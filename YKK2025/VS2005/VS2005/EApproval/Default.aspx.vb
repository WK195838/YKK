
Partial Class _Default
    Inherits System.Web.UI.Page

    Protected Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub


    Protected Sub Button7_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button7.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("EApprovalSheet_02.aspx?pFormNo=007002" & "&pFormSno=" & TextBox2.Text.Trim)
        'Response.Redirect("BusinessTripSheet_01.aspx?pFormNo=003114&pFormSno=0&pStep=1&pSeqNo=1&pApplyID=sl014&pUserID=sl014")


    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("EApprovalSheet_01.aspx?pFormNo=007002" & "&pFormSno=" & TextBox2.Text.Trim & "&pStep=" & TextBox3.Text.Trim & "&pSeqNo=" & TextBox4.Text.Trim & "&pApplyID=" & TextBox5.Text.Trim & "&pUserID=" & TextBox6.Text.Trim)
        'Response.Redirect("BusinessTripSheet_01.aspx?pFormNo=003114&pFormSno=0&pStep=1&pSeqNo=1&pApplyID=sl014&pUserID=sl014")

    End Sub

    Protected Sub Button2_Click1(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("EApprovalSchedule.aspx?pUserID=" & TextBox6.Text.Trim)
        'Response.Redirect("BusinessTripSheet_01.aspx?pFormNo=003114&pFormSno=0&pStep=1&pSeqNo=1&pApplyID=sl014&pUserID=sl014")

    End Sub

    Protected Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("QAModSheet_01.aspx?pFormNo=008003" & "&pFormSno=" & TextBox2.Text.Trim & "&pStep=" & TextBox3.Text.Trim & "&pSeqNo=" & TextBox4.Text.Trim & "&pApplyID=" & TextBox5.Text.Trim & "&pUserID=" & TextBox6.Text.Trim)
        'Response.Redirect("BusinessTripSheet_01.aspx?pFormNo=003114&pFormSno=0&pStep=1&pSeqNo=1&pApplyID=sl014&pUserID=sl014")

    End Sub

    Protected Sub Button4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button4.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("QAModSheet_02.aspx?pFormNo=008003" & "&pFormSno=" & TextBox2.Text.Trim & "&pStep=" & TextBox3.Text.Trim & "&pSeqNo=" & TextBox4.Text.Trim & "&pApplyID=" & TextBox5.Text.Trim & "&pUserID=" & TextBox6.Text.Trim)
        'Response.Redirect("BusinessTripSheet_01.aspx?pFormNo=003114&pFormSno=0&pStep=1&pSeqNo=1&pApplyID=sl014&pUserID=sl014")

    End Sub


End Class
