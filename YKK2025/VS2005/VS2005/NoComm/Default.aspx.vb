
Partial Class _Default
    Inherits System.Web.UI.Page

    Protected Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub


    Protected Sub Button7_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button7.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("NoCommSheet_02.aspx?pFormNo=001172" & "&pFormSno=" & TextBox2.Text.Trim)
        'Response.Redirect("BusinessTripSheet_01.aspx?pFormNo=003114&pFormSno=0&pStep=1&pSeqNo=1&pApplyID=sl014&pUserID=sl014")


    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("NoCommSheet_01.aspx?pFormNo=001172" & "&pFormSno=" & TextBox2.Text.Trim & "&pStep=" & TextBox3.Text.Trim & "&pSeqNo=" & TextBox4.Text.Trim & "&pApplyID=" & TextBox5.Text.Trim & "&pUserID=" & TextBox6.Text.Trim)
        'Response.Redirect("BusinessTripSheet_01.aspx?pFormNo=003114&pFormSno=0&pStep=1&pSeqNo=1&pApplyID=sl014&pUserID=sl014")
        'Response.Redirect("NoCommSheet_01.aspx?pFormNo=001172&pFormSno=0&pStep=1&pSeqNo=1&pApplyID=it013&pUserID=it013&pSales=B&pCust=E8200&pItem=8287259")
    End Sub

    Protected Sub Button2_Click1(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("NoCommList.aspx?pUserID=" & TextBox6.Text.Trim)
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


    Protected Sub Button5_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button5.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("EApprovalRDSheet_01.aspx?pFormNo=007003" & "&pFormSno=" & TextBox2.Text.Trim & "&pStep=" & TextBox3.Text.Trim & "&pSeqNo=" & TextBox4.Text.Trim & "&pApplyID=" & TextBox5.Text.Trim & "&pUserID=" & TextBox6.Text.Trim)
        'Response.Redirect("BusinessTripSheet_01.aspx?pFormNo=003114&pFormSno=0&pStep=1&pSeqNo=1&pApplyID=sl014&pUserID=sl014")

    End Sub

    Protected Sub Button6_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button6.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("EApprovalRDSheet_02.aspx?pFormNo=007003" & "&pFormSno=" & TextBox2.Text.Trim & "&pStep=" & TextBox3.Text.Trim & "&pSeqNo=" & TextBox4.Text.Trim & "&pApplyID=" & TextBox5.Text.Trim & "&pUserID=" & TextBox6.Text.Trim)
        'Response.Redirect("BusinessTripSheet_01.aspx?pFormNo=003114&pFormSno=0&pStep=1&pSeqNo=1&pApplyID=sl014&pUserID=sl014")

    End Sub

    Protected Sub Button8_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button8.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("EApprovalRDSchedule.aspx?pUserID=" & TextBox6.Text.Trim)
        'Response.Redirect("BusinessTripSheet_01.aspx?pFormNo=003114&pFormSno=0&pStep=1&pSeqNo=1&pApplyID=sl014&pUserID=sl014")

    End Sub
End Class
