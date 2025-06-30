
Partial Class _Default_Stock
    Inherits System.Web.UI.Page



 


    Protected Sub Button26_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button26.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("StockInSheet_01.aspx?pFormNo=003112" & "&pFormSno=" & TextBox2.Text.Trim & "&pStep=" & TextBox3.Text.Trim & "&pSeqNo=" & TextBox4.Text.Trim & "&pApplyID=" & TextBox5.Text.Trim & "&pUserID=" & TextBox6.Text.Trim)

    End Sub

    Protected Sub Button27_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button27.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("StockOutSheet_01.aspx?pFormNo=003113" & "&pFormSno=" & TextBox2.Text.Trim & "&pStep=" & TextBox3.Text.Trim & "&pSeqNo=" & TextBox4.Text.Trim & "&pApplyID=" & TextBox5.Text.Trim & "&pUserID=" & TextBox6.Text.Trim)

    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("StockInSheet_04.aspx?pFormNo=003113" & "&pFormSno=" & TextBox2.Text.Trim & "&pStep=" & TextBox3.Text.Trim & "&pSeqNo=" & TextBox4.Text.Trim & "&pApplyID=" & TextBox5.Text.Trim & "&pUserID=" & TextBox6.Text.Trim)
 
    End Sub

    Protected Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("BusinessTripSheet_01.aspx?pFormNo=003114" & "&pFormSno=" & TextBox2.Text.Trim & "&pStep=" & TextBox3.Text.Trim & "&pSeqNo=" & TextBox4.Text.Trim & "&pApplyID=" & TextBox5.Text.Trim & "&pUserID=" & TextBox6.Text.Trim)
        'Response.Redirect("BusinessTripSheet_01.aspx?pFormNo=003114&pFormSno=0&pStep=1&pSeqNo=1&pApplyID=sl014&pUserID=sl014")

    End Sub

    Protected Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("StockInSheet_03.aspx?pFormNo=003113" & "&pFormSno=" & TextBox2.Text.Trim & "&pStep=" & TextBox3.Text.Trim & "&pSeqNo=" & TextBox4.Text.Trim & "&pApplyID=" & TextBox5.Text.Trim & "&pUserID=" & TextBox6.Text.Trim)

    End Sub

    Protected Sub Button4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button4.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("CloseAccountSheet_01.aspx?pFormNo=003115" & "&pFormSno=" & TextBox2.Text.Trim & "&pStep=" & TextBox3.Text.Trim & "&pSeqNo=" & TextBox4.Text.Trim & "&pApplyID=" & TextBox5.Text.Trim & "&pUserID=" & TextBox6.Text.Trim)
        'Response.Redirect("CloseAccountSheet_011.aspx?pFormNo=003115&pFormSno=0&pStep=1&pSeqNo=1&pApplyID=sl014&pUserID=sl014")
    End Sub

    Protected Sub Button5_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button5.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("StockNewCommission.aspx?pUserID=" & TextBox6.Text.Trim)
    End Sub

    Protected Sub Button6_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button6.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("CloseAccountSheet_02.aspx?pFormNo=003115" & "&pFormSno=" & TextBox2.Text.Trim & "&pStep=" & TextBox3.Text.Trim & "&pSeqNo=" & TextBox4.Text.Trim & "&pApplyID=" & TextBox5.Text.Trim & "&pUserID=" & TextBox6.Text.Trim)
        'Response.Redirect("CloseAccountSheet_011.aspx?pFormNo=003115&pFormSno=0&pStep=1&pSeqNo=1&pApplyID=sl014&pUserID=sl014")

    End Sub

    Protected Sub Button7_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button7.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("BusinessTripSheet_02.aspx?pFormNo=003114" & "&pFormSno=" & TextBox2.Text.Trim & "&pStep=" & TextBox3.Text.Trim & "&pSeqNo=" & TextBox4.Text.Trim & "&pApplyID=" & TextBox5.Text.Trim & "&pUserID=" & TextBox6.Text.Trim)
        'Response.Redirect("BusinessTripSheet_01.aspx?pFormNo=003114&pFormSno=0&pStep=1&pSeqNo=1&pApplyID=sl014&pUserID=sl014")

    End Sub

    Protected Sub Button8_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button8.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("ScheduleBC.aspx?pUserID=" & TextBox6.Text.Trim)
    End Sub
End Class
