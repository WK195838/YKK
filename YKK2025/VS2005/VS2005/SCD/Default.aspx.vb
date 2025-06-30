
Partial Class _Default
    Inherits System.Web.UI.Page
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     主程式
    '**
    '*****************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Button3.Attributes("onclick") = "window.open('OpenNoPickerForSampleSheet.aspx','NoPicker','status=0,toolbar=0,width=620,height=650,resizable=yes,scrollbars=yes');"    '選取參考開發資料

    End Sub
    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("http://LocalHost:2790/SCD/CommissionSheet_01.aspx?pFormNo=" & DropDownList1.SelectedValue & "&pFormSno=" & TextBox2.Text.Trim & "&pStep=" & TextBox3.Text.Trim & "&pSeqNo=" & TextBox4.Text.Trim & "&pApplyID=" & TextBox5.Text.Trim)

    End Sub

    Protected Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("http://LocalHost:2790/SCD/CommissionSheet_02.aspx?pFormNo=" & DropDownList1.SelectedValue & "&pFormSno=" & TextBox2.Text.Trim)
    End Sub

    Protected Sub Button4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button4.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("http://LocalHost:2790/SCD/SampleSheet_01.aspx?pFormNo=" & DropDownList1.SelectedValue & "&pFormSno=" & TextBox2.Text.Trim & "&pStep=" & TextBox3.Text.Trim & "&pSeqNo=" & TextBox4.Text.Trim & "&pApplyID=" & TextBox5.Text.Trim)
    End Sub

    Protected Sub Button5_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button5.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("http://LocalHost:2790/SCD/SampleSheet_02.aspx?pFormNo=" & DropDownList1.SelectedValue & "&pFormSno=" & TextBox2.Text.Trim)
    End Sub
End Class
