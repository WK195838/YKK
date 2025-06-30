
Partial Class _Default
    Inherits System.Web.UI.Page
    '--------------------------------------------------------------------------------------------------
    '  Item Register
    '--------------------------------------------------------------------------------------------------
    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ItemRegister.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("ISMSSheet_01.aspx?pFormNo=003101" & "&pFormSno=" & TextBox2.Text.Trim & "&pStep=" & TextBox3.Text.Trim & "&pSeqNo=" & TextBox4.Text.Trim & "&pApplyID=" & TextBox5.Text.Trim & "&pAgentID=")

    End Sub

    Protected Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ItemRegister2.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("ISMSSheet_02.aspx?pFormNo=003101" & "&pFormSno=" & TextBox2.Text.Trim)
    End Sub

    Protected Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ItemRegisterHistory.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("http://localhost/WorkFlow/BefOPList.aspx?pFormNo=" & DropDownList1.SelectedValue & "&pFormSno=" & TextBox2.Text.Trim & "&pStep=" & TextBox3.Text.Trim & "&pSeqNo=" & TextBox4.Text.Trim & "&pApplyID=" & TextBox5.Text.Trim)
    End Sub
    '--------------------------------------------------------------------------------------------------
    '  ZIP Register
    '--------------------------------------------------------------------------------------------------
    Protected Sub ZIPRegister_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ZIPRegister.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("DASW_DISPOSALSheet01.aspx?pFormNo=006001" & "&pFormSno=" & TextBox2.Text.Trim & "&pStep=" & TextBox3.Text.Trim & "&pSeqNo=" & TextBox4.Text.Trim & "&pApplyID=" & TextBox5.Text.Trim & "&pAgentID=")
    End Sub

    Protected Sub ZIPRegister2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ZIPRegister2.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("ItemRegisterZIPSheet_02.aspx?pFormNo=" & DropDownList1.SelectedValue & "&pFormSno=" & TextBox2.Text.Trim)

    End Sub

    Protected Sub ZIPRegisterHistory_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ZIPRegisterHistory.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("http://localhost/WorkFlow/BefOPList.aspx?pFormNo=" & DropDownList1.SelectedValue & "&pFormSno=" & TextBox2.Text.Trim & "&pStep=" & TextBox3.Text.Trim & "&pSeqNo=" & TextBox4.Text.Trim & "&pApplyID=" & TextBox5.Text.Trim)
    End Sub
    '--------------------------------------------------------------------------------------------------
    '  SLD Register
    '--------------------------------------------------------------------------------------------------
    Protected Sub SLDRegister_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SLDRegister.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("ItemRegisterSLDSheet_01.aspx?pFormNo=" & DropDownList1.SelectedValue & "&pFormSno=" & TextBox2.Text.Trim & "&pStep=" & TextBox3.Text.Trim & "&pSeqNo=" & TextBox4.Text.Trim & "&pApplyID=" & TextBox5.Text.Trim & "&pAgentID=")
    End Sub

    Protected Sub SLDRegister2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SLDRegister2.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("ItemRegisterSLDSheet_02.aspx?pFormNo=" & DropDownList1.SelectedValue & "&pFormSno=" & TextBox2.Text.Trim)
    End Sub

    Protected Sub SLDRegisterHistory_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SLDRegisterHistory.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("http://localhost/WorkFlow/BefOPList.aspx?pFormNo=" & DropDownList1.SelectedValue & "&pFormSno=" & TextBox2.Text.Trim & "&pStep=" & TextBox3.Text.Trim & "&pSeqNo=" & TextBox4.Text.Trim & "&pApplyID=" & TextBox5.Text.Trim)
    End Sub
    '--------------------------------------------------------------------------------------------------
    '  CH Register
    '--------------------------------------------------------------------------------------------------
    Protected Sub CHRegister_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CHRegister.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("ItemRegisterCHSheet_01.aspx?pFormNo=" & DropDownList1.SelectedValue & "&pFormSno=" & TextBox2.Text.Trim & "&pStep=" & TextBox3.Text.Trim & "&pSeqNo=" & TextBox4.Text.Trim & "&pApplyID=" & TextBox5.Text.Trim & "&pAgentID=")
    End Sub

    Protected Sub CHRegister2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CHRegister2.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("ItemRegisterCHSheet_02.aspx?pFormNo=" & DropDownList1.SelectedValue & "&pFormSno=" & TextBox2.Text.Trim)
    End Sub

    Protected Sub CHRegisterHistory_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CHRegisterHistory.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("http://localhost/WorkFlow/BefOPList.aspx?pFormNo=" & DropDownList1.SelectedValue & "&pFormSno=" & TextBox2.Text.Trim & "&pStep=" & TextBox3.Text.Trim & "&pSeqNo=" & TextBox4.Text.Trim & "&pApplyID=" & TextBox5.Text.Trim)
    End Sub

    Protected Sub Inq_Commission_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Inq_Commission.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("InqCommission.aspx?pUserID=" & TextBox6.Text.Trim)
    End Sub

    Protected Sub FSLDRegister_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles FSLDRegister.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("ItemRegisterFSLDSheet_01.aspx?pFormNo=" & DropDownList1.SelectedValue & "&pFormSno=" & TextBox2.Text.Trim & "&pStep=" & TextBox3.Text.Trim & "&pSeqNo=" & TextBox4.Text.Trim & "&pApplyID=" & TextBox5.Text.Trim & "&pAgentID=")
    End Sub

    Protected Sub FSLDRegister2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles FSLDRegister2.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("ItemRegisterFSLDSheet_02.aspx?pFormNo=" & DropDownList1.SelectedValue & "&pFormSno=" & TextBox2.Text.Trim)
    End Sub

    Protected Sub FSLDRegisterHistory_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles FSLDRegisterHistory.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("http://localhost/WorkFlow/BefOPList.aspx?pFormNo=" & DropDownList1.SelectedValue & "&pFormSno=" & TextBox2.Text.Trim & "&pStep=" & TextBox3.Text.Trim & "&pSeqNo=" & TextBox4.Text.Trim & "&pApplyID=" & TextBox5.Text.Trim)
    End Sub

    Protected Sub Button18_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button18.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("PriceInforSheet_01.aspx?pFormNo=001161" & "&pFormSno=" & TextBox2.Text.Trim & "&pStep=" & TextBox3.Text.Trim & "&pSeqNo=" & TextBox4.Text.Trim & "&pApplyID=" & TextBox5.Text.Trim)

    End Sub

    Protected Sub Button19_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button19.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("PriceInforSheet_02.aspx?pFormNo=001161" & "&pFormSno=" & TextBox2.Text.Trim & "&pStep=" & TextBox3.Text.Trim & "&pSeqNo=" & TextBox4.Text.Trim & "&pApplyID=" & TextBox5.Text.Trim)

    End Sub

    Protected Sub CaseReviewSheet_01_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CaseReviewSheet_01.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("3S_CaseReviewSheet_01.aspx?pFormNo=" & DropDownList1.SelectedValue & "&pFormSno=" & TextBox2.Text.Trim & "&pStep=" & TextBox3.Text.Trim & "&pSeqNo=" & TextBox4.Text.Trim & "&pApplyID=" & TextBox5.Text.Trim & "&pAgentID=")
    End Sub

    Protected Sub CaseReviewSheet_02_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CaseReviewSheet_02.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        'Response.Redirect("3S_CaseReviewSheet_02.aspx?pFormNo=001101" & "&pFormSno=" & TextBox2.Text.Trim & "&pStep=" & TextBox3.Text.Trim & "&pSeqNo=" & TextBox4.Text.Trim & "&pApplyID=" & TextBox5.Text.Trim)
        Response.Redirect("3S_CaseReviewSheet_02.aspx?pFormNo=001101" & "&pFormSno=" & TextBox2.Text.Trim)

    End Sub

    Protected Sub CaseReviewSheet_H_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CaseReviewSheet_H.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("http://localhost/WorkFlow/BefOPList.aspx?pFormNo=" & DropDownList1.SelectedValue & "&pFormSno=" & TextBox2.Text.Trim & "&pStep=" & TextBox3.Text.Trim & "&pSeqNo=" & TextBox4.Text.Trim & "&pApplyID=" & TextBox5.Text.Trim)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Protected Sub Button1_Click1(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("NewCommission.aspx?" & "&pUserID=" & TextBox5.Text.Trim)
    End Sub
End Class
