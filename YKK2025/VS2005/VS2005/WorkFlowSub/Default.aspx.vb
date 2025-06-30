Imports System.Data
Imports System.Data.OleDb

Partial Class _Default
    Inherits System.Web.UI.Page

    '*****************************************************************
    '**
    '**     自訂涵式庫
    '**
    '*****************************************************************
    Public YKK As New YKK_SPDClass   'YKK SPD系共通涵式

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    End Sub


    Protected Sub CaseReviewSheet_01_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CaseReviewSheet_01.Click
        Response.Cookies("UserID").Value = TextBox6.Text
        Response.Redirect("3S_CaseReviewSheet_01.aspx?pFormNo=" & DropDownList1.SelectedValue & "&pFormSno=" & TextBox2.Text.Trim & "&pStep=" & TextBox3.Text.Trim & "&pSeqNo=" & TextBox4.Text.Trim & "&pApplyID=" & TextBox5.Text.Trim & "&pAgentID=")
    End Sub
End Class
