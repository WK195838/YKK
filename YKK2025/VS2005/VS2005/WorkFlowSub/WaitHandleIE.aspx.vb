Imports System.Diagnostics

Partial Class WaitHandleIE
    Inherits System.Web.UI.Page
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '*****************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '
        'MsgBox(Request.QueryString("pURL"))
        '
        If Request.QueryString("pURL") <> "" Then
            Dim someScript As String = ""
            someScript = "<script language='javascript'>" & "StartUpPage('" & Request.QueryString("pURL") & "')" & ";</script>"
            Page.ClientScript.RegisterStartupScript(Me.GetType(), "onload", someScript)
        End If
        '

    End Sub
End Class
