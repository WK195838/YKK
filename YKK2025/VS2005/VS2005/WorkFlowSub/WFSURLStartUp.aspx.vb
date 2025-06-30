Imports System.Diagnostics

Partial Class WFSURLStartUp
    Inherits System.Web.UI.Page

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '** Start Up YKK-WFS URL
    '*****************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim xURL As String = UCase(Request.QueryString("pURL"))
        If xURL <> "" Then
            xURL = Replace(xURL, "%20", "")
            xURL = Replace(xURL, "%22", "")
            xURL = Replace(xURL, "YKK-WFS://", "HTTP://")

            Response.Write("<script>" & "document.location.href='" & xURL & "'" & ";</script>")
        End If
    End Sub

End Class
