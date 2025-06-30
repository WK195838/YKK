
Partial Class RedirectURL
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim Http As String = System.Configuration.ConfigurationManager.AppSettings("Http")  'Http Address
        Response.Redirect(Http + "/WorkFlow/WaitHandle.aspx?pUserID=" + Request.QueryString("pUserID"))

    End Sub
End Class
