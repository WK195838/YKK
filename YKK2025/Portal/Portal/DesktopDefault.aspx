<%@ Page language="VB" %>
<%@ Import Namespace="DotNetNuke.Common" %>

<script runat="server">

    Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

	Dim strURL As String = Request.Url.ToString().ToLower
	strURL = strURL.Replace("desktopdefault.aspx", glbDefaultPage)            
	Response.Redirect(strURL, True)
	

    End Sub

</script>