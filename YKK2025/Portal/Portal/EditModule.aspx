<%@ Page language="VB" %>
<%@ Import Namespace="DotNetNuke.Framework" %>

<script runat="server">

    Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

	Dim strURL As String = Request.Url.ToString().ToLower
	strURL = strURL.Replace("editmodule.aspx", glbDefaultPage)            
	Response.Redirect(strURL, True)
	
    End Sub

</script>
