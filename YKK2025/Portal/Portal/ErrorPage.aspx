<%@ Page CodeBehind="ErrorPage.aspx.vb" language="vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.Services.Exceptions.ErrorPage" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML LANG="en-US">
    <HEAD>
        <TITLE runat="server" id="Title">Error</TITLE> 
        <LINK id="StyleSheet" runat="server" href="portal.css" type="text/css" rel="stylesheet"></LINK>
        <script src="js/dnncore.js"></script>
    </HEAD>
    <BODY id="Body" runat="server" onscroll="__dnn_bodyscroll()" bottommargin="0" leftmargin="0"
        topmargin="0" rightmargin="0" marginwidth="0" marginheight="0">
        <FORM id="Form" runat="server" enctype="multipart/form-data">
            <ASP:PLACEHOLDER id="ErrorPlaceHolder" runat="server" />
        </FORM>
    </BODY>
</HTML>
