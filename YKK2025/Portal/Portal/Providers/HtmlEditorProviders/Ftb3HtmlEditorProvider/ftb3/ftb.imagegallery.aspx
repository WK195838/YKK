<%@ Page Language="vb" ValidateRequest="false" Trace="false" CodeBehind="ftb.imagegallery.aspx.vb" AutoEventWireup="false" Inherits="DotNetNuke.HtmlEditor.FTBImageGallery" %>
<%@ Register TagPrefix="FTB" Namespace="FreeTextBoxControls" Assembly="FreeTextBox" %>
<html>
<head>
	<title>
		<%= Title %>
	</title>
</head>
<body>
    <form id="Form1" runat="server" enctype="multipart/form-data">  
		<FTB:ImageGallery id="imgGallery" runat="Server" />
	</form>
</body>
</html>
