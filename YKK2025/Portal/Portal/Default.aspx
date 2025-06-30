<%@ Page CodeBehind="Default.aspx.vb" language="vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.Framework.CDefault" %>
<%@ Register TagPrefix="dnn" Namespace="DotNetNuke.Common.Controls" Assembly="DotNetNuke" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD id="Head">
		<TITLE>
			<%= Title %>
		</TITLE>
		<%= Comment %>
		<META NAME="DESCRIPTION" CONTENT="<%= Description %>">
		<META NAME="KEYWORDS" CONTENT="<%= KeyWords %>">
		<META NAME="COPYRIGHT" CONTENT="<%= Copyright %>">
		<META NAME="GENERATOR" CONTENT="<%= Generator %>">
		<META NAME="AUTHOR" CONTENT="<%= Author %>">
		<META NAME="RESOURCE-TYPE" CONTENT="DOCUMENT">
		<META NAME="DISTRIBUTION" CONTENT="GLOBAL">
		<META NAME="ROBOTS" CONTENT="INDEX, FOLLOW">
		<META NAME="REVISIT-AFTER" CONTENT="1 DAYS">
		<META NAME="RATING" CONTENT="GENERAL">
		<META HTTP-EQUIV="PAGE-ENTER" CONTENT="RevealTrans(Duration=0,Transition=1)">
		<style id="StylePlaceholder" runat="server"></style>
		<asp:placeholder id="CSS" runat="server"></asp:placeholder>
		<asp:placeholder id="FAVICON" runat="server"></asp:placeholder>
		<script src="<%= Page.ResolveUrl("js/dnncore.js") %>"></script>
		<asp:placeholder id="phDNNHead" runat="server"></asp:placeholder>
	</HEAD>
	<BODY ID="Body" runat="server" ONSCROLL="__dnn_bodyscroll()" BOTTOMMARGIN="0" LEFTMARGIN="0"
		TOPMARGIN="0" RIGHTMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<noscript></noscript>
		<dnn:Form id="Form" runat="server" ENCTYPE="multipart/form-data" style=”height:100%;>
			<asp:Label ID="SkinError" Runat="server" CssClass="NormalRed" Visible="False"></asp:Label>
			<asp:placeholder id="SkinPlaceHolder" runat="server" />
			<INPUT ID="ScrollTop" runat="server" NAME="ScrollTop" TYPE="hidden"> 
			<INPUT ID="__dnnVariable" runat="server" NAME="__dnnVariable" TYPE="hidden">
		</dnn:Form>
	</BODY>
</HTML>
