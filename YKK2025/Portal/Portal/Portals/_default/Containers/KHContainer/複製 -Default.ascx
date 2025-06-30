<%@ Control language="vb" CodeBehind="~/admin/Containers/container.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.UI.Containers.Container" %>
<%@ Register TagPrefix="dnn" TagName="ACTIONS" Src="~/Admin/Containers/SolPartActions.ascx" %>
<%@ Register TagPrefix="dnn" TagName="TITLE" Src="~/Admin/Containers/Title.ascx" %>

<TABLE CELLPADDING="0" CELLSPACING="0" WIDTH="100%" CLASS="ContainerTable">
    <TR>
	<TD  BGCOLOR="" VISIBLE="False" WIDTH=50><dnn:ACTIONS runat="server" id="dnnACTIONS" /></TD>
<!--	
	<TD  BGCOLOR="#B5A9AD" ALIGN="LEFT" WIDTH="100%" VALIGN="MIDDLE" CLASS="Title" NOWRAP><IMG SRC="<%= SkinPath %>GreyTrans.gif" ALIGN="TEXTTOP" NOWRAP><dnn:TITLE runat="server" id="dnnTITLE" CssClass="TITLE" /></TD>
---->	
	<TD  BGCOLOR="" ALIGN="LEFT" WIDTH="100%" VALIGN="MIDDLE" CLASS="Title" NOWRAP></TD>
    </TR>
    <TR>
        
    <TD ID="ContentPane" RUNAT="server" CLASS="ContainerPane" COLSPAN="2" VALIGN="TOP"></TD>
    </TR>
</TABLE>
<IMG SRC="<%= SkinPath %>1x1.GIF" HEIGHT="3">