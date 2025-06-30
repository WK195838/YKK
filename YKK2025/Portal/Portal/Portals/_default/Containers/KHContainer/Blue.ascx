<%@ Control language="vb" CodeBehind="~/admin/Containers/container.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.UI.Containers.Container" %>
<%@ Register TagPrefix="dnn" TagName="ACTIONS" Src="~/Admin/Containers/SolPartActions.ascx" %>
<%@ Register TagPrefix="dnn" TagName="TITLE" Src="~/Admin/Containers/Title.ascx" %>
<TABLE CELLPADDING="0" CELLSPACING="0" WIDTH="100%" CLASS="ContainerTable">
	
    <TR>
	<TD  VISIBLE="False" WIDTH=50><dnn:ACTIONS runat="server" id="dnnACTIONS" /></TD>
	<TD  BGCOLOR="#9ABEEC" ALIGN="LEFT" WIDTH="100%" VALIGN="MIDDLE" NOWRAP><IMG SRC="<%= SkinPath %>BlueTrans.gif" ALIGN="TEXTTOP"> <dnn:TITLE runat="server" id="dnnTITLE" CssClass="TITLE" /></TD>
    </TR>
    <TR>
        
    <TD ID="ContentPane" RUNAT="server" CLASS="ContainerPane2" COLSPAN="2" VALIGN="TOP"></TD>
    </TR>
</TABLE>
<IMG SRC="<%= SkinPath %>1x1.GIF" HEIGHT="3">
