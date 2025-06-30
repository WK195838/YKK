<%@ Control language="vb" CodeBehind="~/admin/Containers/container.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.UI.Containers.Container" %>
<%@ Register TagPrefix="dnn" TagName="ACTIONS" Src="~/Admin/Containers/SolPartActions.ascx" %>
<%@ Register TagPrefix="dnn" TagName="TITLE1" Src="~/Admin/Containers/Title.ascx" %>
<TABLE CELLPADDING="0" CELLSPACING="0" WIDTH="100%" CLASS="ContainerTable">
	
    <TR>
	<TD  BACKGROUND="<%= SkinPath %>Split.gif" VISIBLE="False" ><dnn:ACTIONS runat="server" id="dnnACTIONS" /></TD>
	<TD  BACKGROUND="<%= SkinPath %>Split.gif" ALIGN="LEFT" WIDTH="100%" VALIGN="MIDDLE" CLASS="Title1" NOWRAP><dnn:TITLE1 runat="server" id="dnnTITLE1" CssClass="TITLE1" /></TD>
    </TR>
    <TR>
        
    <TD ID="ContentPane" RUNAT="server" CLASS="ContainerPane" COLSPAN="2" VALIGN="TOP"></TD>
    </TR>
</TABLE>
<IMG SRC="<%= SkinPath %>1x1.GIF" HEIGHT="3">
