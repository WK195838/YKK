<%@ Control language="vb" CodeBehind="~/admin/Containers/container.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.UI.Containers.Container" %>
<%@ Register TagPrefix="dnn" TagName="ACTIONS" Src="~/Admin/Containers/SolPartActions.ascx" %>
<%@ Register TagPrefix="dnn" TagName="TITLE2" Src="~/Admin/Containers/Title.ascx" %>
<TABLE CELLPADDING="0" CELLSPACING="0" WIDTH="100%" CLASS="ContainerTable">
	
    <TR>
	<TD  BACKGROUND="<%= SkinPath %>BlueSplit.gif" VISIBLE="False" ><dnn:ACTIONS runat="server" id="dnnACTIONS" /></TD>
	<TD  BACKGROUND="<%= SkinPath %>BlueSplit.gif" ALIGN="LEFT" WIDTH="100%" VALIGN="MIDDLE" CLASS="Title2" NOWRAP><dnn:TITLE2 runat="server" id="dnnTITLE2" CssClass="TITLE2" /></TD>
    </TR>
    <TR>
        
    <TD ID="ContentPane" RUNAT="server" CLASS="ContainerPane2" COLSPAN="2" VALIGN="TOP"></TD>
    </TR>
</TABLE>
<IMG SRC="<%= SkinPath %>1x1.GIF" HEIGHT="3">
