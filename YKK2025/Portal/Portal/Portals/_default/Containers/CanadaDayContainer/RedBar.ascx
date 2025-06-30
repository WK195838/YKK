<%@ Control language="vb" CodeBehind="~/admin/Containers/container.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.Container" %>
<%@ Register TagPrefix="dnn" TagName="ACTIONS" Src="~/Admin/Containers/SolPartActions.ascx" %>
<%@ Register TagPrefix="dnn" TagName="TITLE" Src="~/Admin/Containers/Title.ascx" %>
<!-------------------------------------Red Bar Container----------------------------------->
<!---------------------------------------Created By Joie------------------------------------>
<!--------------------------------------portals.my-asp.net--------------------------------------->

<TABLE CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
    <TR>
        <TD HEIGHT="1" COLSPAN="2"><IMG SRC="/Portals/_default/Containers/CanadaDayContainer/spacer.gif" WIDTH="200" HEIGHT="1"></TD>
    </TR>	
    <TR>
	<TD BGCOLOR="#DE2918" VISIBLE="False" HEIGHT="25"><dnn:ACTIONS runat="server" id="dnnACTIONS" /></TD>
	<TD BGCOLOR="#DE2918" ALIGN="LEFT" HEIGHT="25"  WIDTH="100%"><dnn:TITLE runat="server" id="dnnTITLE" CSSClass="Title" /></TD>
    </TR>
    <TR>
        <TD ID="ContentPane" RUNAT="server" STYLE="PADDING:7;" COLSPAN="2"></TD>
    </TR>
</TABLE>

