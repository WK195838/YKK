<%@ Control language="vb" CodeBehind="~/admin/Skins/skin.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.UI.Skins.Skin" %>
<%@ Register TagPrefix="dnn" TagName="LOGIN" Src="~/Admin/Skins/Login.ascx" %>
<%@ Register TagPrefix="dnn" TagName="USER" Src="~/Admin/Skins/User.ascx" %>
<%@ Register TagPrefix="dnn" TagName="LOGO" Src="~/Admin/Skins/Logo.ascx" %>
<%@ Register TagPrefix="dnn" TagName="HELP" Src="~/Admin/Skins/Help.ascx" %>
<%@ Register TagPrefix="dnn" TagName="HOSTNAME" Src="~/Admin/Skins/HostName.ascx" %>
<%@ Register TagPrefix="dnn" TagName="MENU" Src="~/Admin/Skins/SolPartMenu.ascx" %>
<%@ Register TagPrefix="dnn" TagName="CURRENTDATE" Src="~/Admin/Skins/CurrentDate.ascx" %>
<%@ Register TagPrefix="dnn" TagName="BREADCRUMB" Src="~/Admin/Skins/BreadCrumb.ascx" %>
<%@ Register TagPrefix="dnn" TagName="COPYRIGHT" Src="~/Admin/Skins/Copyright.ascx" %>
<%@ Register TagPrefix="dnn" TagName="TERMS" Src="~/Admin/Skins/Terms.ascx" %>
<%@ Register TagPrefix="dnn" TagName="PRIVACY" Src="~/Admin/Skins/Privacy.ascx" %>
<!-- Begin Table Head -->
<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
<TABLE BGCOLOR="#FFFFFF" BORDER="0" CELLPADDING="0" CELLSPACING="0" ALIGN="LEFT">
  <TR> 
    <TD BACKGROUND="<%= SkinPath %>Image44x1.png" VALIGN="TOP" ROWSPAN="3" COLSPAN="1" WIDTH="43" HEIGHT="83" noWrap><IMG SRC="<%= SkinPath %>Image41x1.png" WIDTH="43" HEIGHT="83" BORDER="0"></TD>
    <TD VALIGN="TOP" ROWSPAN="3" COLSPAN="1" WIDTH="18" HEIGHT="83" noWrap><IMG SRC="<%= SkinPath %>Image41x2.png" WIDTH="18" HEIGHT="83" BORDER="0"></TD>
    <TD VALIGN="TOP" ROWSPAN="1" COLSPAN="1" WIDTH="716" HEIGHT="20" noWrap><IMG SRC="<%= SkinPath %>Image41x3.png" WIDTH="716" HEIGHT="20" BORDER="0"></TD>
    <TD ALIGN="RIGHT" ROWSPAN="1" COLSPAN="1" WIDTH="100%" HEIGHT="20" nowrap><FONT Style="FONT-SIZE:8pt" FACE="Verdana" COLOR="#000000"><a href="mailto:joo@ykk.com.tw"><b>Help</b></a></FONT></TD>
  </TR>
  <TR> 
    <TD VALIGN="TOP" ROWSPAN="1" HEIGHT="41" noWrap><IMG SRC="<%= SkinPath %>Image42x1.png" WIDTH="716" HEIGHT="41" BORDER="0"></TD>
    <TD ALIGN="RIGHT" ROWSPAN="1" HEIGHT="41" WIDTH="100%" noWrap><FONT Style="FONT-SIZE:8pt" FACE="Verdana" COLOR="#000000"><dnn:LOGIN runat="server" id="dnnLOGIN" /> / <dnn:USER runat="server" id="dnnUSER" /></FONT></TD>
  </TR>
  <TR> 
    <TD VALIGN="TOP" ROWSPAN="1" COLSPAN="3" HEIGHT="13"> 
      <TABLE WIDTH="100%" HEIGHT="100%" BORDER="0" VALIGN="TOP">
        <TR VALIGN="TOP"> 
          <TD WIDTH="300" HEIGHT="13"><FONT Style="FONT-SIZE:8pt" FACE="Verdana" COLOR="#000000"></TD>
          <TD COLSPAN="2" Align=Left><FONT STYLE="FONT-SIZE:8pt" FACE="Verdana" COLOR="#000000"><dnn:MENU runat="server" id="dnnMENU" /></FONT></TD>
        </TR>
        <TR VALIGN="TOP"> 
          <TD WIDTH="300" HEIGHT="13"><FONT Style="FONT-SIZE:8pt" FACE="Verdana" COLOR="#000000"><dnn:CURRENTDATE runat="server" id="dnnCURRENTDATE" /></FONT></TD>
          <TD WIDTH="50" >&nbsp;</TD>
          <TD Align=Right><FONT Style="FONT-SIZE:8pt" FACE="Verdana" COLOR="#000000"><dnn:LOGO runat="server" id="dnnLOGO" /></FONT></TD>
        </TR>
      </TABLE>
    </TD>
  </TR>
  <TR> 
    <TD BACKGROUND="<%= SkinPath %>Image44x1.png" ROWSPAN="1" COLSPAN="1" WIDTH="43" HEIGHT="100%"> 
    </TD>
    <TD vAlign="TOP" ROWSPAN="1" COLSPAN="4" HEIGHT="100%"> 
      <TABLE WIDTH="100%" HEIGHT="100%" BORDER="0">
        <TR> 
          <TD ID="LeftPane" runat="server" visible="false" vAlign="top" noWrap width="150">&nbsp;</TD>
          <TD ID="ContentPane" runat="server" visible="false" vAlign="top">&nbsp;</TD>
          <TD ID="RightPane" runat="server" visible="false" vAlign="top" noWrap width="150">&nbsp;</TD>
        </TR>
      </TABLE>
    </TD>
  </TR>
  <TR> 
    <TD BACKGROUND="<%= SkinPath %>Image44x1.png" ROWSPAN="1" COLSPAN="1" WIDTH="43" HEIGHT="100%"> 
    </TD>
    <TD style="BORDER-TOP: blue 1px solid" vAlign="top" HEIGHT="20px" ROWSPAN="1" COLSPAN="4"> 
      <TABLE WIDTH="100%" HEIGHT="20px" BORDER="0">
         <TR> 
          <TD vAlign="top" Align="Center" noWrap><FONT Style="FONT-SIZE:7pt" FACE="Verdana" COLOR="#000000"><dnn:COPYRIGHT runat="server" id="dnnCOPYRIGHT" /></FONT></TD>
        </TR>
       <TR> 
          <TD vAlign="top" Align="Center" noWrap><FONT Style="FONT-SIZE:7pt" FACE="Verdana" COLOR="#000000"></FONT></TD>
        </TR>
      </TABLE>
    </TD>
  </TR>
</TABLE>
<!-- End Table Head-->

