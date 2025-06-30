<%@ Control language="vb" CodeBehind="~/admin/Skins/skin.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.Skin" %>
<%@ Register TagPrefix="dnn" TagName="USER" Src="~/Admin/Skins/User.ascx" %>
<%@ Register TagPrefix="dnn" TagName="LOGIN" Src="~/Admin/Skins/Login.ascx" %>
<%@ Register TagPrefix="dnn" TagName="CURRENTDATE" Src="~/Admin/Skins/CurrentDate.ascx" %>
<%@ Register TagPrefix="dnn" TagName="SOLPARTMENU" Src="~/Admin/Skins/SolPartMenu.ascx" %>
<%@ Register TagPrefix="dnn" TagName="BREADCRUMB" Src="~/Admin/Skins/BreadCrumb.ascx" %>
<%@ Register TagPrefix="dnn" TagName="LOGO" Src="~/Admin/Skins/Logo.ascx" %>
<%@ Register TagPrefix="dnn" TagName="PRIVACY" Src="~/Admin/Skins/Privacy.ascx" %>
<%@ Register TagPrefix="dnn" TagName="COPYRIGHT" Src="~/Admin/Skins/Copyright.ascx" %>
<%@ Register TagPrefix="dnn" TagName="TERMS" Src="~/Admin/Skins/Terms.ascx" %>
<table width="99%" height="99%"  border="0" align="center" cellpadding="0" cellspacing="1" class="TabBg">
  <tr>
    <td height="100%" valign="top" bgcolor="#FFFFFF"><TABLE WIDTH=100% height="100%" BORDER=0 align="center" CELLPADDING=0 CELLSPACING=0>
      <TR>
        <TD height="179" COLSPAN=7 valign="top" class="BannerBG"><table width="100%" height="179"  border="0" cellpadding="3" cellspacing="0" class="BannerBackground">
            <tr valign="middle">
              <td class="DateContainer"><img src="/Portals/_default/Skins/Macced/spacer.gif" width="10" height="15" hspace="0" vspace="0" align="left"><dnn:USER runat="server" id="dnnUSER" CssClass="SelectedTab" /> &nbsp;&nbsp;&nbsp;<dnn:LOGIN runat="server" id="dnnLOGIN" CssClass="SelectedTab" /></td>
              <td height="20" align="right"><span class="DateContainer"><dnn:CURRENTDATE runat="server" id="dnnCURRENTDATE" CssClass="SelectedTab" /></span></td>
              <td width="2%"><img src="/Portals/_default/Skins/Macced/spacer.gif" width="1" height="15" hspace="0" vspace="0" border="0" align="right"></td>
            </tr>
            <tr valign="top">
              <td><img src="/Portals/_default/Skins/Macced/spacer.gif" width="10" height="20" hspace="0" vspace="0" border="0" align="left"><dnn:SOLPARTMENU runat="server" id="dnnSOLPARTMENU" /></td>
              <td><dnn:BREADCRUMB runat="server" id="dnnBREADCRUMB" /></td>
              <td>&nbsp;</td>
            </tr>
            <tr>
              <td height="120" valign="middle"><dnn:LOGO runat="server" id="dnnLOGO" /></td>
              <td></td>
            </tr>
        </table>
          </TD>
      </TR>
      <TR>
        <TD height="100%" colspan="7" valign="top" class="MainContentBG">   <TABLE style="BORDER-COLLAPSE: collapse" cellPadding="0" width="100%" border="0" height="100%">
        <TR>
          <TD width="160" height="100%" vAlign="top" class="LeftContent" id="LeftPane" runat="server" visible="false"></TD>
          <TD height="100%" align="middle" vAlign="top" class="ContentPane" id="ContentPane" runat="server" visible="false" ></TD>
          <TD width="160" height="100%" vAlign="top" class="RightPane" id="RightPane" runat="server" visible="false"></TD>
        </TR>
      </TABLE>     </TD>
      </TR>
      <TR>
        <TD COLSPAN=7 valign="top" class="FooterBG"><table width="100%" height="25"  border="0" cellpadding="1" cellspacing="1" class="Footer">
            <tr valign="top">
              <td align="center"><dnn:PRIVACY runat="server" id="dnnPRIVACY" />&nbsp;&nbsp;&nbsp;<dnn:COPYRIGHT runat="server" id="dnnCOPYRIGHT" CssClass="OtherTabs" /> &nbsp;&nbsp;&nbsp;<dnn:TERMS runat="server" id="dnnTERMS" /></td>
              </tr>
        </table></TD>
      </TR>
    </TABLE></td>
  </tr>
</table>

