<%@ Control language="vb" CodeBehind="~/admin/Skins/skin.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.UI.Skins.Skin" %>
<%@ Register TagPrefix="dnn" TagName="LOGO" Src="~/Admin/Skins/Logo.ascx" %>
<%@ Register TagPrefix="dnn" TagName="SEARCH" Src="~/Admin/Skins/Search.ascx" %>
<%@ Register TagPrefix="dnn" TagName="USER" Src="~/Admin/Skins/User.ascx" %>
<%@ Register TagPrefix="dnn" TagName="LOGIN" Src="~/Admin/Skins/Login.ascx" %>
<%@ Register TagPrefix="dnn" TagName="MENU" Src="~/Admin/Skins/SolPartMenu.ascx" %>
<%@ Register TagPrefix="dnn" TagName="COPYRIGHT" Src="~/Admin/Skins/Copyright.ascx" %>
<table class="body" border="0" align="center" cellpadding="0" cellspacing="0" width="100%">
  <tr>
    <td><table width="875" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td><table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td id="logo"><dnn:LOGO runat="server" id="dnnLOGO" /></td>
                <td rowspan="2" align="right" valign="bottom"><table border="0" cellspacing="0" cellpadding="0" id="search">
                    <tr>
                      <td class="searchbox"><dnn:SEARCH runat="server" id="dnnSEARCH" /><img src="<%= SkinPath %>images/icon02.gif" width="13" height="13" class="icon-deco"></td>
                    </tr>
                  </table>
                  <table border="0" cellspacing="0" cellpadding="0" id="nav">
                    <tr>
                      <td><img src="<%= SkinPath %>images/tab_nav01.gif" width="27" height="21"></td>
                      <td class="navbox" nowrap><img src="<%= SkinPath %>images/icon01.gif" width="13" height="13" class="icon-deco"><dnn:USER runat="server" id="dnnUSER" /> 
                        <img src="<%= SkinPath %>images/icon01.gif" width="13" height="13" class="icon-deco"><dnn:LOGIN runat="server" id="dnnLOGIN" /></td>
                    </tr>
                  </table></td>
              </tr>
              <tr>
                <td id="navi"><table height="30" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td class="navibox" align="left"><dnn:MENU runat="server" id="dnnMENU" UseRootBreadcrumbArrow="false" LeftSeparator="&nbsp;" RightSeparator="&nbsp;" LeftSeparatorBreadcrumbCssClass="MainMenu_LeftOn" RightSeparatorBreadcrumbCssClass="MainMenu_RightOn" LeftSeparatorCssClass="MainMenu_Left" RightSeparatorCssClass="MainMenu_Right" RootMenuItemBreadcrumbCssClass="MainMenu_Active" RootMenuItemSelectedCssClass="MainMenu_BreadcrumbActive2" /></td>
                      <td>&nbsp;</td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
        </tr>
      </table>
      <table width="875" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td width="210" valign="top" id="LeftPane" class="LeftPane" runat="server"></td>
          <td width="450" valign="top" id="ContentPane" class="ContentPane" runat="server"></td>
          <td width="210" valign="top" id="RightPane" class="RightPane" runat="server"></td>
        </tr>
      </table>
      <div id="foot">
        <div class="copyr">
          <p>
          <table border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td valign="middle" align="center" id="BottomPane" class="BottomPane" runat="server"></td>
            </tr>
          </table>
          <dnn:COPYRIGHT runat="server" id="dnnCOPYRIGHT" /></p></div>
      </div></td>
  </tr>
</table>

