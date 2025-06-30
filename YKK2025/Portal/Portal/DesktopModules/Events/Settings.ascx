<%@ Register TagPrefix="dnn" TagName="Label" Src="~/controls/LabelControl.ascx" %>
<%@ Control language="vb" CodeBehind="Settings.ascx.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.Modules.Events.Settings" %>
<table cellspacing="0" cellpadding="2" summary="Edit Events Design Table" border="0">
  <tr>
    <td class="SubHead" width="175"><dnn:label id="plView" runat="server" controlname="optView" suffix=":"></dnn:label></td>
    <td valign="bottom">
      <asp:radiobuttonlist id="optView" runat="server" repeatdirection="Horizontal" cssclass="NormalTextBox">
        <asp:listitem resourcekey="List" value="L">List</asp:listitem>
        <asp:listitem resourcekey="Calendar" value="C">Calendar</asp:listitem>
      </asp:radiobuttonlist>
    </td>
  </tr>
  <tr valign="top">
    <td class="SubHead" width="175"><dnn:label id="plWidth" runat="server" controlname="txtWidth" suffix=":"></dnn:label></td>
    <td valign="bottom"><asp:textbox id="txtWidth" runat="server" cssclass="NormalTextBox" columns="5"></asp:textbox>
			<asp:CompareValidator id="valWidth1" runat="server" ErrorMessage="<br>Width must be a number greater than zero"
				CssClass="NormalRed" Operator="DataTypeCheck" Type="Integer" ControlToValidate="txtWidth" Display="Dynamic" resourcekey="valWidth"></asp:CompareValidator>
			<asp:comparevalidator id="valWidth2" runat="server" cssclass="NormalRed" resourcekey="valWidth"
				valuetocompare="0" operator="GreaterThan" display="Dynamic" errormessage="<br>Width must be a number greater than zero"
				controltovalidate="txtWidth"></asp:comparevalidator></td>
  </tr>
  <tr valign="top">
    <td class="SubHead" width="175"><dnn:label id="plHeight" runat="server" controlname="txtHeight" suffix=":"></dnn:label></td>
    <td valign="bottom"><asp:textbox id="txtHeight" runat="server" cssclass="NormalTextBox" columns="5"></asp:textbox>
			<asp:CompareValidator id="valHeight1" runat="server" ErrorMessage="<br>Height must be a number greater than zero"
				CssClass="NormalRed" Operator="DataTypeCheck" Type="Integer" ControlToValidate="txtHeight" Display="Dynamic" resourcekey="valHeight"></asp:CompareValidator>
			<asp:comparevalidator id="valHeight2" runat="server" cssclass="NormalRed" resourcekey="valHeight"
				valuetocompare="0" operator="GreaterThan" display="Dynamic" errormessage="<br>Height must be a number greater than zero"
				controltovalidate="txtHeight"></asp:comparevalidator></td>
  </tr>
</table>
