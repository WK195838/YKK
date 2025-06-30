<%@ Control language="vb" CodeBehind="Settings.ascx.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.Modules.Feedback.Settings" %>
<%@ Register TagPrefix="dnn" TagName="Label" Src="~/controls/LabelControl.ascx" %>
<table cellspacing="0" cellpadding="2" summary="Feedback Settings Design Table" border="0">
  <tr valign="top">
    <td class="SubHead" width="175"><dnn:label id="plSendTo" runat="server" controlname="txtSendTo" suffix=":"></dnn:label></td>
    <td valign="bottom" width="325"><asp:textbox id="txtSendTo" runat="server" width="300px" cssclass="NormalTextBox" columns="35" maxlength="100"></asp:textbox>
			<asp:regularexpressionvalidator id="valSendTo" resourcekey="valSendTo" runat="server" cssclass="NormalRed" controltovalidate="txtSendTo"
				errormessage="<br>Email Must be Valid" validationexpression="[\w\.-]+(\+[\w-]*)?@([\w-]+\.)+[\w-]+" display="Dynamic"></asp:regularexpressionvalidator>
    </td>
  </tr>
  <tr valign="top">
    <td class="SubHead" width="175"><dnn:label id="plWidth" runat="server" controlname="txtWidth" suffix=":"></dnn:label></td>
    <td valign="bottom" width="325"><asp:textbox id="txtWidth" runat="server" width="300px" cssclass="NormalTextBox" columns="35" maxlength="100"></asp:textbox>
	<asp:rangevalidator id="valWidth" resourcekey="valWidth.ErrorMessage" controltovalidate="txtWidth" minimumvalue="0"  maximumvalue="32768" type="Integer"
		display="Dynamic" cssclass="NormalRed" errormessage="<br>Width Must Be A Valid Integer" runat="server" />
    </td>
  </tr>
  <tr valign="top">
    <td class="SubHead" width="175"><dnn:label id="plRows" runat="server" controlname="txtRows" suffix=":"></dnn:label></td>
    <td valign="bottom" width="325"><asp:textbox id="txtRows" runat="server" width="300px"  cssclass="NormalTextBox" columns="35" maxlength="100"></asp:textbox>
	<asp:regularexpressionvalidator id="valRows" resourcekey="valRows.ErrorMessage" controltovalidate="txtRows"
		validationexpression="^[1-9]+[0-9]*$" display="Dynamic" cssclass="NormalRed" errormessage="<br>Rows Must Be A Valid Integer"	runat="server" />
    </td>
  </tr>
</table>

