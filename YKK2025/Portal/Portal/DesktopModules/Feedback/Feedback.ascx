<%@ Control Language="vb" AutoEventWireup="false" Explicit="True" Codebehind="Feedback.ascx.vb" Inherits="DotNetNuke.Modules.Feedback.Feedback" %>
<%@ Register TagPrefix="dnn" TagName="Label" Src="~/controls/LabelControl.ascx" %>
<table cellspacing="0" cellpadding="4" border="0" summary="Feedback Design Table">
  <tr valign="top">
    <td class="SubHead">
      <dnn:label id="plEmail" runat="server" controlname="txtEmail" suffix=":"></dnn:label>
      <asp:textbox id="txtEmail" runat="server" width="150px" cssclass="NormalTextBox" columns="35" maxlength="100"></asp:textbox>
    </td>
  </tr>
  <tr valign="top">
    <td class="SubHead">
      <dnn:label id="plName" runat="server" controlname="txtName" suffix=":"></dnn:label>
      <asp:textbox id="txtName" runat="server" width="150px" cssclass="NormalTextBox" columns="35" maxlength="100"></asp:textbox>
    </td>
  </tr>
  <tr valign="top">
    <td class="SubHead">
      <dnn:label id="plSubject" runat="server" controlname="txtSubject" suffix=":"></dnn:label>
      <asp:textbox id="txtSubject" runat="server" width="150px" cssclass="NormalTextBox" columns="35" maxlength="100"></asp:textbox>
    </td>
  </tr>
  <tr valign="top">
    <td class="SubHead">
      <dnn:label id="plBody" runat="server" controlname="txtBody" suffix=":"></dnn:label>
      <asp:textbox id="txtBody" runat="server" width="150px" columns="25" textmode="Multiline" rows="10" cssclass="NormalTextBox"></asp:textbox>
    </td>
  </tr>
  <tr valign="top">
    <td class="SubHead" nowrap>
      <asp:CheckBox id="chkCopy" Runat="server" cssclass="NormalTextBox"></asp:CheckBox><dnn:label id="plCopy" runat="server" controlname="chkCopy" suffix="?"></dnn:label>
    </td>
  </tr>
  <tr valign="top">
    <td align="middle">
      <asp:linkbutton id="cmdCancel" resourcekey="cmdCancel" runat="server" cssclass="CommandButton" causesvalidation="False">Cancel</asp:linkbutton>
      &nbsp;
      <asp:linkbutton id="cmdSend" resourcekey="cmdSend" runat="server" cssclass="CommandButton" causesvalidation="True">Send</asp:linkbutton>
    </td>
  </tr>
  <tr valign="top">
    <td align="middle" colspan="2">
      <asp:label id="lblMessage" runat="server" cssclass="NormalRed"></asp:label>
    </td>
  </tr>
</table>