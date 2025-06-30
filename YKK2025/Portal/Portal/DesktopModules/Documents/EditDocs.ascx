<%@ Control language="vb" CodeBehind="EditDocs.ascx.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.Modules.Documents.EditDocs" %>
<%@ Register TagPrefix="Portal" TagName="URL" Src="~/controls/URLControl.ascx" %>
<%@ Register TagPrefix="Portal" TagName="Audit" Src="~/controls/ModuleAuditControl.ascx" %>
<%@ Register TagPrefix="Portal" TagName="Tracking" Src="~/controls/URLTrackingControl.ascx" %>
<%@ Register TagPrefix="dnn" TagName="Label" Src="~/controls/LabelControl.ascx" %>
<table width="560" cellspacing="0" cellpadding="0" border="0" summary="Edit Documents Design Table">
  <tr valign="top">
    <td class="SubHead" width="125"><dnn:label id="plName" runat="server" controlname="txtName" suffix=":"></dnn:label></td>
    <td width="400">
      <asp:textbox id="txtName" cssclass="NormalTextBox" width="300" maxlength="150" runat="server" />
      <asp:requiredfieldvalidator resourcekey="Name.ErrorMessage" display="Dynamic" runat="server" errormessage="<br>You Must Enter A Title For The Document" controltovalidate="txtName" id="valName" cssclass="NormalRed" />
    </td>
  </tr>
  <tr>
    <td class="SubHead" width="125"><dnn:label id="plURL" runat="server" controlname="ctlURL" suffix=":"></dnn:label></td>
    <td width="400">
      <portal:url id="ctlURL" runat="server" width="275" showtabs="False" urltype="F" shownewwindow="True" />
    </td>
  </tr>
  <tr height="5"><td></td></tr>
  <tr valign="top">
    <td class="SubHead" width="125"><dnn:label id="plCategory" runat="server" controlname="txtCategory" suffix=":"></dnn:label></td>
    <td width="400">
      <asp:textbox id="txtCategory" cssclass="NormalTextBox" width="300" maxlength="50" runat="server" />
    </td>
  </tr>
</table>
<p>
  <asp:linkbutton id="cmdUpdate" resourcekey="cmdUpdate" runat="server" cssclass="CommandButton" text="Update" borderstyle="none"></asp:linkbutton>&nbsp;
  <asp:linkbutton id="cmdCancel" resourcekey="cmdCancel" runat="server" cssclass="CommandButton" text="Cancel" borderstyle="none" causesvalidation="False"></asp:linkbutton>&nbsp;
  <asp:linkbutton id="cmdDelete" resourcekey="cmdDelete" runat="server" cssclass="CommandButton" text="Delete" borderstyle="none" causesvalidation="False"></asp:linkbutton>
</p>
<portal:audit id="ctlAudit" runat="server" />
<br><br>
<portal:tracking id="ctlTracking" runat="server" />

