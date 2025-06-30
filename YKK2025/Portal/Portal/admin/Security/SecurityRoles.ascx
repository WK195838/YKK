<%@ Register TagPrefix="dnn" TagName="Label" Src="~/controls/LabelControl.ascx" %>
<%@ Control language="vb" CodeBehind="SecurityRoles.ascx.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.Modules.Admin.Security.SecurityRoles" %>
<table class="Settings" cellspacing="2" cellpadding="2" summary="Security Roles Design Table" border="0">
	<tr>
		<td width="650" valign="top">
			<asp:panel id="pnlRoles" runat="server" cssclass="WorkPanel" visible="True">
      <TABLE cellSpacing=4 cellPadding=0 
        summary="Security Roles Design Table">
        <TR>
          <TD colSpan=5>
<asp:Label id=lblTitle Runat="server" cssClass="Head"></asp:Label></TD></TR>
        <TR>
          <TD height=10></TD></TR>
        <TR>
          <TD class=SubHead vAlign=top>
<dnn:label id=plUsers runat="server" suffix="" controlname="cboUsers"></dnn:label>
<dnn:label id=plRoles runat="server" suffix="" controlname="cboRoles"></dnn:label></TD>
          <TD width=20>&nbsp;</TD>
          <TD class=SubHead vAlign=top>
<dnn:label id=plExpiryDate runat="server" suffix="" controlname="txtExpiryDate"></dnn:label></TD>
          <TD class=SubHead vAlign=top>
<asp:hyperlink id=cmdExpiryCalendar cssclass="CommandButton" runat="server" resourcekey="Calendar">Calendar</asp:hyperlink></TD>
          <TD class=SubHead vAlign=top>&nbsp;</TD></TR>
        <TR>
          <TD vAlign=top>
<asp:dropdownlist id=cboUsers cssclass="NormalTextBox" runat="server" autopostback="True" datavaluefield="UserID" datatextfield="FullName" width="200"></asp:dropdownlist>
<asp:dropdownlist id=cboRoles cssclass="NormalTextBox" runat="server" autopostback="True" datavaluefield="RoleID" datatextfield="RoleName" width="200"></asp:dropdownlist></TD>
          <TD width=20>&nbsp;</TD>
          <TD vAlign=top colSpan=2>
<asp:textbox id=txtExpiryDate cssclass="NormalTextBox" runat="server" width="200"></asp:textbox></TD>
          <TD vAlign=top>&nbsp;&nbsp;
<asp:linkbutton id=cmdAdd cssclass="CommandButton" runat="server"></asp:linkbutton></TD></TR>
        <TR>
          <TD height=10></TD></TR></TABLE>
<asp:comparevalidator id=valExpiryDate cssclass="NormalRed" runat="server" resourcekey="valExpiryDate" display="Dynamic" type="Date" operator="DataTypeCheck" errormessage="<br>Invalid expiry date" controltovalidate="txtExpiryDate"></asp:comparevalidator>
			</asp:panel>
		    <asp:checkbox id="chkNotify" resourcekey="SendNotification" runat="server" cssclass="SubHead" text="Send Notification?" textalign="Right"></asp:checkbox>
		</td>
	</tr>
	<tr><td height="25"></td></tr>
	<tr>
		<td>
			<hr noshade size="1">
			<asp:panel id="pnlUserRoles" runat="server" cssclass="WorkPanel" visible="True">
<asp:datagrid id=grdUserRoles runat="server" width="100%" summary="Security Roles Design Table" border="0" cellpadding="4" cellspacing="0" autogeneratecolumns="false" enableviewstate="false" datakeyfield="UserRoleID" ondeletecommand="grdUserRoles_Delete" BorderStyle="None" BorderWidth="0px" GridLines="None">
<Columns>
<asp:TemplateColumn>
<ItemTemplate>
							  <asp:imagebutton id="cmdDeleteUserRole" runat="server" alternatetext="Delete" causesvalidation="False" commandname="Delete" imageurl="~/images/delete.gif" resourcekey="cmdDelete"></asp:imagebutton>
						  
</ItemTemplate>
</asp:TemplateColumn>
<asp:BoundColumn DataField="FullName" HeaderText="UserName">
<HeaderStyle CssClass="NormalBold">
</HeaderStyle>

<ItemStyle CssClass="Normal">
</ItemStyle>
</asp:BoundColumn>
<asp:BoundColumn DataField="RoleName" HeaderText="SecurityRole">
<HeaderStyle CssClass="NormalBold">
</HeaderStyle>

<ItemStyle CssClass="Normal">
</ItemStyle>
</asp:BoundColumn>
<asp:TemplateColumn HeaderText="ExpiryDate">
<HeaderStyle CssClass="NormalBold">
</HeaderStyle>

<ItemTemplate>
							  <asp:label runat="server" text='<%#FormatExpiryDate(DataBinder.Eval(Container.DataItem, "ExpiryDate")) %>' cssclass="Normal" id="Label1" name="Label1"/>
						  
</ItemTemplate>
</asp:TemplateColumn>
</Columns>
</asp:datagrid>
      <HR noShade SIZE=1>
			</asp:panel>  
		</td>
	</tr>
</table>
<p>
  <asp:linkbutton id="cmdCancel" resourcekey="cmdCancel" runat="server" cssclass="CommandButton" text="Cancel" causesvalidation="False"></asp:linkbutton>
</p>
