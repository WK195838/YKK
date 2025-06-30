<%@ Control language="vb" CodeBehind="Register.ascx.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.Modules.Admin.Security.Register" %>
<%@ Register TagPrefix="dnn" TagName="Address" Src="~/controls/Address.ascx"%>
<%@ Register TagPrefix="dnn" TagName="Label" Src="~/controls/LabelControl.ascx" %>
<%@ Register TagPrefix="dnn" TagName="Sectionhead" Src="~/controls/SectionHeadControl.ascx" %>
<%@ Register TagPrefix="dnn" TagName="User" Src="~/controls/User.ascx"%>
<asp:panel id="UserRow" runat="server">
<asp:label id=lblRegistration runat="server" cssclass="Normal" width="600"></asp:label>
<TABLE cellSpacing=1 cellPadding=0 width=700 summary="Register Design Table" 
border=0>
  <TR>
    <TD colSpan=3>
<asp:label id=lblRegister runat="server" cssclass="Normal"></asp:label></TD></TR>
  <TR>
    <TD class=SubHead vAlign=top width=350>
<dnn:user id=userControl runat="server"></dnn:user></TD>
    <TD>&nbsp;</TD>
    <TD class=NormalBold vAlign=top width=350 rowSpan=8>
<dnn:address id=addressUser runat="server"></dnn:address></TD></TR></TABLE><BR>
<dnn:sectionhead id=dshPreferences runat="server" cssclass="Head" section="tblPreferences" includerule="True" resourcekey="Preferences" text="Preferences"></dnn:sectionhead>
<TABLE id=tblPreferences cellSpacing=1 cellPadding=0 width=600 
summary=Preferences runat="server">
  <TR>
    <TD class=SubHead width=225>
<dnn:label id=plLocale runat="server" text="Preferred Language:" controlname="cboLocale"></dnn:label></TD>
    <TD class=NormalBold noWrap>
<asp:dropdownlist id=cboLocale tabIndex=18 runat="server" cssclass="NormalTextBox" width="300"></asp:dropdownlist></TD></TR>
  <TR>
    <TD class=SubHead width=225>
<dnn:label id=plTimeZone runat="server" text="Time Zone:" controlname="cboTimeZone"></dnn:label></TD>
    <TD class=NormalBold noWrap>
<asp:dropdownlist id=cboTimeZone tabIndex=19 runat="server" cssclass="NormalTextBox" width="300"></asp:dropdownlist></TD></TR></TABLE>
</asp:panel>
<br>
<asp:panel id="PasswordManagementRow" runat="server">
<dnn:sectionhead id=dshPassword runat="server" cssclass="Head" section="tblPassword" includerule="True" resourcekey="ChangePassword" text="Change Password"></dnn:sectionhead>
<TABLE id=tblPassword cellSpacing=0 cellPadding=4 width=600 
summary="Password Management" border=0 runat="server">
  <TR vAlign=top height=*>
    <TD id=MessageCell colSpan=2 runat="server"></TD></TR>
  <TR>
    <TD class=SubHead width=175>
<dnn:label id=plOldPassword runat="server" text="Old Password:" controlname="txtOldPassword"></dnn:label></TD>
    <TD class=NormalBold noWrap>
<asp:textbox id=txtOldPassword tabIndex=20 runat="server" cssclass="NormalTextBox" textmode="Password" size="25" maxlength="20"></asp:textbox>&nbsp;*</TD></TR>
  <TR>
    <TD class=SubHead width=175>
<dnn:label id=plNewPassword runat="server" text="New Password:" controlname="txtNewPassword"></dnn:label></TD>
    <TD class=NormalBold noWrap>
<asp:textbox id=txtNewPassword tabIndex=21 runat="server" cssclass="NormalTextBox" textmode="Password" size="25" maxlength="20"></asp:textbox>&nbsp;*</TD></TR>
  <TR>
    <TD class=SubHead width=175>
<dnn:label id=plNewConfirm runat="server" text="Confirm New Password:" controlname="txtNewConfirm"></dnn:label></TD>
    <TD class=NormalBold noWrap>
<asp:textbox id=txtNewConfirm tabIndex=22 runat="server" cssclass="NormalTextBox" textmode="Password" size="25" maxlength="20"></asp:textbox>&nbsp;*</TD></TR>
  <TR>
    <TD colSpan=2>
<asp:linkbutton class=CommandButton id=cmdUpdatePassword runat="server" resourcekey="cmdUpdatePassword" text="Update Password"></asp:linkbutton></TD></TR></TABLE>
</asp:panel>
<br>
<asp:panel id="ServicesRow" runat="server">
<dnn:sectionhead id=dshServices runat="server" cssclass="Head" section="tblServices" includerule="True" resourcekey="Services" text="Membership Services" isExpanded="False"></dnn:sectionhead>
<TABLE id=tblServices cellSpacing=0 cellPadding=4 width=600 
summary="Register Design Table" border=0 runat="server">
  <TR>
    <TD>
<asp:datagrid id=grdServices runat="server" summary="Register Design Table" border="0" cellpadding="5" cellspacing="5" autogeneratecolumns="false" enableviewstate="true">
					<columns>
						<asp:templatecolumn>
							<itemtemplate>
								<asp:HyperLink Text='<%# ServiceText(DataBinder.Eval(Container.DataItem,"Subscribed")) %>' CssClass="CommandButton" NavigateUrl='<%# ServiceURL("RoleID",DataBinder.Eval(Container.DataItem,"RoleID"),DataBinder.Eval(Container.DataItem,"ServiceFee"),DataBinder.Eval(Container.DataItem,"Subscribed")) %>' runat="server" ID="Hyperlink1"/>
							</itemtemplate>
						</asp:templatecolumn>
						<asp:boundcolumn headertext="Name" datafield="RoleName" itemstyle-cssclass="Normal" headerstyle-cssclass="NormalBold" />
						<asp:boundcolumn headertext="Description" datafield="Description" itemstyle-cssclass="Normal" headerstyle-cssclass="NormalBold" />
						<asp:templatecolumn headertext="Fee" headerstyle-cssclass="NormalBold">
							<itemtemplate>
								<asp:Label runat="server" Text='<%#FormatPrice(DataBinder.Eval(Container.DataItem, "ServiceFee")) %>' CssClass="Normal" ID="Label2"/>
							</itemtemplate>
						</asp:templatecolumn>
						<asp:templatecolumn headertext="Every" headerstyle-cssclass="NormalBold">
							<itemtemplate>
								<asp:Label runat="server" Text='<%#FormatPeriod(DataBinder.Eval(Container.DataItem, "BillingPeriod")) %>' CssClass="Normal" ID="Label3"/>
							</itemtemplate>
						</asp:templatecolumn>
						<asp:boundcolumn headertext="Period" datafield="BillingFrequency" itemstyle-cssclass="Normal" headerstyle-cssclass="NormalBold" />
						<asp:templatecolumn headertext="Trial" headerstyle-cssclass="NormalBold">
							<itemtemplate>
								<asp:Label runat="server" Text='<%#FormatPrice(DataBinder.Eval(Container.DataItem, "TrialFee")) %>' CssClass="Normal" ID="Label4"/>
							</itemtemplate>
						</asp:templatecolumn>
						<asp:templatecolumn headertext="Every" headerstyle-cssclass="NormalBold">
							<itemtemplate>
								<asp:Label runat="server" Text='<%#FormatPeriod(DataBinder.Eval(Container.DataItem, "TrialPeriod")) %>' CssClass="Normal" ID="Label5"/>
							</itemtemplate>
						</asp:templatecolumn>
						<asp:boundcolumn headertext="Period" datafield="TrialFrequency" itemstyle-cssclass="Normal" headerstyle-cssclass="NormalBold" />
						<asp:templatecolumn headertext="ExpiryDate" headerstyle-cssclass="NormalBold">
							<itemtemplate>
								<asp:Label runat="server" Text='<%#FormatExpiryDate(DataBinder.Eval(Container.DataItem, "ExpiryDate")) %>' CssClass="Normal" ID="Label1"/>
							</itemtemplate>
						</asp:templatecolumn>
					</columns>
				</asp:datagrid>
<asp:label id=lblServices runat="server" cssclass="Normal" visible="False"></asp:label></TD></TR></TABLE>
</asp:panel>
<p>
	<asp:linkbutton class="CommandButton" id="cmdRegister" runat="server" text="Register"></asp:linkbutton>&nbsp;&nbsp;
	<asp:linkbutton class="CommandButton" id="cmdCancel" resourcekey="cmdCancel" runat="server" causesvalidation="False"
		text="Cancel" borderstyle="none"></asp:linkbutton>&nbsp;&nbsp;
	<asp:linkbutton class="CommandButton" id="cmdUnregister" resourcekey="cmdUnregister" runat="server"
		text="Unregister"></asp:linkbutton>
</p>
