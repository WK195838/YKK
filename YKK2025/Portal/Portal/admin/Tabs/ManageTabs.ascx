<%@ Register TagPrefix="dnn" TagName="SectionHead" Src="~/controls/SectionHeadControl.ascx" %>
<%@ Register TagPrefix="dnn" TagName="Label" Src="~/controls/LabelControl.ascx" %>
<%@ Register TagPrefix="dnn" TagName="URL" Src="~/controls/URLControl.ascx" %>
<%@ Register TagPrefix="Portal" Namespace="DotNetNuke.Security.Permissions.Controls" Assembly="DotNetNuke" %>
<%@ Register TagPrefix="dnn" TagName="Skin" Src="~/controls/SkinControl.ascx" %>
<%@ Control Language="vb" AutoEventWireup="false" Explicit="True" Codebehind="ManageTabs.ascx.vb" Inherits="DotNetNuke.Modules.Admin.Tabs.ManageTabs"%>
<table class="Settings" cellspacing="2" cellpadding="2" summary="Manage Tabs Design Table"
	border="0">
	<tr>
		<td width="560" valign="top">
			<asp:panel id="pnlSettings" runat="server" cssclass="WorkPanel" visible="True">
				<dnn:sectionhead id="dshBasic" cssclass="Head" runat="server" includerule="True" resourcekey="BasicSettings"
					section="tblBasic" text="Basic Settings"></dnn:sectionhead>
				<table id="tblBasic" cellSpacing="0" cellPadding="2" width="525" summary="Basic Settings Design Table"
					border="0" runat="server">
					<tr>
						<td colSpan="2">
							<asp:label id="lblBasicSettingsHelp" cssclass="Normal" runat="server" resourcekey="BasicSettingsHelp"
								enableviewstate="False"></asp:label></td>
					</tr>
					<tr>
						<td width="25"></td>
						<td vAlign="top" width="475">
							<dnn:sectionhead id="dshPage" cssclass="Head" runat="server" resourcekey="PageDetails" section="tblPage"
								text="Page Details"></dnn:sectionhead>
							<table id="tblPage" cellSpacing="2" cellPadding="2" summary="Site Details Design Table"
								border="0" runat="server">
								<tr>
									<td class="SubHead" width="150">
										<dnn:label id="plTabName" runat="server" resourcekey="TabName" controlname="txtTabName" helpkey="TabNameHelp"
											suffix=":"></dnn:label></td>
									<td width="325">
										<asp:textbox id="txtTabName" cssclass="NormalTextBox" runat="server" width="300" maxlength="50"></asp:textbox>
										<asp:requiredfieldvalidator id="valTabName" cssclass="NormalRed" runat="server" resourcekey="valTabName.ErrorMessage"
											controltovalidate="txtTabName" errormessage="<br>Tab Name Is Required" display="Dynamic"></asp:requiredfieldvalidator>
									</td>
								</tr>
								<tr>
									<td class="SubHead" width="150">
										<dnn:label id="plTitle" runat="server" resourcekey="Title" controlname="txtTitle" helpkey="TitleHelp"
											suffix=":"></dnn:label></td>
									<td>
										<asp:textbox id="txtTitle" cssclass="NormalTextBox" runat="server" width="300" maxlength="200"></asp:textbox></td>
								</tr>
								<tr>
									<td class="SubHead" width="150">
										<dnn:label id="plDescription" runat="server" resourcekey="Description" controlname="txtDescription"
											helpkey="DescriptionHelp" suffix=":"></dnn:label></td>
									<td class="NormalTextBox" width="325">
										<asp:textbox id="txtDescription" cssclass="NormalTextBox" runat="server" width="300" maxlength="500"
											rows="3" textmode="MultiLine"></asp:textbox></td>
								</tr>
								<tr>
									<td class="SubHead" width="150">
										<dnn:label id="plKeywords" runat="server" resourcekey="KeyWords" controlname="txtKeyWords"
											helpkey="KeyWordsHelp" suffix=":"></dnn:label></td>
									<td class="NormalTextBox" width="325">
										<asp:textbox id="txtKeyWords" cssclass="NormalTextBox" runat="server" width="300" maxlength="500"
											rows="3" textmode="MultiLine"></asp:textbox></td>
								</tr>
								<tr>
									<td class="SubHead" width="150">
										<dnn:label id="plParentTab" runat="server" resourcekey="ParentTab" controlname="cboTab" helpkey="ParentTabHelp"
											suffix=":"></dnn:label></td>
									<td width="325">
										<asp:dropdownlist id="cboTab" cssclass="NormalTextBox" runat="server" width="300" datavaluefield="TabId"
											datatextfield="TabName"></asp:dropdownlist></td>
								</tr>
								<tr>
									<td class="SubHead" vAlign="top" width="150"><BR>
										<BR>
										<dnn:label id="plPermissions" runat="server" resourcekey="Permissions" controlname="dgPermissions"
											helpkey="PermissionsHelp" suffix=":"></dnn:label></td>
									<td width="325">
										<Portal:TabPermissionsGrid id="dgPermissions" runat="server" ForeColor="Transparent">
											<HeaderStyle Font-Bold="True" HorizontalAlign="Center" ForeColor="Black" BackColor="Transparent"></HeaderStyle>
											<FooterStyle BackColor="White"></FooterStyle>
										</Portal:TabPermissionsGrid></td>
								</tr>
							</table>
						</td>
					</tr>
				</table>
				<BR>
				<dnn:sectionhead id="dshCopy" cssclass="Head" runat="server" includerule="True" resourcekey="Copy"
					section="tblCopy" text="Copy Page"></dnn:sectionhead>
				<table id="tblCopy" cellSpacing="0" cellPadding="2" width="525" summary="Copy Tab Design Table"
					border="0" runat="server">
					<tr>
						<td width="25"></td>
						<td class="SubHead" width="150">
							<dnn:label id="plTemplate" runat="server" resourcekey="CopyModules" controlname="cboTemplate"
								helpkey="CopyModulesHelp" suffix=":"></dnn:label></td>
						<td width="325">
							<asp:dropdownlist id="cboTemplate" cssclass="NormalTextBox" runat="server" width="300" datavaluefield="TabId"
								datatextfield="TabName"></asp:dropdownlist></td>
					</tr>
					<tr>
						<td width="25"></td>
						<td class="SubHead" width="150">
							<dnn:label id="plContent" runat="server" resourcekey="CopyContent" controlname="chkContent"
								helpkey="CopyContentHelp" suffix="?"></dnn:label></td>
						<td width="325">
							<asp:checkbox id="chkContent" runat="server" font-names="Verdana,Arial" font-size="8pt"></asp:checkbox></td>
					</tr>
				</table>
				<BR>
				<dnn:sectionhead id="dshAdvanced" cssclass="Head" runat="server" includerule="True" resourcekey="AdvancedSettings"
					section="tblAdvanced" text="Advanced Settings" isexpanded="False"></dnn:sectionhead>
				<table id="tblAdvanced" cellSpacing="0" cellPadding="2" width="525" summary="Advanced Settings Design Table"
					border="0" runat="server">
					<tr>
						<td colSpan="2">
							<asp:label id="lblAdvancedSettingsHelp" cssclass="Normal" runat="server" resourcekey="AdvancedSettingsHelp"
								enableviewstate="False"></asp:label></td>
					</tr>
					<tr>
						<td width="25"></td>
						<td vAlign="top" width="475">
							<dnn:sectionhead id="dhsAppearance" cssclass="Head" runat="server" resourcekey="Appearance" section="tblAppearance"
								text="Appearance"></dnn:sectionhead>
							<table id="tblAppearance" cellSpacing="2" cellPadding="2" summary="Appearance Design Table"
								border="0" runat="server">
								<tr>
									<td class="SubHead" vAlign="top" width="150">
										<dnn:label id="plIcon" runat="server" resourcekey="Icon" controlname="ctlIcon" helpkey="IconHelp"
											suffix=":"></dnn:label></td>
									<td width="325">
										<dnn:url id="ctlIcon" runat="server" width="300" showlog="False"></dnn:url></td>
								</tr>
								<tr>
									<td class="SubHead" vAlign="top" width="150">
										<dnn:label id="plSkin" runat="server" resourcekey="TabSkin" controlname="ctlSkin" helpkey="TabSkinHelp"
											suffix=":"></dnn:label></td>
									<td vAlign="top" width="325">
										<dnn:skin id="ctlSkin" runat="server"></dnn:skin></td>
								</tr>
								<tr>
									<td class="SubHead" vAlign="top" width="150">
										<dnn:label id="plContainer" runat="server" resourcekey="TabContainer" controlname="ctlContainer"
											helpkey="TabContainerHelp" suffix=":"></dnn:label></td>
									<td vAlign="top" width="325">
										<dnn:skin id="ctlContainer" runat="server"></dnn:skin></td>
								</tr>
								<tr>
									<td class="SubHead" vAlign="top" width="150">
										<dnn:label id="plHidden" runat="server" resourcekey="Hidden" controlname="chkHidden" helpkey="HiddenHelp"
											suffix=":"></dnn:label></td>
									<td width="325">
										<asp:checkbox id="chkHidden" runat="server" font-names="Verdana,Arial" font-size="8pt"></asp:checkbox></td>
								</tr>
								<tr>
									<td class="SubHead" vAlign="top" width="150">
										<dnn:label id="plDisable" runat="server" resourcekey="Disabled" controlname="chkDisableLink"
											helpkey="DisabledHelp" suffix=":"></dnn:label></td>
									<td width="325">
										<asp:checkbox id="chkDisableLink" runat="server" font-names="Verdana,Arial" font-size="8pt"></asp:checkbox></td>
								</tr>
								<tr>
									<td class="SubHead" width="150">
										<dnn:label id="plRefreshInterval" runat="server" columns="10" resourcekey="RefreshInterval" controlname="cboRefreshInterval" helpkey="RefreshInterval.Help"
											suffix=":"></dnn:label></td>
									<td width="325">
										<asp:TextBox id="txtRefreshInterval" cssclass="NormalTextBox" runat="server"/></td>
								</tr>
								<tr>
									<td class="SubHead" vAlign="top" width="150">
										<dnn:label id="plPageHeadText" runat="server" resourcekey="PageHeadText" controlname="txtPageHeadText" helpkey="PageHeadText.Help" suffix=":"></dnn:label></td>
									<td class="NormalTextBox" width="325"><asp:TextBox id="txtPageHeadText" runat="server" cssclass="NormalTextBox" textmode="MultiLine" columns="50" rows="4"/>
										</td>
								</tr>
							</table>
							<BR>
							<dnn:sectionhead id="dshOther" cssclass="Head" runat="server" resourcekey="OtherSettings" section="tblOther"
								text="Other Settings"></dnn:sectionhead>
							<table id="tblOther" cellSpacing="2" cellPadding="2" summary="Appearance Design Table"
								border="0" runat="server">
								<tr>
									<td class="SubHead" vAlign="top" width="150">
										<dnn:label id="plStartDate" runat="server" text="Start Date:" controlname="txtStartDate"></dnn:label></td>
									<td>
										<asp:textbox id="txtStartDate" cssclass="NormalTextBox" runat="server" width="120" maxlength="11"
											columns="30"></asp:textbox>&nbsp;
										<asp:hyperlink id="cmdStartCalendar" resourcekey="Calendar" cssclass="CommandButton" runat="server">Calendar</asp:hyperlink>
										<asp:CompareValidator id="valtxtStartDate" resourcekey="valStartDate.ErrorMessage" ControlToValidate="txtStartDate"
											Display="Dynamic" Runat="server" ErrorMessage="<br>Invalid Start Date" Type="Date" Operator="DataTypeCheck"></asp:CompareValidator></td>
								</tr>
								<tr>
									<td class="SubHead" vAlign="top" width="150">
										<dnn:label id="plEndDate" runat="server" text="End Date:" controlname="txtEndDate"></dnn:label></td>
									<td>
										<asp:textbox id="txtEndDate" cssclass="NormalTextBox" runat="server" width="120" maxlength="11"
											columns="30"></asp:textbox>&nbsp;
										<asp:hyperlink id="cmdEndCalendar" resourcekey="Calendar" cssclass="CommandButton" runat="server">Calendar</asp:hyperlink>
										<asp:CompareValidator id="valtxtEndDate" resourcekey="valEndDate.ErrorMessage" ControlToValidate="txtEndDate"
											Display="Dynamic" Runat="server" ErrorMessage="<br>Invalid End Date" Type="Date" Operator="DataTypeCheck"></asp:CompareValidator></td>
								</tr>
								<tr>
									<td class="SubHead" vAlign="top" width="150">
										<dnn:label id="plURL" runat="server" resourcekey="Url" controlname="ctlURL" helpkey="UrlHelp"
											suffix=":"></dnn:label></td>
									<td class="NormalTextBox" width="325">
										<dnn:url id="ctlURL" runat="server" width="300" showlog="False" showtrack="False" shownone="True"></dnn:url></td>
								</tr>
							</table>
						</td>
					</tr>
				</table>
			</asp:panel>
		</td>
	</tr>
</table>
<p>
	<asp:linkbutton cssclass="CommandButton" id="cmdUpdate" resourcekey="cmdUpdate" runat="server" Text="Update"></asp:linkbutton>&nbsp;&nbsp;
	<asp:linkbutton cssclass="CommandButton" id="cmdCancel" resourcekey="cmdCancel" runat="server" Text="Cancel"
		CausesValidation="False" BorderStyle="none"></asp:linkbutton>&nbsp;&nbsp;
	<asp:linkbutton cssclass="CommandButton" id="cmdDelete" resourcekey="cmdDelete" runat="server" Text="Delete"
		CausesValidation="False" BorderStyle="none"></asp:linkbutton>&nbsp;&nbsp;
	<asp:linkbutton cssclass="CommandButton" id="cmdGoogle" resourcekey="SubmitToGoogle" runat="server"
		text="Submit Tab To Google"></asp:linkbutton>
</p>
