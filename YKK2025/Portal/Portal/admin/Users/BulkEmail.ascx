<%@ Register TagPrefix="dnn" TagName="TextEditor" Src="~/controls/TextEditor.ascx"%>
<%@ Register TagPrefix="dnn" TagName="SectionHead" Src="~/controls/SectionHeadControl.ascx" %>
<%@ Register TagPrefix="dnn" TagName="Label" Src="~/controls/LabelControl.ascx" %>
<%@ Control Inherits="DotNetNuke.Modules.Admin.Users.BulkEmail" CodeBehind="BulkEmail.ascx.vb" language="vb" AutoEventWireup="false" Explicit="True" %>
<%@ Register TagPrefix="dnn" TagName="URLControl" Src="~/controls/URLControl.ascx" %>
<table class="Settings" cellspacing="2" cellpadding="2" summary="Edit Roles Design Table"
	border="0">
	<tr>
		<td width="560" valign="top">
			<asp:panel id="pnlSettings" runat="server" cssclass="WorkPanel" visible="True">
				<dnn:sectionhead id="dshBasic" cssclass="Head" runat="server" text="Basic Settings" section="tblBasic"
					resourcekey="BasicSettings" includerule="True"></dnn:sectionhead>
				<TABLE id="tblBasic" cellSpacing="0" cellPadding="2" width="525" summary="Basic Settings Design Table"
					border="0" runat="server">
					<TR>
						<TD colSpan="2">
							<asp:label id="lblBasicSettingsHelp" cssclass="Normal" runat="server" resourcekey="BasicSettingsDescription"
								enableviewstate="False"></asp:label></TD>
					</TR>
					<TR>
						<TD class="SubHead" width="150">
							<dnn:label id="plRoles" runat="server" suffix=":" controlname="chkRoles"></dnn:label></TD>
						<TD align="center" width="325">
							<asp:checkboxlist id="chkRoles" cssclass="Normal" runat="server" repeatcolumns="2" datavaluefield="RoleName"
								datatextfield="RoleName" width="325"></asp:checkboxlist></TD>
					</TR>
					<TR vAlign="top">
						<TD class="SubHead" width="150">
							<dnn:label id="plEmail" runat="server" suffix=":" controlname="txtEmail"></dnn:label></TD>
						<TD align="center" width="325">
							<asp:textbox id="txtEmail" cssclass="NormalTextBox" runat="server" width="325" rows="3" textmode="MultiLine"></asp:textbox></TD>
					</TR>
					<TR vAlign="top">
						<TD class="SubHead" width="150">
							<dnn:label id="plSubject" runat="server" suffix=":" controlname="txtSubject"></dnn:label></TD>
						<TD width="325">
							<asp:textbox id="txtSubject" cssclass="NormalTextBox" runat="server" width="325" maxlength="100"
								columns="40"></asp:textbox></TD>
					</TR>
				</TABLE>
				<BR>
				<dnn:sectionhead id="dshMessage" cssclass="Head" runat="server" text="Message" section="tblMessage"
					resourcekey="Message" includerule="True"></dnn:sectionhead>
				<TABLE id="tblMessage" cellSpacing="0" cellPadding="2" width="525" summary="Message Design Table"
					border="0" runat="server">
					<TR>
						<TD colSpan="2">
							<asp:label id="lblMessageHelp" cssclass="Normal" runat="server" resourcekey="MessageDescription"
								enableviewstate="False"></asp:label></TD>
					</TR>
					<TR vAlign="top">
						<TD colSpan="2">
							<dnn:texteditor id="teMessage" runat="server" width="550" chooserender="False" choosemode="True"
								height="350" defaultmode="Rich" htmlencode="False" textrendermode="Raw"></dnn:texteditor></TD>
					</TR>
				</TABLE>
				<BR>
				<dnn:sectionhead id="dshAdvanced" cssclass="Head" runat="server" text="Advanced Settings" section="tblAdvanced"
					resourcekey="AdvancedSettings" includerule="True" isexpanded="False"></dnn:sectionhead>
				<TABLE id="tblAdvanced" cellSpacing="0" cellPadding="2" width="525" summary="Message Design Table"
					border="0" runat="server">
					<TR>
						<TD colSpan="2">
							<asp:label id="lblAdvancedSettingsHelp" cssclass="Normal" runat="server" resourcekey="AdvancedSettingsHelp"
								enableviewstate="False"></asp:label></TD>
					</TR>
					<TR vAlign="top">
						<TD class="SubHead" width="150">
							<dnn:label id="plAttachment" runat="server" suffix=":" controlname="cboAttachment"></dnn:label></TD>
						<TD width="325">
							<dnn:URLControl id="ctlAttachment" runat="server" ShowUrls="False" ShowTabs="False" ShowLog="False" ShowTrack="False" ShowUpload="true" required="False"></dnn:URLControl></TD>
					</TR>
					<TR vAlign="top">
						<TD class="SubHead" width="150">
							<dnn:label id="plPriority" runat="server" suffix=":" controlname="cboPriority"></dnn:label></TD>
						<TD width="325">
							<asp:dropdownlist id="cboPriority" cssclass="NormalTextBox" runat="server" width="100">
								<asp:listitem resourcekey="High" value="1">High</asp:listitem>
								<asp:listitem resourcekey="Normal" value="2" selected="True">Normal</asp:listitem>
								<asp:listitem resourcekey="Low" value="3">Low</asp:listitem>
							</asp:dropdownlist></TD>
					</TR>
					<TR vAlign="top">
						<TD class="SubHead" width="150">
							<dnn:label id="plSendMethod" runat="server" suffix=":" controlname="cboSendMethod"></dnn:label></TD>
						<TD width="325">
							<asp:dropdownlist id="cboSendMethod" cssclass="NormalTextBox" runat="server" width="325px">
								<asp:listitem resourcekey="SendTo" value="TO" selected="True">TO: One Message Per Email Address ( Personalized )</asp:listitem>
								<asp:listitem resourcekey="SendBCC" value="BCC">BCC: One Email To Blind Distribution List ( Not Personalized )</asp:listitem>
							</asp:dropdownlist></TD>
					</TR>
					<TR vAlign="top">
						<TD class="SubHead" width="150">
							<dnn:label id="plSendAction" runat="server" suffix=":" controlname="optSendAction"></dnn:label></TD>
						<TD width="325">
							<asp:radiobuttonlist id="optSendAction" cssclass="Normal" runat="server" repeatdirection="Horizontal">
								<asp:listitem resourcekey="Synchronous" value="S">Synchronous</asp:listitem>
								<asp:listitem resourcekey="Asynchronous" value="A" selected="True">Asynchronous</asp:listitem>
							</asp:radiobuttonlist></TD>
					</TR>
				</TABLE>
			</asp:panel>
		</td>
	</tr>
</table>
<p>
	<asp:linkbutton id="cmdSend" resourcekey="cmdSend" text="Send Email" runat="server" cssclass="CommandButton" />
</p>
