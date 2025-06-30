<%@ Control language="vb" CodeBehind="DiscussDetails.ascx.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.Modules.Discussions.DiscussDetails" %>
<%@ Register TagPrefix="dnn" TagName="Label" Src="~/controls/LabelControl.ascx" %>
<%@ Register TagPrefix="dnn" TagName="TextEditor" Src="~/controls/TextEditor.ascx"%>
<asp:panel id="EditPanel" Visible="false" runat="server">
	<TABLE cellSpacing="0" cellPadding="4" width="560" summary="Discuss Details Design Table"
		border="0">
		<TR vAlign="top">
			<TD class="SubHead" width="125">
				<dnn:label id="plTitleField" runat="server" controlname="TitleField" suffix=":"></dnn:label></TD>
			<TD width="400">
				<asp:TextBox id="TitleField" runat="server" cssclass="NormalTextBox" width="400" columns="40"
					maxlength="100"></asp:TextBox>
				<asp:RequiredFieldValidator id="valTitleField" resourcekey="valTitleField.ErrorMessage" ErrorMessage="You Must Enter The Title Of The Message"
					Display="Dynamic" ControlToValidate="TitleField" CssClass="NormalRed" Runat="server"></asp:RequiredFieldValidator></TD>
		</TR>
		<TR vAlign="top">
			<TD class="SubHead" width="125">
				<dnn:label id="plBodyField" runat="server" controlname="BodyField" suffix=":"></dnn:label></TD>
			<TD width="400">
				<dnn:texteditor id="teBodyField" runat="server" width="450" height="250"></dnn:texteditor>
				<asp:RequiredFieldValidator id="valBodyField" resourcekey="valBodyField.ErrorMessage" ErrorMessage="You Must Enter The Body Of The Message"
					Display="Dynamic" ControlToValidate="TitleField" CssClass="NormalRed" Runat="server"></asp:RequiredFieldValidator>
			</TD>
		</TR>
		<TR vAlign="top">
			<TD>&nbsp;
			</TD>
			<TD>
				<asp:LinkButton class="CommandButton" id="cmdSubmit" runat="server" resourcekey="cmdSubmit" Text="Submit">Submit</asp:LinkButton>&nbsp;
				<asp:LinkButton class="CommandButton" id="cmdUpdate" runat="server" Visible="False" resourcekey="cmdUpdate"
					Text="Submit">Update</asp:LinkButton>&nbsp;
				<asp:LinkButton class="CommandButton" id="cmdCancel" runat="server" resourcekey="cmdCancel" Text="Cancel"
					CausesValidation="False"></asp:LinkButton></TD>
		</TR>
	</TABLE>
</asp:panel>
<table ID="tblOriginalMessage" runat="server" width="560" cellspacing="0" cellpadding="4"
	border="0" Summary="Discuss Details Design Table">
	<tr id="rowOriginalMessage" runat="server" valign="top" visible="false">
		<td>
			<hr noShade SIZE="1">
		</td>
	</tr>
	<tr valign="top">
		<td align="left">
			<table cellSpacing="0" cellPadding="4" width="560" border="0" Summary="Discuss Details Design Table">
				<tr>
					<td class="SubHead" valign="top" width="100"><asp:Label id="lblSubject" resourcekey="Subject" runat="server" />:</td>
					<td width="400"><asp:Label id="Subject" CssClass="Normal" Width="400" runat="server" /></td>
				</tr>
				<tr>
					<td class="SubHead" valign="top" width="100"><asp:Label id="lblAuthor" resourcekey="Author" runat="server" />:</td>
					<td width="400"><asp:Label id="CreatedByUser" CssClass="Normal" Width="400" runat="server" /></td>
				</tr>
				<tr>
					<td class="SubHead" valign="top" width="100"><asp:Label id="lblDate" resourcekey="Date" runat="server" />:</td>
					<td><asp:Label id="CreatedDate" CssClass="Normal" Width="400" runat="server" /></td>
				</tr>
				<tr>
					<td class="SubHead" valign="top" width="100"><asp:Label id="lblBody" resourcekey="Body" runat="server" />:</td>
					<td><asp:Label id="Body" CssClass="Normal" Width="400" runat="server" /></td>
				</tr>
				<tr>
					<td colspan="2">&nbsp;</td>
				</tr>
				<tr>
					<td colspan="2">
						<asp:LinkButton class="CommandButton" id="cmdCancel2" resourcekey="cmdCancel" Text="Cancel" CausesValidation="False"
							runat="server" />&nbsp;
						<asp:LinkButton class="CommandButton" id="cmdReply" resourcekey="cmdReply" runat="server" Text="Reply"
							CausesValidation="False"></asp:LinkButton>&nbsp;
						<asp:LinkButton class="CommandButton" id="cmdEdit" resourcekey="cmdEdit" runat="server" Text="Edit"
							CausesValidation="False"></asp:LinkButton>&nbsp;
						<asp:LinkButton class="CommandButton" id="cmdDelete" resourcekey="cmdDelete" runat="server" Text="Delete"
							CausesValidation="False"></asp:LinkButton>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
