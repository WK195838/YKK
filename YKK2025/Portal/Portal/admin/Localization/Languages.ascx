<%@ Control Language="vb" AutoEventWireup="false" Codebehind="Languages.ascx.vb" Inherits="DotNetNuke.Services.Localization.Languages" %>
<%@ Register TagPrefix="dnn" TagName="SectionHead" Src="~/controls/SectionHeadControl.ascx" %>
<%@ Register TagPrefix="dnn" TagName="Label" Src="~/controls/LabelControl.ascx" %>
<dnn:sectionhead id="dshBasic" includerule="False" resourcekey="SupportedLocales" section="tblBasic"
	text="Suported Locales" cssclass="Head" runat="server"></dnn:sectionhead><br>
<TABLE id="tblBasic" cellSpacing="1" cellPadding="1" border="0" runat="server">
	<TR>
		<TD noWrap><asp:datagrid id="dgLocales" runat="server" GridLines="None" CellPadding="4" AutoGenerateColumns="False"
				CssClass="Normal">
				<AlternatingItemStyle Wrap="False"></AlternatingItemStyle>
				<ItemStyle Wrap="False"></ItemStyle>
				<HeaderStyle Font-Bold="True" Wrap="False"></HeaderStyle>
				<Columns>
					<asp:BoundColumn DataField="name" ReadOnly="True" HeaderText="Name"></asp:BoundColumn>
					<asp:BoundColumn DataField="key" ReadOnly="True" HeaderText="Key"></asp:BoundColumn>
					<asp:TemplateColumn HeaderText="Status">
						<ItemTemplate>
							<asp:Label id="lblStatus" runat="server" CssClass="Normal"></asp:Label>
						</ItemTemplate>
					</asp:TemplateColumn>
					<asp:TemplateColumn>
						<ItemTemplate>
							<asp:LinkButton id="cmdDisable" runat="server" CssClass="CommandButton" Text="Disable" CausesValidation="false"
								CommandName="Disable"></asp:LinkButton>
						</ItemTemplate>
					</asp:TemplateColumn>
					<asp:TemplateColumn>
						<ItemTemplate>
							<asp:LinkButton id=cmdDelete runat="server" resourcekey="cmdDelete" CssClass="CommandButton" Text="Delete" CommandName="Delete" CausesValidation="False" Visible="<%# UserInfo.IsSuperUser %>">Delete</asp:LinkButton>
						</ItemTemplate>
					</asp:TemplateColumn>
				</Columns>
			</asp:datagrid>
			<P><asp:checkbox id="chkDeleteFiles" resourcekey="DeleteFiles" runat="server" Text="Delete all resources of this locale, too"
					CssClass="Normal"></asp:checkbox></P>
		</TD>
	</TR>
</TABLE>
<asp:Panel id="pnlAdd" runat="server">
	<P>
		<dnn:sectionhead id="dshAdd" runat="server" cssclass="Head" text="Add New Locale" section="tblAdd"
			resourcekey="AddNewLocale" includerule="False"></dnn:sectionhead><BR>
		<TABLE id="tblAdd" cellSpacing="1" cellPadding="1" border="0" runat="server">
			<TR>
				<TD class="SubHead" vAlign="top" width="75">
					<dnn:label id="lbLocale" runat="server" text="Name" controlname="txtName"></dnn:label></TD>
				<TD vAlign="top" noWrap>
					<asp:dropdownlist id="cboLocales" runat="server" Width="300px"></asp:dropdownlist></TD>
			</TR>
			<TR>
				<TD noWrap></TD>
				<TD noWrap>
					<asp:linkbutton id="cmdAdd" runat="server" resourcekey="Add" CssClass="CommandButton">Add</asp:linkbutton></TD>
			</TR>
		</TABLE>
	</P>
</asp:Panel>
