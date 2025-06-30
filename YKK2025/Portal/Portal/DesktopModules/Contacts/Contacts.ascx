<%@ Control language="vb" Inherits="DotNetNuke.Modules.Contacts.Contacts" CodeBehind="Contacts.ascx.vb" AutoEventWireup="false" Explicit="True" %>
<asp:datagrid id=grdContacts runat="server" CellPadding="4" EnableViewState="false" Summary="This table shows various contacts with information like name, email, and phone number." AutoGenerateColumns="false" BorderWidth="0">
	<Columns>
		<asp:TemplateColumn>
			<ItemTemplate>
				<asp:HyperLink id="editContact" NavigateUrl='<%# EditURL("ItemID",DataBinder.Eval(Container.DataItem,"ItemID")) %>' Visible="<%# IsEditable %>" runat="server"><asp:image id="editContactImage" imageurl="~/images/edit.gif" visible="<%# IsEditable %>" alternatetext="Edit" runat="server" resourcekey="Edit"/></asp:HyperLink>
			</ItemTemplate>
		</asp:TemplateColumn>
		<asp:BoundColumn HeaderText="Name" DataField="Name" ItemStyle-CssClass="Normal" HeaderStyle-Cssclass="NormalBold" />
		<asp:BoundColumn HeaderText="Role" DataField="Role" ItemStyle-CssClass="Normal" HeaderStyle-Cssclass="NormalBold" />
		<asp:TemplateColumn HeaderText="Email" ItemStyle-CssClass="Normal" HeaderStyle-Cssclass="NormalBold">
			<ItemTemplate>
				<asp:Label ID="lblEmail" Runat="server" Text='<%# DisplayEmail(DataBinder.Eval(Container.DataItem, "Email")) %>'></asp:Label>
			</ItemTemplate>
		</asp:TemplateColumn>
		<asp:BoundColumn HeaderText="Telephone" DataField="Contact1" ItemStyle-CssClass="Normal" HeaderStyle-Cssclass="NormalBold" />
		<asp:BoundColumn HeaderText="Telephone2" DataField="Contact2" ItemStyle-CssClass="Normal" HeaderStyle-Cssclass="NormalBold" />
	</Columns>
</asp:datagrid>
