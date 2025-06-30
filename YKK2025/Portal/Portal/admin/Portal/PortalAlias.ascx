<%@ Control language="vb" CodeBehind="PortalAlias.ascx.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.Modules.Admin.Portals.PortalAlias" %>
<asp:DataGrid BorderWidth=0px Width="500" AutoGenerateColumns=false ID="dgPortalAlias" Runat="server">
	<Columns>
		<asp:TemplateColumn>
			<ItemStyle Width="15"/>
			<ItemTemplate>
				<asp:HyperLink NavigateUrl='<%# EditURL("paid",DataBinder.Eval(Container.DataItem,"PortalAliasID")) %>' runat="server" Visible='<% # IsNotCurrent(DataBinder.Eval(Container.DataItem,"PortalAliasID")) %>' ID="Hyperlink1">
<asp:Image ImageUrl="~/images/edit.gif" AlternateText="Edit" runat="server" ID="Hyperlink1Image" resourcekey="Edit"/></asp:HyperLink>
			</ItemTemplate>
		</asp:TemplateColumn>
		<asp:BoundColumn ItemStyle-CssClass="Normal" HeaderStyle-CssClass="NormalBold" DataField="HTTPAlias" HeaderText="HTTP Alias" />
	</Columns>
</asp:DataGrid>
