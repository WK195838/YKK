<%@ Control Inherits="DotNetNuke.Modules.Admin.Security.Roles" CodeBehind="Roles.ascx.vb" language="vb" AutoEventWireup="false" Explicit="True" %>
<asp:datagrid id="grdRoles" Border="0" CellPadding="4" CellSpacing="0" Width="100%" AutoGenerateColumns="false"
	EnableViewState="false" runat="server" summary="Roles Design Table" BorderStyle="None" BorderWidth="0px"
	GridLines="None">
	<Columns>
		<asp:TemplateColumn>
			<ItemStyle Width="20px"></ItemStyle>
			<ItemTemplate>
				<asp:HyperLink NavigateUrl='<%# EditURL("RoleID",DataBinder.Eval(Container.DataItem,"RoleID")) %>' Visible="<%# IsEditable %>" runat="server" ID="Hyperlink1">
					<asp:Image ImageUrl="~/images/edit.gif" AlternateText="Edit" Visible="<%# IsEditable %>" runat="server" ID="Hyperlink1Image" resourcekey="Edit"/>
				</asp:HyperLink>
			</ItemTemplate>
		</asp:TemplateColumn>
		<asp:BoundColumn DataField="RoleName" HeaderText="Name">
			<HeaderStyle CssClass="NormalBold"></HeaderStyle>
			<ItemStyle CssClass="Normal"></ItemStyle>
		</asp:BoundColumn>
		<asp:BoundColumn DataField="Description" HeaderText="Description">
			<HeaderStyle CssClass="NormalBold"></HeaderStyle>
			<ItemStyle CssClass="Normal"></ItemStyle>
		</asp:BoundColumn>
		<asp:TemplateColumn HeaderText="Fee">
			<HeaderStyle CssClass="NormalBold"></HeaderStyle>
			<ItemTemplate>
				<asp:Label runat="server" Text='<%#FormatPrice(DataBinder.Eval(Container.DataItem, "ServiceFee")) %>' CssClass="Normal" ID="Label1"/>
			</ItemTemplate>
		</asp:TemplateColumn>
		<asp:TemplateColumn HeaderText="Every">
			<HeaderStyle CssClass="NormalBold"></HeaderStyle>
			<ItemTemplate>
				<asp:Label runat="server" Text='<%#FormatPeriod(DataBinder.Eval(Container.DataItem, "BillingPeriod")) %>' CssClass="Normal" ID="Label2"/>
			</ItemTemplate>
		</asp:TemplateColumn>
		<asp:BoundColumn DataField="BillingFrequency" HeaderText="Period">
			<HeaderStyle CssClass="NormalBold"></HeaderStyle>
			<ItemStyle CssClass="Normal"></ItemStyle>
		</asp:BoundColumn>
		<asp:TemplateColumn HeaderText="Trial">
			<HeaderStyle CssClass="NormalBold"></HeaderStyle>
			<ItemTemplate>
				<asp:Label runat="server" Text='<%#FormatPrice(DataBinder.Eval(Container.DataItem, "TrialFee")) %>' CssClass="Normal" ID="Label3"/>
			</ItemTemplate>
		</asp:TemplateColumn>
		<asp:TemplateColumn HeaderText="Every">
			<HeaderStyle CssClass="NormalBold"></HeaderStyle>
			<ItemTemplate>
				<asp:Label runat="server" Text='<%#FormatPeriod(DataBinder.Eval(Container.DataItem, "TrialPeriod")) %>' CssClass="Normal" ID="Label4"/>
			</ItemTemplate>
		</asp:TemplateColumn>
		<asp:BoundColumn DataField="TrialFrequency" HeaderText="Period">
			<HeaderStyle CssClass="NormalBold"></HeaderStyle>
			<ItemStyle CssClass="Normal"></ItemStyle>
		</asp:BoundColumn>
		<asp:TemplateColumn HeaderText="Public">
			<HeaderStyle HorizontalAlign="Center" CssClass="NormalBold"></HeaderStyle>
			<ItemStyle HorizontalAlign="Center" CssClass="Normal"></ItemStyle>
			<ItemTemplate>
				<asp:Image Runat="server" ID="imgApproved" ImageUrl="~/images/checked.gif" Visible='<%# DataBinder.Eval(Container.DataItem,"IsPublic")="true" %>'/>
				<asp:Image Runat="server" ID="imgNotApproved" ImageUrl="~/images/unchecked.gif" Visible='<%# DataBinder.Eval(Container.DataItem,"IsPublic")="false" %>'/>
			</ItemTemplate>
		</asp:TemplateColumn>
		<asp:TemplateColumn HeaderText="Auto">
			<HeaderStyle HorizontalAlign="Center" CssClass="NormalBold"></HeaderStyle>
			<ItemStyle HorizontalAlign="Center" CssClass="Normal"></ItemStyle>
			<ItemTemplate>
				<asp:Image Runat="server" ID="Image1" ImageUrl="~/images/checked.gif" Visible='<%# DataBinder.Eval(Container.DataItem,"AutoAssignment")="true" %>'/>
				<asp:Image Runat="server" ID="Image2" ImageUrl="~/images/unchecked.gif" Visible='<%# DataBinder.Eval(Container.DataItem,"AutoAssignment")="false" %>'/>
			</ItemTemplate>
		</asp:TemplateColumn>
	</Columns>
</asp:datagrid>
