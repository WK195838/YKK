<%@ Control Inherits="DotNetNuke.Modules.Admin.Users.UserAccounts" CodeBehind="Users.ascx.vb" language="vb" AutoEventWireup="false" Explicit="True" %>
<%@ Register TagPrefix="dnnsc" Namespace="DotNetNuke.UI.WebControls" Assembly="DotNetNuke" %>
<table width="100%" border="0">
	<tr>
		<td align="left" class="Normal">
			<asp:Label ID="Label1" Runat="server" resourcekey="Search" CssClass="SubHead">Search:</asp:Label><br>
			<asp:TextBox id="txtSearch" Runat="server" />
			<asp:DropDownList id="ddlSearchType" Runat="server">
				<asp:ListItem Value="username" resourcekey="Username.Header">Username</asp:ListItem>
				<asp:ListItem Value="email" resourcekey="Email.Header">Email</asp:ListItem>
			</asp:DropDownList>
			<asp:ImageButton ID="btnSearch" Runat="server" ImageUrl="~/images/icon_search_16px.gif" />
		</td>
		<td align="left" class="Normal">
			<asp:Label ID="Label6" Runat="server" resourcekey="DisplayDate" CssClass="SubHead">Display Date:</asp:Label><br>
			<asp:RadioButtonList id="rblDisplayDate" Runat="server" AutoPostBack="True" RepeatDirection="Horizontal"
				RepeatLayout="Flow">
				<asp:ListItem Value="L" resourcekey="LastLogin.Header">Last Login</asp:ListItem>
				<asp:ListItem Value="C" resourcekey="CreatedDate.Header">Created Date</asp:ListItem>
			</asp:RadioButtonList>
		</td>
		<td align="right" class="Normal">
			<asp:Label ID="Label5" Runat="server" resourcekey="RecordsPage" CssClass="SubHead">Records Per Page:</asp:Label><br>
			<asp:DropDownList id="ddlRecordsPerPage" Runat="server" AutoPostBack="True">
				<asp:ListItem Value="10">10</asp:ListItem>
				<asp:ListItem Value="25">25</asp:ListItem>
				<asp:ListItem Value="50">50</asp:ListItem>
				<asp:ListItem Value="100">100</asp:ListItem>
				<asp:ListItem Value="250">250</asp:ListItem>
			</asp:DropDownList>
		</td>
	</tr>
	<tr>
		<td colspan="3" height="15"></td>
	</tr>
</table>
<asp:Panel ID="plLetterSearch" Runat="server" HorizontalAlign="Center">
	<asp:Repeater id="rptLetterSearch" Runat="server">
		<ItemTemplate>
			<asp:HyperLink runat="server" CssClass="CommandButton" NavigateUrl='<%# FilterURL(Container.DataItem,"1") %>' Text='<%# Container.DataItem %>'>
			</asp:HyperLink>&nbsp;&nbsp;
		</ItemTemplate>
	</asp:Repeater>
</asp:Panel>
<asp:datagrid id="grdUsers" Border="0" CellPadding="4" width="100%" AutoGenerateColumns="false"
	EnableViewState="false" runat="server" summary="Users Design Table" BorderStyle="None" BorderWidth="0px"
	GridLines="None">
	<Columns>
		<asp:TemplateColumn>
			<ItemStyle Width="20px"></ItemStyle>
			<ItemTemplate>
				<asp:HyperLink NavigateUrl='<%# FormatURL("UserID",CType(Container.DataItem, DotNetNuke.Entities.Users.UserInfo).UserID) %>' Visible="<%# IsEditable %>" runat="server" ID="Hyperlink1">
					<asp:Image ImageUrl="~/images/edit.gif" AlternateText="Edit" Visible="<%# IsEditable %>" runat="server" ID="Hyperlink1Image" resourcekey="Edit"/>
				</asp:HyperLink>
			</ItemTemplate>
		</asp:TemplateColumn>
		<asp:TemplateColumn HeaderText="Username">
			<HeaderStyle CssClass="NormalBold"></HeaderStyle>
			<ItemStyle CssClass="Normal"></ItemStyle>
			<ItemTemplate>
				<asp:Label ID="Label3" Runat="server" Text='<%# CType(Container.DataItem, DotNetNuke.Entities.Users.UserInfo).Membership.Username %>'>
				</asp:Label>
			</ItemTemplate>
		</asp:TemplateColumn>
		<asp:TemplateColumn HeaderText="Name">
			<HeaderStyle CssClass="NormalBold"></HeaderStyle>
			<ItemStyle CssClass="Normal"></ItemStyle>
			<ItemTemplate>
				<asp:Label ID="Label2" Runat="server" Text='<%# CType(Container.DataItem, DotNetNuke.Entities.Users.UserInfo).Profile.FullName%>'>
				</asp:Label>
			</ItemTemplate>
		</asp:TemplateColumn>
		<asp:TemplateColumn HeaderText="Address">
			<HeaderStyle CssClass="NormalBold"></HeaderStyle>
			<ItemStyle CssClass="Normal"></ItemStyle>
			<ItemTemplate>
				<asp:Label ID="lblAddress" Runat="server" Text='<%# DisplayAddress(CType(Container.DataItem, DotNetNuke.Entities.Users.UserInfo).Profile.Unit,CType(Container.DataItem, DotNetNuke.Entities.Users.UserInfo).Profile.Street, CType(Container.DataItem, DotNetNuke.Entities.Users.UserInfo).Profile.City, CType(Container.DataItem, DotNetNuke.Entities.Users.UserInfo).Profile.Region, CType(Container.DataItem, DotNetNuke.Entities.Users.UserInfo).Profile.Country, CType(Container.DataItem, DotNetNuke.Entities.Users.UserInfo).Profile.PostalCode) %>'>
				</asp:Label>
			</ItemTemplate>
		</asp:TemplateColumn>
		<asp:TemplateColumn HeaderText="Telephone">
			<HeaderStyle CssClass="NormalBold"></HeaderStyle>
			<ItemStyle CssClass="Normal"></ItemStyle>
			<ItemTemplate>
				<asp:Label ID="Label4" Runat="server" Text='<%# DisplayEmail(CType(Container.DataItem, DotNetNuke.Entities.Users.UserInfo).Profile.Telephone) %>'>
				</asp:Label>
			</ItemTemplate>
		</asp:TemplateColumn>
		<asp:TemplateColumn HeaderText="Email">
			<HeaderStyle CssClass="NormalBold"></HeaderStyle>
			<ItemStyle CssClass="Normal"></ItemStyle>
			<ItemTemplate>
				<asp:Label ID="lblEmail" Runat="server" Text='<%# DisplayEmail(CType(Container.DataItem, DotNetNuke.Entities.Users.UserInfo).Membership.Email) %>'>
				</asp:Label>
			</ItemTemplate>
		</asp:TemplateColumn>
		<asp:TemplateColumn HeaderText="CreatedDate">
			<HeaderStyle CssClass="NormalBold"></HeaderStyle>
			<ItemStyle CssClass="Normal"></ItemStyle>
			<ItemTemplate>
				<asp:Label ID="lblLastLogin" Runat="server" Text='<%# DisplayDate(CType(Container.DataItem, DotNetNuke.Entities.Users.UserInfo).Membership.CreatedDate) %>'>
				</asp:Label>
			</ItemTemplate>
		</asp:TemplateColumn>
		<asp:TemplateColumn HeaderText="LastLogin">
			<HeaderStyle CssClass="NormalBold"></HeaderStyle>
			<ItemStyle CssClass="Normal"></ItemStyle>
			<ItemTemplate>
				<asp:Label ID="Label7" Runat="server" Text='<%# DisplayDate(CType(Container.DataItem, DotNetNuke.Entities.Users.UserInfo).Membership.LastLoginDate) %>'>
				</asp:Label>
			</ItemTemplate>
		</asp:TemplateColumn>
		<asp:TemplateColumn HeaderText="Authorized">
			<HeaderStyle CssClass="NormalBold"></HeaderStyle>
			<ItemStyle CssClass="Normal"></ItemStyle>
			<ItemTemplate>
				<asp:Image Runat="server" ID="imgApproved" ImageUrl="~/images/checked.gif" Visible="<%# CType(Container.DataItem, DotNetNuke.Entities.Users.UserInfo).Membership.Approved=true%>"/>
				<asp:Image Runat="server" ID="imgNotApproved" ImageUrl="~/images/unchecked.gif" Visible="<%# CType(Container.DataItem, DotNetNuke.Entities.Users.UserInfo).Membership.Approved=false%>"/>
			</ItemTemplate>
		</asp:TemplateColumn>
	</Columns>
</asp:datagrid>
<br>
<br>
<dnnsc:PagingControl id="ctlPagingControl" runat="server"></dnnsc:PagingControl>
<p align="center">
	<asp:LinkButton ID="cmdDelete" Runat="server" CssClass="CommandButton" resourcekey="DeleteUnauthorized">Delete Unauthorized Users</asp:LinkButton>
</p>
