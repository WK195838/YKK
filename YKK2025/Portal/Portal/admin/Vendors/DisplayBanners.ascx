<%@ Control language="vb" Inherits="DotNetNuke.Modules.Admin.Vendors.DisplayBanners" CodeBehind="DisplayBanners.ascx.vb" AutoEventWireup="false" Explicit="True" %>
<asp:DataList id=lstBanners runat="server" CellPadding="4" Summary="Banner Design Table" EnableViewState="true" Width="100%">
	<ItemStyle HorizontalAlign="Center" Width="100%" BorderColor="#000000"></ItemStyle>
	<ItemTemplate>
		<asp:Label ID="lblItem" Runat="server" Text='<%# FormatItem(DataBinder.Eval(Container.DataItem,"VendorId"),DataBinder.Eval(Container.DataItem,"BannerId"),DataBinder.Eval(Container.DataItem,"BannerTypeId"),DataBinder.Eval(Container.DataItem,"BannerName"),DataBinder.Eval(Container.DataItem,"ImageFile"),DataBinder.Eval(Container.DataItem,"Description"),DataBinder.Eval(Container.DataItem,"Url"),DataBinder.Eval(Container.DataItem,"Width"),DataBinder.Eval(Container.DataItem,"Height")) %>'></asp:Label>
	</ItemTemplate>
</asp:DataList>
