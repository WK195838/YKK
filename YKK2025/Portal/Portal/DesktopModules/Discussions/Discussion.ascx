<%@ Control language="vb" Inherits="DotNetNuke.Modules.Discussions.Discussion" CodeBehind="Discussion.ascx.vb" AutoEventWireup="false" Explicit="True" %>
<asp:DataList id=lstDiscussions runat="server" Summary="Discussion Design Table" CellPadding="4" DataKeyField="Parent" ItemStyle-Cssclass="Normal">
	<ItemTemplate>
		<asp:ImageButton id="btnSelect" AlternateText="Select" ImageUrl='<%# NodeImage(Cint(DataBinder.Eval(Container.DataItem, "ChildCount"))) %>' CommandName="select" runat="server" resourcekey="AltSelect"/>
		<asp:hyperlink Text='<%# DataBinder.Eval(Container.DataItem, "Title") %>' NavigateUrl='<%# EditUrl("ItemID",DataBinder.Eval(Container.DataItem, "ItemID"),"Edit","itemindex=" & lstDiscussions.Items.Count) %>' runat="server" />( <%# FormatUser(DataBinder.Eval(Container.DataItem,"CreatedByUser"),DataBinder.Eval(Container.DataItem,"CreatedDate", "{0:g}")) %> )
  </ItemTemplate>
	<SelectedItemTemplate>
		<asp:ImageButton id="btnCollapse" AlternateText="Collapse" ImageUrl="~/images/minus.gif" runat="server" CommandName="collapse" resourcekey="AltCollapse"/>
		<asp:hyperlink Text='<%# DataBinder.Eval(Container.DataItem, "Title") %>' NavigateUrl='<%# EditUrl("ItemID",DataBinder.Eval(Container.DataItem, "ItemID"),"Edit","itemindex=" & lstDiscussions.Items.Count) %>' runat="server" />( <%# FormatUser(DataBinder.Eval(Container.DataItem,"CreatedByUser"),DataBinder.Eval(Container.DataItem,"CreatedDate", "{0:g}")) %> )
		<asp:DataList id="DetailList" ItemStyle-Cssclass="Normal" datasource="<%# GetThreadMessages() %>" runat="server">
			<ItemTemplate>
				<br>
				<%# IndentChild(DataBinder.Eval(Container.DataItem, "Indent")) %>
				<asp:hyperlink Text='<%# DataBinder.Eval(Container.DataItem, "Title") %>' NavigateUrl='<%# EditUrl("ItemID",DataBinder.Eval(Container.DataItem, "ItemID"),"Edit","itemindex=" & lstDiscussions.Items.Count) %>' runat="server" />, <%# FormatUser(DataBinder.Eval(Container.DataItem,"CreatedByUser"),DataBinder.Eval(Container.DataItem,"CreatedDate", "{0:g}")) %>
			</ItemTemplate>
		</asp:DataList>
	</SelectedItemTemplate>
</asp:DataList>
