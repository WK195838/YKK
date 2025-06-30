<%@ Control language="vb" Inherits="DotNetNuke.Modules.UserDefinedTable.UserDefinedTable" CodeBehind="UserDefinedTable.ascx.vb" AutoEventWireup="false" Explicit="True" %>
<asp:datagrid id="grdData" runat="server" OnSortCommand="grdData_Sort" AllowSorting="True" ItemStyle-CssClass="Normal"
	HeaderStyle-CssClass="NormalBold" AutoGenerateColumns="False" CellPadding="4" BorderWidth="0">
	<columns>
		<asp:TemplateColumn>
			<itemtemplate>
				<asp:HyperLink NavigateUrl='<%# EditURL("UserDefinedRowId",DataBinder.Eval(Container.DataItem,"UserDefinedRowId")) %>' Visible="<%# IsEditable %>" runat="server" ID="Hyperlink1">
					<asp:Image ID=Hyperlink1Image Runat=server ImageUrl="~/images/edit.gif" AlternateText="Edit" Visible="<%#IsEditable%>" resourcekey="Edit"/>
				</asp:HyperLink>
			</itemtemplate>
		</asp:TemplateColumn>
	</columns>
</asp:datagrid>
