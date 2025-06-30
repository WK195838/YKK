<%@ Control language="vb" Inherits="DotNetNuke.Modules.Documents.Document" CodeBehind="Document.ascx.vb" AutoEventWireup="false" Explicit="True" %>
<asp:datagrid id=grdDocuments runat="server" cellpadding="4" datakeyfield="ItemID" oneditcommand="grdDocuments_Edit" enableviewstate="false" autogeneratecolumns="false" BorderWidth="0" summary="This table shows various documents that can be downloaded.">
	<columns>
		<asp:templatecolumn>
			<itemtemplate>
				<asp:hyperlink id="editLink" navigateurl='<%# EditURL("ItemID",DataBinder.Eval(Container.DataItem,"ItemID")) %>' visible="<%# IsEditable %>" runat="server"><asp:image id="editLinkImage" imageurl="~/images/edit.gif" visible="<%# IsEditable %>" alternatetext="Edit" runat="server" resourcekey="Edit"/></asp:hyperlink>
			</itemtemplate>
		</asp:templatecolumn>
		<asp:boundcolumn headertext="Category" datafield="Category" itemstyle-wrap="false" itemstyle-cssclass="Normal" headerstyle-cssclass="NormalBold"/>
		<asp:templatecolumn headertext="Title" headerstyle-cssclass="NormalBold">
			<itemtemplate>
				<asp:hyperlink id="docLink" text='<%# DataBinder.Eval(Container.DataItem,"Title") %>' navigateurl='<%# FormatURL(DataBinder.Eval(Container.DataItem,"URL"),DataBinder.Eval(Container.DataItem,"TrackClicks")) %>' cssclass="Normal" Target= '<%# IIF(DataBinder.Eval(Container.DataItem,"NewWindow"),"_blank","_self") %>' runat="server" />
			</itemtemplate>
		</asp:templatecolumn>
		<asp:boundcolumn  Visible="False" headertext="Owner" datafield="CreatedByUser" itemstyle-cssclass="Normal" headerstyle-cssclass="NormalBold" />
		<asp:boundcolumn  headertext="LastUpdated" datafield="CreatedDate" dataformatstring="{0:d}" itemstyle-cssclass="Normal" headerstyle-cssclass="NormalBold" />
		<asp:templatecolumn  headertext="Size" headerstyle-cssclass="NormalBold" itemstyle-horizontalalign="Right" headerstyle-horizontalalign="Right">
			<itemtemplate>
				<asp:label id="Label1" text='<%# FormatSize(DataBinder.Eval(Container.DataItem,"Size")) %>' cssclass="Normal" runat="server" />
			</itemtemplate>
		</asp:templatecolumn>
		<asp:templatecolumn>
			<itemtemplate>
				<asp:linkbutton id="lnkDownload" resourcekey="cmdDownLoad" runat="server" cssclass="CommandButton" commandname="Edit">Download</asp:linkbutton>
			</itemtemplate>
		</asp:templatecolumn>
	</columns>
</asp:datagrid>
