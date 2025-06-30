<%@ Control language="vb" Inherits="DotNetNuke.Modules.Events.Events" CodeBehind="Events.ascx.vb" AutoEventWireup="false" Explicit="True" %>
<asp:datalist id="lstEvents" runat="server" EnableViewState="false" CellPadding="4" summary="Events Design Table">
	<itemtemplate>
	    <table cellpadding="2" cellspacing="0" border="0" summary="Events Design Table">
			<tr>
				<td id="colIcon" runat="server" valign="top" align="center" rowspan="3" width='<%# DataBinder.Eval(Container.DataItem,"MaxWidth") %>'>
					<asp:Image ID="imgIcon" AlternateText='<%# DataBinder.Eval(Container.DataItem,"AltText") %>' runat="server" ImageUrl='<%# FormatImage(DataBinder.Eval(Container.DataItem,"IconFile")) %>' Visible='<%# FormatImage(DataBinder.Eval(Container.DataItem,"IconFile")) <> "" %>'></asp:Image>
				</td>
				<td>
					<asp:HyperLink id="editLink" NavigateUrl='<%# EditURL("ItemID",DataBinder.Eval(Container.DataItem,"ItemID")) %>' Visible="<%# IsEditable %>" runat="server"><asp:Image id="editLinkImage" ImageUrl="~/images/edit.gif" Visible="<%# IsEditable %>" AlternateText="Edit" runat="server" resourcekey="Edit" /></asp:HyperLink>
					<asp:Label ID="lblTitle" Runat="server" Cssclass="SubHead" text='<%# DataBinder.Eval(Container.DataItem,"Title") %>'></asp:Label>
				</td>
			</tr>
			<tr>
				<td>
					<asp:Label ID="lblDateTime" Runat="server" Cssclass="SubHead" text='<%# FormatDateTime(DataBinder.Eval(Container.DataItem,"DateTime")) %>'></asp:Label>					
				</td>
			</tr>
			<tr>
				<td>
					<asp:Label ID="lblDescription" Runat="server" CssClass="Normal" text='<%# DataBinder.Eval(Container.DataItem,"Description") %>'></asp:Label>					
				</td>
			</tr>
	    </table>
		<br>
	</ItemTemplate>
</asp:datalist>
<asp:calendar id="calEvents" runat="server" BorderWidth="1" CssClass="Normal" SelectionMode="None" summary="Events Calendar Design Table">
	<dayheaderstyle backcolor="#EEEEEE" cssclass="NormalBold" borderwidth="1"></DayHeaderStyle>
	<daystyle cssclass="Normal" borderwidth="1" verticalalign="Top"></DayStyle>
	<othermonthdaystyle forecolor="#FFFFFF"></OtherMonthDayStyle>
	<titlestyle font-bold="True"></TitleStyle>
	<nextprevstyle cssclass="NormalBold"></NextPrevStyle>
</asp:calendar>
