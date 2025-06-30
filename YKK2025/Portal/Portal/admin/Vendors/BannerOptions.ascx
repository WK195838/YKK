<%@ Control language="vb" CodeBehind="BannerOptions.ascx.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.Modules.Admin.Vendors.BannerOptions" %>
<%@ Register TagPrefix="dnn" TagName="Label" Src="~/controls/LabelControl.ascx" %>
<br>
<table cellSpacing="2" cellPadding="0" width="560" summary="Banner Options Design Table">
	<tr valign="bottom">
		<td class="SubHead" width="125"><dnn:label id="plSource" runat="server" controlname="optSource" suffix=":"></dnn:label></td>
		<td>
			<asp:RadioButtonList id="optSource" runat="server" CssClass="NormalBold" RepeatDirection="Horizontal">
				<asp:ListItem Value="G">Host</asp:ListItem>
				<asp:ListItem Value="L">Site</asp:ListItem>
			</asp:RadioButtonList>
		</td>
	</tr>
	<tr valign="bottom">
		<td class="SubHead" width="125"><dnn:label id="plType" runat="server" controlname="cboType" suffix=":"></dnn:label></td>
		<td>
			<asp:DropDownList ID="cboType" Runat="server" CssClass="NormalTextBox" Width="250px" DataTextField="BannerTypeName" DataValueField="BannerTypeId"></asp:DropDownList>
		</td>
	</tr>
	<tr valign="bottom">
		<td class="SubHead" width="125"><dnn:label id="plGroup" runat="server" controlname="txtGroup" suffix=":"></dnn:label></td>
		<td>
			<asp:TextBox id="txtGroup" Runat="server" CssClass="NormalTextBox" Columns="30" Width="250px"></asp:TextBox> 
		</td>
	</tr>
	<tr valign="bottom">
		<td class="SubHead" width="125"><dnn:label id="plCount" runat="server" controlname="txtCount" suffix=":"></dnn:label></td>
		<td>
			<asp:TextBox id="txtCount" Runat="server" CssClass="NormalTextBox" Columns="30" Width="250px"></asp:TextBox> 
			<asp:RegularExpressionValidator id="valCount" ControlToValidate="txtCount" ValidationExpression="^[0-9]*$" Display="Dynamic" ErrorMessage="<br>Banner Count Must Be A Valid Integer" runat="server" CssClass="NormalRed" />
		</td>
	</tr>
	<tr valign="bottom">
		<td class="SubHead" width="125"><dnn:label id="plOrientation" runat="server" controlname="optOrientation" suffix=":"></dnn:label></td>
		<td>
			<asp:RadioButtonList id="optOrientation" runat="server" CssClass="NormalBold" RepeatDirection="Horizontal">
				<asp:ListItem Value="V">Vertical</asp:ListItem>
				<asp:ListItem Value="H">Horizontal</asp:ListItem>
			</asp:RadioButtonList>
		</td>
	</tr>
	<tr valign="bottom">
		<td class="SubHead" width="125"><dnn:label id="plBorder" runat="server" controlname="txtBorder" suffix=":"></dnn:label></td>
		<td>
			<asp:TextBox id="txtBorder" Runat="server" CssClass="NormalTextBox" Columns="30" Width="250px"></asp:TextBox> 
			<asp:RegularExpressionValidator id="valBorder" ControlToValidate="txtBorder" ValidationExpression="^[0-9]*$" Display="Dynamic" ErrorMessage="<br>Border Width Must Be A Valid Integer" runat="server" CssClass="NormalRed" />
		</td>
	</tr>
	<tr valign="bottom">
		<td class="SubHead" width="125"><dnn:label id="plBorderColor" runat="server" controlname="txtBorderColor" suffix=":"></dnn:label></td>
		<td>
			<asp:TextBox id="txtBorderColor" Runat="server" CssClass="NormalTextBox" Columns="30" Width="250px"></asp:TextBox> 
		</td>
	</tr>
	<tr valign="bottom">
		<td class="SubHead" width="125"><dnn:label id="plRowHeight" runat="server" controlname="txtRowHeight" suffix=":"></dnn:label></td>
		<td>
			<asp:TextBox id="txtRowHeight" Runat="server" CssClass="NormalTextBox" Columns="30" Width="250px"></asp:TextBox> 
			<asp:RegularExpressionValidator id="valRowHeight" ControlToValidate="txtRowHeight" ValidationExpression="^[0-9]*$" Display="Dynamic" ErrorMessage="<br>Row Height Must Be A Valid Integer" runat="server" CssClass="NormalRed" />
		</td>
	</tr>
	<tr valign="bottom">
		<td class="SubHead" width="125"><dnn:label id="plColWidth" runat="server" controlname="txtColWidth" suffix=":"></dnn:label></td>
		<td>
			<asp:TextBox id="txtColWidth" Runat="server" CssClass="NormalTextBox" Columns="30" Width="250px"></asp:TextBox> 
			<asp:RegularExpressionValidator id="valColWidth" ControlToValidate="txtColWidth" ValidationExpression="^[0-9]*$" Display="Dynamic" ErrorMessage="<br>Column Width Must Be A Valid Integer" runat="server" CssClass="NormalRed" />
		</td>
	</tr>
</table>
<p>
	<asp:LinkButton id="cmdUpdate" Text="Update" resourcekey="cmdUpdate" runat="server" class="CommandButton" BorderStyle="none" />
	&nbsp;
	<asp:LinkButton id="cmdCancel" Text="Cancel" resourcekey="cmdCancel" CausesValidation="False" runat="server" class="CommandButton" BorderStyle="none" />
</p>
