<%@ Control language="vb" CodeBehind="EditUserDefinedTable.ascx.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.Modules.UserDefinedTable.EditUserDefinedTable" %>
<br>
<asp:Table ID="tblFields" Runat="server" summary="Edit User Defined Design Table"></asp:Table>
<p>
	<asp:LinkButton id="cmdUpdate" Text="Update" runat="server" resourcekey="cmdUpdate" class="CommandButton" BorderStyle="none" />
	&nbsp;
	<asp:LinkButton id="cmdCancel" Text="Cancel" CausesValidation="False" resourcekey="cmdCancel" runat="server" class="CommandButton"
		BorderStyle="none" />
	&nbsp;
	<asp:LinkButton id="cmdDelete" Text="Delete" CausesValidation="False" resourcekey="cmdDelete" runat="server" class="CommandButton"
		BorderStyle="none" />
</p>
