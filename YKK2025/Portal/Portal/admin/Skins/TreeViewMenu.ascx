<%@ Control language="vb" CodeBehind="TreeViewMenu.ascx.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.UI.Skins.Controls.TreeViewMenu" %>
<%@ Register TagPrefix="dnn" Namespace="DotNetNuke.UI.WebControls" Assembly="DotNetNuke.WebControls" %>
<table id="tblMain" runat="server" border="0" cellpadding="5" cellspacing="0">
	<tr>
		<td id="cellHeader" runat="server" valign="top">
			<table id="tblHeader" runat="server" cellpadding="0" cellspacing="0" border="0" width="100%">
				<tr>
					<td >
						<asp:label id="lblHeader" Runat="Server"></asp:label>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td id="cellBody" runat="server" valign="top">
			<dnn:DNNTree id="DNNTree" runat="server" />
		</td>
	</tr>
</table>
