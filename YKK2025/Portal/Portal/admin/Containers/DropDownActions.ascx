<%@ Control Language="vb" AutoEventWireup="false" Codebehind="DropDownActions.ascx.vb" Inherits="DotNetNuke.UI.Containers.DropDownActions" %>
<script>
	function cmdGo_OnClick(o)
	{
		if (o.selectedIndex > -1)
		{
			var sVal = dnn.getVar('__dnn_CSAction_' + o.id + '_' + o.options[o.selectedIndex].value);
			var bRet = true;
			if (sVal != null && sVal.length > 0)
				eval('bRet = ' + sVal);
			
			if (bRet == false)
				return false;
		}
		return true;
	}
</script>

<table cellpadding="0" cellspacing="0" border="0">
	<tr>
		<td nowrap>
			<asp:dropdownlist id="cboActions" runat="server"></asp:dropdownlist>
			<asp:ImageButton id="cmdGo" runat="server" ImageUrl="~/images/fwd.gif" AlternateText="Go" ToolTip="Go"></asp:ImageButton>
		</td>
	</tr>
</table>
