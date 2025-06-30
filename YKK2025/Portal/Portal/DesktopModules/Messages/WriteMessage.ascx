<%@ Control Inherits="ImportantMessage.WriteMessage" Src="ImportantMessages.vb" Language="vb" %>
<%@ Register TagPrefix="dnn" TagName="Label" Src="~/controls/LabelControl.ascx" %>
<table cellspacing="0" cellpadding="2" border="0" summary="importantmessagesdev edit design table">
	<tr>
		<td class="SubHead" width="200">		
			<dnn:label id="plEditText" controlname="txtEditText" suffix="" runat="server"></dnn:label>
		</td>
		<td>		
			<asp:TextBox id="txtEditText" CssClass="NormalTextBox" runat="server"></asp:TextBox>
		</td>
	</tr>
</table>
<asp:linkbutton id="cmdUpdate" resourcekey="cmdUpdate" CssClass="CommandButton" borderstyle="none" runat="server"></asp:linkbutton>
<asp:linkbutton id="cmdCancel" resourcekey="cmdCancel" CssClass="CommandButton" causesvalidation="False" borderstyle="none" runat="server"></asp:linkbutton>
