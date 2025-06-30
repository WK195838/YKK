<%@ Register TagPrefix="dnn" TagName="Label" Src="~/controls/LabelControl.ascx" %>
<%@ Register TagPrefix="dnn" TagName="TextEditor" Src="~/controls/TextEditor.ascx"%>
<%@ Control Language="vb" AutoEventWireup="false" Codebehind="HitCounterEdit.ascx.vb" Inherits="Boskone.DNN.Modules.HitCounter.HitCounterEdit" targetSchema="http://schemas.microsoft.com/intellisense/ie5"%>
<TABLE id="Table1" cellSpacing="0" cellPadding="0" width="100%" border="1">
	<TR>
		<TD class="optionSideCell" vAlign="top" align="center" colSpan="2"><dnn:label id="lblInstructions" suffix="." runat="server"></dnn:label></TD>
	</TR>
	<TR>
		<TD class="optionEditorCell" vAlign="top" align="center" colSpan="2"><dnn:texteditor id="TextEditor1" runat="server" Width="575" Mode="R" Height="250"></dnn:texteditor></TD>
	</TR>
	<tr>
		<td colSpan="2">
			<table id="table2" cellSpacing="0" cellPadding="0" width="100%">
				<TR>
					<TD colspan="2" class="optionSideCell" align="center" width="50%"><dnn:label id="lblOptions" suffix=":" runat="server"></dnn:label></TD>
				</TR>
				<tr>
					<td class="optionBottomCell" align="center" width="50%"><dnn:label id="lblTimeOut" suffix=":" runat="server"></dnn:label><asp:textbox id="txtTimeOut" runat="server" Width="40px"></asp:textbox>&nbsp;
						<asp:rangevalidator id="TimeOutValidator" runat="server" ErrorMessage="Must be between 1 and 60!" ControlToValidate="txtTimeOut"
							Type="Integer" Display="Dynamic" MinimumValue="1" MaximumValue="60"></asp:rangevalidator></td>
					<td class="optionBottomCell" align="center"><dnn:label id="lblAdminHits" runat="server"></dnn:label><asp:checkbox id="chkAdminHits" runat="server" TextAlign="Left" Text=""></asp:checkbox></td>
				</tr>
				<tr>
					<TD class="optionHeaderCell" align="right"><dnn:label id="lblHead1" suffix=":" runat="server"></dnn:label></TD>
					<td class="optionHeaderCell" align="left"><asp:textbox id="txtCurrentHits" runat="server" Width="48px"></asp:textbox><asp:regularexpressionvalidator id="HitValidator" runat="server" ValidationExpression="^\d+$" ErrorMessage="Must be numeric!"
							ControlToValidate="txtCurrentHits"></asp:regularexpressionvalidator></td>
				</tr>
			</table>
		</td>
	<TR>
		<TD align="center" colSpan="2">
			<asp:linkbutton id="cmdUpdate" resourcekey="cmdUpdate" runat="server" cssclass="CommandButton" text="Update"
				borderstyle="none"></asp:linkbutton>&nbsp;|&nbsp;
			<asp:linkbutton id="cmdCancel" resourcekey="cmdCancel" runat="server" cssclass="CommandButton" text="Cancel"
				borderstyle="none" causesvalidation="False"></asp:linkbutton>&nbsp;
		</TD>
	</TR>
</TABLE>
