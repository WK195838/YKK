<%@ Register TagPrefix="dnn" TagName="Label" Src="~/controls/LabelControl.ascx" %>
<%@ Register TagPrefix="Portal" TagName="Audit" Src="~/controls/ModuleAuditControl.ascx" %>
<%@ Register TagPrefix="dnn" TagName="TextEditor" Src="~/controls/TextEditor.ascx"%>
<%@ Control language="vb" CodeBehind="EditFAQs.ascx.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.Modules.FAQs.EditFAQs" %>
<table width="650" cellspacing="0" cellpadding="0" border="0" summary="Edit FAQs Design Table">
	<tr valign="top">
		<td class="SubHead" width="125"><dnn:label id="plQuestionField" runat="server" controlname="QuestionField" suffix=":"></dnn:label></td>
		<td>
			<dnn:texteditor ControlID="teQuestionField" id="teQuestionField" runat="server" height="200" width="500" />
			<%--<JPW bug=#151 type=add comment=added a RequiredFieldValidator for QuestionField>--%>
			<asp:RequiredFieldValidator ID="valQuestionField" resourcekey="valQuestionField.ErrorMessage" ControlToValidate="teQuestionField"
				CssClass="NormalRed" Display="Dynamic" ErrorMessage="<br>A question is required" Runat="server" />
			<%--</JPW>--%><br><br><br>
		</td>
	</tr>
	<tr valign="top">
		<td class="SubHead" width="125"><dnn:label id="plAnswerField" runat="server" controlname="AnswerField" suffix=":"></dnn:label></td>
		<td>
			<dnn:texteditor ControlID="teAnswerField" id="teAnswerField" runat="server" height="200" width="500" />
			<%--<JPW bug=#151 type=add comment=added a RequiredFieldValidator for AnswerField>--%>
			<asp:RequiredFieldValidator ID="valAnswerField" resourcekey="valAnswerField.ErrorMessage" ControlToValidate="teAnswerField"
				CssClass="NormalRed" Display="Dynamic" ErrorMessage="<br>An answer is required" Runat="server" />
			<%--</JPW>--%>
		</td>
	</tr>
</table>
<p>
	<asp:linkbutton id="cmdUpdate" resourcekey="cmdUpdate" runat="server" cssclass="CommandButton" text="Update"
		borderstyle="none"></asp:linkbutton>&nbsp;
	<asp:linkbutton id="cmdCancel" resourcekey="cmdCancel" runat="server" cssclass="CommandButton" text="Cancel"
		borderstyle="none" causesvalidation="False"></asp:linkbutton>&nbsp;
	<asp:linkbutton id="cmdDelete" resourcekey="cmdDelete" runat="server" cssclass="CommandButton" text="Delete"
		borderstyle="none" causesvalidation="False"></asp:linkbutton>
</p>
<portal:audit id="ctlAudit" runat="server" />
