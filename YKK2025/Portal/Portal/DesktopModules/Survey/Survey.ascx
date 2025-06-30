<%@ Control Language="vb" AutoEventWireup="false" Inherits="DotNetNuke.Modules.Survey.Survey" CodeBehind="Survey.ascx.vb" %>
<br>
<asp:panel id="pnlSurvey" runat="server" visible="False">
<asp:datalist id="lstSurvey" runat="server" cellpadding="4" datakeyfield="SurveyId">
		<ItemTemplate>
			<asp:HyperLink id=cmdEdit1 NavigateUrl='<%# EditURL("SurveyId",DataBinder.Eval(Container.DataItem,"SurveyId")) %>' Visible="<%# IsEditable %>" runat="server"><asp:image id="Image1" imageurl="~/images/edit.gif" visible="<%# IsEditable %>" alternatetext="Edit" runat="server" resourcekey="Edit"/></asp:HyperLink>
			<asp:Label id=lblQuestion1 Text='<%# FormatQuestion(DataBinder.Eval(Container.DataItem,"Question"),lstSurvey.Items.Count + 1) %>' CssClass="NormalBold" Runat="server">
			</asp:Label><BR>
			<asp:radiobuttonlist id="optOptions" visible="False" runat="server" cssclass="Normal" datatextfield="OptionName"
				datavaluefield="SurveyOptionId" repeatdirection="Vertical"></asp:radiobuttonlist>
			<asp:checkboxlist id="chkOptions" visible="False" runat="server" cssclass="Normal" datatextfield="OptionName"
				datavaluefield="SurveyOptionId" repeatdirection="Vertical"></asp:checkboxlist><BR>
		</ItemTemplate>
	</asp:datalist>
<asp:linkbutton id="cmdSubmit" runat="server" cssclass="CommandButton" resourcekey="cmdSubmit">Submit Survey</asp:linkbutton>&nbsp; 
<asp:linkbutton id="cmdResults" runat="server" cssclass="CommandButton" resourcekey="cmdResults">View Results</asp:linkbutton>
</asp:panel>
<asp:panel id="pnlResults" runat="server" visible="False">
	<asp:datalist id="lstResults" runat="server" cellpadding="4" datakeyfield="SurveyId">
		<itemtemplate>
			<asp:HyperLink id="cmdEdit2" ImageUrl="~/images/edit.gif" NavigateUrl='<%# EditURL("SurveyId",DataBinder.Eval(Container.DataItem,"SurveyId")) %>' Visible="<%# IsEditable %>" runat="server" />
			<asp:Label ID="lblQuestion2" Runat="server" CssClass="NormalBold" Text='<%# FormatQuestion(DataBinder.Eval(Container.DataItem,"Question"),lstResults.Items.Count + 1) %>'>
			</asp:Label><br>
			<asp:label id="lblResults" runat="server" cssclass="Normal"></asp:label>
			<br>
		</itemtemplate>
	</asp:datalist>
	<asp:linkbutton id="cmdSurvey" runat="server" cssclass="CommandButton" resourcekey="cmdSurvey">View Survey</asp:linkbutton>
</asp:panel>
