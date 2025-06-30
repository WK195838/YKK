<%@ Register TagPrefix="dnn" TagName="Label" Src="~/controls/LabelControl.ascx" %>
<%@ Register TagPrefix="Portal" TagName="URL" Src="~/controls/URLControl.ascx" %>
<%@ Register TagPrefix="Portal" TagName="Audit" Src="~/controls/ModuleAuditControl.ascx" %>
<%@ Register TagPrefix="dnn" TagName="TextEditor" Src="~/controls/TextEditor.ascx"%>
<%@ Control language="vb" CodeBehind="EditEvents.ascx.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.Modules.Events.EditEvents" %>
<asp:panel id="pnlContent" runat="server">
	<TABLE cellSpacing="0" cellPadding="0" width="600" summary="Edit Events Design Table">
		<TR vAlign="top">
			<TD class="SubHead" width="125">
				<dnn:label id="plTitle" runat="server" controlname="txtTitle" suffix=":"></dnn:label></TD>
			<TD width="450">
				<asp:textbox id="txtTitle" runat="server" maxlength="150" width="390" cssclass="NormalTextBox"
					columns="30"></asp:textbox>
				<asp:requiredfieldvalidator id="valTitle" runat="server" cssclass="NormalRed" resourcekey="valTitle.ErrorMessage"
					controltovalidate="txtTitle" errormessage="Title Is Required" display="Dynamic"></asp:requiredfieldvalidator></TD>
		</TR>
		<TR vAlign="top">
			<TD class="SubHead" width="125">
				<dnn:label id="plDescription" runat="server" controlname="txtDescription" suffix=":"></dnn:label></TD>
			<TD width="450">
				<dnn:texteditor id="teDescription" runat="server" width="450" height="200"></dnn:texteditor>
				<asp:requiredfieldvalidator id="valDescription" runat="server" cssclass="NormalRed" resourcekey="valDescription.ErrorMessage"
					controltovalidate="teDescription" errormessage="Description Is Required" display="Dynamic"></asp:requiredfieldvalidator></TD>
		</TR>
		<TR vAlign="top">
			<TD class="SubHead" width="125">
				<dnn:label id="plImage" runat="server" controlname="cboImage" suffix=":"></dnn:label></TD>
			<TD width="450">
				<portal:url id="ctlImage" runat="server" width="300" showtabs="False" showurls="False" urltype="F"
					showtrack="False" showlog="False" required="False"></portal:url></TD>
		</TR>
		<TR>
			<TD class="SubHead" width="125">
				<dnn:label id="plAlt" runat="server" controlname="txtAlt" suffix=":"></dnn:label></TD>
			<TD width="450">
				<asp:textbox id="txtAlt" runat="server" cssclass="NormalTextBox" columns="50"></asp:textbox>
				<asp:label id="lblvalAltText" EnableViewState="False" Runat="server" CssClass="NormalRed"></asp:label></TD>
		</TR>
		<TR>
			<TD class="SubHead" width="125">
				<dnn:label id="plEvery" runat="server" controlname="txtEvery" suffix=":"></dnn:label></TD>
			<TD width="450">
				<asp:textbox id="txtEvery" runat="server" maxlength="3" cssclass="NormalTextBox" columns="3"></asp:textbox>&nbsp;
				<LABEL style="DISPLAY: none" for="<%=cboPeriod.ClientID%>">Period</LABEL>
				<asp:dropdownlist id="cboPeriod" runat="server" cssclass="NormalTextBox">
					<asp:listitem value=""></asp:listitem>
					<asp:listitem resourcekey="Days" value="D">Day(s)</asp:listitem>
					<asp:listitem resourcekey="Weeks" value="W">Week(s)</asp:listitem>
					<asp:listitem resourcekey="Months" value="M">Month(s)</asp:listitem>
					<asp:listitem resourcekey="Years" value="Y">Year(s)</asp:listitem>
				</asp:dropdownlist>
				<asp:CompareValidator id="valEvery1" runat="server" resourcekey="valEvery" CssClass="NormalRed" Display="Dynamic"
					ControlToValidate="txtEvery" Type="Integer" Operator="DataTypeCheck" ErrorMessage="<br>The frequency must be a number greater than zero"></asp:CompareValidator>
				<asp:comparevalidator id="valEvery2" runat="server" cssclass="NormalRed" resourcekey="valEvery" controltovalidate="txtEvery"
					errormessage="<br>The frequency must be a number greater than zero" display="Dynamic" operator="GreaterThan"
					valuetocompare="0"></asp:comparevalidator></TD>
		</TR>
		<TR>
			<TD class="SubHead" width="125">
				<dnn:label id="plStartDate" runat="server" controlname="txtStartDate" suffix=":"></dnn:label></TD>
			<TD width="450">
				<asp:textbox id="txtStartDate" runat="server" maxlength="20" cssclass="NormalTextBox" columns="20"></asp:textbox>&nbsp;
				<asp:hyperlink id="cmdStartCalendar" runat="server" cssclass="CommandButton" resourcekey="Calendar">Calendar</asp:hyperlink>
				<asp:requiredfieldvalidator id="valStartDate" runat="server" cssclass="NormalRed" resourcekey="valStartDate.ErrorMessage"
					controltovalidate="txtStartDate" errormessage="<br>Start Date Is Required" display="Dynamic"></asp:requiredfieldvalidator>
				<asp:comparevalidator id="valStartDate2" runat="server" cssclass="NormalRed" resourcekey="valStartDate2.ErrorMessage"
					controltovalidate="txtStartDate" errormessage="<br>Invalid start date!" display="Dynamic" type="Date"
					operator="DataTypeCheck"></asp:comparevalidator></TD>
		</TR>
		<TR>
			<TD class="SubHead" width="125">
				<dnn:label id="plTime" runat="server" controlname="txtTime" suffix=":"></dnn:label></TD>
			<TD width="450">
				<asp:textbox id="txtTime" runat="server" maxlength="8" cssclass="NormalTextBox" columns="8"></asp:textbox>
				<asp:label id="lblvalTime" EnableViewState="False" Runat="server" CssClass="NormalRed"></asp:label></TD>
		</TR>
		<TR>
			<TD class="SubHead" width="125">
				<dnn:label id="plExpiryDate" runat="server" controlname="txtExpiryDate" suffix=":"></dnn:label></TD>
			<TD width="450">
				<asp:textbox id="txtExpiryDate" runat="server" maxlength="20" cssclass="NormalTextBox" columns="20"></asp:textbox>&nbsp;
				<asp:hyperlink id="cmdExpiryCalendar" runat="server" cssclass="CommandButton" resourcekey="Calendar">Calendar</asp:hyperlink>
				<asp:comparevalidator id="valExpiryDate" runat="server" cssclass="NormalRed" resourcekey="valExpiryDate.ErrorMessage"
					controltovalidate="txtExpiryDate" errormessage="<br>Invalid expiry date!" display="Dynamic" type="Date"
					operator="DataTypeCheck"></asp:comparevalidator></TD>
		</TR>
	</TABLE>
	<P>
		<asp:linkbutton id="cmdUpdate" runat="server" cssclass="CommandButton" resourcekey="cmdUpdate" text="Update"
			borderstyle="none"></asp:linkbutton>&nbsp;
		<asp:linkbutton id="cmdCancel" runat="server" cssclass="CommandButton" resourcekey="cmdCancel" text="Cancel"
			borderstyle="none" causesvalidation="False"></asp:linkbutton>&nbsp;
		<asp:linkbutton id="cmdDelete" runat="server" cssclass="CommandButton" resourcekey="cmdDelete" text="Delete"
			borderstyle="none" causesvalidation="False"></asp:linkbutton></P>
	<portal:audit id="ctlAudit" runat="server"></portal:audit>
</asp:panel>
