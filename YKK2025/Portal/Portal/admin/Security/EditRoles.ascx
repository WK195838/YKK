<%@ Control language="vb" CodeBehind="EditRoles.ascx.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.Modules.Admin.Security.EditRoles" %>
<%@ Register TagPrefix="dnn" TagName="Label" Src="~/controls/LabelControl.ascx" %>
<%@ Register TagPrefix="dnn" TagName="SectionHead" Src="~/controls/SectionHeadControl.ascx" %>
<table class="Settings" cellspacing="2" cellpadding="2" summary="Edit Roles Design Table" border="0">
  <tr>
    <td width="560" valign="top">
      <asp:panel id="pnlSettings" runat="server" cssclass="WorkPanel" visible="True">
        <dnn:sectionhead id="dshBasic" runat="server"
          cssclass="Head"
          includerule="True"
          resourcekey="BasicSettings"
          section="tblBasic"
          text="Basic Settings"
          />
        <table id="tblBasic" runat="server" cellspacing="0" cellpadding="2" width="525" summary="Basic Settings Design Table" border="0">
          <tr>
            <td colspan="2"><asp:label id="lblBasicSettingsHelp" resourcekey="BasicSettingsDescription" cssclass="Normal" runat="server" enableviewstate="False"></asp:label></td>
          </tr>
	        <tr valign="top">
            <td class="SubHead" width="150"><dnn:label id="plRoleName" resourcekey="RoleName" runat="server" controlname="txtRoleName" suffix=":"></dnn:label></td>
		        <td align="left" width="325">
			        <asp:textbox id="txtRoleName" cssclass="NormalTextBox" width="325" columns="30" maxlength="50" runat="server" />
			        <asp:Label ID="lblRoleName" CssClass="Normal" Runat="server" Visible="False"/>
			        <asp:requiredfieldvalidator id="valRoleName" resourcekey="valRoleName" display="Dynamic" runat="server" errormessage="<br>You Must Enter a Valid Name" controltovalidate="txtRoleName" cssclass="NormalRed" />
		        </td>
	        </tr>
	        <tr valign="top">
            <td class="SubHead" width="150"><dnn:label id="plDescription" resourcekey="Description" runat="server" controlname="txtDescription" suffix=":"></dnn:label></td>
		        <td width="325"><asp:textbox id="txtDescription" cssclass="NormalTextBox" width="325" columns="30" maxlength="1000" runat="server" height="84px" textmode="MultiLine" /></td>
	        </tr>
	        <tr>
            <td class="SubHead" width="150"><dnn:label id="plIsPublic" resourcekey="PublicRole" runat="server" controlname="chkIsPublic"></dnn:label></td>
		        <td width="325"><asp:checkbox id="chkIsPublic" runat="server"></asp:checkbox></td>
	        </tr>
	        <tr>
            <td class="SubHead" width="150"><dnn:label id="plAutoAssignment" resourcekey="AutoAssignment" runat="server" controlname="chkAutoAssignment"></dnn:label></td>
		        <td width="325"><asp:checkbox id="chkAutoAssignment" runat="server"></asp:checkbox></td>
	        </tr>
        </table>
        <br>
        <dnn:sectionhead id="dshAdvanced" runat="server"
          cssclass="Head"
          includerule="True"
          isexpanded="False"
          resourcekey="AdvancedSettings"
          section="tblAdvanced"
          text="Advanced Settings"
          />
        <table id="tblAdvanced" runat="server" cellspacing="0" cellpadding="2" summary="Advanced Settings Design Table"  width="525" border="0">
          <tr>
            <td colspan="2"><asp:label id="lblAdvancedSettingsHelp" resourcekey="AdvancedSettingsHelp" cssclass="Normal" runat="server" enableviewstate="False"></asp:label></td>
          </tr>
	        <tr valign="top">
            <td class="SubHead" width="150"><dnn:label id="plServiceFee" resourcekey="ServiceFee" runat="server" controlname="txtServiceFee" suffix=":"></dnn:label></td>
		        <td width="325">
			        <asp:textbox id="txtServiceFee" cssclass="NormalTextBox" width="100" columns="30" maxlength="50" runat="server" />			
			        <asp:comparevalidator id="valServiceFee1" resourcekey="valServiceFee1" runat="server" controltovalidate="txtServiceFee" errormessage="<br>Service Fee Value Entered Is Not Valid" display="Dynamic" operator="DataTypeCheck" type="Currency" cssclass="NormalRed"></asp:comparevalidator>
			        <asp:comparevalidator id="valServiceFee2" resourcekey="valServiceFee2" runat="server" controltovalidate="txtServiceFee" errormessage="<br>Service Fee Must Be Greater Than or Equal to Zero" display="Dynamic" operator="GreaterThanEqual" valuetocompare="0" cssclass="NormalRed"></asp:comparevalidator>
		        </td>
	        </tr>
	        <tr valign="top">
            <td class="SubHead" width="150"><dnn:label id="plBillingPeriod" resourcekey="BillingPeriod" runat="server" controlname="txtBillingPeriod" suffix=":"></dnn:label></td>
		        <td width="325">
			        <asp:textbox id="txtBillingPeriod" cssclass="NormalTextBox" width="100" columns="30" maxlength="50" runat="server" />&nbsp;&nbsp;
			        <asp:dropdownlist id="cboBillingFrequency" datatextfield="text" datavaluefield="value" runat="server" cssclass="NormalTextBox" width="200px"></asp:dropdownlist>
			        <asp:comparevalidator id="valBillingPeriod1" resourcekey="valBillingPeriod1" runat="server" controltovalidate="txtBillingPeriod" errormessage="<br>Billing Period Value Entered Is Not Valid" display="Dynamic" operator="DataTypeCheck" type="Integer" cssclass="NormalRed"></asp:comparevalidator>
			        <asp:comparevalidator id="valBillingPeriod2" resourcekey="valBillingPeriod2" runat="server" controltovalidate="txtBillingPeriod" errormessage="<br>Billing Period Must Be Greater Than or Equal to Zero" display="Dynamic" operator="GreaterThan" valuetocompare="0" cssclass="NormalRed"></asp:comparevalidator>
		        </td>
	        </tr>
	        <tr valign="top">
            <td class="SubHead" width="150"><dnn:label id="plTrialFee" resourcekey="TrialFee" runat="server" controlname="txtTrialFee" suffix=":"></dnn:label></td>
		        <td width="325">
			        <asp:textbox id="txtTrialFee" cssclass="NormalTextBox" width="325" columns="30" maxlength="50" runat="server" />			
			        <asp:comparevalidator id="valTrialFee1" resourcekey="valTrialFee1" runat="server" controltovalidate="txtTrialFee" errormessage="<br>Trial Fee Value Entered Is Not Valid" display="Dynamic" operator="DataTypeCheck" type="Currency" cssclass="NormalRed"></asp:comparevalidator>
			        <asp:comparevalidator id="valTrialFee2" resourcekey="valTrialFee2" runat="server" controltovalidate="txtTrialFee" errormessage="<br>Trial Fee Must Be Greater Than Zero" display="Dynamic" operator="GreaterThanEqual" valuetocompare="0" cssclass="NormalRed"></asp:comparevalidator>
		        </td>
	        </tr>
	        <tr valign="top">
            <td class="SubHead" width="150"><dnn:label id="plTrialPeriod" resourcekey="TrialPeriod" runat="server" controlname="txtTrialPeriod" suffix=":"></dnn:label></td>
		        <td width="325">
			        <asp:textbox id="txtTrialPeriod" cssclass="NormalTextBox" width="100" columns="30" maxlength="50" runat="server" />&nbsp;&nbsp;
			        <asp:dropdownlist id="cboTrialFrequency" datatextfield="text" datavaluefield="value" runat="server" cssclass="NormalTextBox" width="200px"></asp:dropdownlist>
			        <asp:comparevalidator id="valTrialPeriod1" resourcekey="valTrialPeriod1" runat="server" controltovalidate="txtTrialPeriod" errormessage="<br>Trial Period Value Entered Is Not Valid" display="Dynamic" operator="DataTypeCheck" type="Integer" cssclass="NormalRed"></asp:comparevalidator>
			        <asp:comparevalidator id="valTrialPeriod2" resourcekey="valTrialPeriod2" runat="server" controltovalidate="txtTrialPeriod" errormessage="<br>Trial Period Must Be Greater Than Zero" display="Dynamic" operator="GreaterThan" valuetocompare="0" cssclass="NormalRed"></asp:comparevalidator>
		        </td>
	        </tr>
        </table>
      </asp:panel>
    </td>
  </tr>
</table>
<p>
	<asp:linkbutton id="cmdUpdate" resourcekey="cmdUpdate" runat="server" cssclass="CommandButton" text="Update" borderstyle="none" />
	&nbsp;
	<asp:linkbutton id="cmdCancel" resourcekey="cmdCancel" runat="server" cssclass="CommandButton" text="Cancel" causesvalidation="False" borderstyle="none" />
	&nbsp;
	<asp:linkbutton id="cmdDelete" resourcekey="cmdDelete" runat="server" cssclass="CommandButton" text="Delete" causesvalidation="False" borderstyle="none" />
	&nbsp;
	<asp:linkbutton id="cmdManage" resourcekey="cmdManage" runat="server" cssclass="CommandButton" text="Manage Users" causesvalidation="False" />
</p>

<asp:panel id="pnlNotImplemented" runat="server" visible="False">
  <table cellspacing="0" cellpadding="0" border="0" summary="Edit Roles Design Table">
	  <tr vAlign="top">
		  <td class="SubHead"><label for="<%=txtEvery.ClientID%>">Push Content Every:</label></td>
		  <td>
			  <asp:textbox ID="txtEvery" runat="server" maxlength="3" Columns="3" cssclass="NormalTextBox"></asp:textbox>&nbsp;
			  <label style="display:none;" for="<%=cboPeriod.ClientID%>">Period</label>
			  <asp:DropDownList ID="cboPeriod" Runat="server" cssclass="NormalTextBox">
				  <asp:ListItem Value=""></asp:ListItem>
				  <asp:ListItem Value="D">Day(s)</asp:ListItem>
				  <asp:ListItem Value="W">Week(s)</asp:ListItem>
				  <asp:ListItem Value="M">Month(s)</asp:ListItem>
				  <asp:ListItem Value="Y">Year(s)</asp:ListItem>
			  </asp:DropDownList>
			  &nbsp;<span class="NormalRed">*Not Implemented*</span>
		  </td>
		  <td class="Normal"></td>
	  </tr>
	  <tr vAlign="top">
		  <td class="SubHead"><label for="<%=txtStartDate.ClientID%>">Start Date:</label>
		  </td>
		  <td>
			  <asp:textbox id="txtStartDate" runat="server" maxlength="20" Columns="20" cssclass="NormalTextBox"></asp:textbox>
			  &nbsp;<span class="NormalRed">*Not Implemented*</span>
		  </td>
	  </tr>
  </table>
</asp:panel>

