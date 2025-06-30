<%@ Control language="vb" CodeBehind="EditModuleDefinition.ascx.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.Modules.Admin.ModuleDefinitions.EditModuleDefinition" %>
<%@ Register TagPrefix="dnn" TagName="Label" Src="~/controls/LabelControl.ascx" %>
<%@ Register TagPrefix="Portal" TagName="DualList" Src="~/controls/DualListControl.ascx" %>
<table cellspacing="0" cellpadding="4" border="0" summary="Module Definitions Design Table">
    <tr>
        <td>
            <table id="tabModule" runat="server" cellspacing="0" cellpadding="4" border="0" summary="Module Definitions Design Table">
                <tr>
                    <td class="SubHead" width="150"><dnn:label id="plModuleName" text="Module Name:" controlname="txtModuleName" runat="server" /></td>
                    <td>
                        <asp:textbox id="txtModuleName" cssclass="NormalTextBox" width="390" columns="30" maxlength="150" runat="server" enabled="False" />
                        <asp:requiredfieldvalidator id="valModuleName" display="Dynamic" resourcekey="valModuleName.ErrorMessage" errormessage="<br>You must enter a Name for the Module." controltovalidate="txtModuleName" runat="server" />
                    </td>
                </tr>
                <tr>
                    <td class="SubHead" width="150"><dnn:label id="plFolderName" text="Folder Name:" controlname="txtFolderName" runat="server" /></td>
                    <td>
                        <asp:textbox id="txtFolderName" cssclass="NormalTextBox" width="390" columns="30" maxlength="150" runat="server" enabled="False" />
                        <asp:requiredfieldvalidator id="valFolderName" display="Dynamic" resourcekey="valFolderName.ErrorMessage" errormessage="<br>You must enter a Folder Name for the location of the Module's files." controltovalidate="txtFolderName" runat="server" Width="335px"/>
                    </td>
                </tr>
                <tr>
                    <td class="SubHead" width="150"><dnn:label id="plFriendlyName" text="Friendly Name:" controlname="txtFriendlyName" runat="server" /></td>
                    <td>
                        <asp:textbox id="txtFriendlyName" cssclass="NormalTextBox" width="390" columns="30" maxlength="150" runat="server" />
                    </td>
                </tr>
                <tr>
                    <td class="SubHead" width="150" valign="top"><dnn:label id="plDescription" text="Description:" controlname="txtDescription" runat="server" /></td>
                    <td><asp:textbox id="txtDescription" cssclass="NormalTextBox" width="390" columns="30" textmode="MultiLine" rows="10" maxlength="2000" runat="server" /></td>
                </tr>
                <tr>
                    <td class="SubHead" width="150" valign="top"><dnn:label id="plVersion" text="Version:" controlname="txtVersion" runat="server" /></td>
                    <td><asp:textbox id="txtVersion" cssclass="NormalTextBox" width="390" columns="30" maxlength="150" runat="server" enabled="False" /></td>
                </tr>
                <tr>
                    <td class="SubHead" width="150"><dnn:label id="plBusinessClass" text="Business Controller Class:" controlname="txtBusinessClass" runat="server" /></td>
                    <td>
                        <asp:textbox id="txtBusinessClass" cssclass="NormalTextBox" width="390" columns="30" maxlength="150" runat="server" enabled="False" />
                    </td>
                </tr>
                <tr>
                    <td class="SubHead" width="150" valign="top"><dnn:label id="plPremium" text="Premium?" controlname="chkPremium" runat="server" /></td>
                    <td>
						<asp:checkbox id="chkPremium" runat="server" cssclass="NormalTextBox" AutoPostBack="True"></asp:checkbox><br>
						<portal:duallist id="ctlPortals" runat="server" ListBoxWidth="130" ListBoxHeight="130" DataValueField="PortalID" DataTextField="PortalName" />
					</td>
                </tr>
            </table>
            <p>
                <asp:linkbutton id="cmdUpdate" resourcekey="cmdUpdate" text="Update" runat="server" class="CommandButton" borderstyle="none" />
                &nbsp;
                <asp:linkbutton id="cmdCancel" resourcekey="cmdCancel" text="Cancel" causesvalidation="False" runat="server" class="CommandButton" borderstyle="none" />
                &nbsp;
                <asp:linkbutton id="cmdDelete" resourcekey="cmdDelete" text="Delete" causesvalidation="False" runat="server" class="CommandButton" borderstyle="none" />
            </p>
            <hr>
            <table id="tabDefinitions" runat="server" cellspacing="0" cellpadding="4" border="0" summary="Module Definitions Design Table">
                <tr>
                    <td class="SubHead" width="150" valign="top"><dnn:label id="plDefinitions" text="Definitions:" controlname="cboDefinitions" runat="server" /></td>
                    <td>
                        <asp:dropdownlist id="cboDefinitions" runat="server" width="290px" cssclass="NormalTextBox" datatextfield="FriendlyName" datavaluefield="ModuleDefId" autopostback="True"></asp:dropdownlist>
                        &nbsp;&nbsp;
                        <asp:linkbutton id="cmdDeleteDefinition" resourcekey="cmdDeleteDefinition" text="Delete Definition" runat="server" class="CommandButton" borderstyle="none" />
                    </td>
                </tr>
                <tr>
                    <td class="SubHead" width="150" valign="top"><dnn:label id="plDefinition" text="New Definition:" controlname="txtDefinition" runat="server" /></td>
                    <td>
                        <asp:textbox id="txtDefinition" cssclass="NormalTextBox" width="290px" columns="30" maxlength="128" runat="server" />
                        &nbsp;&nbsp;
                        <asp:linkbutton id="cmdAddDefinition" resourcekey="cmdAddDefinition" text="Add Definition" runat="server" class="CommandButton" borderstyle="none" />
                    </td>
                </tr>
            </table>
            <hr>
            <table id="tabCache" runat="server" cellspacing="0" cellpadding="4" border="0" width="100%" summary="Module Definitions Design Table">
                <tr>
                    <td class="SubHead" width="150" valign="top"><dnn:label id="plCacheTime" text="Default Cache Time:" controlname="txtCacheTime" runat="server" /></td>
                    <td>
                        <asp:textbox id="txtCacheTime" cssclass="NormalTextBox" width="290px" columns="30" maxlength="128" runat="server" />
                        &nbsp;&nbsp;
                        <asp:linkbutton id="cmdUpdateCacheTime" resourcekey="cmdUpdateCacheTime" text="Update Cache Time" runat="server" class="CommandButton" borderstyle="none" />
                    </td>
                </tr>
            </table>
            <hr>
            <table id="tabControls" runat="server" cellspacing="0" cellpadding="4" border="0" width="100%" summary="Module Definitions Design Table">
                <tr>
                    <td colspan="2">
                        <asp:datagrid id="grdControls" runat="server" width="100%" border="0" cellspacing="3" autogeneratecolumns="false" enableviewstate="true" summary="Module Controls Design Table">
                            <columns>
                                <asp:templatecolumn>
                                    <itemstyle width="20px"></itemstyle>
                                    <itemtemplate>
                                        <asp:HyperLink id=Hyperlink1 runat="server" NavigateUrl='<%# FormatURL("modulecontrolid",DataBinder.Eval(Container.DataItem,"ModuleControlId")) %>'>
                                            <asp:image imageurl="~/images/edit.gif" alternatetext="Edit" runat="server" id="Hyperlink1Image" />
                                        </asp:hyperlink>
                                    </itemtemplate>
                                </asp:templatecolumn>
                                <asp:boundcolumn datafield="ControlKey" headertext="Control">
                                    <headerstyle cssclass="NormalBold"></headerstyle>
                                    <itemstyle cssclass="Normal"></itemstyle>
                                </asp:boundcolumn>
                                <asp:boundcolumn datafield="ControlTitle" headertext="Title">
                                    <headerstyle cssclass="NormalBold"></headerstyle>
                                    <itemstyle cssclass="Normal"></itemstyle>
                                </asp:boundcolumn>
                                <asp:boundcolumn datafield="ControlSrc" headertext="Source">
                                    <headerstyle cssclass="NormalBold"></headerstyle>
                                    <itemstyle cssclass="Normal"></itemstyle>
                                </asp:boundcolumn>
                            </columns>
                        </asp:datagrid>
                        <p>
                            <asp:linkbutton id="cmdAddControl" resourcekey="cmdAddControl" text="Add Control" runat="server" class="CommandButton" borderstyle="none" />
                        </p>
                    </td>
                </tr>
            </table>
        </td>
    </tr>
</table>
