<%@ Control Language="vb" AutoEventWireup="false" Codebehind="ListEditor.ascx.vb" Inherits="DotNetNuke.Common.Lists.ListEditor" %>
<%@ Register TagPrefix="dnntv" Namespace="DotNetNuke.UI.WebControls" Assembly="DotNetNuke.WebControls" %>
<%@ Register TagPrefix="dnn" TagName="SectionHead" Src="~/controls/SectionHeadControl.ascx" %>
<%@ Register TagPrefix="dnn" TagName="Label" Src="~/controls/LabelControl.ascx" %>
<TABLE id="tblList" cellSpacing="5" width="660" border="0">
	<TR>
		<TD vAlign="top" width="240"><dnn:sectionhead id="dshTree" includerule="True" section="tblTree" cssclass="Head" text="Lists" resourcekey="Lists"
				runat="server"></dnn:sectionhead>
			<table id="tblTree" cellSpacing="0" cellPadding="0" width="100%" border="0" runat="server">
				<tr>
					<td><dnntv:dnntree id="DNNTree" runat="server" DefaultNodeCssClassOver="Normal" DefaultNodeCssClass="Normal"
							cssClass="Normal"></dnntv:dnntree></td>
				</tr>
				<tr>
					<td><asp:linkbutton id="cmdAddList" text="Add List" resourcekey="AddList" runat="server" CssClass="CommandButton"></asp:linkbutton></td>
				</tr>
			</table>
		</TD>
		<td vAlign="top" width="420"><dnn:sectionhead id="dshDetails" includerule="True" section="tblDetails" cssclass="Head" text="List Entries"
				resourcekey="ListEntries" runat="server"></dnn:sectionhead>
			<table id="tblDetails" cellSpacing="0" cellPadding="0" width="100%" border="0" runat="server">
				<tr id="rowListDetails" runat="server">
					<td class="subhead" width="100%">
						<TABLE id="tblEntryInfo" cellSpacing="0" cellPadding="3" width="100%" border="0" runat="server">
							<TR id="rowListParent" runat="server">
								<TD class="SubHead" width="120"><dnn:label id="plListParent" text="Parent:" runat="server" controlname="lblListParent"></dnn:label></TD>
								<TD><asp:label id="lblListParent" cssclass="Normal" runat="server"></asp:label></TD>
							</TR>
							<TR>
								<TD class="SubHead" width="120"><dnn:label id="plListName" text="List Name:" runat="server" controlname="lblListName"></dnn:label></TD>
								<TD><asp:label id="lblListName" cssclass="Normal" runat="server"></asp:label></TD>
							</TR>
							<TR>
								<TD class="SubHead" width="120"><dnn:label id="plEntryCount" text="Total:" runat="server" controlname="lblEntryCount"></dnn:label></TD>
								<TD><asp:label id="lblEntryCount" cssclass="Normal" runat="server"></asp:label></TD>
							</TR>
							<TR id="rowListCommand" runat="server">
								<TD class="SubHead" colSpan="2"><asp:linkbutton id="cmdAddEntry" text="Add Entry" resourcekey="cmdAddEntry" runat="server" CssClass="CommandButton"></asp:linkbutton>&nbsp;
									<asp:linkbutton id="cmdDeleteList" cssclass="CommandButton" text="Delete List" resourcekey="cmdDeleteList"
										runat="server"></asp:linkbutton></TD>
							</TR>
						</TABLE>
						<hr noShade SIZE="1">
					</td>
				</tr>
				<tr id="rowEntryGrid" runat="server">
					<td class="subhead" width="100%"><asp:datagrid id="grdEntries" runat="server" DataKeyField="EntryID" AutoGenerateColumns="false"
							width="100%" CellPadding="3" Border="0">
							<Columns>
								<asp:TemplateColumn ItemStyle-Width="20">
									<ItemTemplate>
										<asp:ImageButton ID="btnEdit" ImageUrl="~/Images/edit.gif" Runat="server" CssClass="CommandButton"
											AlternateText="Edit" resourcekey="btnEdit.AlternateText" CommandName="edit"></asp:ImageButton>
									</ItemTemplate>
								</asp:TemplateColumn>
								<asp:BoundColumn HeaderText="Text" DataField="Text" ItemStyle-CssClass="Normal" HeaderStyle-Cssclass="NormalBold" />
								<asp:BoundColumn HeaderText="Value" DataField="Value" ItemStyle-CssClass="Normal" HeaderStyle-Cssclass="NormalBold"
									HeaderStyle-HorizontalAlign="center" ItemStyle-HorizontalAlign="Center" />
								<asp:TemplateColumn HeaderText="" ItemStyle-CssClass="Normal" HeaderStyle-Cssclass="NormalBold" HeaderStyle-Width="18px">
									<ItemTemplate>
										<asp:ImageButton ID="btnUp" Visible="<%# EnableSortOrder() %>" ImageUrl="~/Images/up.gif" Runat="server" CssClass="CommandButton" AlternateText="Move entry up" resourcekey="btnUp.AlternateText" CommandName="up">
										</asp:ImageButton>
									</ItemTemplate>
								</asp:TemplateColumn>
								<asp:TemplateColumn HeaderText="" ItemStyle-CssClass="Normal" HeaderStyle-Cssclass="NormalBold" HeaderStyle-Width="18px">
									<ItemTemplate>
										<asp:ImageButton ID="btnDown" Visible="<%# EnableSortOrder() %>" ImageUrl="~/Images/dn.gif" Runat="server" CssClass="CommandButton" AlternateText="Move entry down" resourcekey="btnDown.AlternateText" CommandName="down">
										</asp:ImageButton>
									</ItemTemplate>
								</asp:TemplateColumn>
							</Columns>
						</asp:datagrid></td>
				</tr>
				<tr id="rowEntryEdit" runat="server">
					<td class="subhead">
						<TABLE id="tblEntryEdit" cellSpacing="0" cellPadding="3" width="100%" border="0" runat="server">
							<TR id="rowParentKey" runat="server">
								<TD class="SubHead" width="160"><dnn:label id="plParentKey" text="Parent:" runat="server" controlname="txtParentKey"></dnn:label></TD>
								<TD><asp:textbox id="txtParentKey" cssclass="NormalTextBox" runat="server" width="240" maxlength="100"
										ReadOnly="True"></asp:textbox></TD>
							</TR>
							<TR id="rowSelectList" runat="server">
								<TD class="SubHead" width="160"><dnn:label id="plSelectList" text="Parent List:" runat="server" controlname="ddlSelectList"></dnn:label></TD>
								<TD><asp:dropdownlist id="ddlSelectList" cssclass="NormalTextBox" AutoPostBack="True" Width="240" Runat="server"></asp:dropdownlist></TD>
							</TR>
							<TR id="rowSelectParent" runat="server">
								<TD class="SubHead" width="160"><dnn:label id="plSelectParent" text="Parent Entry:" runat="server" controlname="ddlSelectParent"></dnn:label></TD>
								<TD><asp:dropdownlist id="ddlSelectParent" cssclass="NormalTextBox" Width="240" Runat="server" Enabled="False"></asp:dropdownlist></TD>
							</TR>
							<TR>
								<TD class="SubHead" width="160"><dnn:label id="plEntryName" text="Entry Name:" runat="server" controlname="txtEntryName"></dnn:label></TD>
								<TD><asp:textbox id="txtEntryName" cssclass="NormalTextBox" runat="server" width="240" maxlength="100"></asp:textbox><asp:textbox id="txtEntryID" cssclass="NormalTextBox" runat="server" width="0"></asp:textbox></TD>
							</TR>
							<TR>
								<TD class="SubHead" width="160"><dnn:label id="plEntryText" text="Entry Text:" runat="server" controlname="txtEntryText"></dnn:label></TD>
								<TD><asp:textbox id="txtEntryText" cssclass="NormalTextBox" runat="server" width="240" maxlength="100"></asp:textbox></TD>
							</TR>
							<TR>
								<TD class="SubHead" width="160"><dnn:label id="plEntryValue" text="Entry Value:" runat="server" controlname="txtEntryValue"></dnn:label></TD>
								<TD><asp:textbox id="txtEntryValue" cssclass="NormalTextBox" runat="server" width="240" maxlength="100"></asp:textbox></TD>
							</TR>
							<TR id="rowEnableSortOrder" runat="server">
								<TD class="SubHead" width="160"><dnn:label id="plEnableSortOrder" text="Enable Sort Order:" runat="server" controlname="chkEnableSortOrder"></dnn:label></TD>
								<TD><asp:checkbox id="chkEnableSortOrder" Runat="server"></asp:checkbox></TD>
							</TR>
							<tr>
								<TD class="SubHead" colSpan="2"><asp:linkbutton class="CommandButton" id="cmdSaveEntry" runat="server" resourcekey="cmdSave" text="Save"></asp:linkbutton>&nbsp;
									<asp:linkbutton id="cmdDelete" runat="server" resourcekey="cmdDeleteEntry" text="Delete Entry" cssclass="CommandButton"
										causesvalidation="False" borderstyle="none"></asp:linkbutton>&nbsp;
									<asp:linkbutton id="cmdCancel" runat="server" resourcekey="cmdCancel" text="Cancel" cssclass="CommandButton"
										causesvalidation="False" borderstyle="none"></asp:linkbutton></TD>
							</tr>
						</TABLE>
					</td>
				</tr>
			</table>
		</td>
	</TR>
</TABLE>
