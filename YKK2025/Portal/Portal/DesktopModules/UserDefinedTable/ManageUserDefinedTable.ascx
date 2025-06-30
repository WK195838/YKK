<%@ Control language="vb" CodeBehind="ManageUserDefinedTable.ascx.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.Modules.UserDefinedTable.ManageUserDefinedTable" %>
<%@ Register TagPrefix="dnn" TagName="Label" Src="~/controls/LabelControl.ascx" %>
<asp:datagrid id="grdFields" runat="server" cssclass="Normal" oncancelcommand="grdFields_CancelEdit"
	onupdatecommand="grdFields_Update" oneditcommand="grdFields_Edit" ondeletecommand="grdFields_Delete"
	onitemcommand="grdFields_Move" onitemdatabound="grdFields_Item_Bound" autogeneratecolumns="False"
	cellspacing="0" cellpadding="2" BorderWidth="0" datakeyfield="UserDefinedFieldId" summary="This table is used to define your own table fields of various types such as dates, integer and text fields.">
	<columns>
		<asp:templatecolumn>
			<itemtemplate>
				<asp:imagebutton id="cmdEditUserDefinedField" runat="server" causesvalidation="false" commandname="Edit"
					imageurl="~/images/edit.gif" alternatetext="Edit" resourcekey="cmdEdit"></asp:imagebutton>
				<asp:imagebutton id="cmdDeleteUserDefinedField" runat="server" causesvalidation="false" commandname="Delete"
					imageurl="~/images/delete.gif" alternatetext="Delete" resourcekey="cmdDelete"></asp:imagebutton>
			</itemtemplate>
			<edititemtemplate>
				<asp:imagebutton id="cmdSaveUserDefinedField" runat="server" causesvalidation="false" commandname="Update"
					imageurl="~/images/save.gif" alternatetext="Save" resourcekey="Save"></asp:imagebutton>
				<asp:imagebutton id="cmdCancelUserDefinedField" runat="server" causesvalidation="false" commandname="Cancel"
					imageurl="~/images/cancel.gif" alternatetext="Cancel" resourcekey="cmdCancel"></asp:imagebutton>
			</edititemtemplate>
		</asp:templatecolumn>
		<asp:templatecolumn headertext="Visible" itemstyle-cssclass="Normal" headerstyle-cssclass="NormalBold"
			headerstyle-horizontalalign="Center" itemstyle-horizontalalign="Center">
			<itemtemplate>
				<asp:Image runat="server" ImageUrl='<%# IIf(DataBinder.Eval(Container.DataItem, "Visible") = True, "~/images/checked.gif", "~/images/unchecked.gif") %>' ID="Image2"/>
			</itemtemplate>
			<edititemtemplate>
				<asp:label id="lblCheckBox2" runat="server" />
				<asp:CheckBox runat="server" id="Checkbox2" Checked='<%# IIf(DataBinder.Eval(Container.DataItem, "Visible") = True, "True", "False") %>' />
			</edititemtemplate>
		</asp:templatecolumn>
		<asp:templatecolumn headertext="Title" itemstyle-cssclass="Normal" headerstyle-cssclass="NormalBold">
			<itemtemplate>
				<asp:label runat="server" Text='<%# DataBinder.Eval(Container.DataItem, "FieldTitle") %>' ID="Label1"/>
			</itemtemplate>
			<edititemtemplate>
				<asp:label id="lblFieldTitle" runat="server" />
				<asp:TextBox runat="server" id="txtFieldTitle" Columns="30" MaxLength="50" Text='<%# DataBinder.Eval(Container.DataItem, "FieldTitle") %>' />
			</edititemtemplate>
		</asp:templatecolumn>
		<asp:templatecolumn headertext="Type" itemstyle-cssclass="Normal" headerstyle-cssclass="NormalBold">
			<itemtemplate>
				<asp:label runat="server" Text='<%# GetFieldTypeName(DataBinder.Eval(Container.DataItem, "FieldType")) %>' ID="Label2"/>
			</itemtemplate>
			<edititemtemplate>
				<asp:label id="lblFieldType" runat="server" />
				<asp:DropDownList ID="cboFieldType" Runat="server" CssClass="NormalTextBox" SelectedIndex='<%# GetFieldTypeIndex(DataBinder.Eval(Container.DataItem, "FieldType")) %>'>
					<asp:listitem resourcekey="Text" value="String">Text</asp:listitem>
					<asp:listitem resourcekey="Integer" value="Int32">Integer</asp:listitem>
					<asp:listitem resourcekey="Decimal" value="Decimal">Decimal</asp:listitem>
					<asp:listitem resourcekey="Date" value="DateTime">Date</asp:listitem>
					<asp:listitem resourcekey="TrueFalse" value="Boolean">True/False</asp:listitem>
				</asp:DropDownList>
			</edititemtemplate>
		</asp:templatecolumn>
		<asp:templatecolumn>
			<itemtemplate>
				<asp:imagebutton id="cmdMoveUserDefinedFieldUp" runat="server" causesvalidation="false" commandname="Item"
					commandargument="Up" imageurl="~/images/up.gif" alternatetext="Move Field Up" resourcekey="MoveUp"></asp:imagebutton>
			</itemtemplate>
		</asp:templatecolumn>
		<asp:templatecolumn>
			<itemtemplate>
				<asp:imagebutton id="cmdMoveUserDefinedFieldDown" runat="server" causesvalidation="false" commandname="Item"
					commandargument="Down" imageurl="~/images/dn.gif" alternatetext="Move Field Down" resourcekey="MoveDown"></asp:imagebutton>
			</itemtemplate>
		</asp:templatecolumn>
	</columns>
</asp:datagrid>
<br>
<TABLE id="Table1" cellSpacing="2" cellPadding="1">
	<TR>
		<TD noWrap class="SubHead">
			<dnn:label id="plSort" runat="server" controlname="cboSortField" Suffix=":"></dnn:label></TD>
		<TD noWrap>
			<asp:dropdownlist id="cboSortField" runat="server" cssclass="NormalTextBox" datatextfield="FieldTitle"
				datavaluefield="UserDefinedFieldId" autopostback="True"></asp:dropdownlist></TD>
		<TD noWrap class="SubHead">
			<dnn:label id="plOrder" runat="server" controlname="cboSortOrder" Suffix=":"></dnn:label></TD>
		<TD noWrap>
			<asp:dropdownlist id="cboSortOrder" cssclass="NormalTextBox" runat="server" autopostback="True">
				<asp:listitem resourcekey="Not_Specified" value="">
					<not specified>
				</asp:listitem>
				<asp:listitem resourcekey="Ascending" value="ASC">Ascending</asp:listitem>
				<asp:listitem resourcekey="Descending" value="DESC">Descending</asp:listitem>
			</asp:dropdownlist></TD>
	</TR>
</TABLE>
<P>
	<asp:linkbutton class="CommandButton" id="cmdAddField" runat="server" resourcekey="cmdAddField"
		borderstyle="none" causesvalidation="False" text="Add New Column"></asp:linkbutton>&nbsp;
	<asp:linkbutton class="CommandButton" id="cmdCancel" runat="server" resourcekey="cmdCancel" borderstyle="none"
		causesvalidation="False" text="Cancel"></asp:linkbutton></P>
