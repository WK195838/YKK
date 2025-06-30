<%@ Control language="vb" CodeBehind="WebUpload.ascx.vb" AutoEventWireup="false" Explicit="True" Inherits="DotNetNuke.Modules.Admin.FileSystem.WebUpload" %>
<table class="Settings" cellSpacing="2" cellPadding="2" summary="Web Upload Design Table">
	<tr>
		<td vAlign="top" width="560"><asp:panel id="pnlUpload" visible="True" cssclass="WorkPanel" runat="server">
				<TABLE class="Settings" id="tblUpload" height="151" cellSpacing="2" cellPadding="2" summary="Web Upload Design Table"
					runat="server">
					<TR>
						<TD class="Head" colSpan="2" align="center">
							<asp:label id="lblUploadType" runat="server"></asp:label><br><hr>
						</TD>
					</TR>
					<TR id="trRoot" runat="server" visible="false">
						<TD width="100">
							<asp:label id="lblRootType" runat="server" cssclass="SubHead"></asp:label></TD>
						<TD width="425">
							<asp:label id="lblRootFolder" runat="server" cssclass="Normal"></asp:label></TD>
					</TR>
					<TR>
						<TD colSpan="2">&nbsp;</TD>
					</TR>
					<TR>
						<TD align="center" colSpan="2"><LABEL style="DISPLAY: none" 
            for="<%=cmdBrowse.ClientID%>">Browse Files</LABEL><INPUT id="cmdBrowse" type="file" size="50" name="cmdBrowse" runat="server">&nbsp;&nbsp;
							<asp:linkbutton id="cmdAdd" runat="server" cssclass="CommandButton" Resourcekey="Add">Add</asp:linkbutton></TD>
					</TR>
					<TR id="trFolders" runat="server" visible="false">
						<TD align="center" colSpan="2">
							<asp:DropDownList id="ddlFolders" runat="server" Width="525px">
								<asp:ListItem Value="Root">Root</asp:ListItem>
								<asp:ListItem Value="Files\">Files</asp:ListItem>
								<asp:ListItem Value="Files\SubFolder\">Files\SubFolder</asp:ListItem>
							</asp:DropDownList></TD>
					</TR>
					<TR>
						<TD align="center" colSpan="2"><LABEL style="DISPLAY: none" 
            for="<%=lstFiles.ClientID%>">First Files</LABEL>
							<asp:listbox id="lstFiles" runat="server" cssclass="NormalTextBox" width="525px" rows="5"></asp:listbox></TD>
					</TR>
					<TR id="trUnzip" runat="server" visible="false">
						<TD colSpan="2">
							<asp:checkbox id="chkUnzip" runat="server" cssclass="Normal" textalign="Right" text="Decompress ZIP Files?"
								resourcekey="Decompress"></asp:checkbox></TD>
					</TR>
					<TR>
						<TD align="left" colSpan="2">
							<asp:linkbutton id="cmdUpload" runat="server" cssclass="CommandButton" resourcekey="cmdUpload">Upload File(s)</asp:linkbutton>&nbsp;&nbsp;
							<asp:linkbutton id="cmdRemove" runat="server" cssclass="CommandButton" resourcekey="cmdRemove">Remove</asp:linkbutton>&nbsp;&nbsp;
							<asp:linkbutton id="cmdCancel" runat="server" cssclass="CommandButton" resourcekey="cmdCancel">Cancel</asp:linkbutton></TD>
					</TR>
					<TR>
						<TD align="left" colSpan="2">
							<asp:label id="lblMessage" runat="server" cssclass="Normal" width="500px" enableviewstate="False"></asp:label></TD>
					</TR>
				</TABLE>
				<BR>
				<TABLE id="tblLogs" cellSpacing="0" cellPadding="0" summary="Resource Upload Logs Table"
					runat="server" visible="False">
					<TR>
						<TD>
							<asp:label id="lblLogTitle" runat="server" resourcekey="LogTitle">Resource Upload Logs</asp:label></TD>
					</TR>
					<TR>
						<TD>
							<asp:linkbutton id="cmdReturn1" runat="server" cssclass="CommandButton" resourcekey="cmdReturn">Return to File Manager</asp:linkbutton></TD>
					</TR>
					<TR>
						<TD>&nbsp;</TD>
					</TR>
					<TR>
						<TD>
							<asp:placeholder id="phPaLogs" runat="server"></asp:placeholder></TD>
					</TR>
					<TR>
						<TD>&nbsp;</TD>
					</TR>
					<TR>
						<TD>
							<asp:linkbutton id="cmdReturn2" runat="server" cssclass="CommandButton" resourcekey="cmdReturn">Return to File Manager</asp:linkbutton></TD>
					</TR>
				</TABLE>
			</asp:panel></td>
	</tr>
</table>
