<%@ Register TagPrefix="dnn" TagName="Label" Src="~/controls/LabelControl.ascx" %>
<%@ Register TagPrefix="dnn" TagName="SectionHead" Src="~/controls/SectionHeadControl.ascx" %>
<%@ Register TagPrefix="dnntv" Namespace="DotNetNuke.UI.WebControls" Assembly="DotNetNuke.WebControls" %>
<%@ Control Language="vb" AutoEventWireup="false" Codebehind="LanguageEditor.ascx.vb" Inherits="DotNetNuke.Services.Localization.LanguageEditor" %>
<style type="text/css">  
.Pending { BORDER-LEFT-COLOR: red; BORDER-BOTTOM-COLOR: red; BORDER-TOP-STYLE: solid; BORDER-TOP-COLOR: red; BORDER-RIGHT-STYLE: solid; BORDER-LEFT-STYLE: solid; BORDER-RIGHT-COLOR: red; BORDER-BOTTOM-STYLE: solid }  
</style>
<TABLE id="Table2" cellSpacing="5" width="100%" border="0">
	<TR>
		<TD vAlign="top" noWrap>
			<P><asp:panel id="Panel1" runat="server" Width="195px">
					<dnntv:DNNTree id="DNNTree" runat="server" DefaultNodeCssClassOver="Normal" CssClass="Normal" DefaultNodeCssClass="Normal"></dnntv:DNNTree>
				</asp:panel></P>
		</TD>
		<TD vAlign="top">
			<P>
				<TABLE id="Table1" cellSpacing="1" cellPadding="1" border="0">
					<TR>
						<TD class="SubHead" vAlign="top"><asp:label id="lblSelected" runat="server" resourcekey="SelectedFile" CssClass="Normal">Selected Resource File:</asp:label></TD>
						<TD vAlign="top"><asp:label id="lblResourceFile" runat="server" text="Selected Resource File:" Font-Bold="True"
								CssClass="Normal">Selected Resource File:</asp:label></TD>
					</TR>
					<TR>
						<TD class="SubHead" vAlign="top"><dnn:label id="lbLocales" runat="server" controlname="cboLocales" text="Available Locales"></dnn:label></TD>
						<TD vAlign="top"><asp:dropdownlist id="cboLocales" runat="server" Width="300px" DataTextField="name" DataValueField="key"
								AutoPostBack="True"></asp:dropdownlist></TD>
					</TR>
					<tr>
						<td class="SubHead" vAlign="top" colspan="2">
							<asp:CheckBox id="chkHighlight" runat="server" Text="Highlight Pending Translations" AutoPostBack="True"
								TextAlign="Left" resourcekey="Highlight"></asp:CheckBox>
						</td>
					</tr>
				</TABLE>
			</P>
			<P><asp:panel id="pnlMissing" runat="server" Wrap="True" Visible="False">
					<asp:label id="lblMissing" runat="server" CssClass="Normal" resourcekey="MissingEntries">System Default resource file contains some entries not present in current localized file. This can lead to some values not being translated.</asp:label>
					<BR>
					<asp:linkbutton id="cmdAddMissing" runat="server" CssClass="CommandButton" resourcekey="cmdAddMissing"
						CausesValidation="false">Add Missing Entries</asp:linkbutton>
				</asp:panel>
				<asp:panel id="pnlConfirm" runat="server" Wrap="True" Visible="False">
<asp:label id="lblConfirm" runat="server" CssClass="Normal" resourcekey="ConfirmMessage">You are about to create a custom localized file for this portal. Are you sure you want to overwrite default localized file?</asp:label>&nbsp; 
<asp:linkbutton id="cmdYes" runat="server" CssClass="CommandButton" resourcekey="Yes" CausesValidation="false">Yes</asp:linkbutton>
				</asp:panel></P>
			<P><asp:datagrid id="dgEditor" runat="server" AutoGenerateColumns="False" CellPadding="3" GridLines="None"
					CssClass="Normal">
					<ItemStyle VerticalAlign="Top"></ItemStyle>
					<HeaderStyle Font-Bold="True"></HeaderStyle>
					<Columns>
						<asp:TemplateColumn>
							<ItemTemplate>
								<TABLE cellSpacing="2" cellPadding="0" width="100%" border="0">
									<TR>
										<TD width="100%" bgColor="silver" colSpan="3">
											<asp:label id="Label3" runat="server" CssClass="NormalBold" resourcekey="ResourceName" Font-Bold="True">
												Resource name:</asp:label>
											<asp:Label id="lblName" runat="server" CssClass="Normal">
												<%# DataBinder.Eval(Container, "DataItem.key") %>
											</asp:Label></TD>
									</TR>
									<TR>
										<TD width="300">
											<asp:label id="Label4" runat="server" CssClass="NormalBold" resourcekey="Value" Font-Bold="True">
												Localized Value</asp:label></TD>
										<TD></TD>
										<TD width="100%">
											<TABLE border="0">
												<TR>
													<TD>
														<dnn:sectionhead id=dshDef runat="server" text="" includerule="False" section="divDef" cssclass="Normal" IsExpanded='<%# boolean.parse(FormatText(DataBinder.Eval(Container, "DataItem.key")).length < 150)  %>'>
														</dnn:sectionhead></TD>
													<TD>
														<asp:label id="Label5" runat="server" CssClass="NormalBold" resourcekey="DefaultValue" Font-Bold="True">
												Default Value</asp:label></TD>
												</TR>
											</TABLE>
										</TD>
									</TR>
									<TR>
										<TD vAlign="top" width="300">
											<asp:TextBox id=txtValue Width="300px" runat="server" Height='<%# FormatHeight(DataBinder.Eval(Container, "DataItem.value")) %>' TextMode="MultiLine" Text='<%# server.htmldecode(DataBinder.Eval(Container, "DataItem.value")) %>' CssClass='<%# FormatStyle(DataBinder.Eval(Container, "DataItem.key"),DataBinder.Eval(Container, "DataItem.value")) %>'>
											</asp:TextBox></TD>
										<TD vAlign="top" noWrap>
											<asp:HyperLink id=lnkEdit runat="server" CssClass="CommandButton" NavigateUrl='<%# OpenFullEditor(DataBinder.Eval(Container, "DataItem.key")) %> '>
												<asp:Image runat="server" AlternateText="Edit" ID="imgEdit" ImageUrl="~/images/uprt.gif" resourcekey="cmdEdit"></asp:Image>
											</asp:HyperLink>&nbsp;
										</TD>
										<TD vAlign="top" width="100%">
											<TABLE id="divDef" cellSpacing="0" cellPadding="0" width="100%" border="0" runat="server">
												<TR>
													<TD>
														<asp:Label id="Label1" runat="server" CssClass="Normal">
															<%# FormatText(DataBinder.Eval(Container, "DataItem.key")) %>
														</asp:Label></TD>
												</TR>
											</TABLE>
										</TD>
									</TR>
								</TABLE>
							</ItemTemplate>
						</asp:TemplateColumn>
						<asp:BoundColumn Visible="False" DataField="key"></asp:BoundColumn>
					</Columns>
				</asp:datagrid></P>
			<P><asp:linkbutton id="cmdUpdate" runat="server" CssClass="CommandButton" resourcekey="cmdUpdate">Update</asp:linkbutton>&nbsp;<asp:linkbutton id="cmdCancel" runat="server" CausesValidation="false" CssClass="CommandButton"
					resourcekey="cmdCancel">Cancel</asp:linkbutton>&nbsp;<asp:linkbutton id="cmdDelete" runat="server" CausesValidation="false" CssClass="CommandButton"
					resourcekey="cmdDelete">Delete</asp:linkbutton></P>
		</TD>
	</TR>
</TABLE>
