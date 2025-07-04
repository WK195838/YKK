<%@ Control CodeBehind="LogViewer.ascx.vb" language="vb" AutoEventWireup="false" Explicit="true" Inherits="DotNetNuke.Modules.Admin.Log.LogViewer" targetSchema="http://schemas.microsoft.com/intellisense/ie5"%>
<%@ Register TagPrefix="dnn" TagName="Label" Src="~/controls/LabelControl.ascx" %>
<%@ Register TagPrefix="dnn" TagName="SectionHead" Src="~/controls/SectionHeadControl.ascx" %>
<%@ Register TagPrefix="dnnsc" Namespace="DotNetNuke.UI.WebControls" Assembly="DotNetNuke" %>
<STYLE type="text/css">  
	.Exception { COLOR: #ffffff; BACKGROUND-COLOR: #ff1414 }  
	.ItemCreated { COLOR: #ffffff; BACKGROUND-COLOR: #009900 }  
	.ItemUpdated { COLOR: #ffffff; BACKGROUND-COLOR: #009999 }  
	.ItemDeleted { COLOR: #000000; BACKGROUND-COLOR: #14ffff }  
	.OperationSuccess { COLOR: #ffffff; BACKGROUND-COLOR: #999900 }  
	.OperationFailure { COLOR: #ffffff; BACKGROUND-COLOR: #990000 }  
	.GeneralAdminOperation { COLOR: #ffffff; BACKGROUND-COLOR: #4d0099 }  
	.AdminAlert { COLOR: #ffffff; BACKGROUND-COLOR: #148aff }  
	.HostAlert { COLOR: #ffffff; BACKGROUND-COLOR: #ff8a14 }  
	#floater { PADDING-RIGHT: 0px; PADDING-LEFT: 0px; BACKGROUND: #ffffff; VISIBILITY: hidden; PADDING-BOTTOM: 0px; MARGIN: 0px; WIDTH: 150px; COLOR: #ffffff; PADDING-TOP: 0px; POSITION: absolute; HEIGHT: auto }  
</STYLE>
<!--
Other colors that match the scheme above:
#990000
#994D00
#999900
#4D9900
#99004D
#D60000
#FF1414
#009900
#990099
#14FFFF
#00D6D6
#00994D
#4D0099
#000099
#004D99
#009999
-->
<SCRIPT language="JavaScript" type="text/javascript">
<!--
function CheckExceptions()
{
    var j,isChecked = false;
    if (document.forms[0].item("Exception").length)
        {
            j=document.forms[0].item("Exception").length;
            for (var i=0;i<j;i++)
                {
                    if (document.forms[0].item("Exception")(i).checked==true)
                    {
                        isChecked = true;
                    }
                }
            if (isChecked!=true)
                {
                    alert('Please select at least one exception.');
                }
            return isChecked;
        }
    else 
        {
            if (document.forms[0].item("Exception").checked)
                return true;
            else
                {
                alert('Please select at least one exception.');
                return false;
                }
        }
}

function flipFlop(eTarget) {
    if (document.getElementById(eTarget).style.display=='')
    {
    	document.getElementById(eTarget).style.display='none';
    }
    else
    {
    	document.getElementById(eTarget).style.display='';
    }
}
  
//-->
</SCRIPT>
<ASP:PANEL id="pnlOptions" runat="server">
	<TABLE width="100%" border="0">
		<TR>
			<TD valign="top">
				<DNN:SECTIONHEAD id="dshSettings" runat="server" text="Viewer Settings" section="tblSettings" resourcekey="Settings"
					cssclass="Head"></DNN:SECTIONHEAD>
				<TABLE id="tblSettings" cellspacing="2" cellpadding="2" border="0" runat="server">
					<TR>
						<TD class="SubHead" noWrap align="left" width="110">
							<DNN:LABEL id="plPortalID" runat="server" suffix=":" controlname="ddlPortalid"></DNN:LABEL></TD>
						<TD width="60">
							<ASP:DROPDOWNLIST id="ddlPortalid" runat="server" autopostback="true"></ASP:DROPDOWNLIST></TD>
						<TD width="25">&nbsp;</TD>
						<TD class="SubHead" align="left" width="100">
							<DNN:LABEL id="plLogType" runat="server" suffix=":" controlname="ddlLogType"></DNN:LABEL></TD>
						<TD width="100">
							<ASP:DROPDOWNLIST id="ddlLogType" runat="server" autopostback="true"></ASP:DROPDOWNLIST></TD>
					</TR>
					<TR>
						<TD class="SubHead" noWrap align="left">
							<DNN:LABEL id="plRecordsPage" runat="server" suffix=":" resourcekey="Recordsperpage" cssclass="SubHead"></DNN:LABEL>
							</TD>
						<TD width="25"><ASP:DROPDOWNLIST id="ddlRecordsPerPage" runat="server" autopostback="True">
								<ASP:LISTITEM value="10">10</ASP:LISTITEM>
								<ASP:LISTITEM value="25">25</ASP:LISTITEM>
								<ASP:LISTITEM value="50">50</ASP:LISTITEM>
								<ASP:LISTITEM value="100">100</ASP:LISTITEM>
								<ASP:LISTITEM value="250">250</ASP:LISTITEM>
							</ASP:DROPDOWNLIST></TD>
						<TD width="25">&nbsp;</TD>
						<TD class="SubHead" colspan="2">
							<ASP:CHECKBOX id="chkColorCoding" runat="server" text="Color Coding On" resourcekey="ColorCoding"
								autopostback="true"></ASP:CHECKBOX></TD>
						<TD width="*">&nbsp;</TD>
					</TR>
					<TR>
						<TD colspan="5">
							<TABLE id="tblInstructions" width="100%" border="0" runat="server">
								<TR height="10">
									<TD></TD>
								</TR>
								<TR>
									<TD class="Normal" align="center">
										<ASP:LABEL id="lbClickRow" runat="server" resourcekey="ClickRow">Click on a row for details.</ASP:LABEL><BR>
									</TD>
								</TR>
								<TR>
									<TD class="NormalBold" align="center">
										<ASP:LABEL id="txtShowing" runat="server"></ASP:LABEL></TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
				</TABLE>
			</TD>
			<TD width="*">&nbsp;</TD>
			<TD class="SubHead" valign="top" align="left" width="230" rowspan="3">
				<ASP:PANEL id="pnlLegend" runat="server" horizontalalign="Center">
					<DNN:SECTIONHEAD id="dshLegend" runat="server" text="Color Coding Legend" section="tblLegend" resourcekey="Legend"
						cssclass="Head" isexpanded="False"></DNN:SECTIONHEAD>
					<TABLE id="tblLegend" cellspacing="2" cellpadding="2" bgcolor="#000000" border="0" runat="server">
						<TR>
							<TD bgcolor="#ffffff">
								<TABLE border="0">
									<TR>
										<TD>
											<TABLE cellspacing="1" cellpadding="3" bgcolor="#000000" border="0">
												<TR>
													<TD class="Normal" bgcolor="#ff1414"><IMG height="5" src="images/spacer.gif" width="5"></TD>
												</TR>
											</TABLE>
										</TD>
										<TD class="Normal">
											<ASP:LABEL id="Label1" runat="server" resourcekey="ExceptionCode">Exception</ASP:LABEL></TD>
									</TR>
									<TR>
										<TD>
											<TABLE cellspacing="1" cellpadding="3" bgcolor="#000000" border="0">
												<TR>
													<TD class="Normal" bgcolor="#009900"><IMG height="5" src="images/spacer.gif" width="5"></TD>
												</TR>
											</TABLE>
										</TD>
										<TD class="Normal">
											<ASP:LABEL id="Label2" runat="server" resourcekey="ItemCreatedCode">Item Created</ASP:LABEL></TD>
									</TR>
									<TR>
										<TD>
											<TABLE cellspacing="1" cellpadding="3" bgcolor="#000000" border="0">
												<TR>
													<TD class="Normal" bgcolor="#009999"><IMG height="5" src="images/spacer.gif" width="5"></TD>
												</TR>
											</TABLE>
										</TD>
										<TD class="Normal">
											<ASP:LABEL id="Label3" runat="server" resourcekey="ItemUpdatedCode">Item Updated</ASP:LABEL></TD>
									</TR>
									<TR>
										<TD>
											<TABLE cellspacing="1" cellpadding="3" bgcolor="#000000" border="0">
												<TR>
													<TD class="Normal" bgcolor="#14ffff"><IMG height="5" src="images/spacer.gif" width="5"></TD>
												</TR>
											</TABLE>
										</TD>
										<TD class="Normal">
											<ASP:LABEL id="Label4" runat="server" resourcekey="ItemDeletedCode">Item Deleted</ASP:LABEL></TD>
									</TR>
									<TR>
										<TD>
											<TABLE cellspacing="1" cellpadding="3" bgcolor="#000000" border="0">
												<TR>
													<TD class="Normal" bgcolor="#999900"><IMG height="5" src="images/spacer.gif" width="5"></TD>
												</TR>
											</TABLE>
										</TD>
										<TD class="Normal" nowrap>
											<ASP:LABEL id="Label5" runat="server" resourcekey="SuccessCode">Operation Success</ASP:LABEL></TD>
									</TR>
									<TR>
										<TD>
											<TABLE cellspacing="1" cellpadding="3" bgcolor="#000000" border="0">
												<TR>
													<TD class="Normal" bgcolor="#990000"><IMG height="5" src="images/spacer.gif" width="5"></TD>
												</TR>
											</TABLE>
										</TD>
										<TD class="Normal">
											<ASP:LABEL id="Label6" runat="server" resourcekey="FailureCode">Operation Failure</ASP:LABEL></TD>
									</TR>
									<TR>
										<TD>
											<TABLE cellspacing="1" cellpadding="3" bgcolor="#000000" border="0">
												<TR>
													<TD class="Normal" bgcolor="#4d0099"><IMG height="5" src="images/spacer.gif" width="5"></TD>
												</TR>
											</TABLE>
										</TD>
										<TD class="Normal">
											<ASP:LABEL id="Label7" runat="server" resourcekey="AdminOpCode">General Admin Operation</ASP:LABEL></TD>
									</TR>
									<TR>
										<TD>
											<TABLE cellspacing="1" cellpadding="3" bgcolor="#000000" border="0">
												<TR>
													<TD class="Normal" bgcolor="#148aff"><IMG height="5" src="images/spacer.gif" width="5"></TD>
												</TR>
											</TABLE>
										</TD>
										<TD class="Normal">
											<ASP:LABEL id="Label8" runat="server" resourcekey="AdminAlertCode">Admin Alert</ASP:LABEL></TD>
									</TR>
									<TR>
										<TD>
											<TABLE cellspacing="1" cellpadding="3" bgcolor="#000000" border="0">
												<TR>
													<TD class="Normal" bgcolor="#ff8a14"><IMG height="5" src="images/spacer.gif" width="5"></TD>
												</TR>
											</TABLE>
										</TD>
										<TD class="Normal">
											<ASP:LABEL id="Label9" runat="server" resourcekey="HostAlert">Host Alert</ASP:LABEL></TD>
									</TR>
								</TABLE>
							</TD>
						</TR>
					</TABLE>
				</ASP:PANEL></TD>
		</TR>
	</TABLE>
	<BR>
</ASP:PANEL>
<BR>
<ASP:DATALIST enableviewstate="False" width="100%" runat="server" id="dlLog" cellpadding="0" cellspacing="1"
	backcolor="#CFCFCF">
	<ITEMSTYLE borderwidth="0" />
	<HEADERSTYLE backcolor="#CFCFCF" borderwidth="0" />
	<HEADERTEMPLATE>
		<SPAN class="NormalBold" style="width:20;background:#CFCFCF;">&nbsp;</SPAN>
		<ASP:LABEL id="lblDateHeader" runat="server" resourcekey="Date" class="NormalBold" style="width:150;background:#CFCFCF;">&nbsp;Date</ASP:LABEL>
		<ASP:LABEL id="lblTypeHeader" runat="server" resourcekey="Type" class="NormalBold" style="width:200;background:#CFCFCF;">&nbsp;Log Type</ASP:LABEL>
		<ASP:LABEL id="lblUserHeader" runat="server" resourcekey="Username" class="NormalBold" style="width:100;background:#CFCFCF;">&nbsp;Username</ASP:LABEL>
		<ASP:LABEL id="lblPortalHeader" runat="server" resourcekey="Portal" class="NormalBold" style="width:150;background:#CFCFCF;">&nbsp;Portal</ASP:LABEL>
		<ASP:LABEL id="lblSummaryHeader" runat="server" resourcekey="Summary" class="NormalBold" style="width:270;background:#CFCFCF;">&nbsp;Summary</ASP:LABEL>
	</HEADERTEMPLATE>
	<ITEMTEMPLATE>
		<TABLE width="100%" border="0" cellspacing="0" cellpadding="0">
			<TR CLASS='<%# GetMyLogType(DataBinder.Eval(Container.DataItem,"LogTypeKey")).LogTypeCSSclass %>' >
				<TD width="20" valign="middle" align="center"><INPUT TYPE="checkbox" NAME="Exception" VALUE='<%# CType(Container.DataItem, DotNetNuke.Services.Log.EventLog.LogInfo).LogGUid %>|<%# CType(Container.DataItem, DotNetNuke.Services.Log.EventLog.LogInfo).LogFileid %>'/></TD>
				<TD NOWRAP ONCLICK="flipFlop('<%# CType(Container.DataItem, DotNetNuke.Services.Log.EventLog.LogInfo).LogGUid %>')">
					<SPAN class="Normal" style="width:150;overflow:hidden;">&nbsp;
						<asp:label EnableViewState="False" runat="server" Text='<%# DataBinder.Eval(Container.DataItem,"LogCreateDate") %>' id="lblDate"/></SPAN>
					<SPAN class="Normal" style="width:200;overflow:hidden;">&nbsp;
						<asp:label EnableViewState="False" runat="server" Text='<%# GetMyLogType(DataBinder.Eval(Container.DataItem,"LogTypeKey")).LogTypeFriendlyName %>' id="lblType"/></SPAN>
					<SPAN class="Normal" style="width:100;overflow:hidden;">&nbsp;
						<asp:label EnableViewState="False" runat="server" Text='<%# DataBinder.Eval(Container.DataItem,"LogUserName") %>' id="lblUserName"/></SPAN>
					<SPAN class="Normal" style="width:150;overflow:hidden;">&nbsp;
						<asp:label EnableViewState="False" runat="server" Text='<%# DataBinder.Eval(Container.DataItem,"LogPortalName") %>' id="lblPortal"/></SPAN>
					<SPAN class="Normal" style="width:280;overflow:hidden;">&nbsp;
						<asp:label EnableViewState="False" runat="server" Text='<%# CType(Container.DataItem, DotNetNuke.Services.Log.EventLog.LogInfo).LogProperties.Summary %>' id="lblSummary"/>&nbsp;...</SPAN>
				</TD>
			</TR>
			<TR STYLE="display:none;" ID='<%# CType(Container.DataItem, DotNetNuke.Services.Log.EventLog.LogInfo).LogGUid %>'>
				<TD colspan="2" nowrap bgcolor="#FFFFFF">
					<TABLE width="100%" border="0" bgcolor="#000000" border="0" cellspacing="1" cellpadding="5">
						<TR>
							<TD bgcolor="#CFCFCF">
								<asp:label class="Normal" EnableViewState="False" runat="server" Text='<%# GetPropertiesText(Container.DataItem) %>' id="lblException"/>
							</TD>
						</TR>
					</TABLE>
				</TD>
			</TR>
		</TABLE>
	</ITEMTEMPLATE>
</ASP:DATALIST>
<DNNSC:PAGINGCONTROL id="ctlPagingControlBottom" runat="server"></DNNSC:PAGINGCONTROL>
<P>
	<ASP:LINKBUTTON cssclass="CommandButton" id="btnDelete" resourcekey="btnDelete" runat="server">Delete Selected Exceptions</ASP:LINKBUTTON>
	&nbsp;&nbsp;
	<ASP:LINKBUTTON cssclass="CommandButton" id="btnClear" resourcekey="btnClear" runat="server">Clear Log</ASP:LINKBUTTON>
</P>
<ASP:PANEL id="pnlSendExceptions" runat="server" cssclass="Normal">
	<DNN:SECTIONHEAD id="dshSendExceptions" runat="server" text="Send Exceptions" section="tblSendExceptions"
		resourcekey="SendExceptions" cssclass="Head" isexpanded="False" includerule="True"></DNN:SECTIONHEAD>
	<TABLE id="tblSendExceptions" cellspacing="2" cellpadding="2" width="560" border="0" runat="server">
		<TR>
			<TD colspan="3">
				<ASP:LABEL class="normal" id="lblExceptionsWarning" runat="server" resourcekey="ExceptionsWarning">
          <B>Please note</B>: By using these features 
            below, you <I>may</I> be sending sensitive data over the Internet in clear 
            text (<I>not</I> encrypted). Before sending your exception submission, 
            please review the contents of your exception log to verify that no 
            sensitive data is contained within it. Only the log entries checked above 
            will be sent.
        </ASP:LABEL>
				<HR noshade size="1">
				<ASP:RADIOBUTTONLIST id="optEmailType" runat="server" cssclass="Normal" autopostback="False" repeatdirection="Horizontal">
					<ASP:LISTITEM value="Email" selected="True" resourcekey="ToEmail.Text">To Specified Email Address</ASP:LISTITEM>
				</ASP:RADIOBUTTONLIST></TD>
		</TR>
		<TR>
			<TD class="SubHead" align="left" width="200">
				<DNN:LABEL id="plEmailAddress" runat="server" suffix=":" controlname="txtEmailAddress"></DNN:LABEL></TD>
			<TD width="200">
				<ASP:TEXTBOX id="txtEmailAddress" runat="server"></ASP:TEXTBOX></TD>
			<TD width="*">&nbsp;</TD>
		</TR>
		<TR>
			<TD class="SubHead" align="left" width="200">
				<DNN:LABEL id="plMessage" runat="server" resourcekey="SendMessage" suffix=":" controlname="txtMessage"></DNN:LABEL></TD>
			<TD width="200">
				<ASP:TEXTBOX id="txtMessage" runat="server" textmode="MultiLine" columns="25" rows="6"></ASP:TEXTBOX></TD>
			<TD width="*">&nbsp;</TD>
		</TR>
		<TR>
			<TD colspan="3">
				<ASP:LINKBUTTON id="btnEmail" runat="server" resourcekey="btnEmail" cssclass="CommandButton">Send Selected Exceptions</ASP:LINKBUTTON></TD>
		</TR>
	</TABLE>
</ASP:PANEL>
