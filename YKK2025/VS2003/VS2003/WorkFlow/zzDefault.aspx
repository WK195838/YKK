<%@ Page Language="vb" AutoEventWireup="false" Codebehind="Default.aspx.vb" Inherits="SPD._Default" smartNavigation="False" codePage="950" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>���u�v����ƨt��</title>
		<META http-equiv="Content-Type" content="text/html; charset=big5">
		<meta content="True" name="vs_snapToGrid">
		<meta content="True" name="vs_showGrid">
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body ms_positioning="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT id="FONT1" tabIndex="0" face="�s�ө���"><IMG style="Z-INDEX: 107; LEFT: 8px; POSITION: absolute; TOP: 8px" height="75" alt=""
					src="Images\MainMenu-1.jpg" width="750">
				<DIV title="�n����T�G" style="DISPLAY: inline; Z-INDEX: 101; LEFT: 224px; WIDTH: 88px; COLOR: #ff0000; POSITION: absolute; TOP: 88px; HEIGHT: 24px"
					ms_positioning="FlowLayout">�n����T�G</DIV>
				<DIV title="�ɶ��G" style="DISPLAY: inline; Z-INDEX: 102; LEFT: 312px; WIDTH: 56px; COLOR: #0000ff; POSITION: absolute; TOP: 88px; HEIGHT: 24px"
					ms_positioning="FlowLayout">�ɶ��G</DIV>
				<asp:label id="DLoginTime" style="Z-INDEX: 103; LEFT: 368px; POSITION: absolute; TOP: 88px"
					runat="server" Width="216px" Height="24px" ForeColor="Blue">xxxxxxx</asp:label>
				<DIV title="�a�I�G" style="DISPLAY: inline; Z-INDEX: 104; LEFT: 312px; WIDTH: 56px; COLOR: #0000ff; POSITION: absolute; TOP: 112px; HEIGHT: 24px"
					ms_positioning="FlowLayout">�a�I�G</DIV>
				<asp:label id="DLoginPlace" style="Z-INDEX: 105; LEFT: 368px; POSITION: absolute; TOP: 112px"
					runat="server" Width="144px" Height="24px" ForeColor="Blue">xxxx</asp:label>
				<DIV title="�עޡG" style="DISPLAY: inline; Z-INDEX: 106; LEFT: 312px; WIDTH: 56px; COLOR: #0000ff; POSITION: absolute; TOP: 136px; HEIGHT: 24px"
					ms_positioning="FlowLayout">�עޡG</DIV>
				<asp:label id="DLoginIP" style="Z-INDEX: 108; LEFT: 368px; POSITION: absolute; TOP: 136px"
					runat="server" Width="144px" Height="24px" ForeColor="Blue">xxx.xxx.xxx.xxx</asp:label><IMG style="Z-INDEX: 109; LEFT: 216px; POSITION: absolute; TOP: 184px" height="166" alt=""
					src="Images\Login.jpg" width="345"><IMG style="Z-INDEX: 110; LEFT: 8px; POSITION: absolute; TOP: 416px" height="75" alt=""
					src="Images\MainMenu-2.jpg" width="750">
				<asp:textbox id="DPW" style="Z-INDEX: 112; LEFT: 328px; POSITION: absolute; TOP: 256px" runat="server"
					Width="80px" Height="20px" ForeColor="Blue" BackColor="Yellow" TextMode="Password"></asp:textbox><asp:textbox id="DID" style="Z-INDEX: 111; LEFT: 328px; POSITION: absolute; TOP: 224px" runat="server"
					Width="80px" Height="20px" ForeColor="Blue" BackColor="Yellow"></asp:textbox><asp:button id="DLoginBtn" style="Z-INDEX: 113; LEFT: 416px; POSITION: absolute; TOP: 256px"
					runat="server" Width="40px" ForeColor="Blue" BackColor="White" Text="�n�J"></asp:button><asp:requiredfieldvalidator id="RequiredFieldValidator1" style="Z-INDEX: 114; LEFT: 416px; POSITION: absolute; TOP: 224px"
					runat="server" Width="96px" Height="24px" Display="Dynamic" ControlToValidate="DID" ErrorMessage="������J�b��"></asp:requiredfieldvalidator></FONT><asp:requiredfieldvalidator id="RequiredFieldValidator2" style="Z-INDEX: 115; LEFT: 328px; POSITION: absolute; TOP: 304px"
				runat="server" Width="96px" Height="24px" Display="Dynamic" ControlToValidate="DPW" ErrorMessage="������J�K�X"></asp:requiredfieldvalidator><asp:label id="DMessage" style="Z-INDEX: 116; LEFT: 216px; POSITION: absolute; TOP: 352px"
				runat="server" Width="536px" Height="24px" ForeColor="Red"></asp:label></form>
	</body>
</HTML>
