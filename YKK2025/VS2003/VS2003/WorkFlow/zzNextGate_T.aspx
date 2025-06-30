<%@ Page Language="vb" AspCompat="true" AutoEventWireup="false" Codebehind="zzNextGate_T.aspx.vb" Inherits="SPD.NextGate_T" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>NextGate_T</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體">
				<asp:Label id="Label1" style="Z-INDEX: 103; LEFT: 24px; POSITION: absolute; TOP: 24px" runat="server"
					Width="150px" Height="24px">表單號碼：</asp:Label>
				<asp:TextBox id="DNextGate3" style="Z-INDEX: 116; LEFT: 184px; POSITION: absolute; TOP: 368px"
					runat="server" Width="150px" Height="24px" BackColor="Yellow">100</asp:TextBox>
				<asp:TextBox id="DNextGate2" style="Z-INDEX: 115; LEFT: 184px; POSITION: absolute; TOP: 336px"
					runat="server" Width="150px" Height="24px" BackColor="Yellow">100</asp:TextBox>
				<asp:TextBox id="DNextGate1" style="Z-INDEX: 114; LEFT: 184px; POSITION: absolute; TOP: 304px"
					runat="server" Width="150px" Height="24px" BackColor="Yellow">100</asp:TextBox>
				<asp:Label id="Label7" style="Z-INDEX: 113; LEFT: 24px; POSITION: absolute; TOP: 312px" runat="server"
					Width="150px" Height="24px">簽核者：</asp:Label>
				<asp:TextBox id="DCount" style="Z-INDEX: 112; LEFT: 184px; POSITION: absolute; TOP: 264px" runat="server"
					Width="150px" Height="24px" BackColor="Yellow">100</asp:TextBox>
				<asp:Label id="Label6" style="Z-INDEX: 111; LEFT: 24px; POSITION: absolute; TOP: 272px" runat="server"
					Width="150px" Height="24px">下一簽核者數：</asp:Label>
				<asp:TextBox id="DNextStep" style="Z-INDEX: 110; LEFT: 184px; POSITION: absolute; TOP: 232px"
					runat="server" Width="150px" Height="24px" BackColor="Yellow">100</asp:TextBox>
				<asp:Label id="Label5" style="Z-INDEX: 109; LEFT: 24px; POSITION: absolute; TOP: 240px" runat="server"
					Width="150px" Height="24px">下一工程：</asp:Label>
				<asp:Button id="Button1" style="Z-INDEX: 108; LEFT: 88px; POSITION: absolute; TOP: 168px" runat="server"
					Width="128px" Height="40px" Text="Go"></asp:Button>
				<asp:TextBox id="DApplyID" style="Z-INDEX: 107; LEFT: 184px; POSITION: absolute; TOP: 120px"
					runat="server" Width="150px" Height="24px" BackColor="Yellow">100</asp:TextBox>
				<asp:Label id="Label4" style="Z-INDEX: 106; LEFT: 24px; POSITION: absolute; TOP: 120px" runat="server"
					Width="150px" Height="24px">申請者：</asp:Label>
				<asp:TextBox id="DUserID" style="Z-INDEX: 105; LEFT: 184px; POSITION: absolute; TOP: 88px" runat="server"
					Width="150px" Height="24px" BackColor="Yellow">100</asp:TextBox>
				<asp:Label id="Label2" style="Z-INDEX: 104; LEFT: 24px; POSITION: absolute; TOP: 88px" runat="server"
					Width="150px" Height="24px">申請簽核者：</asp:Label>
				<asp:TextBox id="DFormNo" style="Z-INDEX: 101; LEFT: 184px; POSITION: absolute; TOP: 24px" runat="server"
					Width="150px" Height="24px" BackColor="Yellow">000001</asp:TextBox>
				<asp:TextBox id="DStep" style="Z-INDEX: 102; LEFT: 184px; POSITION: absolute; TOP: 56px" runat="server"
					Width="150px" Height="24px" BackColor="Yellow">1</asp:TextBox>
				<asp:Label id="Label3" style="Z-INDEX: 100; LEFT: 24px; POSITION: absolute; TOP: 56px" runat="server"
					Width="150px" Height="24px">工程代號：</asp:Label></FONT>
		</form>
	</body>
</HTML>
