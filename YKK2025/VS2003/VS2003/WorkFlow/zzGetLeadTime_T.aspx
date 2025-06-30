<%@ Page Language="vb" AspCompat="true" AutoEventWireup="false" Codebehind="zzGetLeadTime_T.aspx.vb" Inherits="SPD.GetLeadTime_T" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>GetLeadTime_T</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體">
				<asp:Label id="Label1" style="Z-INDEX: 105; LEFT: 24px; POSITION: absolute; TOP: 24px" runat="server"
					Width="120px" Height="24px">表單號碼：</asp:Label>
				<asp:TextBox id="DLevel" style="Z-INDEX: 113; LEFT: 152px; POSITION: absolute; TOP: 120px" runat="server"
					Width="150px" Height="24px" BackColor="Yellow"></asp:TextBox>
				<asp:Label id="Label2" style="Z-INDEX: 112; LEFT: 24px; POSITION: absolute; TOP: 128px" runat="server"
					Width="120px" Height="24px">難易度：</asp:Label>
				<asp:TextBox id="DEndTime" style="Z-INDEX: 111; LEFT: 152px; POSITION: absolute; TOP: 216px"
					runat="server" Width="150px" Height="24px" BackColor="Yellow" ForeColor="Red"></asp:TextBox>
				<asp:TextBox id="DStartTime" style="Z-INDEX: 110; LEFT: 152px; POSITION: absolute; TOP: 184px"
					runat="server" Width="150px" Height="24px" BackColor="Yellow" ForeColor="Red"></asp:TextBox>
				<asp:Label id="Label6" style="Z-INDEX: 109; LEFT: 24px; POSITION: absolute; TOP: 216px" runat="server"
					Width="120px" Height="24px">預定完成：</asp:Label>
				<asp:Label id="Label5" style="Z-INDEX: 108; LEFT: 24px; POSITION: absolute; TOP: 184px" runat="server"
					Width="120px" Height="24px">預定開始：</asp:Label>
				<asp:TextBox id="DCTime" style="Z-INDEX: 107; LEFT: 152px; POSITION: absolute; TOP: 88px" runat="server"
					Width="150px" Height="24px" BackColor="Yellow">2006/8/4 9:45:10</asp:TextBox>
				<asp:Label id="Label4" style="Z-INDEX: 106; LEFT: 24px; POSITION: absolute; TOP: 88px" runat="server"
					Width="120px" Height="24px">日期時間：</asp:Label>
				<asp:TextBox id="DFormNo" style="Z-INDEX: 101; LEFT: 152px; POSITION: absolute; TOP: 24px" runat="server"
					Width="150px" Height="24px" BackColor="Yellow">000001</asp:TextBox>
				<asp:Label id="Label3" style="Z-INDEX: 100; LEFT: 24px; POSITION: absolute; TOP: 56px" runat="server"
					Width="120px" Height="24px">工程代號：</asp:Label>
				<asp:TextBox id="DStep" style="Z-INDEX: 103; LEFT: 152px; POSITION: absolute; TOP: 56px" runat="server"
					Width="150px" Height="24px" BackColor="Yellow">1</asp:TextBox>
				<asp:Button id="Button1" style="Z-INDEX: 104; LEFT: 88px; POSITION: absolute; TOP: 256px" runat="server"
					Width="128px" Height="40px" Text="Go"></asp:Button></FONT>
		</form>
	</body>
</HTML>
