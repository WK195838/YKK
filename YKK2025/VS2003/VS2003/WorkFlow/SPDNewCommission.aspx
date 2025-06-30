<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SPDNewCommission.aspx.vb" Inherits="SPD.SPDNewCommission"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>新委託</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體">
				<asp:dropdownlist id="DNewSheet" style="Z-INDEX: 101; LEFT: 8px; POSITION: absolute; TOP: 8px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="24px" Width="256px" AutoPostBack="True"></asp:dropdownlist><asp:dropdownlist id="DLevel" style="Z-INDEX: 106; LEFT: 96px; POSITION: absolute; TOP: 40px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="24px" Width="169px"></asp:dropdownlist><asp:button id="DSimulation" style="Z-INDEX: 104; LEFT: 184px; POSITION: absolute; TOP: 80px"
					runat="server" ForeColor="Blue" BackColor="White" Height="24px" Width="80px" Text="工程模擬"></asp:button><asp:button id="DNew" style="Z-INDEX: 102; LEFT: 96px; POSITION: absolute; TOP: 80px" runat="server"
					ForeColor="Blue" BackColor="White" Height="24px" Width="80px" Text="新委託"></asp:button>
				<asp:TextBox id="TextBox1" style="Z-INDEX: 105; LEFT: 272px; POSITION: absolute; TOP: 8px" runat="server"
					Height="96px" Width="320px" TextMode="MultiLine" ReadOnly="True" BorderStyle="Groove">先選擇委託單之後點選需要的按鈕．　　　　　新委託：發行新的委託單．　　　　　　　　　工程模擬：現在委託後預定何時完成．(需設定難易度)　　      　                     　圖面委託書時如有附樣品請選擇Z</asp:TextBox>
				<asp:Label id="Label1" style="Z-INDEX: 107; LEFT: 40px; POSITION: absolute; TOP: 40px" runat="server"
					Height="24px" Width="48px" Font-Size="11pt">難易度</asp:Label></FONT></form>
	</body>
</HTML>
