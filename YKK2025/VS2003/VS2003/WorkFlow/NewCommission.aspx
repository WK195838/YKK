<%@ Page Language="vb" AutoEventWireup="false" Codebehind="NewCommission.aspx.vb" Inherits="SPD.NewCommission"%>
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
				<asp:dropdownlist id="DNewSheet" style="Z-INDEX: 101; POSITION: absolute; TOP: 8px; LEFT: 8px" runat="server"
					AutoPostBack="True" Width="256px" Height="24px" BackColor="Yellow" ForeColor="Blue"></asp:dropdownlist><asp:dropdownlist id="DLevel" style="Z-INDEX: 106; POSITION: absolute; TOP: 40px; LEFT: 96px" runat="server"
					Width="169px" Height="24px" BackColor="Yellow" ForeColor="Blue"></asp:dropdownlist><asp:button id="DSimulation" style="Z-INDEX: 104; POSITION: absolute; TOP: 80px; LEFT: 184px"
					runat="server" Width="80px" Height="24px" BackColor="White" ForeColor="Blue" Text="工程模擬"></asp:button><asp:button id="DNew" style="Z-INDEX: 102; POSITION: absolute; TOP: 80px; LEFT: 96px" runat="server"
					Width="80px" Height="24px" BackColor="White" ForeColor="Blue" Text="新委託"></asp:button><asp:textbox id="TextBox1" style="Z-INDEX: 105; POSITION: absolute; TOP: 8px; LEFT: 272px" runat="server"
					Width="320px" Height="96px" BorderStyle="Groove" ReadOnly="True" TextMode="MultiLine">先選擇委託單之後點選需要的按鈕．　　　　　新委託：發行新的委託單．　　　　　　　　　工程模擬：現在委託後預定何時完成．(需設定難易度)　　      　                     　圖面委託書時如有附樣品請選擇Z</asp:textbox><asp:label id="Label1" style="Z-INDEX: 107; POSITION: absolute; TOP: 40px; LEFT: 40px" runat="server"
					Width="48px" Height="24px" Font-Size="11pt">難易度</asp:label></FONT></form>
	</body>
</HTML>
