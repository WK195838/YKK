<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SliderPicker.aspx.vb" Inherits="SPD.SliderPicker"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>拉頭代碼選擇</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="SliderForm" method="post" runat="server">
			<FONT face="新細明體">
				<asp:image id="DQASheet" style="Z-INDEX: 101; LEFT: 8px; POSITION: absolute; TOP: 8px" runat="server"
					Height="243px" Width="278px" ImageUrl="Images\SliderCodeSheet.jpg"></asp:image><asp:textbox id="DContent" style="Z-INDEX: 106; LEFT: 16px; POSITION: absolute; TOP: 120px" runat="server"
					Height="120px" Width="260px" TextMode="MultiLine" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove"></asp:textbox><asp:imagebutton id="BDown" style="Z-INDEX: 105; LEFT: 16px; POSITION: absolute; TOP: 88px" runat="server"
					Height="27px" Width="37px" ImageUrl="Images\arrow1d.jpg"></asp:imagebutton><asp:imagebutton id="BClose" style="Z-INDEX: 102; LEFT: 192px; POSITION: absolute; TOP: 96px" runat="server"
					Height="20px" Width="84px" ImageUrl="Images\close.gif"></asp:imagebutton><asp:textbox id="DSlider1" style="Z-INDEX: 104; LEFT: 16px; POSITION: absolute; TOP: 52px" runat="server"
					Height="20px" Width="136px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove"></asp:textbox><asp:textbox id="DSlider2" style="Z-INDEX: 103; LEFT: 162px; POSITION: absolute; TOP: 52px" runat="server"
					Height="20px" Width="117px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove"></asp:textbox></FONT></form>
	</body>
</HTML>
