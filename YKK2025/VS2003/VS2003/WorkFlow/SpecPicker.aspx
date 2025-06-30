<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SpecPicker.aspx.vb" Inherits="SPD.SpecPicker" aspCompat="True"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>規格選擇</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="SpecForm" method="post" runat="server">
			<FONT face="新細明體">
				<asp:imagebutton style="Z-INDEX: 107; POSITION: absolute; TOP: 96px; LEFT: 240px" id="BClose" runat="server"
					ImageUrl="Images\close.gif" Width="84px" Height="20px"></asp:imagebutton><asp:imagebutton style="Z-INDEX: 106; POSITION: absolute; TOP: 88px; LEFT: 16px" id="BDown" runat="server"
					ImageUrl="Images\arrow1d.jpg" Width="37px" Height="27px"></asp:imagebutton><asp:textbox style="Z-INDEX: 105; POSITION: absolute; TOP: 120px; LEFT: 16px" id="DSpec" runat="server"
					Width="304px" Height="120px" TextMode="MultiLine" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow"></asp:textbox><asp:dropdownlist style="Z-INDEX: 104; POSITION: absolute; TOP: 52px; LEFT: 184px" id="DBody" runat="server"
					Width="136px" Height="20px" ForeColor="Blue" BackColor="Yellow" AutoPostBack="True">
					<asp:ListItem Value="3">CF</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 103; POSITION: absolute; TOP: 52px; LEFT: 98px" id="DChainType"
					runat="server" Width="78px" Height="20px" ForeColor="Blue" BackColor="Yellow" AutoPostBack="True">
					<asp:ListItem Value="3">CF</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 102; POSITION: absolute; TOP: 52px; LEFT: 16px" id="Dsize" runat="server"
					Width="76px" Height="20px" ForeColor="Blue" BackColor="Yellow" AutoPostBack="True">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:image style="Z-INDEX: 101; POSITION: absolute; TOP: 8px; LEFT: 8px" id="DQASheet" runat="server"
					ImageUrl="Images\SpecSheet.jpg" Width="320px" Height="244px"></asp:image>
				<asp:HyperLink style="Z-INDEX: 108; POSITION: absolute; TOP: 96px; LEFT: 80px" id="LChainType"
					runat="server" Height="16px" Width="81px" Target="_blank" NavigateUrl="Images\台灣YKK特殊引手兼用胴體型別(更新).xls">可兼用型別</asp:HyperLink></FONT></form>
	</body>
</HTML>
