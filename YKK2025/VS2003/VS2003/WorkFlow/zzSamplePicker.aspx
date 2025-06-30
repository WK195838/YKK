<%@ Page Language="vb" AutoEventWireup="false" Codebehind="zzSamplePicker.aspx.vb" Inherits="SPD.SamplePicker"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>SamplePicker</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="SampleForm" method="post" runat="server">
			<FONT face="·s²Ó©úÅé">
				<asp:image id="DPriceSheet" style="Z-INDEX: 100; LEFT: 8px; POSITION: absolute; TOP: 8px" runat="server"
					Height="243px" Width="362px" ImageUrl="Images\SampleSheet_001.jpg"></asp:image>
				<asp:dropdownlist id="Dsize" style="Z-INDEX: 109; LEFT: 16px; POSITION: absolute; TOP: 52px" runat="server"
					Width="48px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DChainType" style="Z-INDEX: 105; LEFT: 64px; POSITION: absolute; TOP: 52px"
					runat="server" Width="62px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="3">CF</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DBody" style="Z-INDEX: 107; LEFT: 126px; POSITION: absolute; TOP: 52px" runat="server"
					Width="70px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="3">CF</asp:ListItem>
				</asp:dropdownlist>
				<asp:textbox id="DContent" style="Z-INDEX: 108; LEFT: 16px; POSITION: absolute; TOP: 120px" runat="server"
					Width="344px" Height="120px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" MaxLength="255"
					TextMode="MultiLine"></asp:textbox>
				<asp:textbox id="DQty" style="Z-INDEX: 106; LEFT: 288px; POSITION: absolute; TOP: 52px" runat="server"
					Height="20px" Width="76px" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue"></asp:textbox>
				<asp:textbox id="DColor" style="Z-INDEX: 103; LEFT: 204px; POSITION: absolute; TOP: 52px" runat="server"
					Height="20px" Width="76px" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue"></asp:textbox>
				<asp:imagebutton id="BDown" style="Z-INDEX: 102; LEFT: 16px; POSITION: absolute; TOP: 88px" runat="server"
					Height="27px" Width="37px" ImageUrl="Images\arrow1d.jpg"></asp:imagebutton>
				<asp:imagebutton id="BClose" style="Z-INDEX: 101; LEFT: 280px; POSITION: absolute; TOP: 96px" runat="server"
					Height="20px" Width="84px" ImageUrl="Images\close.gif"></asp:imagebutton></FONT>
		</form>
	</body>
</HTML>
