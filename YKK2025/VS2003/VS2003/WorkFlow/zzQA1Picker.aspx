<%@ Page Language="vb" AutoEventWireup="false" Codebehind="zzQA1Picker.aspx.vb" Inherits="SPD.QA1Picker"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>QA1Picker</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<script language="javascript" type="text/javascript">
			function CalendarPicker(strField)
			{
				window.open('DatePicker.aspx?field=' + strField,'calendarPopup','width=250,height=190,resizable=yes');
			}
		</script>
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="QAForm1" method="post" runat="server">
			<FONT face="新細明體">
				<asp:image id="DQASheet" style="Z-INDEX: 100; LEFT: 8px; POSITION: absolute; TOP: 8px" runat="server"
					Height="311px" Width="677px" ImageUrl="Images\QASheet_001.jpg"></asp:image><asp:dropdownlist id="Dsize" style="Z-INDEX: 121; LEFT: 162px; POSITION: absolute; TOP: 52px" runat="server"
					Height="20px" Width="48px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DChainType" style="Z-INDEX: 104; LEFT: 210px; POSITION: absolute; TOP: 52px"
					runat="server" Height="20px" Width="62px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">CF</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DBody" style="Z-INDEX: 107; LEFT: 272px; POSITION: absolute; TOP: 52px" runat="server"
					Height="20px" Width="70px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">CF</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DContent" style="Z-INDEX: 120; LEFT: 16px; POSITION: absolute; TOP: 184px" runat="server"
					Height="128px" Width="660px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" TextMode="MultiLine" MaxLength="255"></asp:textbox><asp:dropdownlist id="DCPSC" style="Z-INDEX: 118; LEFT: 604px; POSITION: absolute; TOP: 120px" runat="server"
					Height="20px" Width="72px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="P" Selected="True">PASS</asp:ListItem>
					<asp:ListItem Value="F">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DDry" style="Z-INDEX: 117; LEFT: 352px; POSITION: absolute; TOP: 120px" runat="server"
					Height="20px" Width="72px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="P" Selected="True">PASS</asp:ListItem>
					<asp:ListItem Value="F">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DWater" style="Z-INDEX: 116; LEFT: 268px; POSITION: absolute; TOP: 120px" runat="server"
					Height="20px" Width="72px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="P" Selected="True">PASS</asp:ListItem>
					<asp:ListItem Value="F">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DKensin" style="Z-INDEX: 115; LEFT: 184px; POSITION: absolute; TOP: 120px" runat="server"
					Height="20px" Width="72px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="P" Selected="True">PASS</asp:ListItem>
					<asp:ListItem Value="F">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DNyuryoku" style="Z-INDEX: 114; LEFT: 100px; POSITION: absolute; TOP: 120px"
					runat="server" Height="20px" Width="72px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="P" Selected="True">PASS</asp:ListItem>
					<asp:ListItem Value="F">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DMityaku" style="Z-INDEX: 113; LEFT: 520px; POSITION: absolute; TOP: 120px"
					runat="server" Height="20px" Width="60px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove"></asp:textbox><asp:textbox id="DYellow" style="Z-INDEX: 112; LEFT: 436px; POSITION: absolute; TOP: 120px" runat="server"
					Height="20px" Width="60px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove"></asp:textbox><asp:dropdownlist id="DKyoudo" style="Z-INDEX: 111; LEFT: 16px; POSITION: absolute; TOP: 120px" runat="server"
					Height="20px" Width="72px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="P" Selected="True">PASS</asp:ListItem>
					<asp:ListItem Value="F">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DGentani" style="Z-INDEX: 110; LEFT: 562px; POSITION: absolute; TOP: 52px" runat="server"
					Height="20px" Width="92px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove"></asp:textbox><asp:textbox id="DSurface1" style="Z-INDEX: 109; LEFT: 436px; POSITION: absolute; TOP: 52px"
					runat="server" Height="20px" Width="58px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove"></asp:textbox><asp:textbox id="DSurface2" style="Z-INDEX: 106; LEFT: 496px; POSITION: absolute; TOP: 52px"
					runat="server" Height="20px" Width="58px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove"></asp:textbox><asp:dropdownlist id="DAssembly" style="Z-INDEX: 108; LEFT: 352px; POSITION: absolute; TOP: 52px"
					runat="server" Height="20px" Width="74px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="正">正</asp:ListItem>
					<asp:ListItem Value="反">反</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DDate" style="Z-INDEX: 105; LEFT: 18px; POSITION: absolute; TOP: 52px" runat="server"
					Height="20px" Width="112px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True"></asp:textbox><asp:button id="BDate" style="Z-INDEX: 103; LEFT: 130px; POSITION: absolute; TOP: 52px" runat="server"
					Height="20px" Width="20px" Text="....."></asp:button><asp:imagebutton id="BClose" style="Z-INDEX: 102; LEFT: 592px; POSITION: absolute; TOP: 160px" runat="server"
					Height="20px" Width="84px" ImageUrl="Images\close.gif"></asp:imagebutton><asp:imagebutton id="BDown" style="Z-INDEX: 101; LEFT: 16px; POSITION: absolute; TOP: 152px" runat="server"
					Height="27px" Width="37px" ImageUrl="Images\arrow1d.jpg"></asp:imagebutton></FONT></form>
	</body>
</HTML>
