<%@ Page Language="vb" AutoEventWireup="false" Codebehind="IT_TroubleShooterSheet_02.aspx.vb" Inherits="SPD.IT_TroubleShooterSheet_02"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>故障檢測申請書</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<script language="javascript" type="text/javascript">
		</script>
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" encType="multipart/form-data" runat="server">
			<FONT face="新細明體">
				<asp:dropdownlist id="DSystem" style="Z-INDEX: 119; LEFT: 456px; POSITION: absolute; TOP: 152px" runat="server"
					Height="20px" Width="224px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="要" Selected="True">要</asp:ListItem>
					<asp:ListItem Value="不要">不要</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DPriority" style="Z-INDEX: 120; LEFT: 120px; POSITION: absolute; TOP: 152px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Width="224px" Height="20px">
					<asp:ListItem Value="要" Selected="True">要</asp:ListItem>
					<asp:ListItem Value="不要">不要</asp:ListItem>
				</asp:dropdownlist>
				<asp:textbox id="DBDays" style="Z-INDEX: 118; LEFT: 608px; POSITION: absolute; TOP: 352px" runat="server"
					Height="20px" Width="70px" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue">DBDays</asp:textbox><asp:textbox id="DDepoName" style="Z-INDEX: 117; LEFT: 120px; POSITION: absolute; TOP: 120px"
					runat="server" Width="40px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">台北</asp:textbox><asp:textbox id="DDepoCode" style="Z-INDEX: 115; LEFT: 160px; POSITION: absolute; TOP: 120px"
					runat="server" Width="25px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">01</asp:textbox><asp:textbox id="DDivision" style="Z-INDEX: 105; LEFT: 192px; POSITION: absolute; TOP: 120px"
					runat="server" Width="88px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DDivision</asp:textbox><asp:textbox id="DDivisionCode" style="Z-INDEX: 106; LEFT: 280px; POSITION: absolute; TOP: 120px"
					runat="server" Width="63px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DDivisionCode</asp:textbox><asp:textbox id="DJobCode" style="Z-INDEX: 114; LEFT: 600px; POSITION: absolute; TOP: 120px"
					runat="server" Width="80px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DJobCode</asp:textbox><asp:textbox id="DEmpID" style="Z-INDEX: 113; LEFT: 600px; POSITION: absolute; TOP: 88px" runat="server"
					Width="80px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DEmpID</asp:textbox><asp:textbox id="DBEndDate" style="Z-INDEX: 112; LEFT: 336px; POSITION: absolute; TOP: 352px"
					runat="server" Width="176px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">DBEndDate</asp:textbox><asp:textbox id="DBStartDate" style="Z-INDEX: 111; LEFT: 120px; POSITION: absolute; TOP: 352px"
					runat="server" Width="184px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">DBStartDate</asp:textbox><asp:textbox id="DRemark" style="Z-INDEX: 110; LEFT: 120px; POSITION: absolute; TOP: 382px" runat="server"
					Width="560px" Height="90px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" TextMode="MultiLine" MaxLength="250">DRemark</asp:textbox><asp:hyperlink id="LRefFile" style="Z-INDEX: 109; LEFT: 120px; POSITION: absolute; TOP: 288px"
					runat="server" Width="72px" Height="8px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">參考附件</asp:hyperlink><asp:textbox id="DTarget" style="Z-INDEX: 108; LEFT: 120px; POSITION: absolute; TOP: 184px" runat="server"
					Width="560px" Height="88px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" TextMode="MultiLine" MaxLength="250">DTarget</asp:textbox><asp:textbox id="DName" style="Z-INDEX: 107; LEFT: 456px; POSITION: absolute; TOP: 88px" runat="server"
					Width="144px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DName</asp:textbox><asp:dropdownlist id="DEngineer" style="Z-INDEX: 104; LEFT: 120px; POSITION: absolute; TOP: 318px"
					runat="server" Width="224px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="要" Selected="True">要</asp:ListItem>
					<asp:ListItem Value="不要">不要</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DJobTitle" style="Z-INDEX: 103; LEFT: 456px; POSITION: absolute; TOP: 120px"
					runat="server" Width="144px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DJobTitle</asp:textbox><asp:textbox id="DDate" style="Z-INDEX: 102; LEFT: 120px; POSITION: absolute; TOP: 88px" runat="server"
					Width="224px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DDate</asp:textbox><asp:image id="DTroubleShooterSheet1" style="Z-INDEX: 100; LEFT: 8px; POSITION: absolute; TOP: 8px"
					runat="server" Width="678px" Height="472px" ImageUrl="Images\IT_TroubleShooterSheet_001.jpg"></asp:image><asp:textbox id="DNo" style="Z-INDEX: 101; LEFT: 8px; POSITION: absolute; TOP: 9px" runat="server"
					Width="216px" Height="16px" ForeColor="Black" BackColor="White" BorderStyle="None" ReadOnly="True" Font-Names="Times New Roman"> 123</asp:textbox></FONT></form>
		</FONT></FORM>
	</body>
</HTML>
