<%@ Page Language="vb" AutoEventWireup="false" Codebehind="zzQA2Picker.aspx.vb" Inherits="SPD.QA2Picker"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>QA2Picker</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<script language="javascript" type="text/javascript">
			function CalendarPicker(strField)
			{
				window.open('DatePicker.aspx?field=' + strField,'calendarPopup','width=250,height=190,resizable=yes');
			}
		</script>
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="QAForm2" method="post" runat="server">
			<FONT face="·s²Ó©úÅé">
				<asp:image id="DQASheet" style="Z-INDEX: 100; LEFT: 8px; POSITION: absolute; TOP: 8px" runat="server"
					ImageUrl="Images\QASheet_002.jpg" Width="593px" Height="243px"></asp:image>
				<asp:textbox id="DContent" style="Z-INDEX: 120; LEFT: 16px; POSITION: absolute; TOP: 120px" runat="server"
					Height="122px" Width="576px" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue" TextMode="MultiLine"
					MaxLength="255"></asp:textbox>
				<asp:textbox id="DRemark" style="Z-INDEX: 109; LEFT: 226px; POSITION: absolute; TOP: 52px" runat="server"
					Width="368px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove"></asp:textbox>
				<asp:dropdownlist id="DResult" style="Z-INDEX: 108; LEFT: 120px; POSITION: absolute; TOP: 52px" runat="server"
					Width="97px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="PASS" Selected="True">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:textbox id="DDate" style="Z-INDEX: 106; LEFT: 16px; POSITION: absolute; TOP: 52px" runat="server"
					Width="72px" Height="20px" ForeColor="Blue" BackColor="Yellow" ReadOnly="True" BorderStyle="Groove"></asp:textbox>
				<asp:button id="BDate" style="Z-INDEX: 103; LEFT: 88px; POSITION: absolute; TOP: 52px" runat="server"
					Width="20px" Height="20px" Text="....."></asp:button>
				<asp:imagebutton id="BDown" style="Z-INDEX: 105; LEFT: 16px; POSITION: absolute; TOP: 88px" runat="server"
					ImageUrl="Images\arrow1d.jpg" Width="37px" Height="27px"></asp:imagebutton>
				<asp:imagebutton id="BClose" style="Z-INDEX: 102; LEFT: 512px; POSITION: absolute; TOP: 96px" runat="server"
					ImageUrl="Images\close.gif" Width="84px" Height="20px"></asp:imagebutton></FONT>
		</form>
	</body>
</HTML>
