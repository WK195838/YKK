<%@ Page Language="vb" AspCompat="true" AutoEventWireup="false" Codebehind="zzSendMail_T.aspx.vb" Inherits="SPD.SendMail_T"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>SendMail_T</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="·s²Ó©úÅé">
				<asp:Button id="Button1" style="Z-INDEX: 101; LEFT: 8px; POSITION: absolute; TOP: 8px" runat="server"
					Text="Go" Width="128px" Height="40px"></asp:Button>
				<asp:DropDownList id="DropDownList1" style="Z-INDEX: 102; LEFT: 152px; POSITION: absolute; TOP: 16px"
					runat="server" Height="24px" Width="168px" ForeColor="Blue" Enabled="False" BackColor="White"></asp:DropDownList>
				<asp:Calendar id="Calendar1" style="Z-INDEX: 103; LEFT: 16px; POSITION: absolute; TOP: 64px" runat="server"
					Height="250px" Width="330px" BackColor="White" ForeColor="Black" BorderStyle="Solid" NextPrevFormat="ShortMonth"
					CellSpacing="1" Font-Size="9pt" Font-Names="Verdana" BorderColor="Black">
					<TodayDayStyle ForeColor="White" BackColor="#999999"></TodayDayStyle>
					<DayStyle BackColor="#CCCCCC"></DayStyle>
					<NextPrevStyle Font-Size="8pt" Font-Bold="True" ForeColor="White"></NextPrevStyle>
					<DayHeaderStyle Font-Size="8pt" Font-Bold="True" Height="8pt" ForeColor="#333333"></DayHeaderStyle>
					<SelectedDayStyle ForeColor="White" BackColor="#333399"></SelectedDayStyle>
					<TitleStyle Font-Size="12pt" Font-Bold="True" Height="12pt" ForeColor="White" BackColor="#333399"></TitleStyle>
					<OtherMonthDayStyle ForeColor="#999999"></OtherMonthDayStyle>
				</asp:Calendar></FONT>
		</form>
	</body>
</HTML>
