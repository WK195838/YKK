<%@ Page Language="vb" AutoEventWireup="false" Codebehind="CalendarList_TP2.aspx.vb" Inherits="SPD.CalendarList_TP2"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>行事曆</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體">
				<asp:calendar id="cal" style="Z-INDEX: 101; POSITION: absolute; TOP: 8px; LEFT: 8px" runat="server"
					Height="376px" Width="672px" DayHeaderStyle-BackColor="#FFCC66" onDayRender="cal_DayRender"
					BorderColor="Black" Font-Name="Verdana" Font-Size="10pt" TitleStyle-BackColor="#FFCC66" TitleStyle-Font-Size="12px"
					TitleStyle-Font-Bold="true" DayStyle-VerticalAlign="Top" DayStyle-Height="50px" DayStyle-Width="14%"
					SelectedDayStyle-BackColor="Navy" ShowNextPrev="True" NextPrevFormat="ShortMonth" BorderStyle="Solid"
					BackColor="White" ForeColor="Black" CellSpacing="1" Font-Names="Verdana" SelectionMode="None">
					<TodayDayStyle Font-Bold="True" ForeColor="Maroon"></TodayDayStyle>
					<DayStyle Height="50px" Width="14%" VerticalAlign="Middle" BackColor="#CCCCCC"></DayStyle>
					<NextPrevStyle Font-Size="8pt" Font-Bold="True" ForeColor="White"></NextPrevStyle>
					<DayHeaderStyle Font-Size="8pt" Font-Bold="True" Height="8pt" ForeColor="#333333" BackColor="White"></DayHeaderStyle>
					<SelectedDayStyle ForeColor="White" BackColor="#333399"></SelectedDayStyle>
					<TitleStyle Font-Size="12pt" Font-Bold="True" Height="12pt" ForeColor="White" BackColor="#333399"></TitleStyle>
					<OtherMonthDayStyle ForeColor="#999999"></OtherMonthDayStyle>
				</asp:calendar></FONT></form>
	</body>
</HTML>
