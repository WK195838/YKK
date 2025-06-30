<%@ Page Language="vb" AutoEventWireup="false" Codebehind="HR_OverTimeSheet_02.aspx.vb" Inherits="SPD.HR_OverTimeSheet_02"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>加班申請書</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<script language="javascript" type="text/javascript">
			function ShowCardTime()
			{
				window.open('HR_CardTimeList.aspx?pWorkDate=' + document.Form1.DOverTimeDate.value + '&pDepoID=' + document.Form1.DDepoCode.value + '&pEmpID=' + document.Form1.DEmpID.value, 'CardTimeSheet','width=290,height=150,resizable=yes');
			}
		</script>
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" encType="multipart/form-data" runat="server">
			<FONT face="新細明體">
				<asp:button id="BCardTime" style="Z-INDEX: 138; LEFT: 456px; POSITION: absolute; TOP: 182px"
					runat="server" Width="94px" Height="24px" BackColor="#FFFFC0" Text="刷卡記錄"></asp:button>
				<asp:textbox id="DCVacation" style="Z-INDEX: 139; LEFT: 560px; POSITION: absolute; TOP: 152px"
					runat="server" BackColor="Yellow" Height="20px" Width="120px" ReadOnly="True" BorderStyle="Groove"
					ForeColor="Blue">DCVacation</asp:textbox>
				<asp:textbox id="DBM" style="Z-INDEX: 112; LEFT: 604px; POSITION: absolute; TOP: 250px; TEXT-ALIGN: right"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px" BorderStyle="Groove"
					ReadOnly="True">DBM</asp:textbox>
				<asp:textbox id="DTraffic" style="Z-INDEX: 137; LEFT: 330px; POSITION: absolute; TOP: 184px"
					runat="server" ReadOnly="True" BorderStyle="Groove" Width="118px" Height="20px" BackColor="Yellow"
					ForeColor="Blue">DTraffic</asp:textbox>
				<asp:textbox id="DDepoName" style="Z-INDEX: 136; LEFT: 120px; POSITION: absolute; TOP: 152px"
					runat="server" ReadOnly="True" BorderStyle="Groove" Width="40px" Height="20px" BackColor="Yellow"
					ForeColor="Blue">台北</asp:textbox>
				<asp:textbox id="DDepoCode" style="Z-INDEX: 135; LEFT: 160px; POSITION: absolute; TOP: 152px"
					runat="server" ReadOnly="True" BorderStyle="Groove" Width="25px" Height="20px" BackColor="Yellow"
					ForeColor="Blue">01</asp:textbox>
				<asp:textbox id="DSalaryYM" style="Z-INDEX: 134; LEFT: 376px; POSITION: absolute; TOP: 86px"
					runat="server" ReadOnly="True" BorderStyle="Groove" Width="72px" Height="20px" BackColor="Yellow"
					ForeColor="Blue">2008/10</asp:textbox>
				<asp:textbox id="DAStartM" style="Z-INDEX: 133; LEFT: 224px; POSITION: absolute; TOP: 282px; TEXT-ALIGN: right"
					runat="server" BorderStyle="Groove" Width="60px" Height="20px" BackColor="Yellow" ForeColor="Blue">DAStartM</asp:textbox>
				<asp:textbox id="DAStartH" style="Z-INDEX: 132; LEFT: 140px; POSITION: absolute; TOP: 282px; TEXT-ALIGN: right"
					runat="server" BorderStyle="Groove" Width="60px" Height="20px" BackColor="Yellow" ForeColor="Blue">DAStartH</asp:textbox>
				<asp:textbox id="DBStartM" style="Z-INDEX: 131; LEFT: 224px; POSITION: absolute; TOP: 250px; TEXT-ALIGN: right"
					runat="server" BorderStyle="Groove" Width="60px" Height="20px" BackColor="Yellow" ForeColor="Blue">DBStartM</asp:textbox>
				<asp:textbox id="DBStartH" style="Z-INDEX: 130; LEFT: 140px; POSITION: absolute; TOP: 250px; TEXT-ALIGN: right"
					runat="server" BorderStyle="Groove" Width="60px" Height="20px" BackColor="Yellow" ForeColor="Blue">DBStartH</asp:textbox>
				<asp:textbox id="DBEndM" style="Z-INDEX: 129; LEFT: 412px; POSITION: absolute; TOP: 250px; TEXT-ALIGN: right"
					runat="server" BorderStyle="Groove" Width="60px" Height="20px" BackColor="Yellow" ForeColor="Blue">DBEndM</asp:textbox>
				<asp:textbox id="DBEndH" style="Z-INDEX: 128; LEFT: 330px; POSITION: absolute; TOP: 250px; TEXT-ALIGN: right"
					runat="server" BorderStyle="Groove" Width="60px" Height="20px" BackColor="Yellow" ForeColor="Blue">DBEndH</asp:textbox>
				<asp:textbox id="DAEndM" style="Z-INDEX: 127; LEFT: 412px; POSITION: absolute; TOP: 282px; TEXT-ALIGN: right"
					runat="server" BorderStyle="Groove" Width="60px" Height="20px" BackColor="Yellow" ForeColor="Blue">DAEndM</asp:textbox>
				<asp:textbox id="DAEndH" style="Z-INDEX: 126; LEFT: 330px; POSITION: absolute; TOP: 282px; TEXT-ALIGN: right"
					runat="server" BorderStyle="Groove" Width="60px" Height="20px" BackColor="Yellow" ForeColor="Blue">DAEndH</asp:textbox>
				<asp:textbox id="DFood" style="Z-INDEX: 125; LEFT: 120px; POSITION: absolute; TOP: 184px" runat="server"
					ReadOnly="True" BorderStyle="Groove" Width="96px" Height="20px" BackColor="Yellow" ForeColor="Blue">DFood</asp:textbox><asp:textbox id="DDateType" style="Z-INDEX: 124; LEFT: 272px; POSITION: absolute; TOP: 86px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="102px" BorderStyle="Groove" ReadOnly="True">DDateType</asp:textbox><asp:textbox id="DJobCode" style="Z-INDEX: 123; LEFT: 640px; POSITION: absolute; TOP: 120px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="40px" BorderStyle="Groove" ReadOnly="True">DJobCode</asp:textbox><asp:textbox id="DFCM" style="Z-INDEX: 122; LEFT: 602px; POSITION: absolute; TOP: 382px; TEXT-ALIGN: right"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px" BorderStyle="Groove">DFCM</asp:textbox><asp:textbox id="DFCH" style="Z-INDEX: 121; LEFT: 518px; POSITION: absolute; TOP: 382px; TEXT-ALIGN: right"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px" BorderStyle="Groove">DFCH</asp:textbox><asp:textbox id="DFBM" style="Z-INDEX: 120; LEFT: 414px; POSITION: absolute; TOP: 382px; TEXT-ALIGN: right"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px" BorderStyle="Groove">DFBM</asp:textbox><asp:textbox id="DFBH" style="Z-INDEX: 119; LEFT: 328px; POSITION: absolute; TOP: 382px; TEXT-ALIGN: right"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px" BorderStyle="Groove">DFBH</asp:textbox><asp:textbox id="DFAM2" style="Z-INDEX: 118; LEFT: 224px; POSITION: absolute; TOP: 382px; TEXT-ALIGN: right"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px" BorderStyle="Groove">DFAM2</asp:textbox><asp:textbox id="DFAH2" style="Z-INDEX: 117; LEFT: 162px; POSITION: absolute; TOP: 382px; TEXT-ALIGN: right"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="40px" BorderStyle="Groove">DFAH2</asp:textbox><asp:textbox id="DFAH1" style="Z-INDEX: 116; LEFT: 162px; POSITION: absolute; TOP: 350px; TEXT-ALIGN: right"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="40px" BorderStyle="Groove">DFAH1</asp:textbox><asp:textbox id="DFAM1" style="Z-INDEX: 115; LEFT: 224px; POSITION: absolute; TOP: 350px; TEXT-ALIGN: right"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px" BorderStyle="Groove">DFAM1</asp:textbox><asp:textbox id="DAM" style="Z-INDEX: 114; LEFT: 604px; POSITION: absolute; TOP: 282px; TEXT-ALIGN: right"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px" BorderStyle="Groove" ReadOnly="True">DAM</asp:textbox><asp:textbox id="DAH" style="Z-INDEX: 113; LEFT: 520px; POSITION: absolute; TOP: 282px; TEXT-ALIGN: right"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px" BorderStyle="Groove" ReadOnly="True">DAH</asp:textbox><asp:textbox id="DBH" style="Z-INDEX: 111; LEFT: 520px; POSITION: absolute; TOP: 250px; TEXT-ALIGN: right"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px" BorderStyle="Groove" ReadOnly="True">DBH</asp:textbox><asp:textbox id="DDivisionCode" style="Z-INDEX: 109; LEFT: 280px; POSITION: absolute; TOP: 152px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="67px" BorderStyle="Groove" ReadOnly="True">DDivisionCode</asp:textbox><asp:textbox id="DDivision" style="Z-INDEX: 108; LEFT: 192px; POSITION: absolute; TOP: 152px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="88px" BorderStyle="Groove" ReadOnly="True">DDivision</asp:textbox><asp:textbox id="DJobTitle" style="Z-INDEX: 107; LEFT: 560px; POSITION: absolute; TOP: 120px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="78px" BorderStyle="Groove" ReadOnly="True">DJobTitle</asp:textbox><asp:textbox id="DEmpID" style="Z-INDEX: 106; LEFT: 336px; POSITION: absolute; TOP: 120px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="20px" Width="114px" BorderStyle="Groove" ReadOnly="True">DEmpID</asp:textbox><asp:textbox id="DName" style="Z-INDEX: 105; LEFT: 120px; POSITION: absolute; TOP: 120px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="20px" Width="216px" BorderStyle="Groove" ReadOnly="True">DName</asp:textbox><asp:textbox id="DDate" style="Z-INDEX: 104; LEFT: 560px; POSITION: absolute; TOP: 86px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="20px" Width="121px" BorderStyle="Groove" ReadOnly="True">DDate</asp:textbox><asp:textbox id="DFReason" style="Z-INDEX: 102; LEFT: 120px; POSITION: absolute; TOP: 416px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="118px" Width="560px" BorderStyle="Groove" TextMode="MultiLine" MaxLength="240">DFReason</asp:textbox><asp:image id="DOverTimeSheet1" style="Z-INDEX: 100; LEFT: 8px; POSITION: absolute; TOP: 6px"
					runat="server" Height="695px" Width="680px" ImageUrl="Images\HR_OverTimeSheet_02.jpg"></asp:image><asp:textbox id="DOverTimeDate" style="Z-INDEX: 103; LEFT: 120px; POSITION: absolute; TOP: 86px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="150px" BorderStyle="Groove" ReadOnly="True">DOverTimeDate</asp:textbox><asp:textbox id="DNo" style="Z-INDEX: 101; LEFT: 8px; POSITION: absolute; TOP: 9px" runat="server"
					ForeColor="Black" BackColor="White" Height="16px" Width="216px" BorderStyle="None" ReadOnly="True" Font-Names="Times New Roman"> 123</asp:textbox></FONT></form>
		</FONT></FORM>
	</body>
</HTML>
