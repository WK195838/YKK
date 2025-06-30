<%@ Page Language="vb" AutoEventWireup="false" Codebehind="HR_TimeOffSheet_02.aspx.vb" Inherits="SPD.HR_TimeOffSheet_02"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>請假申請書</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<script language="javascript" type="text/javascript">
			//--Calendar------------------------------------
			function ShowCardTime()
			{
				window.open('HR_CardTimeList01.aspx?pMonth=' + document.Form1.DSalaryYM.value + '&pDepoID=' + document.Form1.DDepoCode.value + '&pEmpID=' + document.Form1.DEmpID.value, 'CardTimeSheet','width=290,height=600,resizable=yes,scrollbar=yes');
			}

			function ShowVacation()
			{
				window.open('http://10.245.1.10/WorkFlowSub/HR_VacationList01.aspx?pMonth=' + document.Form1.DSalaryYM.value + '&pDepoID=' + document.Form1.DDepoCode.value + '&pEmpID=' + document.Form1.DEmpID.value + '&pNo=' + document.Form1.DNo.value + '&pVacation=' + document.Form1.DVacation.value + '&pDieType=' + document.Form1.DDieType.value, 'VacationSheet','width=520,height=600,resizable=yes,scrollbar=yes');
			}
		</script>
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" encType="multipart/form-data" runat="server">
			<FONT face="新細明體">
				<asp:button id="BVARecord" style="Z-INDEX: 131; POSITION: absolute; TOP: 216px; LEFT: 592px"
					runat="server" Text="請假記錄" BackColor="#FFFFC0" Height="24px" Width="85px" CausesValidation="False"></asp:button><asp:textbox id="DAEndH" style="Z-INDEX: 169; POSITION: absolute; TOP: 350px; LEFT: 472px" runat="server"
					BackColor="Yellow" Height="20px" Width="40px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">2008/10/10</asp:textbox><asp:textbox id="DBEndH" style="Z-INDEX: 168; POSITION: absolute; TOP: 318px; LEFT: 472px" runat="server"
					BackColor="Yellow" Height="20px" Width="40px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">2008/10/10</asp:textbox><asp:textbox id="DAStartH" style="Z-INDEX: 166; POSITION: absolute; TOP: 350px; LEFT: 240px"
					runat="server" BackColor="Yellow" Height="20px" Width="40px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">2008/10/10</asp:textbox><asp:textbox id="DBStartH" style="Z-INDEX: 165; POSITION: absolute; TOP: 318px; LEFT: 240px"
					runat="server" BackColor="Yellow" Height="20px" Width="40px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">2008/10/10</asp:textbox><asp:textbox id="DDieType" style="Z-INDEX: 164; POSITION: absolute; TOP: 218px; LEFT: 120px"
					runat="server" BackColor="Yellow" Height="20px" Width="184px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">DDieType</asp:textbox><asp:textbox id="DVacation" style="Z-INDEX: 163; POSITION: absolute; TOP: 186px; LEFT: 120px"
					runat="server" BackColor="Yellow" Height="20px" Width="184px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">DVacation</asp:textbox><asp:textbox id="DAfter" style="Z-INDEX: 162; POSITION: absolute; TOP: 152px; LEFT: 120px" runat="server"
					BackColor="Yellow" Height="20px" Width="100px" BorderStyle="Groove" ForeColor="Blue">DAfter</asp:textbox><asp:textbox id="DName" style="Z-INDEX: 161; POSITION: absolute; TOP: 88px; LEFT: 456px" runat="server"
					BackColor="Yellow" Height="20px" Width="144px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">DName</asp:textbox><asp:textbox id="DAEndDate" style="Z-INDEX: 160; POSITION: absolute; TOP: 350px; LEFT: 352px"
					runat="server" BackColor="Yellow" Height="20px" Width="120px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">2008/10/10</asp:textbox><asp:textbox id="DAStartDate" style="Z-INDEX: 159; POSITION: absolute; TOP: 350px; LEFT: 120px"
					runat="server" BackColor="Yellow" Height="20px" Width="120px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">2008/10/10</asp:textbox><asp:textbox id="DBEndDate" style="Z-INDEX: 158; POSITION: absolute; TOP: 318px; LEFT: 352px"
					runat="server" BackColor="Yellow" Height="20px" Width="120px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">2008/10/10</asp:textbox><asp:textbox id="DBStartDate" style="Z-INDEX: 157; POSITION: absolute; TOP: 318px; LEFT: 120px"
					runat="server" BackColor="Yellow" Height="20px" Width="120px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">2008/10/10</asp:textbox><asp:hyperlink id="LOTNo4" style="Z-INDEX: 156; POSITION: absolute; TOP: 280px; LEFT: 328px" runat="server"
					Height="8px" Width="64px" Font-Names="Times New Roman" Font-Size="Small" Target="_blank" NavigateUrl="BoardEdit.aspx">06020000000086</asp:hyperlink><asp:hyperlink id="LOTNo2" style="Z-INDEX: 155; POSITION: absolute; TOP: 280px; LEFT: 144px" runat="server"
					Height="8px" Width="64px" Font-Names="Times New Roman" Font-Size="Small" Target="_blank" NavigateUrl="BoardEdit.aspx">06020000000086</asp:hyperlink><asp:hyperlink id="LOTNo5" style="Z-INDEX: 154; POSITION: absolute; TOP: 256px; LEFT: 512px" runat="server"
					Height="8px" Width="64px" Font-Names="Times New Roman" Font-Size="Small" Target="_blank" NavigateUrl="BoardEdit.aspx">06020000000086</asp:hyperlink><asp:hyperlink id="LOTNo3" style="Z-INDEX: 153; POSITION: absolute; TOP: 256px; LEFT: 328px" runat="server"
					Height="8px" Width="64px" Font-Names="Times New Roman" Font-Size="Small" Target="_blank" NavigateUrl="BoardEdit.aspx">06020000000086</asp:hyperlink><asp:hyperlink id="LOTNo1" style="Z-INDEX: 152; POSITION: absolute; TOP: 256px; LEFT: 144px" runat="server"
					Height="8px" Width="64px" Font-Names="Times New Roman" Target="_blank" NavigateUrl="BoardEdit.aspx">06020000000086</asp:hyperlink><asp:textbox id="DOTHours5" style="Z-INDEX: 150; POSITION: absolute; TEXT-ALIGN: right; TOP: 254px; LEFT: 632px"
					runat="server" BackColor="Yellow" Height="20px" Width="32px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">DOTHours</asp:textbox><asp:textbox id="Label5" style="Z-INDEX: 149; POSITION: absolute; TOP: 254px; LEFT: 496px" runat="server"
					BackColor="#FFFFC0" Height="16px" Width="12px" BorderStyle="None" ReadOnly="True" ForeColor="Black" Font-Names="Times New Roman">5.</asp:textbox><asp:textbox id="DOTHours4" style="Z-INDEX: 148; POSITION: absolute; TEXT-ALIGN: right; TOP: 278px; LEFT: 448px"
					runat="server" BackColor="Yellow" Height="20px" Width="32px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">DOTHours</asp:textbox><asp:textbox id="Label4" style="Z-INDEX: 147; POSITION: absolute; TOP: 278px; LEFT: 312px" runat="server"
					BackColor="#FFFFC0" Height="16px" Width="12px" BorderStyle="None" ReadOnly="True" ForeColor="Black" Font-Names="Times New Roman">4.</asp:textbox><asp:textbox id="DOTHours3" style="Z-INDEX: 146; POSITION: absolute; TEXT-ALIGN: right; TOP: 254px; LEFT: 448px"
					runat="server" BackColor="Yellow" Height="20px" Width="32px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">DOTHours</asp:textbox><asp:textbox id="DOTHours2" style="Z-INDEX: 145; POSITION: absolute; TEXT-ALIGN: right; TOP: 278px; LEFT: 264px"
					runat="server" BackColor="Yellow" Height="20px" Width="32px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">DOTHours</asp:textbox><asp:textbox id="DOTHours1" style="Z-INDEX: 144; POSITION: absolute; TEXT-ALIGN: right; TOP: 254px; LEFT: 264px"
					runat="server" BackColor="Yellow" Height="20px" Width="32px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">10.5</asp:textbox><asp:textbox id="Label3" style="Z-INDEX: 143; POSITION: absolute; TOP: 254px; LEFT: 312px" runat="server"
					BackColor="#FFFFC0" Height="16px" Width="12px" BorderStyle="None" ReadOnly="True" ForeColor="Black" Font-Names="Times New Roman">3.</asp:textbox><asp:textbox id="Label2" style="Z-INDEX: 142; POSITION: absolute; TOP: 278px; LEFT: 128px" runat="server"
					BackColor="#FFFFC0" Height="16px" Width="12px" BorderStyle="None" ReadOnly="True" ForeColor="Black" Font-Names="Times New Roman">2.</asp:textbox><asp:textbox id="Label1" style="Z-INDEX: 141; POSITION: absolute; TOP: 254px; LEFT: 128px" runat="server"
					BackColor="#FFFFC0" Height="16px" Width="12px" BorderStyle="None" ReadOnly="True" ForeColor="Black" Font-Names="Times New Roman">1.</asp:textbox><asp:textbox id="DOTHours" style="Z-INDEX: 135; POSITION: absolute; TEXT-ALIGN: right; TOP: 288px; LEFT: 600px"
					runat="server" BackColor="Yellow" Height="20px" Width="40px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">0</asp:textbox><asp:textbox id="DADays" style="Z-INDEX: 134; POSITION: absolute; TEXT-ALIGN: right; TOP: 350px; LEFT: 600px"
					runat="server" BackColor="Yellow" Height="20px" Width="40px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">0</asp:textbox><asp:button id="BCardTime" style="Z-INDEX: 130; POSITION: absolute; TOP: 216px; LEFT: 502px"
					runat="server" Text="刷卡記錄" BackColor="#FFFFC0" Height="24px" Width="85px" CausesValidation="False"></asp:button>
				<asp:textbox id="DVDays" style="Z-INDEX: 131; POSITION: absolute; TOP: 218px; LEFT: -500px" runat="server"
					BackColor="Yellow" Height="20px" Width="75px" BorderStyle="Groove" ForeColor="Blue">DVDays</asp:textbox>
				<asp:textbox id="DVDaysBlank" style="Z-INDEX: 101; POSITION: absolute; TOP: 216px; LEFT: 320px"
					runat="server" BackColor="#FFFF80" Height="24px" Width="80px" BorderStyle="None" ReadOnly="True"
					ForeColor="Black" Font-Names="Times New Roman"></asp:textbox>
				<asp:textbox id="DSalary" style="Z-INDEX: 128; POSITION: absolute; TOP: 186px; LEFT: 606px" runat="server"
					BackColor="Yellow" Height="20px" Width="75px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">DSalary</asp:textbox><asp:textbox id="DEvidence" style="Z-INDEX: 127; POSITION: absolute; TOP: 186px; LEFT: 416px"
					runat="server" BackColor="Yellow" Height="20px" Width="75px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">DEvidence</asp:textbox><asp:textbox id="DTimeOffAgent" style="Z-INDEX: 126; POSITION: absolute; TOP: 152px; LEFT: 562px"
					runat="server" BackColor="Yellow" Height="20px" Width="119px" BorderStyle="Groove" ForeColor="Blue">DTimeOffAgent</asp:textbox><asp:textbox id="DJobAgent" style="Z-INDEX: 125; POSITION: absolute; TOP: 152px; LEFT: 332px"
					runat="server" BackColor="Yellow" Height="20px" Width="118px" BorderStyle="Groove" ForeColor="Blue">DJobAgent</asp:textbox><asp:textbox id="DDepoCode" style="Z-INDEX: 124; POSITION: absolute; TOP: 120px; LEFT: 496px"
					runat="server" BackColor="Yellow" Height="20px" Width="25px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">01</asp:textbox><asp:textbox id="DDepoName" style="Z-INDEX: 123; POSITION: absolute; TOP: 120px; LEFT: 456px"
					runat="server" BackColor="Yellow" Height="20px" Width="40px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">台北</asp:textbox><asp:textbox id="DSalaryYM" style="Z-INDEX: 122; POSITION: absolute; TOP: 88px; LEFT: 264px"
					runat="server" BackColor="Yellow" Height="20px" Width="81px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">2008/10</asp:textbox><asp:textbox id="DJobCode" style="Z-INDEX: 114; POSITION: absolute; TOP: 120px; LEFT: 264px"
					runat="server" BackColor="Yellow" Height="20px" Width="81px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">DJobCode</asp:textbox><asp:textbox id="DBDays" style="Z-INDEX: 112; POSITION: absolute; TEXT-ALIGN: right; TOP: 318px; LEFT: 600px"
					runat="server" BackColor="Yellow" Height="20px" Width="40px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">0</asp:textbox><asp:textbox id="DDivisionCode" style="Z-INDEX: 111; POSITION: absolute; TOP: 120px; LEFT: 616px"
					runat="server" BackColor="Yellow" Height="20px" Width="65px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">DDivisionCode</asp:textbox><asp:textbox id="DDivision" style="Z-INDEX: 110; POSITION: absolute; TOP: 120px; LEFT: 528px"
					runat="server" BackColor="Yellow" Height="20px" Width="88px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">DDivision</asp:textbox><asp:textbox id="DJobTitle" style="Z-INDEX: 109; POSITION: absolute; TOP: 120px; LEFT: 120px"
					runat="server" BackColor="Yellow" Height="20px" Width="144px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">DJobTitle</asp:textbox><asp:textbox id="DEmpID" style="Z-INDEX: 108; POSITION: absolute; TOP: 88px; LEFT: 600px" runat="server"
					BackColor="Yellow" Height="20px" Width="81px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">DEmpID</asp:textbox><asp:textbox id="DDate" style="Z-INDEX: 107; POSITION: absolute; TOP: 88px; LEFT: 120px" runat="server"
					BackColor="Yellow" Height="20px" Width="144px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">DDate</asp:textbox><asp:textbox id="DFReason" style="Z-INDEX: 102; POSITION: absolute; TOP: 384px; LEFT: 120px"
					runat="server" BackColor="Yellow" Height="122px" Width="560px" BorderStyle="Groove" ForeColor="Blue" TextMode="MultiLine" MaxLength="240">DFReason</asp:textbox><asp:image id="DTimeOffSheet1" style="Z-INDEX: 100; POSITION: absolute; TOP: 8px; LEFT: 8px"
					runat="server" Height="505px" Width="680px" ImageUrl="Images\HR_TimeOffSheet_01.jpg"></asp:image><asp:textbox id="DNo" style="Z-INDEX: 101; POSITION: absolute; TOP: 9px; LEFT: 8px" runat="server"
					BackColor="White" Height="16px" Width="216px" BorderStyle="None" ReadOnly="True" ForeColor="Black" Font-Names="Times New Roman"> 123</asp:textbox></FONT></form>
		</FONT></FORM>
	</body>
</HTML>
