<%@ Page Language="vb" AutoEventWireup="false" Codebehind="HR_TimeOffSheet_01.aspx.vb" Inherits="SPD.HR_TimeOffSheet_01" aspCompat="True"%>
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
			function CalendarPicker(strField1, strField2, strField3) {
				window.open('VacationDatePicker.aspx?field1=' + strField1 + '&field2=' + strField2 + '&field3=' + strField3,'CalendarPopup','width=250,height=190,resizable=yes');
			}

			function ShowCardTime()
			{
				window.open('HR_CardTimeList01.aspx?pMonth=' + document.Form1.DSalaryYM.value + '&pDepoID=' + document.Form1.DDepoCode.value + '&pEmpID=' + document.Form1.DEmpID.value, 'CardTimeSheet','width=290,height=600,resizable=yes,scrollbar=yes');
			}

			function ShowVacation()
			{
				window.open('http://10.245.1.6/WorkFlowSub/HR_VacationList01.aspx?pMonth=' + document.Form1.DSalaryYM.value + '&pDepoID=' + document.Form1.DDepoCode.value + '&pEmpID=' + document.Form1.DEmpID.value + '&pNo=' + document.Form1.DNo.value + '&pVacation=' + document.Form1.DVacation.value + '&pDieType=' + document.Form1.DDieType.value, 'VacationSheet','width=520,height=600,resizable=yes,scrollbar=yes');
			}

			function ShowOverTime()
			{
				window.open('http://10.245.1.10/WorkFlowSub1/HR_OverTimeList01.aspx?pMonth=' + document.Form1.DSalaryYM.value + '&pDepoID=' + document.Form1.DDepoCode.value + '&pEmpID=' + document.Form1.DEmpID.value, 'VacationSheet','width=290,height=800,resizable=yes,scrollbar=yes');
			}

			function ShowAOverTime()
			{
				window.open('http://10.245.1.6/WorkFlowSub/HR_OverTimeAndTimeOffList.aspx?pMonth=' + document.Form1.DSalaryYM.value + '&pDepoID=' + document.Form1.DDepoCode.value + '&pEmpID=' + document.Form1.DEmpID.value, 'OverTimeVacationSheet','width=710,height=500,resizable=yes,scrollbar=yes');
			}

		    function Button(F, MSG) {
				//alert(F);
				document.cookie="RunBOK=False";
				document.cookie="RunBNG1=False";
				document.cookie="RunBNG2=False";
				document.cookie="RunBSAVE=False";

				answer = confirm("是否執行[" + MSG + "]作業嗎？ 請確認....");
				if (answer) {
					//OK Button
					//FOK=document.getElementById("BOK");
					//if(FOK!=null) document.Form1.BOK.disabled=true;  	
					//NG-1 Button
					//FNG1=document.getElementById("BNG1");
					//if(FNG1!=null) document.Form1.BNG1.disabled=true;  	
					//NG-2 Button
					//FNG2=document.getElementById("BNG2");
					//if(FNG2!=null) document.Form1.BNG2.disabled=true;  	
					//Save Button
					//FSAVE=document.getElementById("BSAVE");
					//if(FSAVE!=null) document.Form1.BSAVE.disabled=true;  	
						
					if (F=="OK")   document.cookie="RunBOK=True";
					if (F=="NG1")  document.cookie="RunBNG1=True";
					if (F=="NG2")  document.cookie="RunBNG2=True";
					if (F=="SAVE") document.cookie="RunBSAVE=True";
				}
			}
		</script>
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" encType="multipart/form-data" runat="server">
			<FONT face="新細明體">
				<asp:button id="BVARecord" style="Z-INDEX: 133; POSITION: absolute; TOP: 216px; LEFT: 560px"
					runat="server" Text="請假" BackColor="#FFFFC0" Height="24px" Width="60px" CausesValidation="False"></asp:button>
				<asp:textbox id="DInDate" style="Z-INDEX: 175; POSITION: absolute; TOP: 120px; LEFT: 264px" runat="server"
					Width="81px" Height="20px" BackColor="LightGray" ForeColor="Blue" ReadOnly="True" BorderStyle="Groove">2005/04/06</asp:textbox><asp:button id="BOverTime" style="Z-INDEX: 173; POSITION: absolute; TOP: 216px; LEFT: 622px"
					runat="server" Text="調休" BackColor="#FFFFC0" Height="24px" Width="60px" CausesValidation="False"></asp:button><asp:button id="BADays" style="Z-INDEX: 172; POSITION: absolute; TOP: 350px; LEFT: 540px" runat="server"
					Text="→" Height="24px" Width="30px" CausesValidation="False"></asp:button><asp:button id="BBDays" style="Z-INDEX: 171; POSITION: absolute; TOP: 318px; LEFT: 540px" runat="server"
					Text="→" Height="24px" Width="30px" CausesValidation="False"></asp:button><asp:button id="BAEndDate" style="Z-INDEX: 170; POSITION: absolute; TOP: 352px; LEFT: 448px"
					runat="server" Text="....." Height="20px" Width="20px" CausesValidation="False"></asp:button><asp:button id="BAStartDate" style="Z-INDEX: 169; POSITION: absolute; TOP: 352px; LEFT: 216px"
					runat="server" Text="....." Height="20px" Width="20px" CausesValidation="False"></asp:button><asp:textbox id="DAEndDate" style="Z-INDEX: 168; POSITION: absolute; TOP: 350px; LEFT: 352px"
					runat="server" BackColor="Yellow" Height="20px" Width="95px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">2008/10/10</asp:textbox><asp:textbox id="DAStartDate" style="Z-INDEX: 167; POSITION: absolute; TOP: 350px; LEFT: 120px"
					runat="server" BackColor="Yellow" Height="20px" Width="95px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">2008/10/10</asp:textbox><asp:button id="BBEndDate" style="Z-INDEX: 166; POSITION: absolute; TOP: 320px; LEFT: 448px"
					runat="server" Text="....." Height="20px" Width="20px" CausesValidation="False"></asp:button><asp:textbox id="DBEndDate" style="Z-INDEX: 165; POSITION: absolute; TOP: 318px; LEFT: 352px"
					runat="server" BackColor="Yellow" Height="20px" Width="95px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">2008/10/10</asp:textbox><asp:button id="BBStartDate" style="Z-INDEX: 164; POSITION: absolute; TOP: 320px; LEFT: 216px"
					runat="server" Text="....." Height="20px" Width="20px" CausesValidation="False"></asp:button><asp:textbox id="DBStartDate" style="Z-INDEX: 163; POSITION: absolute; TOP: 318px; LEFT: 120px"
					runat="server" BackColor="Yellow" Height="20px" Width="95px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">2008/10/10</asp:textbox><asp:dropdownlist id="DAfter" style="Z-INDEX: 162; POSITION: absolute; TOP: 152px; LEFT: 120px" runat="server"
					BackColor="Yellow" Height="20px" Width="100px" ForeColor="Blue" AutoPostBack="True"></asp:dropdownlist><asp:dropdownlist id="DName" style="Z-INDEX: 161; POSITION: absolute; TOP: 88px; LEFT: 456px" runat="server"
					BackColor="Yellow" Height="20px" Width="144px" ForeColor="Blue" AutoPostBack="True"></asp:dropdownlist><asp:hyperlink id="LOTNo4" style="Z-INDEX: 160; POSITION: absolute; TOP: 280px; LEFT: 328px" runat="server"
					Height="8px" Width="64px" Font-Names="Times New Roman" Font-Size="Small" Target="_blank" NavigateUrl="BoardEdit.aspx">06020000000086</asp:hyperlink><asp:hyperlink id="LOTNo2" style="Z-INDEX: 159; POSITION: absolute; TOP: 280px; LEFT: 144px" runat="server"
					Height="8px" Width="64px" Font-Names="Times New Roman" Font-Size="Small" Target="_blank" NavigateUrl="BoardEdit.aspx">06020000000086</asp:hyperlink><asp:hyperlink id="LOTNo5" style="Z-INDEX: 158; POSITION: absolute; TOP: 256px; LEFT: 512px" runat="server"
					Height="8px" Width="64px" Font-Names="Times New Roman" Font-Size="Small" Target="_blank" NavigateUrl="BoardEdit.aspx">06020000000086</asp:hyperlink><asp:hyperlink id="LOTNo3" style="Z-INDEX: 157; POSITION: absolute; TOP: 256px; LEFT: 328px" runat="server"
					Height="8px" Width="64px" Font-Names="Times New Roman" Font-Size="Small" Target="_blank" NavigateUrl="BoardEdit.aspx">06020000000086</asp:hyperlink><asp:hyperlink id="LOTNo1" style="Z-INDEX: 156; POSITION: absolute; TOP: 256px; LEFT: 144px" runat="server"
					Height="8px" Width="64px" Font-Names="Times New Roman" Target="_blank" NavigateUrl="BoardEdit.aspx">06020000000086</asp:hyperlink><asp:button id="BOTRecord" style="Z-INDEX: 155; POSITION: absolute; TOP: 280px; LEFT: 488px"
					runat="server" Text="加班調休" BackColor="#FFFFC0" Height="24px" Width="85px" CausesValidation="False" Enabled="False"></asp:button><asp:textbox id="DOTHours5" style="Z-INDEX: 154; POSITION: absolute; TEXT-ALIGN: right; TOP: 254px; LEFT: 632px"
					runat="server" BackColor="Yellow" Height="20px" Width="32px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">DOTHours</asp:textbox><asp:textbox id="Label5" style="Z-INDEX: 153; POSITION: absolute; TOP: 254px; LEFT: 496px" runat="server"
					BackColor="#FFFFC0" Height="16px" Width="12px" BorderStyle="None" ReadOnly="True" ForeColor="Black" Font-Names="Times New Roman">5.</asp:textbox><asp:textbox id="DOTHours4" style="Z-INDEX: 152; POSITION: absolute; TEXT-ALIGN: right; TOP: 278px; LEFT: 448px"
					runat="server" BackColor="Yellow" Height="20px" Width="32px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">DOTHours</asp:textbox><asp:textbox id="Label4" style="Z-INDEX: 151; POSITION: absolute; TOP: 278px; LEFT: 312px" runat="server"
					BackColor="#FFFFC0" Height="16px" Width="12px" BorderStyle="None" ReadOnly="True" ForeColor="Black" Font-Names="Times New Roman">4.</asp:textbox><asp:textbox id="DOTHours3" style="Z-INDEX: 150; POSITION: absolute; TEXT-ALIGN: right; TOP: 254px; LEFT: 448px"
					runat="server" BackColor="Yellow" Height="20px" Width="32px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">DOTHours</asp:textbox><asp:textbox id="DOTHours2" style="Z-INDEX: 149; POSITION: absolute; TEXT-ALIGN: right; TOP: 278px; LEFT: 264px"
					runat="server" BackColor="Yellow" Height="20px" Width="32px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">DOTHours</asp:textbox><asp:textbox id="DOTHours1" style="Z-INDEX: 148; POSITION: absolute; TEXT-ALIGN: right; TOP: 254px; LEFT: 264px"
					runat="server" BackColor="Yellow" Height="20px" Width="32px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">10.5</asp:textbox><asp:textbox id="Label3" style="Z-INDEX: 147; POSITION: absolute; TOP: 254px; LEFT: 312px" runat="server"
					BackColor="#FFFFC0" Height="16px" Width="12px" BorderStyle="None" ReadOnly="True" ForeColor="Black" Font-Names="Times New Roman">3.</asp:textbox><asp:textbox id="Label2" style="Z-INDEX: 146; POSITION: absolute; TOP: 278px; LEFT: 128px" runat="server"
					BackColor="#FFFFC0" Height="16px" Width="12px" BorderStyle="None" ReadOnly="True" ForeColor="Black" Font-Names="Times New Roman">2.</asp:textbox><asp:textbox id="Label1" style="Z-INDEX: 145; POSITION: absolute; TOP: 254px; LEFT: 128px" runat="server"
					BackColor="#FFFFC0" Height="16px" Width="12px" BorderStyle="None" ReadOnly="True" ForeColor="Black" Font-Names="Times New Roman">1.</asp:textbox><asp:textbox id="DOTNo4" style="Z-INDEX: 144; POSITION: absolute; TOP: 278px; LEFT: 328px" runat="server"
					BackColor="Yellow" Height="20px" Width="115px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">06020000000086</asp:textbox><asp:textbox id="DOTNo5" style="Z-INDEX: 143; POSITION: absolute; TOP: 254px; LEFT: 512px" runat="server"
					BackColor="Yellow" Height="20px" Width="115px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">06020000000086</asp:textbox><asp:textbox id="DOTNo3" style="Z-INDEX: 142; POSITION: absolute; TOP: 254px; LEFT: 328px" runat="server"
					BackColor="Yellow" Height="20px" Width="115px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">06020000000086</asp:textbox><asp:textbox id="DOTNo2" style="Z-INDEX: 141; POSITION: absolute; TOP: 278px; LEFT: 144px" runat="server"
					BackColor="Yellow" Height="20px" Width="115px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">06020000000086</asp:textbox><asp:textbox id="DOTNo1" style="Z-INDEX: 140; POSITION: absolute; TOP: 254px; LEFT: 144px" runat="server"
					BackColor="Yellow" Height="20px" Width="115px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">06020000000086</asp:textbox><asp:textbox id="DOTHours" style="Z-INDEX: 139; POSITION: absolute; TEXT-ALIGN: right; TOP: 288px; LEFT: 600px"
					runat="server" BackColor="Yellow" Height="20px" Width="40px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">0</asp:textbox><asp:textbox id="DADays" style="Z-INDEX: 138; POSITION: absolute; TEXT-ALIGN: right; TOP: 350px; LEFT: 600px"
					runat="server" BackColor="Yellow" Height="20px" Width="40px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">0</asp:textbox><asp:dropdownlist id="DAEndH" style="Z-INDEX: 137; POSITION: absolute; TOP: 350px; LEFT: 472px" runat="server"
					BackColor="Yellow" Height="20px" Width="46px" ForeColor="Blue"></asp:dropdownlist><asp:dropdownlist id="DAStartH" style="Z-INDEX: 136; POSITION: absolute; TOP: 350px; LEFT: 240px"
					runat="server" BackColor="Yellow" Height="20px" Width="46px" ForeColor="Blue"></asp:dropdownlist><asp:dropdownlist id="DBEndH" style="Z-INDEX: 135; POSITION: absolute; TOP: 318px; LEFT: 472px" runat="server"
					BackColor="Yellow" Height="20px" Width="46px" ForeColor="Blue"></asp:dropdownlist><asp:dropdownlist id="DBStartH" style="Z-INDEX: 134; POSITION: absolute; TOP: 318px; LEFT: 240px"
					runat="server" BackColor="Yellow" Height="20px" Width="46px" ForeColor="Blue"></asp:dropdownlist><asp:button id="BCardTime" style="Z-INDEX: 132; POSITION: absolute; TOP: 216px; LEFT: 498px"
					runat="server" Text="刷卡" BackColor="#FFFFC0" Height="24px" Width="60px" CausesValidation="False"></asp:button>
				<asp:textbox id="DVDays" style="Z-INDEX: 131; POSITION: absolute; TOP: 218px; LEFT: -500px" runat="server"
					BackColor="Yellow" Height="20px" Width="75px" BorderStyle="Groove" ForeColor="Blue">DVDays</asp:textbox>
				<asp:textbox id="DVDaysBlank" style="Z-INDEX: 101; POSITION: absolute; TOP: 216px; LEFT: 320px"
					runat="server" BackColor="#FFFF80" Height="24px" Width="80px" BorderStyle="None" ReadOnly="True"
					ForeColor="Black" Font-Names="Times New Roman"></asp:textbox>
				<asp:textbox id="DSalary" style="Z-INDEX: 130; POSITION: absolute; TOP: 186px; LEFT: 606px" runat="server"
					BackColor="Yellow" Height="20px" Width="75px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">DSalary</asp:textbox><asp:textbox id="DEvidence" style="Z-INDEX: 129; POSITION: absolute; TOP: 186px; LEFT: 416px"
					runat="server" BackColor="Yellow" Height="20px" Width="75px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">DEvidence</asp:textbox><asp:textbox id="DTimeOffAgent" style="Z-INDEX: 128; POSITION: absolute; TOP: 152px; LEFT: 562px"
					runat="server" BackColor="Yellow" Height="20px" Width="119px" BorderStyle="Groove" ForeColor="Blue">DTimeOffAgent</asp:textbox><asp:textbox id="DDepoCode" style="Z-INDEX: 127; POSITION: absolute; TOP: 120px; LEFT: 496px"
					runat="server" BackColor="Yellow" Height="20px" Width="25px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">01</asp:textbox><asp:textbox id="DDepoName" style="Z-INDEX: 126; POSITION: absolute; TOP: 120px; LEFT: 456px"
					runat="server" BackColor="Yellow" Height="20px" Width="40px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">台北</asp:textbox><asp:textbox id="DSalaryYM" style="Z-INDEX: 125; POSITION: absolute; TOP: 88px; LEFT: 264px"
					runat="server" BackColor="Yellow" Height="20px" Width="81px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">2008/10</asp:textbox><asp:label id="DHistoryLabel" style="Z-INDEX: 123; POSITION: absolute; TOP: 736px; LEFT: 8px"
					runat="server" Height="16px" Width="64px" ForeColor="Blue" Font-Names="新細明體" Font-Size="11pt">核定履歷</asp:label><asp:textbox id="DReasonDesc" style="Z-INDEX: 122; POSITION: absolute; TOP: 640px; LEFT: 168px"
					runat="server" BackColor="Yellow" Height="59px" Width="424px" BorderStyle="Groove" ForeColor="Blue" TextMode="MultiLine" Visible="False">DReasonDesc</asp:textbox><asp:textbox id="DReason" style="Z-INDEX: 121; POSITION: absolute; TOP: 608px; LEFT: 240px" runat="server"
					BackColor="Yellow" Height="20px" Width="352px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue" Visible="False">DReason</asp:textbox><asp:dropdownlist id="DReasonCode" style="Z-INDEX: 120; POSITION: absolute; TOP: 608px; LEFT: 168px"
					runat="server" BackColor="Yellow" Height="20px" Width="64px" ForeColor="Blue" AutoPostBack="True" Visible="False">
					<asp:ListItem Value="01">01</asp:ListItem>
					<asp:ListItem Value="02">02</asp:ListItem>
				</asp:dropdownlist><asp:image id="DDelay" style="Z-INDEX: 119; POSITION: absolute; TOP: 600px; LEFT: 8px" runat="server"
					Height="107px" Width="593px" Visible="False" ImageUrl="Images\Sheet_Delay.jpg"></asp:image><asp:textbox id="DDecideDesc" style="Z-INDEX: 118; POSITION: absolute; TOP: 528px; LEFT: 56px"
					runat="server" BackColor="Yellow" Height="56px" Width="536px" BorderStyle="Groove" ForeColor="Blue" TextMode="MultiLine">DDecideDesc</asp:textbox><asp:image id="DDescSheet" style="Z-INDEX: 117; POSITION: absolute; TOP: 520px; LEFT: 8px"
					runat="server" Height="75px" Width="593px" ImageUrl="Images\MapSheet_004.jpg"></asp:image><asp:textbox id="DJobCode" style="Z-INDEX: 116; POSITION: absolute; TOP: 120px; LEFT: 200px"
					runat="server" BackColor="Yellow" Height="20px" Width="64px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">DJobCode</asp:textbox><asp:datagrid id="DataGrid9" style="Z-INDEX: 115; POSITION: absolute; TOP: 752px; LEFT: 8px" runat="server"
					BackColor="White" Height="100px" Width="780px" BorderStyle="None" Font-Size="9pt" AutoGenerateColumns="False" CellPadding="4" BorderWidth="1px" BorderColor="#CC9966">
					<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
					<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
					<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
					<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
					<Columns>
						<asp:BoundColumn DataField="StepNameDesc" ReadOnly="True" HeaderText="工程">
							<HeaderStyle Width="170px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="DecideName" ReadOnly="True" HeaderText="擔當">
							<HeaderStyle Width="50px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="AgentName" ReadOnly="True" HeaderText="代理/兼職">
							<HeaderStyle Width="70px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="FlowTypeDesc" ReadOnly="True" HeaderText="類別">
							<HeaderStyle Width="50px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="StsDesc" ReadOnly="True" HeaderText="處理結果">
							<HeaderStyle Width="50px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="DecideDescA" ReadOnly="True" HeaderText="說明">
							<HeaderStyle Width="200px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="Description" ReadOnly="True" HeaderText="核定時間">
							<HeaderStyle Width="150px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
					</Columns>
					<PagerStyle NextPageText="下一頁" PrevPageText="上一頁" HorizontalAlign="Center" ForeColor="#993300"
						BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
				</asp:datagrid><asp:textbox id="DJobAgent" style="Z-INDEX: 174; POSITION: absolute; TOP: 152px; LEFT: 330px"
					runat="server" BackColor="Yellow" Height="20px" Width="119px" BorderStyle="Groove" ForeColor="Blue">DJobAgent</asp:textbox><asp:textbox id="DBDays" style="Z-INDEX: 114; POSITION: absolute; TEXT-ALIGN: right; TOP: 318px; LEFT: 600px"
					runat="server" BackColor="Yellow" Height="20px" Width="40px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">0</asp:textbox><asp:dropdownlist id="DDieType" style="Z-INDEX: 113; POSITION: absolute; TOP: 218px; LEFT: 120px"
					runat="server" BackColor="Yellow" Height="20px" Width="184px" ForeColor="Blue" AutoPostBack="True" Enabled="False"></asp:dropdownlist><asp:dropdownlist id="DVacation" style="Z-INDEX: 112; POSITION: absolute; TOP: 186px; LEFT: 120px"
					runat="server" BackColor="Yellow" Height="20px" Width="184px" ForeColor="Blue" AutoPostBack="True"></asp:dropdownlist><asp:textbox id="DDivisionCode" style="Z-INDEX: 111; POSITION: absolute; TOP: 120px; LEFT: 616px"
					runat="server" BackColor="Yellow" Height="20px" Width="65px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">DDivisionCode</asp:textbox><asp:textbox id="DDivision" style="Z-INDEX: 110; POSITION: absolute; TOP: 120px; LEFT: 528px"
					runat="server" BackColor="Yellow" Height="20px" Width="88px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">DDivision</asp:textbox><asp:textbox id="DJobTitle" style="Z-INDEX: 109; POSITION: absolute; TOP: 120px; LEFT: 120px"
					runat="server" BackColor="Yellow" Height="20px" Width="80px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">DJobTitle</asp:textbox><asp:textbox id="DEmpID" style="Z-INDEX: 108; POSITION: absolute; TOP: 88px; LEFT: 600px" runat="server"
					BackColor="Yellow" Height="20px" Width="81px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">DEmpID</asp:textbox><asp:textbox id="DDate" style="Z-INDEX: 107; POSITION: absolute; TOP: 88px; LEFT: 120px" runat="server"
					BackColor="Yellow" Height="20px" Width="144px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue">DDate</asp:textbox><asp:textbox id="DFReason" style="Z-INDEX: 102; POSITION: absolute; TOP: 384px; LEFT: 120px"
					runat="server" BackColor="Yellow" Height="88px" Width="560px" BorderStyle="Groove" ForeColor="Blue" TextMode="MultiLine" MaxLength="240">DFReason</asp:textbox><asp:image id="DTimeOffSheet1" style="Z-INDEX: 100; POSITION: absolute; TOP: 8px; LEFT: 8px"
					runat="server" Height="505px" Width="680px" ImageUrl="Images\HR_TimeOffSheet_02.jpg"></asp:image><asp:textbox id="DNo" style="Z-INDEX: 101; POSITION: absolute; TOP: 9px; LEFT: 8px" runat="server"
					BackColor="White" Height="16px" Width="216px" BorderStyle="None" ReadOnly="True" ForeColor="Black" Font-Names="Times New Roman"> 123</asp:textbox></FONT><INPUT id="BSAVE" style="Z-INDEX: 103; POSITION: absolute; WIDTH: 80px; HEIGHT: 28px; TOP: 712px; LEFT: 336px"
				onclick="Button('SAVE');" type="button" value="儲存" name="Button1" runat="server"><INPUT id="BNG2" style="Z-INDEX: 104; POSITION: absolute; WIDTH: 80px; HEIGHT: 28px; TOP: 712px; LEFT: 424px"
				onclick="Button('NG2');" type="button" value="NG2" name="Button1" runat="server"><INPUT id="BNG1" style="Z-INDEX: 105; POSITION: absolute; WIDTH: 80px; HEIGHT: 28px; TOP: 712px; LEFT: 512px"
				onclick="Button('NG1');" type="button" value="NG1" name="Button1" runat="server"><INPUT id="BOK" style="Z-INDEX: 106; POSITION: absolute; WIDTH: 80px; HEIGHT: 28px; TOP: 712px; LEFT: 600px"
				onclick="Button('OK');" type="button" value="OK" name="Button2" runat="server">
			</FONT></form>
	</body>
</HTML>
