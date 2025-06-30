<%@ Page Language="vb" AutoEventWireup="false" Codebehind="HR_AwaySheet_01.aspx.vb" Inherits="SPD.HR_AwaySheet_01" aspCompat="True"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>外出申請書</title>
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
				<asp:button id="BADays" style="Z-INDEX: 144; POSITION: absolute; TOP: 216px; LEFT: 480px" runat="server"
					Text="→" Height="24px" Width="30px" CausesValidation="False"></asp:button>
				<asp:button id="BCardTime" style="Z-INDEX: 147; POSITION: absolute; TOP: 150px; LEFT: 596px"
					runat="server" CausesValidation="False" Width="85px" Height="24px" Text="刷卡記錄" BackColor="#FFFFC0"></asp:button><asp:textbox id="DBDay" style="Z-INDEX: 146; POSITION: absolute; TEXT-ALIGN: right; TOP: 186px; LEFT: 536px"
					runat="server" Height="20px" Width="40px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue" BackColor="Yellow">0</asp:textbox><asp:textbox id="DADay" style="Z-INDEX: 145; POSITION: absolute; TEXT-ALIGN: right; TOP: 218px; LEFT: 536px"
					runat="server" Height="20px" Width="40px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue" BackColor="Yellow">0</asp:textbox><asp:button id="BBDays" style="Z-INDEX: 143; POSITION: absolute; TOP: 184px; LEFT: 480px" runat="server"
					Text="→" Height="24px" Width="30px" CausesValidation="False"></asp:button><asp:button id="BAEndDate" style="Z-INDEX: 142; POSITION: absolute; TOP: 218px; LEFT: 384px"
					runat="server" Text="....." Height="20px" Width="20px" CausesValidation="False"></asp:button><asp:button id="BAStartDate" style="Z-INDEX: 141; POSITION: absolute; TOP: 218px; LEFT: 192px"
					runat="server" Text="....." Height="20px" Width="20px" CausesValidation="False"></asp:button><asp:textbox id="DAEndDate" style="Z-INDEX: 140; POSITION: absolute; TOP: 218px; LEFT: 312px"
					runat="server" Height="20px" Width="72px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue" BackColor="Yellow">2008/10/10</asp:textbox><asp:textbox id="DAStartDate" style="Z-INDEX: 139; POSITION: absolute; TOP: 218px; LEFT: 120px"
					runat="server" Height="20px" Width="72px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue" BackColor="Yellow">2008/10/10</asp:textbox><asp:button id="BBEndDate" style="Z-INDEX: 138; POSITION: absolute; TOP: 186px; LEFT: 384px"
					runat="server" Text="....." Height="20px" Width="20px" CausesValidation="False"></asp:button><asp:textbox id="DBEndDate" style="Z-INDEX: 137; POSITION: absolute; TOP: 186px; LEFT: 312px"
					runat="server" Height="20px" Width="72px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue" BackColor="Yellow">2008/10/10</asp:textbox><asp:button id="BBStartDate" style="Z-INDEX: 136; POSITION: absolute; TOP: 186px; LEFT: 192px"
					runat="server" Text="....." Height="20px" Width="20px" CausesValidation="False"></asp:button><asp:textbox id="DBStartDate" style="Z-INDEX: 135; POSITION: absolute; TOP: 186px; LEFT: 120px"
					runat="server" Height="20px" Width="72px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue" BackColor="Yellow">2008/10/10</asp:textbox><asp:dropdownlist id="DPlace" style="Z-INDEX: 134; POSITION: absolute; TOP: 152px; LEFT: 120px" runat="server"
					Height="20px" Width="144px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="中壢第１工廠">中壢第１工廠</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DName" style="Z-INDEX: 133; POSITION: absolute; TOP: 88px; LEFT: 456px" runat="server"
					Height="20px" Width="144px" ForeColor="Blue" BackColor="Yellow" AutoPostBack="True"></asp:dropdownlist><asp:textbox id="DAHour" style="Z-INDEX: 132; POSITION: absolute; TEXT-ALIGN: right; TOP: 218px; LEFT: 600px"
					runat="server" Height="20px" Width="40px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue" BackColor="Yellow">0</asp:textbox><asp:dropdownlist id="DAEndH" style="Z-INDEX: 131; POSITION: absolute; TOP: 218px; LEFT: 408px" runat="server"
					Height="20px" Width="46px" ForeColor="Blue" BackColor="Yellow"></asp:dropdownlist><asp:dropdownlist id="DAStartH" style="Z-INDEX: 130; POSITION: absolute; TOP: 218px; LEFT: 216px"
					runat="server" Height="20px" Width="46px" ForeColor="Blue" BackColor="Yellow"></asp:dropdownlist><asp:dropdownlist id="DBEndH" style="Z-INDEX: 128; POSITION: absolute; TOP: 186px; LEFT: 408px" runat="server"
					Height="20px" Width="46px" ForeColor="Blue" BackColor="Yellow"></asp:dropdownlist><asp:dropdownlist id="DBStartH" style="Z-INDEX: 127; POSITION: absolute; TOP: 186px; LEFT: 216px"
					runat="server" Height="20px" Width="46px" ForeColor="Blue" BackColor="Yellow"></asp:dropdownlist><asp:textbox id="DJobAgent" style="Z-INDEX: 126; POSITION: absolute; TOP: 152px; LEFT: 456px"
					runat="server" Height="20px" Width="136px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" MaxLength="20">DJobAgent</asp:textbox><asp:textbox id="DOtherPlace" style="Z-INDEX: 125; POSITION: absolute; TOP: 152px; LEFT: 264px"
					runat="server" Height="20px" Width="81px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" MaxLength="20">DOtherPlace</asp:textbox><asp:textbox id="DDepoCode" style="Z-INDEX: 124; POSITION: absolute; TOP: 120px; LEFT: 496px"
					runat="server" Height="20px" Width="25px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue" BackColor="Yellow">01</asp:textbox><asp:textbox id="DDepoName" style="Z-INDEX: 123; POSITION: absolute; TOP: 120px; LEFT: 456px"
					runat="server" Height="20px" Width="40px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue" BackColor="Yellow">台北</asp:textbox><asp:textbox id="DSalaryYM" style="Z-INDEX: 122; POSITION: absolute; TOP: 88px; LEFT: 264px"
					runat="server" Height="20px" Width="81px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue" BackColor="Yellow">2008/10</asp:textbox><asp:label id="DHistoryLabel" style="Z-INDEX: 121; POSITION: absolute; TOP: 600px; LEFT: 8px"
					runat="server" Height="16px" Width="64px" ForeColor="Blue" Font-Names="新細明體" Font-Size="11pt">核定履歷</asp:label><asp:textbox id="DReasonDesc" style="Z-INDEX: 120; POSITION: absolute; TOP: 504px; LEFT: 168px"
					runat="server" Height="59px" Width="424px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" TextMode="MultiLine" Visible="False">DReasonDesc</asp:textbox><asp:textbox id="DReason" style="Z-INDEX: 119; POSITION: absolute; TOP: 472px; LEFT: 240px" runat="server"
					Height="20px" Width="352px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue" BackColor="Yellow" Visible="False">DReason</asp:textbox><asp:dropdownlist id="DReasonCode" style="Z-INDEX: 118; POSITION: absolute; TOP: 472px; LEFT: 168px"
					runat="server" Height="20px" Width="64px" ForeColor="Blue" BackColor="Yellow" AutoPostBack="True" Visible="False">
					<asp:ListItem Value="01">01</asp:ListItem>
					<asp:ListItem Value="02">02</asp:ListItem>
				</asp:dropdownlist><asp:image id="DDelay" style="Z-INDEX: 117; POSITION: absolute; TOP: 464px; LEFT: 8px" runat="server"
					Height="107px" Width="593px" Visible="False" ImageUrl="Images\Sheet_Delay.jpg"></asp:image><asp:textbox id="DDecideDesc" style="Z-INDEX: 116; POSITION: absolute; TOP: 392px; LEFT: 56px"
					runat="server" Height="56px" Width="536px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" TextMode="MultiLine">DDecideDesc</asp:textbox><asp:image id="DDescSheet" style="Z-INDEX: 115; POSITION: absolute; TOP: 384px; LEFT: 8px"
					runat="server" Height="75px" Width="593px" ImageUrl="Images\MapSheet_004.jpg"></asp:image><asp:textbox id="DJobCode" style="Z-INDEX: 114; POSITION: absolute; TOP: 120px; LEFT: 264px"
					runat="server" Height="20px" Width="81px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue" BackColor="Yellow">DJobCode</asp:textbox><asp:datagrid id="DataGrid9" style="Z-INDEX: 113; POSITION: absolute; TOP: 616px; LEFT: 8px" runat="server"
					Height="100px" Width="780px" BorderStyle="None" BackColor="White" Font-Size="9pt" AutoGenerateColumns="False" CellPadding="4" BorderWidth="1px" BorderColor="#CC9966">
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
				</asp:datagrid><asp:textbox id="DBHour" style="Z-INDEX: 112; POSITION: absolute; TEXT-ALIGN: right; TOP: 186px; LEFT: 600px"
					runat="server" Height="20px" Width="40px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue" BackColor="Yellow">0</asp:textbox><asp:textbox id="DDivisionCode" style="Z-INDEX: 111; POSITION: absolute; TOP: 120px; LEFT: 616px"
					runat="server" Height="20px" Width="65px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue" BackColor="Yellow">DDivisionCode</asp:textbox><asp:textbox id="DDivision" style="Z-INDEX: 110; POSITION: absolute; TOP: 120px; LEFT: 528px"
					runat="server" Height="20px" Width="88px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue" BackColor="Yellow">DDivision</asp:textbox><asp:textbox id="DJobTitle" style="Z-INDEX: 109; POSITION: absolute; TOP: 120px; LEFT: 120px"
					runat="server" Height="20px" Width="144px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue" BackColor="Yellow">DJobTitle</asp:textbox><asp:textbox id="DEmpID" style="Z-INDEX: 108; POSITION: absolute; TOP: 88px; LEFT: 600px" runat="server"
					Height="20px" Width="81px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue" BackColor="Yellow">DEmpID</asp:textbox><asp:textbox id="DDate" style="Z-INDEX: 107; POSITION: absolute; TOP: 88px; LEFT: 120px" runat="server"
					Height="20px" Width="144px" BorderStyle="Groove" ReadOnly="True" ForeColor="Blue" BackColor="Yellow">DDate</asp:textbox><asp:textbox id="DFReason" style="Z-INDEX: 102; POSITION: absolute; TOP: 250px; LEFT: 120px"
					runat="server" Height="122px" Width="560px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" MaxLength="240" TextMode="MultiLine">DFReason</asp:textbox><asp:image id="DAwaySheet1" style="Z-INDEX: 100; POSITION: absolute; TOP: 8px; LEFT: 8px" runat="server"
					Height="372px" Width="680px" ImageUrl="Images\HR_AwaySheet_01.jpg"></asp:image><asp:textbox id="DNo" style="Z-INDEX: 101; POSITION: absolute; TOP: 9px; LEFT: 8px" runat="server"
					Height="16px" Width="216px" BorderStyle="None" ReadOnly="True" ForeColor="Black" BackColor="White" Font-Names="Times New Roman"> 123</asp:textbox></FONT><INPUT id="BSAVE" style="Z-INDEX: 103; POSITION: absolute; WIDTH: 80px; HEIGHT: 28px; TOP: 576px; LEFT: 336px"
				onclick="Button('SAVE');" type="button" value="儲存" name="Button1" runat="server"><INPUT id="BNG2" style="Z-INDEX: 104; POSITION: absolute; WIDTH: 80px; HEIGHT: 28px; TOP: 576px; LEFT: 424px"
				onclick="Button('NG2');" type="button" value="NG2" name="Button1" runat="server"><INPUT id="BNG1" style="Z-INDEX: 105; POSITION: absolute; WIDTH: 80px; HEIGHT: 28px; TOP: 576px; LEFT: 512px"
				onclick="Button('NG1');" type="button" value="NG1" name="Button1" runat="server"><INPUT id="BOK" style="Z-INDEX: 106; POSITION: absolute; WIDTH: 80px; HEIGHT: 28px; TOP: 576px; LEFT: 600px"
				onclick="Button('OK');" type="button" value="OK" name="Button2" runat="server"></form>
		</FONT></FORM>
	</body>
</HTML>
