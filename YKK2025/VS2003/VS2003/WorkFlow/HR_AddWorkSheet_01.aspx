<%@ Page Language="vb" AutoEventWireup="false" Codebehind="HR_AddWorkSheet_01.aspx.vb" Inherits="SPD.HR_AddWorkSheet_01" aspCompat="True"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>補工申請書</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<script language="javascript" type="text/javascript">
			function CalendarPicker(strField1, strField2, strField3) {
				window.open('VacationDatePicker.aspx?field1=' + strField1 + '&field2=' + strField2 + '&field3=' + strField3,'CalendarPopup','width=250,height=190,resizable=yes');
			}

			function ShowCardTime()
			{
				window.open('HR_CardTimeList.aspx?pWorkDate=' + document.Form1.DWorkDate.value + '&pDepoID=' + document.Form1.DDepoCode.value + '&pEmpID=' + document.Form1.DEmpID.value, 'CardTimeSheet','width=290,height=150,resizable=yes');
			}

			function ShowCardTimeA()
			{
				window.open('HR_CardTimeList.aspx?pWorkDate=' + document.Form1.DAWorkDate.value + '&pDepoID=' + document.Form1.DDepoCode.value + '&pEmpID=' + document.Form1.DEmpID.value, 'CardTimeSheet','width=290,height=150,resizable=yes');
			}

			function ShowOverTime()
			{
				window.open('http://10.245.1.6/WorkFlowSub/HR_OverTimeList.aspx?pMonth=' + document.Form1.DSalaryYM.value + '&pDepoID=' + document.Form1.DDepoCode.value + '&pEmpID=' + document.Form1.DEmpID.value, 'OverTimeSheet','width=320,height=400,resizable=yes');
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
				<asp:textbox id="DWM" style="Z-INDEX: 119; POSITION: absolute; TEXT-ALIGN: right; TOP: 186px; LEFT: 600px"
					runat="server" ReadOnly="True" BorderStyle="Groove" Width="60px" Height="20px" BackColor="Yellow"
					ForeColor="Blue">DWM</asp:textbox><asp:dropdownlist id="DAddWorkType" style="Z-INDEX: 145; POSITION: absolute; TOP: 152px; LEFT: 456px"
					runat="server" Width="120px" Height="20px" BackColor="Yellow" ForeColor="Blue"></asp:dropdownlist><asp:button id="BACardTime" style="Z-INDEX: 144; POSITION: absolute; TOP: 218px; LEFT: 248px"
					runat="server" Width="94px" Height="24px" BackColor="#FFFFC0" CausesValidation="False" Text="刷卡記錄"></asp:button><asp:button id="BAWorkDate" style="Z-INDEX: 143; POSITION: absolute; TOP: 218px; LEFT: 224px"
					runat="server" Width="20px" Height="20px" CausesValidation="False" Text="....."></asp:button><asp:textbox id="DAWorkDate" style="Z-INDEX: 142; POSITION: absolute; TOP: 218px; LEFT: 120px"
					runat="server" ReadOnly="True" BorderStyle="Groove" Width="104px" Height="20px" BackColor="Yellow" ForeColor="Blue">2008/10/10</asp:textbox><asp:button id="BOverTime" style="Z-INDEX: 141; POSITION: absolute; TOP: 152px; LEFT: 584px"
					runat="server" Width="96px" Height="24px" BackColor="#FFFFC0" CausesValidation="False" Text="加班記錄"></asp:button><asp:dropdownlist id="DName" style="Z-INDEX: 139; POSITION: absolute; TOP: 88px; LEFT: 456px" runat="server"
					Width="160px" Height="20px" BackColor="Yellow" ForeColor="Blue" AutoPostBack="True"></asp:dropdownlist><asp:button id="BCardTime" style="Z-INDEX: 138; POSITION: absolute; TOP: 152px; LEFT: 248px"
					runat="server" Width="94px" Height="24px" BackColor="#FFFFC0" CausesValidation="False" Text="刷卡記錄"></asp:button><asp:textbox id="DDepoCode" style="Z-INDEX: 137; POSITION: absolute; TOP: 120px; LEFT: 496px"
					runat="server" ReadOnly="True" BorderStyle="Groove" Width="25px" Height="20px" BackColor="Yellow" ForeColor="Blue">01</asp:textbox><asp:textbox id="DDepoName" style="Z-INDEX: 136; POSITION: absolute; TOP: 120px; LEFT: 456px"
					runat="server" ReadOnly="True" BorderStyle="Groove" Width="40px" Height="20px" BackColor="Yellow" ForeColor="Blue">台北</asp:textbox><asp:textbox id="DSalaryYM" style="Z-INDEX: 135; POSITION: absolute; TOP: 88px; LEFT: 224px"
					runat="server" ReadOnly="True" BorderStyle="Groove" Width="64px" Height="20px" BackColor="Yellow" ForeColor="Blue">2008/10</asp:textbox><asp:label id="DHistoryLabel" style="Z-INDEX: 134; POSITION: absolute; TOP: 632px; LEFT: 8px"
					runat="server" Width="64px" Height="16px" ForeColor="Blue" Font-Names="新細明體" Font-Size="11pt">核定履歷</asp:label><asp:textbox id="DReasonDesc" style="Z-INDEX: 133; POSITION: absolute; TOP: 536px; LEFT: 168px"
					runat="server" BorderStyle="Groove" Width="424px" Height="59px" BackColor="Yellow" ForeColor="Blue" Visible="False" TextMode="MultiLine">DReasonDesc</asp:textbox><asp:textbox id="DReason" style="Z-INDEX: 132; POSITION: absolute; TOP: 504px; LEFT: 240px" runat="server"
					ReadOnly="True" BorderStyle="Groove" Width="352px" Height="20px" BackColor="Yellow" ForeColor="Blue" Visible="False">DReason</asp:textbox><asp:dropdownlist id="DReasonCode" style="Z-INDEX: 131; POSITION: absolute; TOP: 504px; LEFT: 168px"
					runat="server" Width="64px" Height="20px" BackColor="Yellow" ForeColor="Blue" AutoPostBack="True" Visible="False">
					<asp:ListItem Value="01">01</asp:ListItem>
					<asp:ListItem Value="02">02</asp:ListItem>
				</asp:dropdownlist><asp:image id="DDelay" style="Z-INDEX: 130; POSITION: absolute; TOP: 496px; LEFT: 8px" runat="server"
					Width="593px" Height="107px" Visible="False" ImageUrl="Images\Sheet_Delay.jpg"></asp:image><asp:textbox id="DDecideDesc" style="Z-INDEX: 129; POSITION: absolute; TOP: 424px; LEFT: 56px"
					runat="server" BorderStyle="Groove" Width="536px" Height="56px" BackColor="Yellow" ForeColor="Blue" TextMode="MultiLine">DDecideDesc</asp:textbox><asp:image id="DDescSheet" style="Z-INDEX: 128; POSITION: absolute; TOP: 416px; LEFT: 8px"
					runat="server" Width="593px" Height="75px" ImageUrl="Images\MapSheet_004.jpg"></asp:image><asp:textbox id="DJobCode" style="Z-INDEX: 127; POSITION: absolute; TOP: 120px; LEFT: 224px"
					runat="server" ReadOnly="True" BorderStyle="Groove" Width="64px" Height="20px" BackColor="Yellow" ForeColor="Blue" Visible="False">DJobCode</asp:textbox><asp:datagrid id="DataGrid9" style="Z-INDEX: 126; POSITION: absolute; TOP: 648px; LEFT: 8px" runat="server"
					BorderStyle="None" Width="780px" Height="100px" BackColor="White" Font-Size="9pt" BorderColor="#CC9966" BorderWidth="1px" CellPadding="4" AutoGenerateColumns="False">
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
				</asp:datagrid><asp:textbox id="DAM" style="Z-INDEX: 125; POSITION: absolute; TEXT-ALIGN: right; TOP: 251px; LEFT: 600px"
					runat="server" ReadOnly="True" BorderStyle="Groove" Width="60px" Height="20px" BackColor="Yellow" ForeColor="Blue">DAM</asp:textbox><asp:textbox id="DAH" style="Z-INDEX: 124; POSITION: absolute; TEXT-ALIGN: right; TOP: 251px; LEFT: 520px"
					runat="server" ReadOnly="True" BorderStyle="Groove" Width="60px" Height="20px" BackColor="Yellow" ForeColor="Blue">DAH</asp:textbox><asp:dropdownlist id="DAEndM" style="Z-INDEX: 123; POSITION: absolute; TOP: 251px; LEFT: 392px" runat="server"
					Width="60px" Height="20px" BackColor="Yellow" ForeColor="Blue" AutoPostBack="True">
					<asp:ListItem Value="0" Selected="True">00</asp:ListItem>
					<asp:ListItem Value="1">01</asp:ListItem>
					<asp:ListItem Value="2">02</asp:ListItem>
					<asp:ListItem Value="3">03</asp:ListItem>
					<asp:ListItem Value="4">04</asp:ListItem>
					<asp:ListItem Value="5">05</asp:ListItem>
					<asp:ListItem Value="6">06</asp:ListItem>
					<asp:ListItem Value="7">07</asp:ListItem>
					<asp:ListItem Value="8">08</asp:ListItem>
					<asp:ListItem Value="9">09</asp:ListItem>
					<asp:ListItem Value="10">10</asp:ListItem>
					<asp:ListItem Value="11">11</asp:ListItem>
					<asp:ListItem Value="12">12</asp:ListItem>
					<asp:ListItem Value="13">13</asp:ListItem>
					<asp:ListItem Value="14">14</asp:ListItem>
					<asp:ListItem Value="15">15</asp:ListItem>
					<asp:ListItem Value="16">16</asp:ListItem>
					<asp:ListItem Value="17">17</asp:ListItem>
					<asp:ListItem Value="18">18</asp:ListItem>
					<asp:ListItem Value="19">19</asp:ListItem>
					<asp:ListItem Value="20">20</asp:ListItem>
					<asp:ListItem Value="21">21</asp:ListItem>
					<asp:ListItem Value="22">22</asp:ListItem>
					<asp:ListItem Value="23">23</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DAEndH" style="Z-INDEX: 122; POSITION: absolute; TOP: 251px; LEFT: 310px" runat="server"
					Width="60px" Height="20px" BackColor="Yellow" ForeColor="Blue" AutoPostBack="True">
					<asp:ListItem Value="0" Selected="True">00</asp:ListItem>
					<asp:ListItem Value="1">01</asp:ListItem>
					<asp:ListItem Value="2">02</asp:ListItem>
					<asp:ListItem Value="3">03</asp:ListItem>
					<asp:ListItem Value="4">04</asp:ListItem>
					<asp:ListItem Value="5">05</asp:ListItem>
					<asp:ListItem Value="6">06</asp:ListItem>
					<asp:ListItem Value="7">07</asp:ListItem>
					<asp:ListItem Value="8">08</asp:ListItem>
					<asp:ListItem Value="9">09</asp:ListItem>
					<asp:ListItem Value="10">10</asp:ListItem>
					<asp:ListItem Value="11">11</asp:ListItem>
					<asp:ListItem Value="12">12</asp:ListItem>
					<asp:ListItem Value="13">13</asp:ListItem>
					<asp:ListItem Value="14">14</asp:ListItem>
					<asp:ListItem Value="15">15</asp:ListItem>
					<asp:ListItem Value="16">16</asp:ListItem>
					<asp:ListItem Value="17">17</asp:ListItem>
					<asp:ListItem Value="18">18</asp:ListItem>
					<asp:ListItem Value="19">19</asp:ListItem>
					<asp:ListItem Value="20">20</asp:ListItem>
					<asp:ListItem Value="21">21</asp:ListItem>
					<asp:ListItem Value="22">22</asp:ListItem>
					<asp:ListItem Value="23">23</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DAStartM" style="Z-INDEX: 121; POSITION: absolute; TOP: 251px; LEFT: 202px"
					runat="server" Width="60px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="0" Selected="True">00</asp:ListItem>
					<asp:ListItem Value="10">10</asp:ListItem>
					<asp:ListItem Value="20">20</asp:ListItem>
					<asp:ListItem Value="30">30</asp:ListItem>
					<asp:ListItem Value="40">40</asp:ListItem>
					<asp:ListItem Value="50">50</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DAStartH" style="Z-INDEX: 120; POSITION: absolute; TOP: 251px; LEFT: 120px"
					runat="server" Width="60px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="0" Selected="True">00</asp:ListItem>
					<asp:ListItem Value="1">01</asp:ListItem>
					<asp:ListItem Value="2">02</asp:ListItem>
					<asp:ListItem Value="3">03</asp:ListItem>
					<asp:ListItem Value="4">04</asp:ListItem>
					<asp:ListItem Value="5">05</asp:ListItem>
					<asp:ListItem Value="6">06</asp:ListItem>
					<asp:ListItem Value="7">07</asp:ListItem>
					<asp:ListItem Value="8">08</asp:ListItem>
					<asp:ListItem Value="9">09</asp:ListItem>
					<asp:ListItem Value="10">10</asp:ListItem>
					<asp:ListItem Value="11">11</asp:ListItem>
					<asp:ListItem Value="12">12</asp:ListItem>
					<asp:ListItem Value="13">13</asp:ListItem>
					<asp:ListItem Value="14">14</asp:ListItem>
					<asp:ListItem Value="15">15</asp:ListItem>
					<asp:ListItem Value="16">16</asp:ListItem>
					<asp:ListItem Value="17">17</asp:ListItem>
					<asp:ListItem Value="18">18</asp:ListItem>
					<asp:ListItem Value="19">19</asp:ListItem>
					<asp:ListItem Value="20">20</asp:ListItem>
					<asp:ListItem Value="21">21</asp:ListItem>
					<asp:ListItem Value="22">22</asp:ListItem>
					<asp:ListItem Value="23">23</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DWH" style="Z-INDEX: 118; POSITION: absolute; TEXT-ALIGN: right; TOP: 186px; LEFT: 520px"
					runat="server" ReadOnly="True" BorderStyle="Groove" Width="60px" Height="20px" BackColor="Yellow" ForeColor="Blue">DWH</asp:textbox><asp:dropdownlist id="DWEndM" style="Z-INDEX: 117; POSITION: absolute; TOP: 186px; LEFT: 392px" runat="server"
					Width="60px" Height="20px" BackColor="Yellow" ForeColor="Blue" AutoPostBack="True">
					<asp:ListItem Value="0" Selected="True">00</asp:ListItem>
					<asp:ListItem Value="1">01</asp:ListItem>
					<asp:ListItem Value="2">02</asp:ListItem>
					<asp:ListItem Value="3">03</asp:ListItem>
					<asp:ListItem Value="4">04</asp:ListItem>
					<asp:ListItem Value="5">05</asp:ListItem>
					<asp:ListItem Value="6">06</asp:ListItem>
					<asp:ListItem Value="7">07</asp:ListItem>
					<asp:ListItem Value="8">08</asp:ListItem>
					<asp:ListItem Value="9">09</asp:ListItem>
					<asp:ListItem Value="10">10</asp:ListItem>
					<asp:ListItem Value="11">11</asp:ListItem>
					<asp:ListItem Value="12">12</asp:ListItem>
					<asp:ListItem Value="13">13</asp:ListItem>
					<asp:ListItem Value="14">14</asp:ListItem>
					<asp:ListItem Value="15">15</asp:ListItem>
					<asp:ListItem Value="16">16</asp:ListItem>
					<asp:ListItem Value="17">17</asp:ListItem>
					<asp:ListItem Value="18">18</asp:ListItem>
					<asp:ListItem Value="19">19</asp:ListItem>
					<asp:ListItem Value="20">20</asp:ListItem>
					<asp:ListItem Value="21">21</asp:ListItem>
					<asp:ListItem Value="22">22</asp:ListItem>
					<asp:ListItem Value="23">23</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DWEndH" style="Z-INDEX: 116; POSITION: absolute; TOP: 186px; LEFT: 310px" runat="server"
					Width="60px" Height="20px" BackColor="Yellow" ForeColor="Blue" AutoPostBack="True">
					<asp:ListItem Value="0" Selected="True">00</asp:ListItem>
					<asp:ListItem Value="1">01</asp:ListItem>
					<asp:ListItem Value="2">02</asp:ListItem>
					<asp:ListItem Value="3">03</asp:ListItem>
					<asp:ListItem Value="4">04</asp:ListItem>
					<asp:ListItem Value="5">05</asp:ListItem>
					<asp:ListItem Value="6">06</asp:ListItem>
					<asp:ListItem Value="7">07</asp:ListItem>
					<asp:ListItem Value="8">08</asp:ListItem>
					<asp:ListItem Value="9">09</asp:ListItem>
					<asp:ListItem Value="10">10</asp:ListItem>
					<asp:ListItem Value="11">11</asp:ListItem>
					<asp:ListItem Value="12">12</asp:ListItem>
					<asp:ListItem Value="13">13</asp:ListItem>
					<asp:ListItem Value="14">14</asp:ListItem>
					<asp:ListItem Value="15">15</asp:ListItem>
					<asp:ListItem Value="16">16</asp:ListItem>
					<asp:ListItem Value="17">17</asp:ListItem>
					<asp:ListItem Value="18">18</asp:ListItem>
					<asp:ListItem Value="19">19</asp:ListItem>
					<asp:ListItem Value="20">20</asp:ListItem>
					<asp:ListItem Value="21">21</asp:ListItem>
					<asp:ListItem Value="22">22</asp:ListItem>
					<asp:ListItem Value="23">23</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DWStartM" style="Z-INDEX: 115; POSITION: absolute; TOP: 186px; LEFT: 202px"
					runat="server" Width="60px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="0" Selected="True">00</asp:ListItem>
					<asp:ListItem Value="10">10</asp:ListItem>
					<asp:ListItem Value="20">20</asp:ListItem>
					<asp:ListItem Value="30">30</asp:ListItem>
					<asp:ListItem Value="40">40</asp:ListItem>
					<asp:ListItem Value="50">50</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DWStartH" style="Z-INDEX: 114; POSITION: absolute; TOP: 186px; LEFT: 120px"
					runat="server" Width="60px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="0" Selected="True">00</asp:ListItem>
					<asp:ListItem Value="1">01</asp:ListItem>
					<asp:ListItem Value="2">02</asp:ListItem>
					<asp:ListItem Value="3">03</asp:ListItem>
					<asp:ListItem Value="4">04</asp:ListItem>
					<asp:ListItem Value="5">05</asp:ListItem>
					<asp:ListItem Value="6">06</asp:ListItem>
					<asp:ListItem Value="7">07</asp:ListItem>
					<asp:ListItem Value="8">08</asp:ListItem>
					<asp:ListItem Value="9">09</asp:ListItem>
					<asp:ListItem Value="10">10</asp:ListItem>
					<asp:ListItem Value="11">11</asp:ListItem>
					<asp:ListItem Value="12">12</asp:ListItem>
					<asp:ListItem Value="13">13</asp:ListItem>
					<asp:ListItem Value="14">14</asp:ListItem>
					<asp:ListItem Value="15">15</asp:ListItem>
					<asp:ListItem Value="16">16</asp:ListItem>
					<asp:ListItem Value="17">17</asp:ListItem>
					<asp:ListItem Value="18">18</asp:ListItem>
					<asp:ListItem Value="19">19</asp:ListItem>
					<asp:ListItem Value="20">20</asp:ListItem>
					<asp:ListItem Value="21">21</asp:ListItem>
					<asp:ListItem Value="22">22</asp:ListItem>
					<asp:ListItem Value="23">23</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DDivisionCode" style="Z-INDEX: 113; POSITION: absolute; TOP: 120px; LEFT: 616px"
					runat="server" ReadOnly="True" BorderStyle="Groove" Width="67px" Height="20px" BackColor="Yellow" ForeColor="Blue">DDivisionCode</asp:textbox><asp:textbox id="DDivision" style="Z-INDEX: 112; POSITION: absolute; TOP: 120px; LEFT: 528px"
					runat="server" ReadOnly="True" BorderStyle="Groove" Width="88px" Height="20px" BackColor="Yellow" ForeColor="Blue">DDivision</asp:textbox><asp:textbox id="DJobTitle" style="Z-INDEX: 111; POSITION: absolute; TOP: 120px; LEFT: 120px"
					runat="server" ReadOnly="True" BorderStyle="Groove" Width="104px" Height="20px" BackColor="Yellow" ForeColor="Blue">DJobTitle</asp:textbox><asp:textbox id="DEmpID" style="Z-INDEX: 110; POSITION: absolute; TOP: 88px; LEFT: 616px" runat="server"
					ReadOnly="True" BorderStyle="Groove" Width="67px" Height="20px" BackColor="Yellow" ForeColor="Blue">DEmpID</asp:textbox><asp:textbox id="DDate" style="Z-INDEX: 109; POSITION: absolute; TOP: 88px; LEFT: 120px" runat="server"
					ReadOnly="True" BorderStyle="Groove" Width="104px" Height="20px" BackColor="Yellow" ForeColor="Blue">DDate</asp:textbox><asp:textbox id="DFReason" style="Z-INDEX: 102; POSITION: absolute; TOP: 286px; LEFT: 120px"
					runat="server" BorderStyle="Groove" Width="560px" Height="118px" BackColor="Yellow" ForeColor="Blue" TextMode="MultiLine" MaxLength="240">DFReason</asp:textbox><asp:image id="DAddWorkSheet1" style="Z-INDEX: 100; POSITION: absolute; TOP: 8px; LEFT: 8px"
					runat="server" Width="680px" Height="406px" ImageUrl="Images\HR_AddWorkSheet_02.jpg"></asp:image><asp:button id="BWorkDate" style="Z-INDEX: 104; POSITION: absolute; TOP: 152px; LEFT: 224px"
					runat="server" Width="20px" Height="20px" CausesValidation="False" Text="....."></asp:button><asp:textbox id="DWorkDate" style="Z-INDEX: 103; POSITION: absolute; TOP: 152px; LEFT: 120px"
					runat="server" ReadOnly="True" BorderStyle="Groove" Width="104px" Height="20px" BackColor="Yellow" ForeColor="Blue">2008/10/10</asp:textbox><asp:textbox id="DNo" style="Z-INDEX: 101; POSITION: absolute; TOP: 9px; LEFT: 8px" runat="server"
					ReadOnly="True" BorderStyle="None" Width="216px" Height="16px" BackColor="White" ForeColor="Black" Font-Names="Times New Roman"> 123</asp:textbox></FONT><INPUT id="BSAVE" style="Z-INDEX: 105; POSITION: absolute; WIDTH: 80px; HEIGHT: 28px; TOP: 608px; LEFT: 336px"
				onclick="Button('SAVE');" type="button" value="儲存" name="Button1" runat="server"><INPUT id="BNG2" style="Z-INDEX: 106; POSITION: absolute; WIDTH: 80px; HEIGHT: 28px; TOP: 608px; LEFT: 424px"
				onclick="Button('NG2');" type="button" value="NG2" name="Button1" runat="server"><INPUT id="BNG1" style="Z-INDEX: 107; POSITION: absolute; WIDTH: 80px; HEIGHT: 28px; TOP: 608px; LEFT: 512px"
				onclick="Button('NG1');" type="button" value="NG1" name="Button1" runat="server"><INPUT id="BOK" style="Z-INDEX: 108; POSITION: absolute; WIDTH: 80px; HEIGHT: 28px; TOP: 608px; LEFT: 600px"
				onclick="Button('OK');" type="button" value="OK" name="Button2" runat="server"></form>
		</FONT></FORM>
	</body>
</HTML>
