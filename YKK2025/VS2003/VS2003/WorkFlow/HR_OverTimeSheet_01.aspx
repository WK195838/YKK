<%@ Page Language="vb" AutoEventWireup="false" Codebehind="HR_OverTimeSheet_01.aspx.vb" Inherits="SPD.HR_OverTimeSheet_01" aspCompat="True"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>加班申請書</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<script language="javascript" type="text/javascript">
			function CalendarPicker(strField1, strField2, strField3, strField4) {
				window.open('OverTimeDatePicker.aspx?field1=' + strField1 + '&field2=' + strField2 + '&field3=' + strField3 + '&field4=' + strField4,'CalendarPopup','width=250,height=190,resizable=yes');
			}

			function ShowCardTime()
			{
				window.open('HR_CardTimeList.aspx?pWorkDate=' + document.Form1.DOverTimeDate.value + '&pDepoID=' + document.Form1.DDepoCode.value + '&pEmpID=' + document.Form1.DEmpID.value, 'CardTimeSheet','width=290,height=150,resizable=yes');
			}

			function ShowOverTime()
			{
				window.open('http://10.245.1.6/WorkFlowSub/HR_OverTimeList.aspx?pMonth=' + document.Form1.DSalaryYM.value + '&pDepoID=' + document.Form1.DDepoCode.value + '&pEmpID=' + document.Form1.DEmpID.value, 'OverTimeSheet','width=400,height=600,resizable=yes');
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
		<form id="Form1" encType="multipart/form-data" method="post" runat="server">
			<FONT face="新細明體">
				<asp:textbox style="Z-INDEX: 120; POSITION: absolute; TEXT-ALIGN: right; TOP: 250px; LEFT: 604px"
					id="DBM" runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px" BorderStyle="Groove"
					ReadOnly="True">DBM</asp:textbox><asp:button style="Z-INDEX: 153; POSITION: absolute; TOP: 182px; LEFT: 560px" id="BOverTime"
					runat="server" BackColor="#FFFFC0" Height="24px" Width="114px" Text="加班記錄" CausesValidation="False"></asp:button><asp:dropdownlist style="Z-INDEX: 152; POSITION: absolute; TOP: 120px; LEFT: 120px" id="DName" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="20px" Width="188px" AutoPostBack="True"></asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 151; POSITION: absolute; TOP: 152px; LEFT: 560px" id="DCVacation"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="121px">
					<asp:ListItem Value="要" Selected="True">要</asp:ListItem>
					<asp:ListItem Value="不要">不要</asp:ListItem>
				</asp:dropdownlist><asp:button style="Z-INDEX: 149; POSITION: absolute; TOP: 182px; LEFT: 456px" id="BCardTime"
					runat="server" BackColor="#FFFFC0" Height="24px" Width="94px" Text="刷卡記錄" CausesValidation="False"></asp:button><asp:dropdownlist style="Z-INDEX: 148; POSITION: absolute; TOP: 184px; LEFT: 336px" id="DTraffic"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="112px">
					<asp:ListItem Value="要" Selected="True">要</asp:ListItem>
					<asp:ListItem Value="不要">不要</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 147; POSITION: absolute; TOP: 152px; LEFT: 160px" id="DDepoCode"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="25px" BorderStyle="Groove" ReadOnly="True">01</asp:textbox><asp:textbox style="Z-INDEX: 146; POSITION: absolute; TOP: 152px; LEFT: 120px" id="DDepoName"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="40px" BorderStyle="Groove" ReadOnly="True">台北</asp:textbox><asp:textbox style="Z-INDEX: 145; POSITION: absolute; TOP: 86px; LEFT: 374px" id="DSalaryYM"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="74px" BorderStyle="Groove" ReadOnly="True">2008/10</asp:textbox><asp:label style="Z-INDEX: 144; POSITION: absolute; TOP: 920px; LEFT: 8px" id="DHistoryLabel"
					runat="server" ForeColor="Blue" Height="16px" Width="64px" Font-Size="11pt" Font-Names="新細明體">核定履歷</asp:label><asp:textbox style="Z-INDEX: 143; POSITION: absolute; TOP: 86px; LEFT: 264px" id="DDateType"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="108px" BorderStyle="Groove" ReadOnly="True">DDateType</asp:textbox><asp:textbox style="Z-INDEX: 142; POSITION: absolute; TOP: 816px; LEFT: 168px" id="DReasonDesc"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="59px" Width="424px" BorderStyle="Groove" TextMode="MultiLine" Visible="False">DReasonDesc</asp:textbox><asp:textbox style="Z-INDEX: 141; POSITION: absolute; TOP: 784px; LEFT: 240px" id="DReason" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="20px" Width="352px" BorderStyle="Groove" ReadOnly="True" Visible="False">DReason</asp:textbox><asp:dropdownlist style="Z-INDEX: 140; POSITION: absolute; TOP: 784px; LEFT: 168px" id="DReasonCode"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="64px" AutoPostBack="True" Visible="False">
					<asp:ListItem Value="01">01</asp:ListItem>
					<asp:ListItem Value="02">02</asp:ListItem>
				</asp:dropdownlist><asp:image style="Z-INDEX: 139; POSITION: absolute; TOP: 776px; LEFT: 8px" id="DDelay" runat="server"
					Height="107px" Width="593px" Visible="False" ImageUrl="Images\Sheet_Delay.jpg"></asp:image><asp:textbox style="Z-INDEX: 138; POSITION: absolute; TOP: 712px; LEFT: 56px" id="DDecideDesc"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="56px" Width="536px" BorderStyle="Groove" TextMode="MultiLine">DDecideDesc</asp:textbox><asp:image style="Z-INDEX: 137; POSITION: absolute; TOP: 704px; LEFT: 8px" id="DDescSheet"
					runat="server" Height="75px" Width="593px" ImageUrl="Images\MapSheet_004.jpg"></asp:image><asp:textbox style="Z-INDEX: 136; POSITION: absolute; TOP: 120px; LEFT: 640px" id="DJobCode"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="40px" BorderStyle="Groove" ReadOnly="True" Visible="False">DJobCode</asp:textbox><asp:textbox style="Z-INDEX: 135; POSITION: absolute; TEXT-ALIGN: right; TOP: 381px; LEFT: 602px"
					id="DFCM" runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px" BorderStyle="Groove">DFCM</asp:textbox><asp:datagrid style="Z-INDEX: 134; POSITION: absolute; TOP: 936px; LEFT: 8px" id="DataGrid9" runat="server"
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
				</asp:datagrid><asp:textbox style="Z-INDEX: 133; POSITION: absolute; TEXT-ALIGN: right; TOP: 381px; LEFT: 518px"
					id="DFCH" runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px" BorderStyle="Groove">DFCH</asp:textbox><asp:textbox style="Z-INDEX: 132; POSITION: absolute; TEXT-ALIGN: right; TOP: 381px; LEFT: 414px"
					id="DFBM" runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px" BorderStyle="Groove">DFBM</asp:textbox><asp:textbox style="Z-INDEX: 131; POSITION: absolute; TEXT-ALIGN: right; TOP: 381px; LEFT: 328px"
					id="DFBH" runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px" BorderStyle="Groove">DFBH</asp:textbox><asp:textbox style="Z-INDEX: 130; POSITION: absolute; TEXT-ALIGN: right; TOP: 381px; LEFT: 224px"
					id="DFAM2" runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px" BorderStyle="Groove">DFAM2</asp:textbox><asp:textbox style="Z-INDEX: 129; POSITION: absolute; TEXT-ALIGN: right; TOP: 381px; LEFT: 162px"
					id="DFAH2" runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="40px" BorderStyle="Groove">DFAH2</asp:textbox><asp:textbox style="Z-INDEX: 128; POSITION: absolute; TEXT-ALIGN: right; TOP: 350px; LEFT: 162px"
					id="DFAH1" runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="40px" BorderStyle="Groove">DFAH1</asp:textbox><asp:textbox style="Z-INDEX: 127; POSITION: absolute; TEXT-ALIGN: right; TOP: 350px; LEFT: 224px"
					id="DFAM1" runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px" BorderStyle="Groove">DFAM1</asp:textbox><asp:textbox style="Z-INDEX: 126; POSITION: absolute; TEXT-ALIGN: right; TOP: 282px; LEFT: 604px"
					id="DAM" runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px" BorderStyle="Groove" ReadOnly="True">DAM</asp:textbox><asp:textbox style="Z-INDEX: 125; POSITION: absolute; TEXT-ALIGN: right; TOP: 282px; LEFT: 520px"
					id="DAH" runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px" BorderStyle="Groove" ReadOnly="True">DAH</asp:textbox><asp:dropdownlist style="Z-INDEX: 124; POSITION: absolute; TOP: 282px; LEFT: 414px" id="DAEndM" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px" AutoPostBack="True">
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
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 123; POSITION: absolute; TOP: 282px; LEFT: 330px" id="DAEndH" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px" AutoPostBack="True">
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
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 122; POSITION: absolute; TOP: 282px; LEFT: 224px" id="DAStartM"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px">
					<asp:ListItem Value="0" Selected="True">00</asp:ListItem>
					<asp:ListItem Value="10">10</asp:ListItem>
					<asp:ListItem Value="20">20</asp:ListItem>
					<asp:ListItem Value="30">30</asp:ListItem>
					<asp:ListItem Value="40">40</asp:ListItem>
					<asp:ListItem Value="50">50</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 121; POSITION: absolute; TOP: 282px; LEFT: 142px" id="DAStartH"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px">
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
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 119; POSITION: absolute; TEXT-ALIGN: right; TOP: 250px; LEFT: 520px"
					id="DBH" runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px" BorderStyle="Groove"
					ReadOnly="True">DBH</asp:textbox><asp:dropdownlist style="Z-INDEX: 118; POSITION: absolute; TOP: 250px; LEFT: 414px" id="DBEndM" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px" AutoPostBack="True">
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
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 117; POSITION: absolute; TOP: 250px; LEFT: 330px" id="DBEndH" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px" AutoPostBack="True">
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
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 116; POSITION: absolute; TOP: 250px; LEFT: 224px" id="DBStartM"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px">
					<asp:ListItem Value="0" Selected="True">00</asp:ListItem>
					<asp:ListItem Value="10">10</asp:ListItem>
					<asp:ListItem Value="20">20</asp:ListItem>
					<asp:ListItem Value="30">30</asp:ListItem>
					<asp:ListItem Value="40">40</asp:ListItem>
					<asp:ListItem Value="50">50</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 115; POSITION: absolute; TOP: 250px; LEFT: 142px" id="DBStartH"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px">
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
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 114; POSITION: absolute; TOP: 184px; LEFT: 120px" id="DFood" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="20px" Width="96px">
					<asp:ListItem Value="要" Selected="True">要</asp:ListItem>
					<asp:ListItem Value="不要">不要</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 113; POSITION: absolute; TOP: 152px; LEFT: 280px" id="DDivisionCode"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="67px" BorderStyle="Groove" ReadOnly="True">DDivisionCode</asp:textbox><asp:textbox style="Z-INDEX: 112; POSITION: absolute; TOP: 152px; LEFT: 192px" id="DDivision"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="88px" BorderStyle="Groove" ReadOnly="True">DDivision</asp:textbox><asp:textbox style="Z-INDEX: 111; POSITION: absolute; TOP: 120px; LEFT: 560px" id="DJobTitle"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="78px" BorderStyle="Groove" ReadOnly="True">DJobTitle</asp:textbox><asp:textbox style="Z-INDEX: 110; POSITION: absolute; TOP: 120px; LEFT: 312px" id="DEmpID" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="20px" Width="136px" BorderStyle="Groove" ReadOnly="True">DEmpID</asp:textbox><asp:textbox style="Z-INDEX: 109; POSITION: absolute; TOP: 86px; LEFT: 560px" id="DDate" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="20px" Width="121px" BorderStyle="Groove" ReadOnly="True">DDate</asp:textbox><asp:textbox style="Z-INDEX: 102; POSITION: absolute; TOP: 416px; LEFT: 120px" id="DFReason"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="118px" Width="560px" BorderStyle="Groove" TextMode="MultiLine" MaxLength="240">DFReason</asp:textbox><asp:image style="Z-INDEX: 100; POSITION: absolute; TOP: 6px; LEFT: 8px" id="DOverTimeSheet1"
					runat="server" Height="695px" Width="680px" ImageUrl="Images\HR_OverTimeSheet_02.jpg"></asp:image><asp:button style="Z-INDEX: 104; POSITION: absolute; TOP: 86px; LEFT: 240px" id="BOverTimeDate"
					runat="server" Height="20px" Width="20px" Text="....." CausesValidation="False"></asp:button><asp:textbox style="Z-INDEX: 103; POSITION: absolute; TOP: 86px; LEFT: 120px" id="DOverTimeDate"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="120px" BorderStyle="Groove" ReadOnly="True">2008/10/10</asp:textbox><asp:textbox style="Z-INDEX: 101; POSITION: absolute; TOP: 9px; LEFT: 8px" id="DNo" runat="server"
					ForeColor="Black" BackColor="White" Height="16px" Width="216px" BorderStyle="None" ReadOnly="True" Font-Names="Times New Roman"> 123</asp:textbox><asp:textbox style="Z-INDEX: 101; POSITION: absolute; TOP: 48px; LEFT: 8px" id="D46Hours" runat="server"
					ForeColor="Black" BackColor="White" Height="28px" Width="472px" BorderStyle="None" ReadOnly="True" Font-Names="Times New Roman" BorderColor="#000000"> 當月平日：已核=???H　未核=???H　計= ???H</asp:textbox></FONT><INPUT style="Z-INDEX: 105; POSITION: absolute; WIDTH: 80px; HEIGHT: 28px; TOP: 896px; LEFT: 336px"
				id="BSAVE" onclick="Button('SAVE');" value="儲存" type="button" name="Button1" runat="server"><INPUT style="Z-INDEX: 106; POSITION: absolute; WIDTH: 80px; HEIGHT: 28px; TOP: 896px; LEFT: 424px"
				id="BNG2" onclick="Button('NG2');" value="NG2" type="button" name="Button1" runat="server"><INPUT style="Z-INDEX: 107; POSITION: absolute; WIDTH: 80px; HEIGHT: 28px; TOP: 896px; LEFT: 512px"
				id="BNG1" onclick="Button('NG1');" value="NG1" type="button" name="Button1" runat="server"><INPUT style="Z-INDEX: 108; POSITION: absolute; WIDTH: 80px; HEIGHT: 28px; TOP: 896px; LEFT: 600px"
				id="BOK" onclick="Button('OK');" value="OK" type="button" name="Button2" runat="server">
			<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
			<!-- ++  隱藏欄位                                                                      ++ -->
			<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ --><asp:textbox style="Z-INDEX: 128; POSITION: absolute; TEXT-ALIGN: right; TOP: 552px; LEFT: 162px"
				id="DFPRAH1" runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="40px" BorderStyle="Groove">0</asp:textbox><asp:textbox style="Z-INDEX: 127; POSITION: absolute; TEXT-ALIGN: right; TOP: 552px; LEFT: 224px"
				id="DFPRAM1" runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px" BorderStyle="Groove">0</asp:textbox><asp:textbox style="Z-INDEX: 129; POSITION: absolute; TEXT-ALIGN: right; TOP: 584px; LEFT: 162px"
				id="DFPRAH2" runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="40px" BorderStyle="Groove">0</asp:textbox><asp:textbox style="Z-INDEX: 130; POSITION: absolute; TEXT-ALIGN: right; TOP: 584px; LEFT: 224px"
				id="DFPRAM2" runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px" BorderStyle="Groove">0</asp:textbox><asp:textbox style="Z-INDEX: 131; POSITION: absolute; TEXT-ALIGN: right; TOP: 584px; LEFT: 328px"
				id="DFPRBH" runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px" BorderStyle="Groove">0</asp:textbox><asp:textbox style="Z-INDEX: 132; POSITION: absolute; TEXT-ALIGN: right; TOP: 584px; LEFT: 414px"
				id="DFPRBM" runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px" BorderStyle="Groove">0</asp:textbox><asp:textbox style="Z-INDEX: 133; POSITION: absolute; TEXT-ALIGN: right; TOP: 584px; LEFT: 518px"
				id="DFPRCH" runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px" BorderStyle="Groove">0</asp:textbox><asp:textbox style="Z-INDEX: 133; POSITION: absolute; TEXT-ALIGN: right; TOP: 584px; LEFT: 608px"
				id="DFPRCM" runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px" BorderStyle="Groove">0</asp:textbox><asp:textbox style="Z-INDEX: 133; POSITION: absolute; TEXT-ALIGN: right; TOP: 584px; LEFT: 608px"
				id="DFPRTAX15H" runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px" BorderStyle="Groove">0</asp:textbox><asp:textbox style="Z-INDEX: 133; POSITION: absolute; TEXT-ALIGN: right; TOP: 584px; LEFT: 608px"
				id="DFPRTAX15M" runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px" BorderStyle="Groove">0</asp:textbox><asp:textbox style="Z-INDEX: 133; POSITION: absolute; TEXT-ALIGN: right; TOP: 584px; LEFT: 608px"
				id="DFPRTAX167H" runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px" BorderStyle="Groove">0</asp:textbox><asp:textbox style="Z-INDEX: 133; POSITION: absolute; TEXT-ALIGN: right; TOP: 584px; LEFT: 608px"
				id="DFPRTAX167M" runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px" BorderStyle="Groove">0</asp:textbox><asp:textbox style="Z-INDEX: 133; POSITION: absolute; TEXT-ALIGN: right; TOP: 584px; LEFT: 608px"
				id="DFPRTAX20H" runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px" BorderStyle="Groove">0</asp:textbox><asp:textbox style="Z-INDEX: 133; POSITION: absolute; TEXT-ALIGN: right; TOP: 584px; LEFT: 608px"
				id="DFPRTAX20M" runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px" BorderStyle="Groove">0</asp:textbox><asp:textbox style="Z-INDEX: 133; POSITION: absolute; TEXT-ALIGN: right; TOP: 584px; LEFT: 608px"
				id="DFPRTAX267H" runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px" BorderStyle="Groove">0</asp:textbox><asp:textbox style="Z-INDEX: 133; POSITION: absolute; TEXT-ALIGN: right; TOP: 584px; LEFT: 608px"
				id="DFPRTAX267M" runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px" BorderStyle="Groove">0</asp:textbox>
			<asp:textbox style="Z-INDEX: 133; POSITION: absolute; TEXT-ALIGN: right; TOP: 560px; LEFT: 608px"
				id="DAgentID" runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px"
				BorderStyle="Groove">ID</asp:textbox>
			<asp:textbox style="Z-INDEX: 133; POSITION: absolute; TEXT-ALIGN: right; TOP: 560px; LEFT: 518px"
				id="DSalaryYM1" runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="60px"
				BorderStyle="Groove">YM</asp:textbox></FONT></form>
		<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ --> 
		</FORM>
	</body>
</HTML>
