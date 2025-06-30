<%@ Page Language="vb" AutoEventWireup="false" Codebehind="IT_ChangeSystemSheet_01.aspx.vb" Inherits="SPD.IT_ChangeSystemSheet_01" aspCompat="True"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>系統變更申請書</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<script language="javascript" type="text/javascript">
			//--Calendar------------------------------------
			function CalendarPicker(strField)
			{
				window.open('DatePicker.aspx?field=' + strField,'calendarPopup','width=250,height=190,resizable=yes');
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
				<asp:button id="BFinishDate" style="Z-INDEX: 122; LEFT: 320px; POSITION: absolute; TOP: 152px"
					runat="server" Width="20px" Height="20px" CausesValidation="False" Text="....."></asp:button>
				<asp:dropdownlist id="DSystem" style="Z-INDEX: 138; LEFT: 456px; POSITION: absolute; TOP: 152px" runat="server"
					Height="20px" Width="224px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="要" Selected="True">要</asp:ListItem>
					<asp:ListItem Value="不要">不要</asp:ListItem>
				</asp:dropdownlist>
				<asp:textbox id="DBDays" style="Z-INDEX: 137; LEFT: 608px; POSITION: absolute; TOP: 448px" runat="server"
					Height="20px" Width="70px" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue">DBDays</asp:textbox><asp:textbox id="DDepoName" style="Z-INDEX: 136; LEFT: 120px; POSITION: absolute; TOP: 120px"
					runat="server" Width="40px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">台北</asp:textbox><asp:textbox id="DDepoCode" style="Z-INDEX: 135; LEFT: 160px; POSITION: absolute; TOP: 120px"
					runat="server" Width="25px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">01</asp:textbox><asp:textbox id="DDivision" style="Z-INDEX: 111; LEFT: 192px; POSITION: absolute; TOP: 120px"
					runat="server" Width="88px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DDivision</asp:textbox><asp:textbox id="DDivisionCode" style="Z-INDEX: 112; LEFT: 280px; POSITION: absolute; TOP: 120px"
					runat="server" Width="63px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DDivisionCode</asp:textbox><asp:textbox id="DJobCode" style="Z-INDEX: 134; LEFT: 600px; POSITION: absolute; TOP: 120px"
					runat="server" Width="80px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DJobCode</asp:textbox><asp:textbox id="DEmpID" style="Z-INDEX: 133; LEFT: 600px; POSITION: absolute; TOP: 88px" runat="server"
					Width="80px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DEmpID</asp:textbox><asp:button id="BBEndDate" style="Z-INDEX: 132; LEFT: 488px; POSITION: absolute; TOP: 448px"
					runat="server" Width="20px" Height="20px" CausesValidation="False" Text="....."></asp:button><asp:textbox id="DBEndDate" style="Z-INDEX: 131; LEFT: 336px; POSITION: absolute; TOP: 448px"
					runat="server" Width="150px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">DBEndDate</asp:textbox><asp:button id="BBStartDate" style="Z-INDEX: 130; LEFT: 272px; POSITION: absolute; TOP: 448px"
					runat="server" Width="20px" Height="20px" CausesValidation="False" Text="....."></asp:button><asp:textbox id="DBStartDate" style="Z-INDEX: 129; LEFT: 120px; POSITION: absolute; TOP: 448px"
					runat="server" Width="150px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">DBStartDate</asp:textbox><asp:textbox id="DITEffect" style="Z-INDEX: 128; LEFT: 120px; POSITION: absolute; TOP: 480px"
					runat="server" Width="560px" Height="90px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" TextMode="MultiLine" MaxLength="250">DITEffect</asp:textbox><asp:dropdownlist id="DDevelopItem" style="Z-INDEX: 127; LEFT: 456px; POSITION: absolute; TOP: 416px"
					runat="server" Width="224px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="要" Selected="True">要</asp:ListItem>
					<asp:ListItem Value="不要">不要</asp:ListItem>
				</asp:dropdownlist><asp:hyperlink id="LRefFile" style="Z-INDEX: 125; LEFT: 120px; POSITION: absolute; TOP: 384px"
					runat="server" Width="72px" Height="8px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">參考附件</asp:hyperlink><INPUT id="DRefFile" style="Z-INDEX: 116; LEFT: 120px; WIDTH: 560px; POSITION: absolute; TOP: 384px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					type="file" size="50" name="File1" runat="server"><asp:textbox id="DEffect" style="Z-INDEX: 124; LEFT: 120px; POSITION: absolute; TOP: 282px" runat="server"
					Width="560px" Height="88px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" TextMode="MultiLine" MaxLength="250">DEffect</asp:textbox><asp:textbox id="DTarget" style="Z-INDEX: 123; LEFT: 120px; POSITION: absolute; TOP: 184px" runat="server"
					Width="560px" Height="88px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" TextMode="MultiLine" MaxLength="250">DTarget</asp:textbox><asp:textbox id="DFinishDate" style="Z-INDEX: 121; LEFT: 120px; POSITION: absolute; TOP: 152px"
					runat="server" Width="198px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DFinishDate</asp:textbox><asp:textbox id="DName" style="Z-INDEX: 120; LEFT: 456px; POSITION: absolute; TOP: 88px" runat="server"
					Width="144px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DName</asp:textbox><asp:label id="DHistoryLabel" style="Z-INDEX: 119; LEFT: 8px; POSITION: absolute; TOP: 792px"
					runat="server" Width="64px" Height="16px" ForeColor="Blue" Font-Size="11pt" Font-Names="新細明體">核定履歷</asp:label><asp:textbox id="DReasonDesc" style="Z-INDEX: 118; LEFT: 168px; POSITION: absolute; TOP: 696px"
					runat="server" Width="424px" Height="59px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" TextMode="MultiLine" Visible="False">DReasonDesc</asp:textbox><asp:textbox id="DReason" style="Z-INDEX: 117; LEFT: 240px; POSITION: absolute; TOP: 664px" runat="server"
					Width="352px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True" Visible="False">DReason</asp:textbox><asp:dropdownlist id="DReasonCode" style="Z-INDEX: 115; LEFT: 168px; POSITION: absolute; TOP: 664px"
					runat="server" Width="64px" Height="20px" ForeColor="Blue" BackColor="Yellow" Visible="False" AutoPostBack="True">
					<asp:ListItem Value="01">01</asp:ListItem>
					<asp:ListItem Value="02">02</asp:ListItem>
				</asp:dropdownlist><asp:image id="DDelay" style="Z-INDEX: 114; LEFT: 8px; POSITION: absolute; TOP: 656px" runat="server"
					Width="593px" Height="107px" Visible="False" ImageUrl="Images\Sheet_Delay.jpg"></asp:image><asp:textbox id="DDecideDesc" style="Z-INDEX: 113; LEFT: 56px; POSITION: absolute; TOP: 592px"
					runat="server" Width="536px" Height="56px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" TextMode="MultiLine">DDecideDesc</asp:textbox><asp:image id="DDescSheet" style="Z-INDEX: 110; LEFT: 8px; POSITION: absolute; TOP: 584px"
					runat="server" Width="593px" Height="75px" ImageUrl="Images\MapSheet_004.jpg"></asp:image><asp:datagrid id="DataGrid9" style="Z-INDEX: 109; LEFT: 8px; POSITION: absolute; TOP: 808px" runat="server"
					Width="780px" Height="100px" BackColor="White" BorderStyle="None" Font-Size="9pt" AutoGenerateColumns="False" CellPadding="4" BorderWidth="1px" BorderColor="#CC9966">
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
						<asp:BoundColumn DataField="AgentName" ReadOnly="True" HeaderText="代理">
							<HeaderStyle Width="50px"></HeaderStyle>
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
							<HeaderStyle Width="220px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="Description" ReadOnly="True" HeaderText="核定時間">
							<HeaderStyle Width="150px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
					</Columns>
					<PagerStyle NextPageText="下一頁" PrevPageText="上一頁" HorizontalAlign="Center" ForeColor="#993300"
						BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
				</asp:datagrid><asp:dropdownlist id="DEngineer" style="Z-INDEX: 108; LEFT: 120px; POSITION: absolute; TOP: 416px"
					runat="server" Width="224px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="要" Selected="True">要</asp:ListItem>
					<asp:ListItem Value="不要">不要</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DJobTitle" style="Z-INDEX: 107; LEFT: 456px; POSITION: absolute; TOP: 120px"
					runat="server" Width="144px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DJobTitle</asp:textbox><asp:textbox id="DDate" style="Z-INDEX: 106; LEFT: 120px; POSITION: absolute; TOP: 88px" runat="server"
					Width="224px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DDate</asp:textbox><asp:image id="DChangeSystemSheet1" style="Z-INDEX: 100; LEFT: 8px; POSITION: absolute; TOP: 8px"
					runat="server" Width="678px" Height="570px" ImageUrl="Images\IT_ChangeSystemSheet_001.jpg"></asp:image><asp:textbox id="DNo" style="Z-INDEX: 101; LEFT: 8px; POSITION: absolute; TOP: 9px" runat="server"
					Width="216px" Height="16px" ForeColor="Black" BackColor="White" BorderStyle="None" ReadOnly="True" Font-Names="Times New Roman"> 123</asp:textbox></FONT><INPUT id="BSAVE" style="Z-INDEX: 102; LEFT: 336px; WIDTH: 80px; POSITION: absolute; TOP: 768px; HEIGHT: 28px"
				onclick="Button('SAVE');" type="button" value="儲存" name="Button1" runat="server"><INPUT id="BNG2" style="Z-INDEX: 103; LEFT: 424px; WIDTH: 80px; POSITION: absolute; TOP: 768px; HEIGHT: 28px"
				onclick="Button('NG2');" type="button" value="NG2" name="Button1" runat="server"><INPUT id="BNG1" style="Z-INDEX: 104; LEFT: 512px; WIDTH: 80px; POSITION: absolute; TOP: 768px; HEIGHT: 28px"
				onclick="Button('NG1');" type="button" value="NG1" name="Button1" runat="server"><INPUT id="BOK" style="Z-INDEX: 105; LEFT: 600px; WIDTH: 80px; POSITION: absolute; TOP: 768px; HEIGHT: 28px"
				onclick="Button('OK');" type="button" value="OK" name="Button2" runat="server"></form>
		</FONT></FORM>
	</body>
</HTML>
