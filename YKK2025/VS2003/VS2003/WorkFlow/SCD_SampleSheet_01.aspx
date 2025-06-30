<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SCD_SampleSheet_01.aspx.vb" Inherits="SPD.SCD_SampleSheet_01" aspCompat="True"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>開發見本委託書</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<script language="javascript" type="text/javascript">
			//--Calendar------------------------------------
			function DataPicker(strField)
			{
				window.open('DevelopDataPicker.aspx?field=' + strField,'DataPopup','status=yes,width=300,height=350,resizable=yes');
			}
            //--Proc Button---------------------------------		
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
				<asp:label id="DHistoryLabel" style="Z-INDEX: 117; POSITION: absolute; TOP: 1288px; LEFT: 8px"
					runat="server" Font-Names="新細明體" Font-Size="11pt" Width="64px" Height="16px" ForeColor="Blue">核定履歷</asp:label><INPUT id="DQCFILE5" style="Z-INDEX: 177; POSITION: absolute; BACKGROUND-COLOR: #ffff00; WIDTH: 108px; HEIGHT: 20px; TOP: 678px; LEFT: 640px"
					type="file" size="1" name="FILE1" runat="server">
				<asp:hyperlink id="LQCFILE5" style="Z-INDEX: 178; POSITION: absolute; TOP: 680px; LEFT: 644px"
					runat="server" Font-Size="12pt" Height="8px" NavigateUrl="" Target="_blank">其它</asp:hyperlink><asp:label id="D1Other" style="Z-INDEX: 173; POSITION: absolute; TOP: 860px; LEFT: 159px" runat="server"
					Font-Size="12px">Other1</asp:label><asp:label id="D2Other" style="Z-INDEX: 172; POSITION: absolute; TOP: 884px; LEFT: 159px" runat="server"
					Font-Size="12px">Other2</asp:label><asp:textbox id="DRNO" style="Z-INDEX: 171; POSITION: absolute; TOP: 40px; LEFT: 16px" runat="server"
					Width="109px" Height="20px" ForeColor="#0000C0" ReadOnly="True" BorderStyle="None" BackColor="Yellow">DRNO</asp:textbox><asp:button id="BRNO" style="Z-INDEX: 170; POSITION: absolute; TOP: 101px; LEFT: 511px" runat="server"
					Width="20px" Height="20px" CausesValidation="False" Text="....."></asp:button><asp:dropdownlist id="DWF7Name" style="Z-INDEX: 169; POSITION: absolute; TOP: 1011px; LEFT: 638px"
					runat="server" Font-Size="9pt" Width="74px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="無">無</asp:ListItem>
					<asp:ListItem Value="田邊公一郎" Selected="True">田邊公一郎</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DWF6Name" style="Z-INDEX: 168; POSITION: absolute; TOP: 1011px; LEFT: 538px"
					runat="server" Font-Size="9pt" Width="74px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="無">無</asp:ListItem>
					<asp:ListItem Value="田邊公一郎" Selected="True">田邊公一郎</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DWF5Name" style="Z-INDEX: 167; POSITION: absolute; TOP: 1011px; LEFT: 438px"
					runat="server" Font-Size="9pt" Width="74px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="無">無</asp:ListItem>
					<asp:ListItem Value="田邊公一郎" Selected="True">田邊公一郎</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DWF4Name" style="Z-INDEX: 166; POSITION: absolute; TOP: 1011px; LEFT: 337px"
					runat="server" Font-Size="9pt" Width="74px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="無">無</asp:ListItem>
					<asp:ListItem Value="田邊公一郎" Selected="True">田邊公一郎</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DWF3Name" style="Z-INDEX: 165; POSITION: absolute; TOP: 1011px; LEFT: 238px"
					runat="server" Font-Size="9pt" Width="74px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="無">無</asp:ListItem>
					<asp:ListItem Value="田邊公一郎" Selected="True">田邊公一郎</asp:ListItem>
				</asp:dropdownlist><asp:hyperlink id="LQCFILE1" style="Z-INDEX: 160; POSITION: absolute; TOP: 680px; LEFT: 140px"
					runat="server" Font-Size="12pt" Height="8px" NavigateUrl="" Target="_blank">吋法檔案</asp:hyperlink><asp:dropdownlist id="DWF7" style="Z-INDEX: 159; POSITION: absolute; TOP: 1038px; LEFT: 638px" runat="server"
					Font-Size="9pt" Width="74px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="無">無</asp:ListItem>
					<asp:ListItem Value="田邊公一郎" Selected="True">田邊公一郎</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DWF6" style="Z-INDEX: 158; POSITION: absolute; TOP: 1038px; LEFT: 538px" runat="server"
					Font-Size="9pt" Width="74px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="無">無</asp:ListItem>
					<asp:ListItem Value="田邊公一郎" Selected="True">田邊公一郎</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DWF5" style="Z-INDEX: 157; POSITION: absolute; TOP: 1038px; LEFT: 438px" runat="server"
					Font-Size="9pt" Width="74px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="無">無</asp:ListItem>
					<asp:ListItem Value="田邊公一郎" Selected="True">田邊公一郎</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DWF4" style="Z-INDEX: 156; POSITION: absolute; TOP: 1038px; LEFT: 337px" runat="server"
					Font-Size="9pt" Width="74px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="無">無</asp:ListItem>
					<asp:ListItem Value="田邊公一郎" Selected="True">田邊公一郎</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DWF3" style="Z-INDEX: 155; POSITION: absolute; TOP: 1038px; LEFT: 238px" runat="server"
					Font-Size="9pt" Width="74px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="無">無</asp:ListItem>
					<asp:ListItem Value="田邊公一郎" Selected="True">田邊公一郎</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DWF2" style="Z-INDEX: 154; POSITION: absolute; TOP: 1038px; LEFT: 137px" runat="server"
					Font-Size="9pt" Width="74px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="無">無</asp:ListItem>
					<asp:ListItem Value="田邊公一郎" Selected="True">田邊公一郎</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DWF1" style="Z-INDEX: 109; POSITION: absolute; TOP: 1038px; LEFT: 38px" runat="server"
					Font-Size="9pt" Width="74px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="無">無</asp:ListItem>
					<asp:ListItem Value="田邊公一郎" Selected="True">田邊公一郎</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DOP6" style="Z-INDEX: 153; POSITION: absolute; TOP: 947px; LEFT: 658px" runat="server"
					Width="74px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="Yellow">OP6</asp:textbox><asp:textbox id="DOP5" style="Z-INDEX: 152; POSITION: absolute; TOP: 947px; LEFT: 558px" runat="server"
					Width="74px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="Yellow">OP5</asp:textbox><asp:textbox id="DOP4" style="Z-INDEX: 151; POSITION: absolute; TOP: 947px; LEFT: 458px" runat="server"
					Width="74px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="Yellow">OP4</asp:textbox><asp:textbox id="DOP3" style="Z-INDEX: 150; POSITION: absolute; TOP: 947px; LEFT: 358px" runat="server"
					Width="74px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="Yellow">OP3</asp:textbox><asp:textbox id="DOP2" style="Z-INDEX: 149; POSITION: absolute; TOP: 947px; LEFT: 258px" runat="server"
					Width="74px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="Yellow">OP2</asp:textbox><asp:textbox id="DOP1" style="Z-INDEX: 148; POSITION: absolute; TOP: 947px; LEFT: 157px" runat="server"
					Width="74px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="Yellow">OP1</asp:textbox><asp:textbox id="DCITEM" style="Z-INDEX: 147; POSITION: absolute; TOP: 901px; LEFT: 277px" runat="server"
					Width="312px" Height="19px" ForeColor="Blue" BorderStyle="Groove" BackColor="Yellow">CITEM</asp:textbox><asp:textbox id="DO2ITEM" style="Z-INDEX: 145; POSITION: absolute; TOP: 878px; LEFT: 277px" runat="server"
					Width="312px" Height="19px" ForeColor="Blue" BorderStyle="Groove" BackColor="Yellow"></asp:textbox><asp:textbox id="DO1ITEM" style="Z-INDEX: 146; POSITION: absolute; TOP: 856px; LEFT: 277px" runat="server"
					Width="312px" Height="19px" ForeColor="Blue" BorderStyle="Groove" BackColor="Yellow"></asp:textbox><asp:textbox id="DCDITEM" style="Z-INDEX: 144; POSITION: absolute; TOP: 833px; LEFT: 277px" runat="server"
					Width="312px" Height="19px" ForeColor="Blue" BorderStyle="Groove" BackColor="Yellow">CDITEM</asp:textbox><asp:textbox id="DCSITEM" style="Z-INDEX: 143; POSITION: absolute; TOP: 810px; LEFT: 277px" runat="server"
					Width="312px" Height="19px" ForeColor="Blue" BorderStyle="Groove" BackColor="Yellow">CSITEM</asp:textbox><asp:textbox id="DTDRITEM" style="Z-INDEX: 142; POSITION: absolute; TOP: 764px; LEFT: 497px"
					runat="server" Width="92px" Height="19px" ForeColor="Blue" BorderStyle="Groove" BackColor="Yellow">TDRITEM</asp:textbox><asp:textbox id="DTSRITEM" style="Z-INDEX: 141; POSITION: absolute; TOP: 741px; LEFT: 497px"
					runat="server" Width="92px" Height="19px" ForeColor="Blue" BorderStyle="Groove" BackColor="Yellow">TSRITEM</asp:textbox><asp:textbox id="DTNRITEM" style="Z-INDEX: 140; POSITION: absolute; TOP: 718px; LEFT: 497px"
					runat="server" Width="92px" Height="19px" ForeColor="Blue" BorderStyle="Groove" BackColor="Yellow">TNRITEM</asp:textbox><asp:textbox id="DCNITEM" style="Z-INDEX: 139; POSITION: absolute; TOP: 787px; LEFT: 277px" runat="server"
					Width="312px" Height="19px" ForeColor="Blue" BorderStyle="Groove" BackColor="Yellow">CNITEM</asp:textbox><asp:textbox id="DTDLITEM" style="Z-INDEX: 138; POSITION: absolute; TOP: 764px; LEFT: 277px"
					runat="server" Width="92px" Height="19px" ForeColor="Blue" BorderStyle="Groove" BackColor="Yellow">TDLITEM</asp:textbox><asp:textbox id="DTSLITEM" style="Z-INDEX: 137; POSITION: absolute; TOP: 741px; LEFT: 277px"
					runat="server" Width="92px" Height="19px" ForeColor="Blue" BorderStyle="Groove" BackColor="Yellow">TSLITEM</asp:textbox><asp:textbox id="DTNLITEM" style="Z-INDEX: 136; POSITION: absolute; TOP: 718px; LEFT: 277px"
					runat="server" Width="92px" Height="19px" ForeColor="Blue" BorderStyle="Groove" BackColor="Yellow">TNLITEM</asp:textbox><INPUT id="DQCFILE1" style="Z-INDEX: 131; POSITION: absolute; BACKGROUND-COLOR: #ffff00; WIDTH: 116px; HEIGHT: 20px; TOP: 678px; LEFT: 137px"
					type="file" size="1" name="FILE1" runat="server">
				<asp:hyperlink id="LQCFILE2" style="Z-INDEX: 161; POSITION: absolute; TOP: 680px; LEFT: 260px"
					runat="server" Font-Size="12pt" Height="8px" NavigateUrl="" Target="_blank">強度檔案</asp:hyperlink><asp:hyperlink id="LQCFILE3" style="Z-INDEX: 163; POSITION: absolute; TOP: 680px; LEFT: 380px"
					runat="server" Font-Size="12pt" Height="8px" NavigateUrl="" Target="_blank">往覆測試檔案</asp:hyperlink><asp:hyperlink id="LQCFILE4" style="Z-INDEX: 162; POSITION: absolute; TOP: 680px; LEFT: 496px"
					runat="server" Font-Size="12pt" Height="8px" NavigateUrl="" Target="_blank">仕樣書.組織圖檔案</asp:hyperlink><INPUT id="DQCFILE4" style="Z-INDEX: 135; POSITION: absolute; BACKGROUND-COLOR: #ffff00; WIDTH: 124px; HEIGHT: 20px; TOP: 678px; LEFT: 504px"
					type="file" size="1" name="FILE1" runat="server"> <INPUT id="DQCFILE3" style="Z-INDEX: 132; POSITION: absolute; BACKGROUND-COLOR: #ffff00; WIDTH: 108px; HEIGHT: 20px; TOP: 678px; LEFT: 384px"
					type="file" size="1" name="FILE1" runat="server"> <INPUT id="DQCFILE2" style="Z-INDEX: 134; POSITION: absolute; BACKGROUND-COLOR: #ffff00; WIDTH: 112px; HEIGHT: 20px; TOP: 678px; LEFT: 260px"
					type="file" size="1" name="FILE1" runat="server">
				<asp:textbox id="DOTHER" style="Z-INDEX: 130; POSITION: absolute; TOP: 624px; LEFT: 140px" runat="server"
					Width="608px" Height="48px" ForeColor="Blue" BorderStyle="Groove" BackColor="Yellow" MaxLength="240"
					TextMode="MultiLine">OTHER</asp:textbox><asp:textbox id="DTHCOL" style="Z-INDEX: 129; POSITION: absolute; TOP: 568px; LEFT: 140px" runat="server"
					Width="608px" Height="48px" ForeColor="Blue" BorderStyle="Groove" BackColor="Yellow" MaxLength="240" TextMode="MultiLine">THCOL</asp:textbox><asp:textbox id="DCCOL" style="Z-INDEX: 128; POSITION: absolute; TOP: 540px; LEFT: 140px" runat="server"
					Width="608px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="Yellow">CCOL</asp:textbox><asp:textbox id="DECOL" style="Z-INDEX: 127; POSITION: absolute; TOP: 511px; LEFT: 140px" runat="server"
					Width="608px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="Yellow">ECOL</asp:textbox><asp:textbox id="DTACOL" style="Z-INDEX: 126; POSITION: absolute; TOP: 400px; LEFT: 140px" runat="server"
					Width="608px" Height="48px" ForeColor="Blue" BorderStyle="Groove" BackColor="Yellow" MaxLength="240" TextMode="MultiLine">TACOL</asp:textbox><asp:textbox id="DDEVPRD" style="Z-INDEX: 125; POSITION: absolute; TOP: 372px; LEFT: 580px" runat="server"
					Width="168px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="Yellow">DEVPRD</asp:textbox><asp:textbox id="DDEVNO" style="Z-INDEX: 124; POSITION: absolute; TOP: 372px; LEFT: 360px" runat="server"
					Width="112px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="Yellow">DEVNO</asp:textbox><asp:textbox id="DTAWIDTH" style="Z-INDEX: 123; POSITION: absolute; TOP: 372px; LEFT: 140px"
					runat="server" Width="80px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="Yellow">TAWIDTH</asp:textbox><INPUT id="DSAMPLEFILE" style="Z-INDEX: 122; POSITION: absolute; BACKGROUND-COLOR: #ffff00; WIDTH: 728px; HEIGHT: 20px; TOP: 204px; LEFT: 20px"
					type="file" size="102" name="File1" runat="server">
				<asp:image id="LSAMPLEFILE" style="Z-INDEX: 121; POSITION: absolute; TOP: 202px; LEFT: 20px"
					runat="server" Width="728px" Height="160px" BorderStyle="Groove" ImageUrl="F:\DMF04006-DS2W.jpg"></asp:image><asp:textbox id="DCODENO" style="Z-INDEX: 120; POSITION: absolute; TOP: 150px; LEFT: 544px" runat="server"
					Width="208px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="Yellow">CODENO</asp:textbox><asp:textbox id="DITEM" style="Z-INDEX: 119; POSITION: absolute; TOP: 150px; LEFT: 340px" runat="server"
					Width="192px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="Yellow">ITEM</asp:textbox><asp:textbox id="DSIZENO" style="Z-INDEX: 118; POSITION: absolute; TOP: 150px; LEFT: 140px" runat="server"
					Width="192px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="Yellow">SIZENO</asp:textbox><asp:textbox id="DReasonDesc" style="Z-INDEX: 116; POSITION: absolute; TOP: 1192px; LEFT: 167px"
					runat="server" Width="424px" Height="56px" ForeColor="Blue" BorderStyle="Groove" BackColor="Yellow" TextMode="MultiLine" Visible="False">DReasonDesc</asp:textbox><asp:textbox id="DReason" style="Z-INDEX: 115; POSITION: absolute; TOP: 1160px; LEFT: 239px"
					runat="server" Width="352px" Height="20px" ForeColor="Blue" ReadOnly="True" BorderStyle="Groove" BackColor="Yellow" Visible="False">DReason</asp:textbox><asp:dropdownlist id="DReasonCode" style="Z-INDEX: 114; POSITION: absolute; TOP: 1160px; LEFT: 167px"
					runat="server" Width="64px" Height="20px" ForeColor="Blue" BackColor="Yellow" Visible="False" AutoPostBack="True">
					<asp:ListItem Value="01">01</asp:ListItem>
					<asp:ListItem Value="02">02</asp:ListItem>
				</asp:dropdownlist><asp:image id="DDelay" style="Z-INDEX: 113; POSITION: absolute; TOP: 1152px; LEFT: 11px" runat="server"
					Width="593px" Height="107px" ImageUrl="Images\Sheet_Delay.jpg" Visible="False"></asp:image><asp:textbox id="DDecideDesc" style="Z-INDEX: 112; POSITION: absolute; TOP: 1088px; LEFT: 57px"
					runat="server" Width="536px" Height="56px" ForeColor="Blue" BorderStyle="Groove" BackColor="Yellow" TextMode="MultiLine">DDecideDesc</asp:textbox><asp:image id="DDescSheet" style="Z-INDEX: 111; POSITION: absolute; TOP: 1080px; LEFT: 11px"
					runat="server" Width="593px" Height="75px" ImageUrl="Images\MapSheet_004.jpg"></asp:image><asp:datagrid id="DataGrid9" style="Z-INDEX: 110; POSITION: absolute; TOP: 1312px; LEFT: 2px"
					runat="server" Font-Size="9pt" Width="780px" Height="100px" BorderStyle="None" BackColor="White" BorderColor="#CC9966" BorderWidth="1px" CellPadding="4" AutoGenerateColumns="False">
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
				</asp:datagrid><asp:textbox id="DTALINE" style="Z-INDEX: 103; POSITION: absolute; TOP: 456px; LEFT: 140px" runat="server"
					Width="608px" Height="48px" ForeColor="Blue" BorderStyle="Groove" BackColor="Yellow" MaxLength="240" TextMode="MultiLine">TALINE</asp:textbox><asp:image id="DSampleSheet1" style="Z-INDEX: 101; POSITION: absolute; TOP: 8px; LEFT: 11px"
					runat="server" ImageUrl="Images\SCD_SampleSheet_011.jpg"></asp:image><asp:image id="DSampleSheet2" style="Z-INDEX: 100; POSITION: absolute; TOP: 704px; LEFT: 8px"
					runat="server" ImageUrl="Images\SCD_SampleSheet_02.jpg"></asp:image><asp:textbox id="DDATE" style="Z-INDEX: 164; POSITION: absolute; TOP: 100px; LEFT: 640px" runat="server"
					Width="109px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="Yellow">DATE</asp:textbox><asp:textbox id="DAPPBUYER" style="Z-INDEX: 104; POSITION: absolute; TOP: 100px; LEFT: 140px"
					runat="server" Width="370px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="Yellow">APPBUYER</asp:textbox><asp:textbox id="DNo" style="Z-INDEX: 102; POSITION: absolute; TOP: 9px; LEFT: 17px" runat="server"
					Font-Names="Times New Roman" Width="216px" Height="16px" ForeColor="Black" ReadOnly="True" BorderStyle="None" BackColor="White"> 123</asp:textbox><INPUT id="BSAVE" style="Z-INDEX: 105; POSITION: absolute; WIDTH: 80px; HEIGHT: 28px; TOP: 1264px; LEFT: 416px"
					onclick="Button('SAVE');" type="button" value="儲存" name="Button1" runat="server"><INPUT id="BNG2" style="Z-INDEX: 106; POSITION: absolute; WIDTH: 80px; HEIGHT: 28px; TOP: 1264px; LEFT: 504px"
					onclick="Button('NG2');" type="button" value="NG2" name="Button1" runat="server"><INPUT id="BNG1" style="Z-INDEX: 107; POSITION: absolute; WIDTH: 80px; HEIGHT: 28px; TOP: 1264px; LEFT: 592px"
					onclick="Button('NG1');" type="button" value="NG1" name="Button1" runat="server"><INPUT id="BOK" style="Z-INDEX: 108; POSITION: absolute; WIDTH: 80px; HEIGHT: 28px; TOP: 1264px; LEFT: 680px"
					onclick="Button('OK');" type="button" value="OK" name="Button2" runat="server">
				<INPUT id="Hidden1" style="Z-INDEX: 174; POSITION: absolute; TOP: 316px; LEFT: 848px" type="hidden"
					name="Hidden1" runat="server"><INPUT id="Hidden2" style="Z-INDEX: 175; POSITION: absolute; TOP: 348px; LEFT: 852px" type="hidden"
					name="Hidden2" runat="server"> </FONT>
		</form>
	</body>
</HTML>
