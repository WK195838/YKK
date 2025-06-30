<%@ Page Language="vb" AutoEventWireup="false" Codebehind="ImportCTSheet_01.aspx.vb" Inherits="SPD.ImportCTSheet_01" aspCompat="True"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>ImportCTSheet_01</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<script language="javascript" type="text/javascript">
			function CalendarPicker(strField)
			{
				window.open('DatePicker.aspx?field=' + strField,'CalendarPopup','width=250,height=190,resizable=yes');
			}
			function DevNoPicker(strField,FormNo)
			{
				window.open('DevNoPicker.aspx?field=' + strField + '&pFormNo=' + FormNo,'DevNoPopup','width=250,height=360,resizable=yes');
			}
			function MapPicker(strField,FormNo)
			{
				window.open('MapPicker_01.aspx?field=' + strField + '&pFormNo=' + FormNo,'MapPopup','width=168,height=360,resizable=yes');
			}
			function ModifySheet()
			{
				if (document.Form1.DOFormNo.value != "" && document.Form1.DOFormSno.value != "") {
					if (document.Form1.DNFormNo.value == "") {
						wNFormNo="900011";
					} else {
						wNFormNo=document.Form1.DNFormNo.value;
					}
					if (document.Form1.DNFormSno.value == "") {
						wNFormSno=0;
					} else {
						wNFormSno=document.Form1.DNFormSno.value;
					}

					var CookieStr=(document.cookie+';');
					var pos = CookieStr.indexOf('Step=');
					if(pos<0){
						wStep=0;
					} else {
						wStep=document.cookie.substring(pos+5,CookieStr.indexOf(';',pos+1));
					}
					//alert(wNFormNo);
					//alert(wNFormSno);
					//alert(document.Form1.DOFormNo.value);
					//alert(document.Form1.DOFormSno.value);
					//alert(wStep);
					window.open('ImportModSheet_01.aspx?pFormNo=' + wNFormNo + '&pFormSno=' + wNFormSno + '&pOFormNo=' + document.Form1.DOFormNo.value + '&pOFormSno=' + document.Form1.DOFormSno.value + '&pStep=' + wStep,'ModifySheet','width=630,height=580,resizable=yes,scrollbars=yes, menubar=yes');
				}
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
		    function OpenPrintSheet(URL) {
				window.open(URL,'PrintSheet','width=630,height=580,resizable=yes,scrollbars=yes, menubar=yes');
			}
		    function OpenSimulationSheet(URL) {
				window.open(URL,'Simulation','width=630,height=580,resizable=yes,scrollbars=yes, menubar=yes');
			}
		</script>
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體">
				<asp:image id="DOPContactSheet" style="Z-INDEX: 100; LEFT: 8px; POSITION: absolute; TOP: 8px"
					runat="server" ImageUrl="Images\ImportCTSheet_001.jpg" Width="594px" Height="419px"></asp:image>
				<asp:textbox id="DSpec" style="Z-INDEX: 143; LEFT: 384px; POSITION: absolute; TOP: 224px" runat="server"
					Height="20px" Width="208px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">DSpec</asp:textbox><asp:dropdownlist id="DTarget" style="Z-INDEX: 142; LEFT: 120px; POSITION: absolute; TOP: 224px" runat="server"
					Width="256px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DOPReadyDesc" style="Z-INDEX: 141; LEFT: 240px; POSITION: absolute; TOP: 584px"
					runat="server" Width="96px" Height="20px" BackColor="White" ForeColor="Red" BorderStyle="None">需閱讀工程履歷</asp:textbox><asp:textbox id="DOPReady" style="Z-INDEX: 140; LEFT: 336px; POSITION: absolute; TOP: 584px"
					runat="server" Width="48px" Height="20px" BackColor="GreenYellow" ForeColor="Red" BorderStyle="Groove" ReadOnly="True"></asp:textbox><asp:textbox id="DReady" style="Z-INDEX: 139; LEFT: 496px; POSITION: absolute; TOP: 192px" runat="server"
					Width="48px" Height="20px" BackColor="GreenYellow" ForeColor="Red" BorderStyle="Groove" ReadOnly="True"></asp:textbox><asp:imagebutton id="BFlow" style="Z-INDEX: 138; LEFT: 8px; POSITION: absolute; TOP: 32px" runat="server"
					ImageUrl="Images\Flow-01.gif" Width="16px" Height="16px"></asp:imagebutton><asp:imagebutton id="BPrint" style="Z-INDEX: 137; LEFT: 8px; POSITION: absolute; TOP: 8px" runat="server"
					ImageUrl="Images\Print.gif" Width="16px" Height="16px"></asp:imagebutton><asp:hyperlink id="LNFormNo" style="Z-INDEX: 132; LEFT: 432px; POSITION: absolute; TOP: 192px"
					runat="server" Width="48px" Height="8px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">新委託</asp:hyperlink><asp:button id="BModify" style="Z-INDEX: 131; LEFT: 432px; POSITION: absolute; TOP: 192px" runat="server"
					Width="56px" Height="20px" Text="新委託"></asp:button><asp:hyperlink id="LOFormNo" style="Z-INDEX: 130; LEFT: 240px; POSITION: absolute; TOP: 192px"
					runat="server" Width="48px" Height="8px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">原委託</asp:hyperlink><asp:textbox id="DNFormSno" style="Z-INDEX: 129; LEFT: 376px; POSITION: absolute; TOP: 192px"
					runat="server" Width="50px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True" AutoPostBack="True">DNFormSno</asp:textbox><asp:textbox id="DNFormNo" style="Z-INDEX: 128; LEFT: 312px; POSITION: absolute; TOP: 192px"
					runat="server" Width="60px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True" AutoPostBack="True">DNFormNo</asp:textbox><asp:hyperlink id="LAttachFile" style="Z-INDEX: 127; LEFT: 552px; POSITION: absolute; TOP: 400px"
					runat="server" Width="32px" Height="8px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">附件</asp:hyperlink><asp:textbox id="DOFormNo" style="Z-INDEX: 126; LEFT: 120px; POSITION: absolute; TOP: 192px"
					runat="server" Width="60px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True" AutoPostBack="True">DOFormNo</asp:textbox><asp:button id="BIn" style="Z-INDEX: 125; LEFT: 272px; POSITION: absolute; TOP: 90px" runat="server"
					Width="20px" Height="20px" Text="....."></asp:button><asp:textbox id="DReasonDesc" style="Z-INDEX: 124; LEFT: 168px; POSITION: absolute; TOP: 656px"
					runat="server" Width="424px" Height="59px" BackColor="Gold" ForeColor="Blue" BorderStyle="Groove" TextMode="MultiLine" Visible="False">DReasonDesc</asp:textbox><asp:textbox id="DReason" style="Z-INDEX: 123; LEFT: 240px; POSITION: absolute; TOP: 624px" runat="server"
					Width="352px" Height="20px" BackColor="Gold" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True" Visible="False">DReason</asp:textbox><asp:dropdownlist id="DReasonCode" style="Z-INDEX: 122; LEFT: 168px; POSITION: absolute; TOP: 624px"
					runat="server" Width="64px" Height="20px" BackColor="Gold" ForeColor="Blue" AutoPostBack="True" Visible="False">
					<asp:ListItem Value="01">01</asp:ListItem>
					<asp:ListItem Value="02">02</asp:ListItem>
				</asp:dropdownlist><asp:hyperlink id="LBefOP" style="Z-INDEX: 121; LEFT: 168px; POSITION: absolute; TOP: 584px" runat="server"
					Width="144px" Height="8px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">工程履歷</asp:hyperlink><asp:textbox id="DAEndTime" style="Z-INDEX: 120; LEFT: 440px; POSITION: absolute; TOP: 552px"
					runat="server" Width="152px" Height="20px" BackColor="Gold" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True" Font-Size="9pt">DAEndTime</asp:textbox><asp:textbox id="DAStartTime" style="Z-INDEX: 119; LEFT: 168px; POSITION: absolute; TOP: 552px"
					runat="server" Width="152px" Height="20px" BackColor="Gold" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True" Font-Size="9pt">DAStartTime</asp:textbox><asp:textbox id="DBEndTime" style="Z-INDEX: 118; LEFT: 440px; POSITION: absolute; TOP: 518px"
					runat="server" Width="152px" Height="20px" BackColor="Gold" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True" Font-Size="9pt">DBEndTime</asp:textbox><asp:textbox id="DBStartTime" style="Z-INDEX: 117; LEFT: 168px; POSITION: absolute; TOP: 518px"
					runat="server" Width="152px" Height="20px" BackColor="Gold" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True" Font-Size="9pt">DBStartTime</asp:textbox><asp:textbox id="DDecideDesc" style="Z-INDEX: 116; LEFT: 56px; POSITION: absolute; TOP: 440px"
					runat="server" Width="536px" Height="56px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" TextMode="MultiLine">DDecideDesc</asp:textbox><asp:image id="DDelay" style="Z-INDEX: 115; LEFT: 8px; POSITION: absolute; TOP: 616px" runat="server"
					ImageUrl="Images\Sheet_Delay.jpg" Width="593px" Height="107px" Visible="False"></asp:image><asp:image id="DDelivery" style="Z-INDEX: 114; LEFT: 8px; POSITION: absolute; TOP: 504px" runat="server"
					ImageUrl="Images\Sheet_Delivery.jpg" Width="593px" Height="110px"></asp:image><asp:image id="DDescSheet" style="Z-INDEX: 113; LEFT: 8px; POSITION: absolute; TOP: 432px"
					runat="server" ImageUrl="Images\MapSheet_004.jpg" Width="593px" Height="75px"></asp:image><asp:textbox id="DFormSno" style="Z-INDEX: 112; LEFT: 8px; POSITION: absolute; TOP: 720px" runat="server"
					Width="97px" Height="20px" BackColor="White" ForeColor="Blue" BorderStyle="None">單號：123</asp:textbox><INPUT id="DAttachFile" style="Z-INDEX: 111; LEFT: 120px; WIDTH: 424px; POSITION: absolute; TOP: 398px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					type="file" size="51" name="File1" runat="server">
				<asp:textbox id="DDReason" style="Z-INDEX: 110; LEFT: 120px; POSITION: absolute; TOP: 328px"
					runat="server" Width="472px" Height="56px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove"
					TextMode="MultiLine">DReason</asp:textbox><asp:textbox id="DContent" style="Z-INDEX: 109; LEFT: 120px; POSITION: absolute; TOP: 256px"
					runat="server" Width="472px" Height="56px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" TextMode="MultiLine">DContent</asp:textbox><asp:textbox id="DOFormSno" style="Z-INDEX: 108; LEFT: 184px; POSITION: absolute; TOP: 192px"
					runat="server" Width="50px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True" AutoPostBack="True">DOFormSno</asp:textbox><asp:textbox id="DSliderCode" style="Z-INDEX: 107; LEFT: 120px; POSITION: absolute; TOP: 158px"
					runat="server" Width="472px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DSliderCode</asp:textbox><asp:textbox id="DNo" style="Z-INDEX: 105; LEFT: 120px; POSITION: absolute; TOP: 90px" runat="server"
					Width="152px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DNo</asp:textbox><asp:textbox id="DDate" style="Z-INDEX: 103; LEFT: 416px; POSITION: absolute; TOP: 90px" runat="server"
					Width="152px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True">DDate</asp:textbox><asp:button id="BDate" style="Z-INDEX: 104; LEFT: 568px; POSITION: absolute; TOP: 90px" runat="server"
					Width="20px" Height="20px" Text="....."></asp:button><asp:dropdownlist id="DPerson" style="Z-INDEX: 102; LEFT: 416px; POSITION: absolute; TOP: 124px" runat="server"
					Width="176px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DDivision" style="Z-INDEX: 101; LEFT: 120px; POSITION: absolute; TOP: 124px"
					runat="server" Width="176px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist></FONT><INPUT id="BSAVE" style="Z-INDEX: 133; LEFT: 256px; WIDTH: 80px; POSITION: absolute; TOP: 728px; HEIGHT: 28px"
				onclick="Button('SAVE');" type="button" value="儲存" name="BSAVE" runat="server"><INPUT id="BNG2" style="Z-INDEX: 134; LEFT: 344px; WIDTH: 80px; POSITION: absolute; TOP: 728px; HEIGHT: 28px"
				onclick="Button('NG2');" type="button" value="NG2" name="Button1" runat="server"><INPUT id="BNG1" style="Z-INDEX: 135; LEFT: 432px; WIDTH: 80px; POSITION: absolute; TOP: 728px; HEIGHT: 28px"
				onclick="Button('NG1');" type="button" value="NG1" name="Button1" runat="server"><INPUT id="BOK" style="Z-INDEX: 136; LEFT: 520px; WIDTH: 80px; POSITION: absolute; TOP: 728px; HEIGHT: 28px"
				onclick="Button('OK');" type="button" value="OK" name="BOK" runat="server"></form>
	</body>
</HTML>
