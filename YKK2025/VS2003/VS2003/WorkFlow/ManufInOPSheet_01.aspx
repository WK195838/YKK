<%@ Page Language="vb" AutoEventWireup="false" Codebehind="ManufInOPSheet_01.aspx.vb" Inherits="SPD.ManufInOPSheet_01" aspCompat="True"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>內製工程連絡單</title>
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
				window.open('DevNoPicker.aspx?field=' + strField + '&pFormNo=' + FormNo,'DevNoPopup','width=180,height=360,resizable=yes');
			}
			function MapPicker(strField,FormNo)
			{
				window.open('MapPicker_01.aspx?field=' + strField + '&pFormNo=' + FormNo,'MapPopup','width=168,height=360,resizable=yes');
			}
			function ModifySheet()
			{
				if (document.Form1.DOFormNo.value != "" && document.Form1.DOFormSno.value != "") {
					if (document.Form1.DNFormNo.value == "") {
						wNFormNo="900001";
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
					window.open('ManufInModSheet_01.aspx?pFormNo=' + wNFormNo + '&pFormSno=' + wNFormSno + '&pOFormNo=' + document.Form1.DOFormNo.value + '&pOFormSno=' + document.Form1.DOFormSno.value + '&pStep=' + wStep,'ModifySheet','width=630,height=580,resizable=yes,scrollbars=yes, menubar=yes');
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
					FOK=document.getElementById("BOK");
					if(FOK!=null) document.Form1.BOK.disabled=true;  	
					//NG-1 Button
					FNG1=document.getElementById("BNG1");
					if(FNG1!=null) document.Form1.BNG1.disabled=true;  	
					//NG-2 Button
					FNG2=document.getElementById("BNG2");
					if(FNG2!=null) document.Form1.BNG2.disabled=true;  	
					//Save Button
					FSAVE=document.getElementById("BSAVE");
					if(FSAVE!=null) document.Form1.BSAVE.disabled=true;  	
						
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
				<asp:image id="DOPContactSheet" style="Z-INDEX: 100; POSITION: absolute; TOP: 8px; LEFT: 8px"
					runat="server" ImageUrl="Images\OPContactSheet_001.jpg" Width="593px" Height="454px"></asp:image>
				<asp:textbox id="DQCLT" style="Z-INDEX: 316; POSITION: absolute; TOP: 228px; LEFT: 520px" runat="server"
					Height="18px" Width="72px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox>
				<asp:label id="Label27" style="Z-INDEX: 315; POSITION: absolute; TOP: 228px; LEFT: 470px" runat="server"
					Height="10px" Width="48px">QC-L/T</asp:label>
				<asp:textbox id="DLevel" style="Z-INDEX: 150; POSITION: absolute; TOP: 158px; LEFT: 416px" runat="server"
					Height="20px" Width="176px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DLevel</asp:textbox>
				<asp:textbox id="DOPReadyDesc" style="Z-INDEX: 149; POSITION: absolute; TOP: 616px; LEFT: 240px"
					runat="server" Height="20px" Width="96px" BorderStyle="None" ForeColor="Red" BackColor="White">需閱讀工程履歷</asp:textbox>
				<asp:textbox id="DOPReady" style="Z-INDEX: 148; POSITION: absolute; TOP: 616px; LEFT: 336px"
					runat="server" Height="20px" Width="48px" BorderStyle="Groove" ForeColor="Red" BackColor="GreenYellow"
					ReadOnly="True"></asp:textbox><asp:textbox id="DReady" style="Z-INDEX: 147; POSITION: absolute; TOP: 226px; LEFT: 408px" runat="server"
					Width="48px" Height="20px" ReadOnly="True" BackColor="GreenYellow" ForeColor="Red" BorderStyle="Groove"></asp:textbox><asp:textbox id="DMakeCAM" style="Z-INDEX: 146; POSITION: absolute; TOP: 260px; LEFT: 496px"
					runat="server" Width="99px" Height="20px" ReadOnly="True" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DMakeCAM</asp:textbox><asp:textbox id="DSuppiler" style="Z-INDEX: 145; POSITION: absolute; TOP: 192px; LEFT: 496px"
					runat="server" Width="99px" Height="20px" ReadOnly="True" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DSuppiler</asp:textbox><asp:imagebutton id="BFlow" style="Z-INDEX: 144; POSITION: absolute; TOP: 32px; LEFT: 8px" runat="server"
					ImageUrl="Images\Flow-01.gif" Width="16px" Height="16px"></asp:imagebutton><asp:imagebutton id="BPrint" style="Z-INDEX: 143; POSITION: absolute; TOP: 8px; LEFT: 8px" runat="server"
					ImageUrl="Images\Print.gif" Width="16px" Height="16px"></asp:imagebutton><asp:hyperlink id="LNFormNo" style="Z-INDEX: 138; POSITION: absolute; TOP: 228px; LEFT: 352px"
					runat="server" Width="48px" Height="8px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">新委託</asp:hyperlink><asp:button id="BModify" style="Z-INDEX: 137; POSITION: absolute; TOP: 226px; LEFT: 352px" runat="server"
					Width="52px" Height="20px" Text="新委託"></asp:button><asp:hyperlink id="LOFormNo" style="Z-INDEX: 136; POSITION: absolute; TOP: 228px; LEFT: 208px"
					runat="server" Width="48px" Height="8px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">原委託</asp:hyperlink><asp:hyperlink id="LMapNo" style="Z-INDEX: 135; POSITION: absolute; TOP: 194px; LEFT: 344px" runat="server"
					Width="40px" Height="8px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">連結</asp:hyperlink><asp:textbox id="DNFormSno" style="Z-INDEX: 134; POSITION: absolute; TOP: 226px; LEFT: 312px"
					runat="server" Width="36px" Height="20px" ReadOnly="True" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" AutoPostBack="True">DNFormSno</asp:textbox><asp:textbox id="DNFormNo" style="Z-INDEX: 133; POSITION: absolute; TOP: 226px; LEFT: 264px"
					runat="server" Width="48px" Height="20px" ReadOnly="True" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" AutoPostBack="True">DNFormNo</asp:textbox><asp:hyperlink id="LAttachFile" style="Z-INDEX: 132; POSITION: absolute; TOP: 436px; LEFT: 552px"
					runat="server" Width="32px" Height="8px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">附件</asp:hyperlink><asp:textbox id="DOFormNo" style="Z-INDEX: 131; POSITION: absolute; TOP: 226px; LEFT: 120px"
					runat="server" Width="48px" Height="20px" ReadOnly="True" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" AutoPostBack="True">DOFormNo</asp:textbox><asp:button id="BIn" style="Z-INDEX: 130; POSITION: absolute; TOP: 90px; LEFT: 272px" runat="server"
					Width="20px" Height="20px" Text="....."></asp:button><asp:textbox id="DReasonDesc" style="Z-INDEX: 129; POSITION: absolute; TOP: 688px; LEFT: 168px"
					runat="server" Width="424px" Height="59px" BackColor="Gold" ForeColor="Blue" BorderStyle="Groove" TextMode="MultiLine" Visible="False">DReasonDesc</asp:textbox><asp:textbox id="DReason" style="Z-INDEX: 128; POSITION: absolute; TOP: 656px; LEFT: 240px" runat="server"
					Width="352px" Height="20px" ReadOnly="True" BackColor="Gold" ForeColor="Blue" BorderStyle="Groove" Visible="False">DReason</asp:textbox><asp:dropdownlist id="DReasonCode" style="Z-INDEX: 127; POSITION: absolute; TOP: 656px; LEFT: 168px"
					runat="server" Width="64px" Height="20px" BackColor="Gold" ForeColor="Blue" AutoPostBack="True" Visible="False">
					<asp:ListItem Value="01">01</asp:ListItem>
					<asp:ListItem Value="02">02</asp:ListItem>
				</asp:dropdownlist><asp:hyperlink id="LBefOP" style="Z-INDEX: 126; POSITION: absolute; TOP: 618px; LEFT: 168px" runat="server"
					Width="144px" Height="8px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">工程履歷</asp:hyperlink><asp:textbox id="DAEndTime" style="Z-INDEX: 125; POSITION: absolute; TOP: 582px; LEFT: 440px"
					runat="server" Width="152px" Height="20px" ReadOnly="True" BackColor="Gold" ForeColor="Blue" BorderStyle="Groove" Font-Size="9pt">DAEndTime</asp:textbox><asp:textbox id="DAStartTime" style="Z-INDEX: 124; POSITION: absolute; TOP: 582px; LEFT: 168px"
					runat="server" Width="152px" Height="20px" ReadOnly="True" BackColor="Gold" ForeColor="Blue" BorderStyle="Groove" Font-Size="9pt">DAStartTime</asp:textbox><asp:textbox id="DBEndTime" style="Z-INDEX: 123; POSITION: absolute; TOP: 548px; LEFT: 440px"
					runat="server" Width="152px" Height="20px" ReadOnly="True" BackColor="Gold" ForeColor="Blue" BorderStyle="Groove" Font-Size="9pt">DBEndTime</asp:textbox><asp:textbox id="DBStartTime" style="Z-INDEX: 122; POSITION: absolute; TOP: 548px; LEFT: 168px"
					runat="server" Width="152px" Height="20px" ReadOnly="True" BackColor="Gold" ForeColor="Blue" BorderStyle="Groove" Font-Size="9pt">DBStartTime</asp:textbox><asp:textbox id="DDecideDesc" style="Z-INDEX: 121; POSITION: absolute; TOP: 472px; LEFT: 56px"
					runat="server" Width="536px" Height="56px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" TextMode="MultiLine">DDecideDesc</asp:textbox><asp:image id="DDelay" style="Z-INDEX: 120; POSITION: absolute; TOP: 648px; LEFT: 8px" runat="server"
					ImageUrl="Images\Sheet_Delay.jpg" Width="593px" Height="107px" Visible="False"></asp:image><asp:image id="DDelivery" style="Z-INDEX: 119; POSITION: absolute; TOP: 536px; LEFT: 8px" runat="server"
					ImageUrl="Images\Sheet_Delivery.jpg" Width="593px" Height="110px"></asp:image><asp:image id="DDescSheet" style="Z-INDEX: 118; POSITION: absolute; TOP: 464px; LEFT: 8px"
					runat="server" ImageUrl="Images\MapSheet_004.jpg" Width="593px" Height="75px"></asp:image><asp:textbox id="DFormSno" style="Z-INDEX: 117; POSITION: absolute; TOP: 752px; LEFT: 8px" runat="server"
					Width="97px" Height="20px" BackColor="White" ForeColor="Blue" BorderStyle="None">單號：123</asp:textbox><INPUT id="DAttachFile" style="Z-INDEX: 116; POSITION: absolute; BACKGROUND-COLOR: #ffff00; WIDTH: 424px; HEIGHT: 20px; TOP: 434px; LEFT: 120px"
					type="file" size="51" name="File1" runat="server">
				<asp:textbox id="DDReason" style="Z-INDEX: 115; POSITION: absolute; TOP: 364px; LEFT: 120px"
					runat="server" Width="472px" Height="56px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove"
					TextMode="MultiLine">DReason</asp:textbox><asp:textbox id="DContent" style="Z-INDEX: 114; POSITION: absolute; TOP: 294px; LEFT: 120px"
					runat="server" Width="472px" Height="56px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" TextMode="MultiLine">DContent</asp:textbox><asp:textbox id="DTarget" style="Z-INDEX: 113; POSITION: absolute; TOP: 260px; LEFT: 120px" runat="server"
					Width="286px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DTarget</asp:textbox><asp:textbox id="DOFormSno" style="Z-INDEX: 112; POSITION: absolute; TOP: 226px; LEFT: 168px"
					runat="server" Width="36px" Height="20px" ReadOnly="True" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" AutoPostBack="True">DOFormSno</asp:textbox><asp:textbox id="DMapNo" style="Z-INDEX: 111; POSITION: absolute; TOP: 192px; LEFT: 120px" runat="server"
					Width="176px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DMapNo</asp:textbox><asp:button id="BOMapNo" style="Z-INDEX: 108; POSITION: absolute; TOP: 192px; LEFT: 296px" runat="server"
					Width="20px" Height="20px" Text="原"></asp:button><asp:button id="BMMapNo" style="Z-INDEX: 110; POSITION: absolute; TOP: 192px; LEFT: 318px" runat="server"
					Width="20px" Height="20px" Text="修"></asp:button><asp:textbox id="DSliderCode" style="Z-INDEX: 107; POSITION: absolute; TOP: 158px; LEFT: 120px"
					runat="server" Width="176px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DSliderCode</asp:textbox><asp:textbox id="DNo" style="Z-INDEX: 105; POSITION: absolute; TOP: 90px; LEFT: 120px" runat="server"
					Width="152px" Height="20px" ReadOnly="True" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DNo</asp:textbox><asp:textbox id="DDate" style="Z-INDEX: 103; POSITION: absolute; TOP: 90px; LEFT: 416px" runat="server"
					Width="152px" Height="20px" ReadOnly="True" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DDate</asp:textbox><asp:button id="BDate" style="Z-INDEX: 104; POSITION: absolute; TOP: 90px; LEFT: 568px" runat="server"
					Width="20px" Height="20px" Text="....."></asp:button><asp:dropdownlist id="DPerson" style="Z-INDEX: 102; POSITION: absolute; TOP: 124px; LEFT: 416px" runat="server"
					Width="176px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DDivision" style="Z-INDEX: 101; POSITION: absolute; TOP: 124px; LEFT: 120px"
					runat="server" Width="176px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist></FONT><INPUT id="BSAVE" style="Z-INDEX: 139; POSITION: absolute; WIDTH: 80px; HEIGHT: 28px; TOP: 760px; LEFT: 256px"
				onclick="Button('SAVE');" type="button" value="儲存" name="BSAVE" runat="server"><INPUT id="BNG2" style="Z-INDEX: 140; POSITION: absolute; WIDTH: 80px; HEIGHT: 28px; TOP: 760px; LEFT: 344px"
				onclick="Button('NG2');" type="button" value="NG2" name="Button1" runat="server"><INPUT id="BNG1" style="Z-INDEX: 141; POSITION: absolute; WIDTH: 80px; HEIGHT: 28px; TOP: 760px; LEFT: 432px"
				onclick="Button('NG1');" type="button" value="NG1" name="Button1" runat="server"><INPUT id="BOK" style="Z-INDEX: 142; POSITION: absolute; WIDTH: 80px; HEIGHT: 28px; TOP: 760px; LEFT: 520px"
				onclick="Button('OK');" type="button" value="OK" name="BOK" runat="server"></form>
	</body>
</HTML>
