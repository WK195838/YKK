<%@ Page Language="vb" AutoEventWireup="false" Codebehind="ColorAppendSheet_01.aspx.vb" Inherits="SPD.ColorAppendSheet_01" aspCompat="True"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>ColorAppendSheet_01</title>
		<meta content="False" name="vs_snapToGrid">
		<meta content="True" name="vs_showGrid">
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
				<asp:image id="DColorAppendSheet" style="Z-INDEX: 103; LEFT: 8px; POSITION: absolute; TOP: -12px"
					runat="server" ImageUrl="Images\ColorAppendSheet_003.jpg" Height="440px" Width="632px"></asp:image>
				<asp:dropdownlist id="DDivision" style="Z-INDEX: 142; LEFT: 120px; POSITION: absolute; TOP: 124px"
					runat="server" Width="192px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist>
				<asp:textbox id="DLevel" style="Z-INDEX: 141; LEFT: 434px; POSITION: absolute; TOP: 158px" runat="server"
					Width="198px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">DLevel</asp:textbox><asp:imagebutton id="BPrint" style="Z-INDEX: 140; LEFT: 8px; POSITION: absolute; TOP: 8px" runat="server"
					ImageUrl="Images\Print.gif" Height="16px" Width="16px"></asp:imagebutton><asp:imagebutton id="BFlow" style="Z-INDEX: 123; LEFT: 8px; POSITION: absolute; TOP: 32px" runat="server"
					ImageUrl="Images\Flow-01.gif" Height="16px" Width="16px"></asp:imagebutton><asp:hyperlink id="LForCastFile" style="Z-INDEX: 122; LEFT: 120px; POSITION: absolute; TOP: 397px"
					runat="server" Height="8px" Width="64px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">報價單</asp:hyperlink><asp:textbox id="DPrice" style="Z-INDEX: 121; LEFT: 434px; POSITION: absolute; TOP: 364px" runat="server"
					Height="20px" Width="145px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DPrice</asp:textbox><asp:textbox id="DPullerPrice" style="Z-INDEX: 120; LEFT: 120px; POSITION: absolute; TOP: 364px"
					runat="server" Height="20px" Width="145px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DPullerPrice</asp:textbox><INPUT id="DForCastFile" style="Z-INDEX: 119; LEFT: 120px; WIDTH: 512px; POSITION: absolute; TOP: 397px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					type="file" size="66" name="File1" runat="server">
				<asp:button id="BImport" style="Z-INDEX: 118; LEFT: 304px; POSITION: absolute; TOP: 90px" runat="server"
					Height="20px" Width="20px" Text="進"></asp:button><asp:button id="BOut" style="Z-INDEX: 117; LEFT: 280px; POSITION: absolute; TOP: 90px" runat="server"
					Height="20px" Width="20px" Text="外"></asp:button><asp:hyperlink id="LOFormNo" style="Z-INDEX: 116; LEFT: 256px; POSITION: absolute; TOP: 226px"
					runat="server" Height="8px" Width="64px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">原委託</asp:hyperlink><asp:textbox id="DOFormNo" style="Z-INDEX: 115; LEFT: 120px; POSITION: absolute; TOP: 226px"
					runat="server" Height="20px" Width="72px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" AutoPostBack="True" ReadOnly="True">DOFormNo</asp:textbox><asp:button id="BIn" style="Z-INDEX: 114; LEFT: 256px; POSITION: absolute; TOP: 90px" runat="server"
					Height="20px" Width="20px" Text="內"></asp:button><asp:textbox id="DManufFlow" style="Z-INDEX: 113; LEFT: 120px; POSITION: absolute; TOP: 293px"
					runat="server" Height="56px" Width="512px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" TextMode="MultiLine">DManufFlow</asp:textbox><asp:textbox id="DColorItem" style="Z-INDEX: 112; LEFT: 120px; POSITION: absolute; TOP: 260px"
					runat="server" Height="20px" Width="512px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DColorItem</asp:textbox><asp:textbox id="DOFormSno" style="Z-INDEX: 111; LEFT: 200px; POSITION: absolute; TOP: 226px"
					runat="server" Height="20px" Width="48px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" AutoPostBack="True" ReadOnly="True">DOFormSno</asp:textbox><asp:textbox id="DMapNo" style="Z-INDEX: 110; LEFT: 120px; POSITION: absolute; TOP: 192px" runat="server"
					Height="20px" Width="200px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DMapNo</asp:textbox><asp:textbox id="DSliderCode" style="Z-INDEX: 109; LEFT: 120px; POSITION: absolute; TOP: 158px"
					runat="server" Height="20px" Width="200px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DSliderCode</asp:textbox><asp:textbox id="DNo" style="Z-INDEX: 108; LEFT: 120px; POSITION: absolute; TOP: 90px" runat="server"
					Height="20px" Width="136px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DNo</asp:textbox><asp:textbox id="DDate" style="Z-INDEX: 106; LEFT: 434px; POSITION: absolute; TOP: 90px" runat="server"
					Height="20px" Width="176px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True">DDate</asp:textbox><asp:button id="BDate" style="Z-INDEX: 107; LEFT: 610px; POSITION: absolute; TOP: 90px" runat="server"
					Height="20px" Width="20px" Text="....."></asp:button><asp:dropdownlist id="DPerson" style="Z-INDEX: 104; LEFT: 434px; POSITION: absolute; TOP: 124px" runat="server"
					Height="20px" Width="192px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist></FONT><asp:textbox id="DOPReadyDesc" style="Z-INDEX: 135; LEFT: 240px; POSITION: absolute; TOP: 584px"
				runat="server" Height="20px" Width="96px" BackColor="White" ForeColor="Red" BorderStyle="None">需閱讀工程履歷</asp:textbox><asp:textbox id="DOPReady" style="Z-INDEX: 124; LEFT: 336px; POSITION: absolute; TOP: 584px"
				runat="server" Height="20px" Width="48px" BackColor="GreenYellow" ForeColor="Red" BorderStyle="Groove" ReadOnly="True"></asp:textbox><asp:textbox id="DReasonDesc" style="Z-INDEX: 125; LEFT: 168px; POSITION: absolute; TOP: 656px"
				runat="server" Height="59px" Width="424px" BackColor="Gold" ForeColor="Blue" BorderStyle="Groove" TextMode="MultiLine" Visible="False">DReasonDesc</asp:textbox><asp:textbox id="DReason" style="Z-INDEX: 126; LEFT: 240px; POSITION: absolute; TOP: 624px" runat="server"
				Height="20px" Width="352px" BackColor="Gold" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True" Visible="False">DReason</asp:textbox><asp:dropdownlist id="DReasonCode" style="Z-INDEX: 127; LEFT: 168px; POSITION: absolute; TOP: 624px"
				runat="server" Height="20px" Width="64px" BackColor="Gold" ForeColor="Blue" AutoPostBack="True" Visible="False">
				<asp:ListItem Value="01">01</asp:ListItem>
				<asp:ListItem Value="02">02</asp:ListItem>
			</asp:dropdownlist><asp:hyperlink id="LBefOP" style="Z-INDEX: 128; LEFT: 168px; POSITION: absolute; TOP: 584px" runat="server"
				Height="8px" Width="144px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">工程履歷</asp:hyperlink><asp:textbox id="DAEndTime" style="Z-INDEX: 139; LEFT: 440px; POSITION: absolute; TOP: 547px"
				runat="server" Height="20px" Width="152px" BackColor="Gold" ForeColor="Blue" Font-Size="9pt" BorderStyle="Groove" ReadOnly="True">DAEndTime</asp:textbox><asp:textbox id="DAStartTime" style="Z-INDEX: 137; LEFT: 168px; POSITION: absolute; TOP: 547px"
				runat="server" Height="20px" Width="152px" BackColor="Gold" ForeColor="Blue" Font-Size="9pt" BorderStyle="Groove" ReadOnly="True">DAStartTime</asp:textbox><asp:textbox id="DBEndTime" style="Z-INDEX: 138; LEFT: 440px; POSITION: absolute; TOP: 514px"
				runat="server" Height="20px" Width="152px" BackColor="Gold" ForeColor="Blue" Font-Size="9pt" BorderStyle="Groove" ReadOnly="True">DBEndTime</asp:textbox><asp:textbox id="DBStartTime" style="Z-INDEX: 136; LEFT: 168px; POSITION: absolute; TOP: 514px"
				runat="server" Height="20px" Width="152px" BackColor="Gold" ForeColor="Blue" Font-Size="9pt" BorderStyle="Groove" ReadOnly="True">DBStartTime</asp:textbox><asp:textbox id="DDecideDesc" style="Z-INDEX: 129; LEFT: 56px; POSITION: absolute; TOP: 440px"
				runat="server" Height="56px" Width="536px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" TextMode="MultiLine">DDecideDesc</asp:textbox><asp:image id="DDelay" style="Z-INDEX: 100; LEFT: 8px; POSITION: absolute; TOP: 616px" runat="server"
				ImageUrl="Images\Sheet_Delay.jpg" Height="107px" Width="593px" Visible="False"></asp:image><asp:image id="DDelivery" style="Z-INDEX: 101; LEFT: 8px; POSITION: absolute; TOP: 504px" runat="server"
				ImageUrl="Images\Sheet_Delivery.jpg" Height="110px" Width="593px"></asp:image><asp:textbox id="DFormSno" style="Z-INDEX: 130; LEFT: 8px; POSITION: absolute; TOP: 720px" runat="server"
				Height="20px" Width="97px" BackColor="White" ForeColor="Blue" BorderStyle="None">單號：123</asp:textbox><INPUT id="BSAVE" style="Z-INDEX: 131; LEFT: 256px; WIDTH: 80px; POSITION: absolute; TOP: 728px; HEIGHT: 28px"
				onclick="Button('SAVE');" type="button" value="儲存" name="BSAVE" runat="server"><INPUT id="BNG2" style="Z-INDEX: 132; LEFT: 344px; WIDTH: 80px; POSITION: absolute; TOP: 728px; HEIGHT: 28px"
				onclick="Button('NG2');" type="button" value="NG2" name="Button1" runat="server"><INPUT id="BNG1" style="Z-INDEX: 133; LEFT: 432px; WIDTH: 80px; POSITION: absolute; TOP: 728px; HEIGHT: 28px"
				onclick="Button('NG1');" type="button" value="NG1" name="Button1" runat="server"><INPUT id="BOK" style="Z-INDEX: 134; LEFT: 520px; WIDTH: 80px; POSITION: absolute; TOP: 728px; HEIGHT: 28px"
				onclick="Button('OK');" type="button" value="OK" name="BOK" runat="server">
			<asp:image id="DDescSheet" style="Z-INDEX: 102; LEFT: 8px; POSITION: absolute; TOP: 432px"
				runat="server" ImageUrl="Images\MapSheet_004.jpg" Height="75px" Width="593px"></asp:image></form>
	</body>
</HTML>
