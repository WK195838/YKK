<%@ Page Language="vb" AutoEventWireup="false" Codebehind="AppendSpecSheet_02.aspx.vb" Inherits="SPD.AppendSpecSheet_02"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>AppendSpecSheet_02</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
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
				<asp:image id="DAppendSpecSheet" style="Z-INDEX: 100; LEFT: 8px; POSITION: absolute; TOP: 8px"
					runat="server" ImageUrl="Images\AppendSpecSheet_004.jpg" Width="635px" Height="689px"></asp:image>
				<asp:textbox id="DCPSC" style="Z-INDEX: 152; LEFT: 440px; POSITION: absolute; TOP: 192px" runat="server"
					Height="20px" Width="192px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow">DCPSC</asp:textbox>
				<asp:hyperlink style="Z-INDEX: 151; LEFT: 448px; POSITION: absolute; TOP: 600px" id="LAssemblerFile"
					runat="server" Height="8px" Width="80px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">組立判定書</asp:hyperlink>
				<asp:textbox style="Z-INDEX: 150; LEFT: 440px; POSITION: absolute; TOP: 224px" id="DAssembler"
					runat="server" Height="20px" Width="193px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DAssembler</asp:textbox><asp:textbox id="DBuyer" style="Z-INDEX: 148; LEFT: 120px; POSITION: absolute; TOP: 158px" runat="server"
					Width="200px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow">DBuyer</asp:textbox><asp:hyperlink id="LRefFile" style="Z-INDEX: 147; LEFT: 120px; POSITION: absolute; TOP: 668px"
					runat="server" Width="64px" Height="8px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">其他附件</asp:hyperlink><asp:hyperlink id="LContactFile" style="Z-INDEX: 146; LEFT: 120px; POSITION: absolute; TOP: 634px"
					runat="server" Width="64px" Height="8px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">切結書</asp:hyperlink><asp:textbox id="DPriceA1" style="Z-INDEX: 145; LEFT: 16px; POSITION: absolute; TOP: 496px" runat="server"
					Width="166px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:dropdownlist id="DPriceA2" style="Z-INDEX: 116; LEFT: 184px; POSITION: absolute; TOP: 496px"
					runat="server" Width="40px" Height="12px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceA3" style="Z-INDEX: 117; LEFT: 224px; POSITION: absolute; TOP: 496px"
					runat="server" Width="64px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DPriceB1" style="Z-INDEX: 118; LEFT: 296px; POSITION: absolute; TOP: 496px"
					runat="server" Width="166px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:dropdownlist id="DPriceB2" style="Z-INDEX: 119; LEFT: 464px; POSITION: absolute; TOP: 496px"
					runat="server" Width="40px" Height="12px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceB3" style="Z-INDEX: 120; LEFT: 504px; POSITION: absolute; TOP: 496px"
					runat="server" Width="64px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DPriceC1" style="Z-INDEX: 121; LEFT: 16px; POSITION: absolute; TOP: 514px" runat="server"
					Width="166px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:dropdownlist id="DPriceC2" style="Z-INDEX: 122; LEFT: 184px; POSITION: absolute; TOP: 514px"
					runat="server" Width="40px" Height="12px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceC3" style="Z-INDEX: 123; LEFT: 224px; POSITION: absolute; TOP: 514px"
					runat="server" Width="64px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DPriceD1" style="Z-INDEX: 124; LEFT: 296px; POSITION: absolute; TOP: 514px"
					runat="server" Width="166px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:dropdownlist id="DPriceD2" style="Z-INDEX: 125; LEFT: 464px; POSITION: absolute; TOP: 514px"
					runat="server" Width="40px" Height="12px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceD3" style="Z-INDEX: 126; LEFT: 504px; POSITION: absolute; TOP: 514px"
					runat="server" Width="64px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DPriceE1" style="Z-INDEX: 127; LEFT: 16px; POSITION: absolute; TOP: 532px" runat="server"
					Width="166px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:dropdownlist id="DPriceE2" style="Z-INDEX: 128; LEFT: 184px; POSITION: absolute; TOP: 532px"
					runat="server" Width="40px" Height="12px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceE3" style="Z-INDEX: 129; LEFT: 224px; POSITION: absolute; TOP: 532px"
					runat="server" Width="64px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DPriceF1" style="Z-INDEX: 130; LEFT: 296px; POSITION: absolute; TOP: 532px"
					runat="server" Width="166px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:dropdownlist id="DPriceF2" style="Z-INDEX: 131; LEFT: 464px; POSITION: absolute; TOP: 532px"
					runat="server" Width="40px" Height="12px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceF3" style="Z-INDEX: 132; LEFT: 504px; POSITION: absolute; TOP: 532px"
					runat="server" Width="64px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DPriceG1" style="Z-INDEX: 133; LEFT: 16px; POSITION: absolute; TOP: 550px" runat="server"
					Width="166px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:dropdownlist id="DPriceG2" style="Z-INDEX: 134; LEFT: 184px; POSITION: absolute; TOP: 550px"
					runat="server" Width="40px" Height="12px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceG3" style="Z-INDEX: 135; LEFT: 224px; POSITION: absolute; TOP: 550px"
					runat="server" Width="64px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DPriceH1" style="Z-INDEX: 136; LEFT: 296px; POSITION: absolute; TOP: 550px"
					runat="server" Width="166px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:dropdownlist id="DPriceH2" style="Z-INDEX: 137; LEFT: 464px; POSITION: absolute; TOP: 550px"
					runat="server" Width="40px" Height="12px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceH3" style="Z-INDEX: 138; LEFT: 504px; POSITION: absolute; TOP: 550px"
					runat="server" Width="64px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DPriceI1" style="Z-INDEX: 139; LEFT: 16px; POSITION: absolute; TOP: 568px" runat="server"
					Width="166px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:dropdownlist id="DPriceI2" style="Z-INDEX: 140; LEFT: 184px; POSITION: absolute; TOP: 568px"
					runat="server" Width="40px" Height="12px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceI3" style="Z-INDEX: 141; LEFT: 224px; POSITION: absolute; TOP: 568px"
					runat="server" Width="64px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DPriceJ1" style="Z-INDEX: 142; LEFT: 296px; POSITION: absolute; TOP: 568px"
					runat="server" Width="166px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:dropdownlist id="DPriceJ2" style="Z-INDEX: 143; LEFT: 464px; POSITION: absolute; TOP: 568px"
					runat="server" Width="40px" Height="12px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceJ3" style="Z-INDEX: 144; LEFT: 504px; POSITION: absolute; TOP: 568px"
					runat="server" Width="64px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DSellVendor" style="Z-INDEX: 115; LEFT: 440px; POSITION: absolute; TOP: 158px"
					runat="server" Width="192px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow">DSellVendor</asp:textbox><asp:hyperlink id="LOFormNo" style="Z-INDEX: 114; LEFT: 256px; POSITION: absolute; TOP: 260px"
					runat="server" Width="64px" Height="8px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">原委託</asp:hyperlink><asp:hyperlink id="LQAFile" style="Z-INDEX: 113; LEFT: 120px; POSITION: absolute; TOP: 600px" runat="server"
					Width="64px" Height="8px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">測試報告</asp:hyperlink><asp:textbox id="DOFormNo" style="Z-INDEX: 112; LEFT: 120px; POSITION: absolute; TOP: 260px"
					runat="server" Width="72px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" ReadOnly="True" AutoPostBack="True">DOFormNo</asp:textbox><asp:textbox id="DFormSno" style="Z-INDEX: 111; LEFT: 8px; POSITION: absolute; TOP: 696px" runat="server"
					Width="97px" Height="20px" BorderStyle="None" ForeColor="Blue" BackColor="White">單號：123</asp:textbox>
				<asp:textbox id="DDescription" style="Z-INDEX: 110; LEFT: 120px; POSITION: absolute; TOP: 394px"
					runat="server" Width="512px" Height="56px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow"
					TextMode="MultiLine">DDescription</asp:textbox><asp:textbox id="DContent" style="Z-INDEX: 109; LEFT: 120px; POSITION: absolute; TOP: 326px"
					runat="server" Width="512px" Height="56px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" TextMode="MultiLine">DContent</asp:textbox><asp:textbox id="DTarget" style="Z-INDEX: 108; LEFT: 120px; POSITION: absolute; TOP: 292px" runat="server"
					Width="512px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow">DTarget</asp:textbox><asp:textbox id="DOFormSno" style="Z-INDEX: 107; LEFT: 200px; POSITION: absolute; TOP: 260px"
					runat="server" Width="48px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" ReadOnly="True" AutoPostBack="True">DOFormSno</asp:textbox><asp:textbox id="DMapNo" style="Z-INDEX: 106; LEFT: 120px; POSITION: absolute; TOP: 224px" runat="server"
					Width="200px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow">DMapNo</asp:textbox><asp:textbox id="DSliderCode" style="Z-INDEX: 105; LEFT: 120px; POSITION: absolute; TOP: 192px"
					runat="server" Width="200px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow">DSliderCode</asp:textbox><asp:textbox id="DNo" style="Z-INDEX: 104; LEFT: 120px; POSITION: absolute; TOP: 90px" runat="server"
					Width="200px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" ReadOnly="True">DNo</asp:textbox><asp:textbox id="DDate" style="Z-INDEX: 103; LEFT: 440px; POSITION: absolute; TOP: 90px" runat="server"
					Width="192px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" ReadOnly="True">DDate</asp:textbox><asp:dropdownlist id="DPerson" style="Z-INDEX: 102; LEFT: 440px; POSITION: absolute; TOP: 124px" runat="server"
					Width="192px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DDivision" style="Z-INDEX: 101; LEFT: 120px; POSITION: absolute; TOP: 124px"
					runat="server" Width="200px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist></FONT></form>
	</body>
</HTML>
