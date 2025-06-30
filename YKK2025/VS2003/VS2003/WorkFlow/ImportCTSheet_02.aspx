<%@ Page Language="vb" AutoEventWireup="false" Codebehind="ImportCTSheet_02.aspx.vb" Inherits="SPD.ImportCTSheet_02"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>ImportCTSheet_02</title>
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
				<asp:image id="DOPContactSheet" style="Z-INDEX: 100; LEFT: 8px; POSITION: absolute; TOP: 8px"
					runat="server" ImageUrl="Images\ImportCTSheet_001.jpg" Width="594px" Height="419px"></asp:image>
				<asp:textbox id="DSpec" style="Z-INDEX: 143; LEFT: 392px; POSITION: absolute; TOP: 224px" runat="server"
					Height="20px" Width="200px" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue">DSpec</asp:textbox>
				<asp:textbox id="DPerson" style="Z-INDEX: 119; LEFT: 416px; POSITION: absolute; TOP: 124px" runat="server"
					Height="20px" Width="176px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DPerson</asp:textbox>
				<asp:textbox id="DDivision" style="Z-INDEX: 118; LEFT: 120px; POSITION: absolute; TOP: 124px"
					runat="server" Height="20px" Width="176px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove"
					ReadOnly="True">DDivision</asp:textbox>
				<asp:textbox id="DTarget" style="Z-INDEX: 116; LEFT: 120px; POSITION: absolute; TOP: 224px" runat="server"
					Height="20px" Width="256px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DTarget</asp:textbox><asp:imagebutton id="BFlow" style="Z-INDEX: 115; LEFT: 8px; POSITION: absolute; TOP: 32px" runat="server"
					ImageUrl="Images\Flow-01.gif" Width="16px" Height="16px"></asp:imagebutton><asp:imagebutton id="BPrint" style="Z-INDEX: 114; LEFT: 8px; POSITION: absolute; TOP: 8px" runat="server"
					ImageUrl="Images\Print.gif" Width="16px" Height="16px"></asp:imagebutton><asp:hyperlink id="LNFormNo" style="Z-INDEX: 113; LEFT: 432px; POSITION: absolute; TOP: 192px"
					runat="server" Width="48px" Height="8px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">新委託</asp:hyperlink><asp:hyperlink id="LOFormNo" style="Z-INDEX: 112; LEFT: 240px; POSITION: absolute; TOP: 192px"
					runat="server" Width="48px" Height="8px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">原委託</asp:hyperlink><asp:textbox id="DNFormSno" style="Z-INDEX: 111; LEFT: 376px; POSITION: absolute; TOP: 192px"
					runat="server" Width="50px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True" AutoPostBack="True">DNFormSno</asp:textbox><asp:textbox id="DNFormNo" style="Z-INDEX: 110; LEFT: 312px; POSITION: absolute; TOP: 192px"
					runat="server" Width="60px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True" AutoPostBack="True">DNFormNo</asp:textbox><asp:hyperlink id="LAttachFile" style="Z-INDEX: 109; LEFT: 120px; POSITION: absolute; TOP: 400px"
					runat="server" Width="32px" Height="8px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">附件</asp:hyperlink><asp:textbox id="DOFormNo" style="Z-INDEX: 108; LEFT: 120px; POSITION: absolute; TOP: 192px"
					runat="server" Width="60px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True" AutoPostBack="True">DOFormNo</asp:textbox><asp:textbox id="DFormSno" style="Z-INDEX: 107; LEFT: 8px; POSITION: absolute; TOP: 432px" runat="server"
					Width="97px" Height="20px" BackColor="White" ForeColor="Blue" BorderStyle="None">單號：123</asp:textbox>
				<asp:textbox id="DDReason" style="Z-INDEX: 106; LEFT: 120px; POSITION: absolute; TOP: 328px"
					runat="server" Width="472px" Height="56px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove"
					TextMode="MultiLine">DReason</asp:textbox><asp:textbox id="DContent" style="Z-INDEX: 105; LEFT: 120px; POSITION: absolute; TOP: 256px"
					runat="server" Width="472px" Height="56px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" TextMode="MultiLine">DContent</asp:textbox><asp:textbox id="DOFormSno" style="Z-INDEX: 104; LEFT: 184px; POSITION: absolute; TOP: 192px"
					runat="server" Width="50px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True" AutoPostBack="True">DOFormSno</asp:textbox><asp:textbox id="DSliderCode" style="Z-INDEX: 103; LEFT: 120px; POSITION: absolute; TOP: 158px"
					runat="server" Width="472px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DSliderCode</asp:textbox><asp:textbox id="DNo" style="Z-INDEX: 102; LEFT: 120px; POSITION: absolute; TOP: 90px" runat="server"
					Width="176px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DNo</asp:textbox><asp:textbox id="DDate" style="Z-INDEX: 101; LEFT: 416px; POSITION: absolute; TOP: 90px" runat="server"
					Width="176px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True">DDate</asp:textbox></FONT></form>
	</body>
</HTML>
