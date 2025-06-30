<%@ Page Language="vb" AutoEventWireup="false" Codebehind="zzOPContactOutSheet_01.aspx.vb" Inherits="SPD.OPContactOutSheet_01" aspCompat="True"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>OPContactOutSheet_01</title>
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
			function MapPicker(strField)
			{
				window.open('MapPicker.aspx?field=' + strField,'MapPopup','width=168,height=360,resizable=yes');
			}
			function ModifySheet()
			{
				if (document.Form1.DOFormNo.value != "" && document.Form1.DOFormSno.value != "") {
					if (document.Form1.DNFormNo.value == "") {
						wNFormNo="900003";
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
					window.open('ModManufOutSheet_01.aspx?pFormNo=' + wNFormNo + '&pFormSno=' + wNFormSno + '&pOFormNo=' + document.Form1.DOFormNo.value + '&pOFormSno=' + document.Form1.DOFormSno.value + '&pStep=' + wStep,'ModifySheet','width=620,height=580,scrollbars=yes,resizable=yes');
				}
			}
		</script>
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<FORM id="Form1" method="post" runat="server">
			<FONT face="新細明體">
				<asp:image id="DOPContactSheet" style="Z-INDEX: 100; LEFT: 8px; POSITION: absolute; TOP: 8px"
					runat="server" Height="454px" Width="593px" ImageUrl="Images\OPContactOutSheet_001.jpg"></asp:image>
				<asp:hyperlink id="LNFormNo" style="Z-INDEX: 141; LEFT: 496px; POSITION: absolute; TOP: 228px"
					runat="server" Height="8px" Width="48px" Target="_blank" NavigateUrl="BoardEdit.aspx" Font-Size="12pt">新委託</asp:hyperlink>
				<asp:button id="BModify" style="Z-INDEX: 140; LEFT: 496px; POSITION: absolute; TOP: 226px" runat="server"
					Height="20px" Width="56px" Text="新委託"></asp:button>
				<asp:hyperlink id="LOFormNo" style="Z-INDEX: 139; LEFT: 272px; POSITION: absolute; TOP: 228px"
					runat="server" Height="8px" Width="48px" Target="_blank" NavigateUrl="BoardEdit.aspx" Font-Size="12pt">原委託</asp:hyperlink>
				<asp:hyperlink id="LMapNo" style="Z-INDEX: 138; LEFT: 346px; POSITION: absolute; TOP: 194px" runat="server"
					Height="8px" Width="40px" Target="_blank" NavigateUrl="BoardEdit.aspx" Font-Size="12pt">連結</asp:hyperlink>
				<asp:textbox id="DNFormSno" style="Z-INDEX: 137; LEFT: 416px; POSITION: absolute; TOP: 226px"
					runat="server" Height="20px" Width="72px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow"
					ReadOnly="True" AutoPostBack="True">DNFormSno</asp:textbox>
				<asp:textbox id="DNFormNo" style="Z-INDEX: 136; LEFT: 344px; POSITION: absolute; TOP: 226px"
					runat="server" Height="20px" Width="72px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow"
					ReadOnly="True" AutoPostBack="True">DNFormNo</asp:textbox>
				<asp:hyperlink id="LAttachFile" style="Z-INDEX: 135; LEFT: 552px; POSITION: absolute; TOP: 436px"
					runat="server" Height="8px" Width="32px" Target="_blank" NavigateUrl="BoardEdit.aspx" Font-Size="12pt">附件</asp:hyperlink>
				<asp:textbox id="DOFormNo" style="Z-INDEX: 134; LEFT: 120px; POSITION: absolute; TOP: 226px"
					runat="server" Height="20px" Width="72px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow"
					ReadOnly="True" AutoPostBack="True">DOFormNo</asp:textbox>
				<asp:button id="BIn" style="Z-INDEX: 132; LEFT: 272px; POSITION: absolute; TOP: 90px" runat="server"
					Height="20px" Width="20px" Text="....."></asp:button>
				<asp:textbox id="DReasonDesc" style="Z-INDEX: 131; LEFT: 168px; POSITION: absolute; TOP: 688px"
					runat="server" Height="59px" Width="424px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow"
					TextMode="MultiLine">DReasonDesc</asp:textbox>
				<asp:textbox id="DReason" style="Z-INDEX: 130; LEFT: 240px; POSITION: absolute; TOP: 656px" runat="server"
					Height="20px" Width="352px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" ReadOnly="True">DReason</asp:textbox>
				<asp:dropdownlist id="DReasonCode" style="Z-INDEX: 129; LEFT: 168px; POSITION: absolute; TOP: 656px"
					runat="server" Height="20px" Width="64px" ForeColor="Blue" BackColor="Yellow" AutoPostBack="True">
					<asp:ListItem Value="01">01</asp:ListItem>
					<asp:ListItem Value="02">02</asp:ListItem>
				</asp:dropdownlist>
				<asp:hyperlink id="LBefOP" style="Z-INDEX: 128; LEFT: 168px; POSITION: absolute; TOP: 618px" runat="server"
					Height="8px" Width="144px" Target="_blank" NavigateUrl="BoardEdit.aspx" Font-Size="12pt">工程履歷</asp:hyperlink>
				<asp:textbox id="DAEndTime" style="Z-INDEX: 127; LEFT: 440px; POSITION: absolute; TOP: 582px"
					runat="server" Height="20px" Width="152px" Font-Size="9pt" BorderStyle="Groove" ForeColor="Blue"
					BackColor="White" ReadOnly="True">DAEndTime</asp:textbox>
				<asp:textbox id="DAStartTime" style="Z-INDEX: 126; LEFT: 168px; POSITION: absolute; TOP: 582px"
					runat="server" Height="20px" Width="152px" Font-Size="9pt" BorderStyle="Groove" ForeColor="Blue"
					BackColor="White" ReadOnly="True">DAStartTime</asp:textbox>
				<asp:textbox id="DBEndTime" style="Z-INDEX: 125; LEFT: 440px; POSITION: absolute; TOP: 548px"
					runat="server" Height="20px" Width="152px" Font-Size="9pt" BorderStyle="Groove" ForeColor="Blue"
					BackColor="White" ReadOnly="True">DBEndTime</asp:textbox>
				<asp:textbox id="DBStartTime" style="Z-INDEX: 124; LEFT: 168px; POSITION: absolute; TOP: 548px"
					runat="server" Height="20px" Width="152px" Font-Size="9pt" BorderStyle="Groove" ForeColor="Blue"
					BackColor="White" ReadOnly="True">DBStartTime</asp:textbox>
				<asp:textbox id="DDecideDesc" style="Z-INDEX: 123; LEFT: 56px; POSITION: absolute; TOP: 472px"
					runat="server" Height="56px" Width="536px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow"
					TextMode="MultiLine">DDecideDesc</asp:textbox>
				<asp:image id="DDelay" style="Z-INDEX: 122; LEFT: 8px; POSITION: absolute; TOP: 648px" runat="server"
					Height="107px" Width="593px" ImageUrl="Images\Sheet_Delay.jpg"></asp:image>
				<asp:image id="DDelivery" style="Z-INDEX: 121; LEFT: 8px; POSITION: absolute; TOP: 536px" runat="server"
					Height="110px" Width="593px" ImageUrl="Images\Sheet_Delivery.jpg"></asp:image>
				<asp:image id="DDescSheet" style="Z-INDEX: 120; LEFT: 8px; POSITION: absolute; TOP: 464px"
					runat="server" Height="75px" Width="593px" ImageUrl="Images\MapSheet_004.jpg"></asp:image>
				<asp:textbox id="DFormSno" style="Z-INDEX: 119; LEFT: 8px; POSITION: absolute; TOP: 752px" runat="server"
					Height="20px" Width="97px" BorderStyle="None" ForeColor="Blue" BackColor="White">單號：123</asp:textbox>
				<asp:button id="BSave" style="Z-INDEX: 111; LEFT: 272px; POSITION: absolute; TOP: 752px" runat="server"
					Height="32px" Width="105px" Text="儲存"></asp:button>
				<asp:button id="BNG" style="Z-INDEX: 108; LEFT: 384px; POSITION: absolute; TOP: 752px" runat="server"
					Height="32px" Width="105px" Text="NG"></asp:button>
				<asp:button id="BOK" style="Z-INDEX: 106; LEFT: 496px; POSITION: absolute; TOP: 752px" runat="server"
					Height="32px" Width="105px" Text="OK"></asp:button><INPUT id="DAttachFile" style="Z-INDEX: 118; LEFT: 120px; WIDTH: 424px; POSITION: absolute; TOP: 434px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					type="file" size="51" name="File1" runat="server">
				<asp:textbox id="DDReason" style="Z-INDEX: 117; LEFT: 120px; POSITION: absolute; TOP: 364px"
					runat="server" Height="56px" Width="472px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow"
					TextMode="MultiLine">DReason</asp:textbox>
				<asp:textbox id="DContent" style="Z-INDEX: 116; LEFT: 120px; POSITION: absolute; TOP: 294px"
					runat="server" Height="56px" Width="472px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow"
					TextMode="MultiLine">DContent</asp:textbox>
				<asp:textbox id="DTarget" style="Z-INDEX: 115; LEFT: 120px; POSITION: absolute; TOP: 260px" runat="server"
					Height="20px" Width="472px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow">DTarget</asp:textbox>
				<asp:textbox id="DOFormSno" style="Z-INDEX: 114; LEFT: 192px; POSITION: absolute; TOP: 226px"
					runat="server" Height="20px" Width="72px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow"
					ReadOnly="True" AutoPostBack="True">DOFormSno</asp:textbox>
				<asp:textbox id="DMapNo" style="Z-INDEX: 113; LEFT: 120px; POSITION: absolute; TOP: 192px" runat="server"
					Height="20px" Width="176px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" AutoPostBack="True">DMapNo</asp:textbox>
				<asp:button id="BOMapNo" style="Z-INDEX: 109; LEFT: 296px; POSITION: absolute; TOP: 192px" runat="server"
					Height="20px" Width="20px" Text="原"></asp:button>
				<asp:button id="BMMapNo" style="Z-INDEX: 112; LEFT: 318px; POSITION: absolute; TOP: 192px" runat="server"
					Height="20px" Width="20px" Text="修"></asp:button>
				<asp:dropdownlist id="DLevel" style="Z-INDEX: 110; LEFT: 416px; POSITION: absolute; TOP: 158px" runat="server"
					Height="20px" Width="176px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist>
				<asp:textbox id="DSliderCode" style="Z-INDEX: 107; LEFT: 120px; POSITION: absolute; TOP: 158px"
					runat="server" Height="20px" Width="176px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow">DSliderCode</asp:textbox>
				<asp:textbox id="DNo" style="Z-INDEX: 105; LEFT: 120px; POSITION: absolute; TOP: 90px" runat="server"
					Height="20px" Width="152px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" ReadOnly="True">DNo</asp:textbox>
				<asp:textbox id="DDate" style="Z-INDEX: 103; LEFT: 416px; POSITION: absolute; TOP: 90px" runat="server"
					Height="20px" Width="152px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" ReadOnly="True">DDate</asp:textbox>
				<asp:button id="BDate" style="Z-INDEX: 104; LEFT: 568px; POSITION: absolute; TOP: 90px" runat="server"
					Height="20px" Width="20px" Text="....."></asp:button>
				<asp:dropdownlist id="DPerson" style="Z-INDEX: 102; LEFT: 416px; POSITION: absolute; TOP: 124px" runat="server"
					Height="20px" Width="176px" ForeColor="Blue" BackColor="Yellow" AutoPostBack="True">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DDivision" style="Z-INDEX: 101; LEFT: 120px; POSITION: absolute; TOP: 124px"
					runat="server" Height="20px" Width="176px" ForeColor="Blue" BackColor="Yellow" AutoPostBack="True">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist></FONT></FORM>
	</body>
</HTML>
