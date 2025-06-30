<%@ Page Language="vb" AutoEventWireup="false" Codebehind="zzContactSheet_01.aspx.vb" Inherits="SPD.ContactSheet_01" aspCompat="True"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>ContactSheet_01</title>
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
			function MapPicker(strField)
			{
				window.open('MapPicker.aspx?field=' + strField,'MapPopup','width=168,height=360,resizable=yes');
			}
			function ModifySheet()
			{
				if (document.Form1.DOFormNo.value != "" && document.Form1.DOFormSno.value != "") {
					if (document.Form1.DNFormNo.value == "") {
						wNFormNo="900002";
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
					window.open('ModManufSheet_02.aspx?pFormNo=' + wNFormNo + '&pFormSno=' + wNFormSno + '&pOFormNo=' + document.Form1.DOFormNo.value + '&pOFormSno=' + document.Form1.DOFormSno.value + '&pStep=' + wStep,'ModifySheet','width=620,height=580,scrollbars=yes,resizable=yes');
				}
			}
		</script>
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<FORM id="Form1" method="post" runat="server">
			<FONT face="新細明體">
				<asp:image id="DContactSheet" style="Z-INDEX: 100; LEFT: 8px; POSITION: absolute; TOP: 8px"
					runat="server" ImageUrl="Images\ContactSheet_001.jpg" Width="593px" Height="454px"></asp:image>
				<asp:hyperlink id="LNFormNo" style="Z-INDEX: 142; LEFT: 496px; POSITION: absolute; TOP: 228px"
					runat="server" Height="8px" Width="48px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">新委託</asp:hyperlink>
				<asp:button id="BModify" style="Z-INDEX: 141; LEFT: 496px; POSITION: absolute; TOP: 228px" runat="server"
					Height="20px" Width="56px" Text="新委託"></asp:button>
				<asp:hyperlink id="LOFormNo" style="Z-INDEX: 140; LEFT: 272px; POSITION: absolute; TOP: 228px"
					runat="server" Width="48px" Height="8px" Target="_blank" NavigateUrl="BoardEdit.aspx" Font-Size="12pt">原委託</asp:hyperlink>
				<asp:hyperlink id="LMapNo" style="Z-INDEX: 139; LEFT: 346px; POSITION: absolute; TOP: 194px" runat="server"
					Width="40px" Height="8px" Target="_blank" NavigateUrl="BoardEdit.aspx" Font-Size="12pt">連結</asp:hyperlink>
				<asp:textbox id="DNFormSno" style="Z-INDEX: 138; LEFT: 416px; POSITION: absolute; TOP: 226px"
					runat="server" Width="72px" Height="20px" AutoPostBack="True" ReadOnly="True" BackColor="Yellow"
					ForeColor="Blue" BorderStyle="Groove">DNFormSno</asp:textbox>
				<asp:textbox id="DNFormNo" style="Z-INDEX: 136; LEFT: 344px; POSITION: absolute; TOP: 226px"
					runat="server" Width="72px" Height="20px" AutoPostBack="True" ReadOnly="True" BackColor="Yellow"
					ForeColor="Blue" BorderStyle="Groove">DNFormNo</asp:textbox>
				<asp:hyperlink id="LAttachFile" style="Z-INDEX: 135; LEFT: 552px; POSITION: absolute; TOP: 436px"
					runat="server" Width="32px" Height="8px" Target="_blank" NavigateUrl="BoardEdit.aspx" Font-Size="12pt">附件</asp:hyperlink>
				<asp:textbox id="DOFormNo" style="Z-INDEX: 134; LEFT: 120px; POSITION: absolute; TOP: 226px"
					runat="server" Width="72px" Height="20px" AutoPostBack="True" ReadOnly="True" BackColor="Yellow"
					ForeColor="Blue" BorderStyle="Groove">DOFormNo</asp:textbox>
				<asp:button id="BIn" style="Z-INDEX: 132; LEFT: 272px; POSITION: absolute; TOP: 90px" runat="server"
					Width="20px" Height="20px" Text="....."></asp:button>
				<asp:textbox id="DReasonDesc" style="Z-INDEX: 131; LEFT: 168px; POSITION: absolute; TOP: 688px"
					runat="server" Width="424px" Height="59px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove"
					TextMode="MultiLine">DReasonDesc</asp:textbox>
				<asp:textbox id="DReason" style="Z-INDEX: 130; LEFT: 240px; POSITION: absolute; TOP: 656px" runat="server"
					Width="352px" Height="20px" ReadOnly="True" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DReason</asp:textbox>
				<asp:dropdownlist id="DReasonCode" style="Z-INDEX: 129; LEFT: 168px; POSITION: absolute; TOP: 656px"
					runat="server" Width="64px" Height="20px" AutoPostBack="True" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="01">01</asp:ListItem>
					<asp:ListItem Value="02">02</asp:ListItem>
				</asp:dropdownlist>
				<asp:hyperlink id="LBefOP" style="Z-INDEX: 128; LEFT: 168px; POSITION: absolute; TOP: 618px" runat="server"
					Width="144px" Height="8px" Target="_blank" NavigateUrl="BoardEdit.aspx" Font-Size="12pt">工程履歷</asp:hyperlink>
				<asp:textbox id="DAEndTime" style="Z-INDEX: 127; LEFT: 440px; POSITION: absolute; TOP: 582px"
					runat="server" Width="152px" Height="20px" Font-Size="9pt" ReadOnly="True" BackColor="White"
					ForeColor="Blue" BorderStyle="Groove">DAEndTime</asp:textbox>
				<asp:textbox id="DAStartTime" style="Z-INDEX: 126; LEFT: 168px; POSITION: absolute; TOP: 582px"
					runat="server" Width="152px" Height="20px" Font-Size="9pt" ReadOnly="True" BackColor="White"
					ForeColor="Blue" BorderStyle="Groove">DAStartTime</asp:textbox>
				<asp:textbox id="DBEndTime" style="Z-INDEX: 125; LEFT: 440px; POSITION: absolute; TOP: 548px"
					runat="server" Width="152px" Height="20px" Font-Size="9pt" ReadOnly="True" BackColor="White"
					ForeColor="Blue" BorderStyle="Groove">DBEndTime</asp:textbox>
				<asp:textbox id="DBStartTime" style="Z-INDEX: 124; LEFT: 168px; POSITION: absolute; TOP: 548px"
					runat="server" Width="152px" Height="20px" Font-Size="9pt" ReadOnly="True" BackColor="White"
					ForeColor="Blue" BorderStyle="Groove">DBStartTime</asp:textbox>
				<asp:textbox id="DDecideDesc" style="Z-INDEX: 123; LEFT: 56px; POSITION: absolute; TOP: 472px"
					runat="server" Width="536px" Height="56px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove"
					TextMode="MultiLine">DDecideDesc</asp:textbox>
				<asp:image id="DDelay" style="Z-INDEX: 122; LEFT: 8px; POSITION: absolute; TOP: 648px" runat="server"
					ImageUrl="Images\Sheet_Delay.jpg" Width="593px" Height="107px"></asp:image>
				<asp:image id="DDelivery" style="Z-INDEX: 121; LEFT: 8px; POSITION: absolute; TOP: 536px" runat="server"
					ImageUrl="Images\Sheet_Delivery.jpg" Width="593px" Height="110px"></asp:image>
				<asp:image id="DDescSheet" style="Z-INDEX: 120; LEFT: 8px; POSITION: absolute; TOP: 464px"
					runat="server" ImageUrl="Images\MapSheet_004.jpg" Width="593px" Height="75px"></asp:image>
				<asp:textbox id="DFormSno" style="Z-INDEX: 119; LEFT: 8px; POSITION: absolute; TOP: 752px" runat="server"
					Width="97px" Height="20px" BackColor="White" ForeColor="Blue" BorderStyle="None">單號：123</asp:textbox>
				<asp:button id="BSave" style="Z-INDEX: 111; LEFT: 272px; POSITION: absolute; TOP: 752px" runat="server"
					Width="105px" Height="32px" Text="儲存"></asp:button>
				<asp:button id="BNG" style="Z-INDEX: 108; LEFT: 384px; POSITION: absolute; TOP: 752px" runat="server"
					Width="105px" Height="32px" Text="NG"></asp:button>
				<asp:button id="BOK" style="Z-INDEX: 106; LEFT: 496px; POSITION: absolute; TOP: 752px" runat="server"
					Width="105px" Height="32px" Text="OK"></asp:button><INPUT id="DAttachFile" style="Z-INDEX: 118; LEFT: 120px; WIDTH: 424px; POSITION: absolute; TOP: 434px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					type="file" size="51" name="File1" runat="server">
				<asp:textbox id="DDReason" style="Z-INDEX: 117; LEFT: 120px; POSITION: absolute; TOP: 364px"
					runat="server" Width="472px" Height="56px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove"
					TextMode="MultiLine">DReason</asp:textbox>
				<asp:textbox id="DContent" style="Z-INDEX: 116; LEFT: 120px; POSITION: absolute; TOP: 294px"
					runat="server" Width="472px" Height="56px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove"
					TextMode="MultiLine">DContent</asp:textbox>
				<asp:textbox id="DTarget" style="Z-INDEX: 115; LEFT: 120px; POSITION: absolute; TOP: 260px" runat="server"
					Width="472px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DTarget</asp:textbox>
				<asp:textbox id="DOFormSno" style="Z-INDEX: 114; LEFT: 192px; POSITION: absolute; TOP: 226px"
					runat="server" Width="72px" Height="20px" AutoPostBack="True" ReadOnly="True" BackColor="Yellow"
					ForeColor="Blue" BorderStyle="Groove">DOFormSno</asp:textbox>
				<asp:textbox id="DMapNo" style="Z-INDEX: 113; LEFT: 120px; POSITION: absolute; TOP: 192px" runat="server"
					Width="176px" Height="20px" AutoPostBack="True" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DMapNo</asp:textbox>
				<asp:button id="BOMapNo" style="Z-INDEX: 109; LEFT: 296px; POSITION: absolute; TOP: 192px" runat="server"
					Width="20px" Height="20px" Text="原"></asp:button>
				<asp:button id="BMMapNo" style="Z-INDEX: 112; LEFT: 318px; POSITION: absolute; TOP: 192px" runat="server"
					Width="20px" Height="20px" Text="修"></asp:button>
				<asp:dropdownlist id="DLevel" style="Z-INDEX: 110; LEFT: 416px; POSITION: absolute; TOP: 158px" runat="server"
					Width="176px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist>
				<asp:textbox id="DSliderCode" style="Z-INDEX: 107; LEFT: 120px; POSITION: absolute; TOP: 158px"
					runat="server" Width="176px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DSliderCode</asp:textbox>
				<asp:textbox id="DNo" style="Z-INDEX: 105; LEFT: 120px; POSITION: absolute; TOP: 90px" runat="server"
					Width="152px" Height="20px" ReadOnly="True" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DNo</asp:textbox>
				<asp:textbox id="DDate" style="Z-INDEX: 103; LEFT: 416px; POSITION: absolute; TOP: 90px" runat="server"
					Width="152px" Height="20px" ReadOnly="True" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DDate</asp:textbox>
				<asp:button id="BDate" style="Z-INDEX: 104; LEFT: 568px; POSITION: absolute; TOP: 90px" runat="server"
					Width="20px" Height="20px" Text="....."></asp:button>
				<asp:dropdownlist id="DPerson" style="Z-INDEX: 102; LEFT: 416px; POSITION: absolute; TOP: 124px" runat="server"
					Width="176px" Height="20px" AutoPostBack="True" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DDivision" style="Z-INDEX: 101; LEFT: 120px; POSITION: absolute; TOP: 124px"
					runat="server" Width="176px" Height="20px" AutoPostBack="True" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist>
			</FONT>
		</FORM>
	</body>
</HTML>
