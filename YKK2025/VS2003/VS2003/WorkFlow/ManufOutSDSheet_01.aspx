<%@ Page Language="vb" AutoEventWireup="false" Codebehind="ManufOutSDSheet_01.aspx.vb" Inherits="SPD.ManufOutSDSheet_01" aspCompat="True"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>外注拉頭細目</title>
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
	<BODY ms_positioning="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體">
				<asp:image id="DManufOutSDSheet" style="Z-INDEX: 100; LEFT: 8px; POSITION: absolute; TOP: 8px"
					runat="server" ImageUrl="Images\SliderDetailSheet_002.jpg" Width="593px" Height="314px"></asp:image>
				<asp:button id="BImport" style="Z-INDEX: 138; LEFT: 272px; POSITION: absolute; TOP: 86px" runat="server"
					Height="20px" Width="24px" Text="進"></asp:button>
				<asp:hyperlink id="LSliderFile" style="Z-INDEX: 137; LEFT: 120px; POSITION: absolute; TOP: 264px"
					runat="server" Height="8px" Width="80px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">拉頭細目表</asp:hyperlink>
				<asp:textbox id="DOPReadyDesc" style="Z-INDEX: 136; LEFT: 240px; POSITION: absolute; TOP: 472px"
					runat="server" Height="20px" Width="104px" BackColor="White" ForeColor="Red" BorderStyle="None">需閱讀工程履歷</asp:textbox>
				<asp:textbox id="DOPReady" style="Z-INDEX: 135; LEFT: 344px; POSITION: absolute; TOP: 472px"
					runat="server" Height="20px" Width="48px" BackColor="GreenYellow" ForeColor="Red" BorderStyle="Groove"
					ReadOnly="True"></asp:textbox>
				<asp:textbox id="DContent" style="Z-INDEX: 134; LEFT: 120px; POSITION: absolute; TOP: 188px"
					runat="server" Height="56px" Width="472px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow"
					TextMode="MultiLine">DContent</asp:textbox>
				<asp:imagebutton id="BFlow" style="Z-INDEX: 133; LEFT: 8px; POSITION: absolute; TOP: 32px" runat="server"
					Height="16px" Width="16px" ImageUrl="Images\Flow-01.gif"></asp:imagebutton>
				<asp:imagebutton id="BPrint" style="Z-INDEX: 128; LEFT: 8px; POSITION: absolute; TOP: 8px" runat="server"
					Height="16px" Width="16px" ImageUrl="Images\Print.gif"></asp:imagebutton><asp:textbox id="DFormSno" style="Z-INDEX: 115; LEFT: 8px; POSITION: absolute; TOP: 608px" runat="server"
					Width="97px" Height="20px" BackColor="White" ForeColor="Blue" BorderStyle="None">單號：123</asp:textbox><asp:dropdownlist id="DReasonCode" style="Z-INDEX: 127; LEFT: 168px; POSITION: absolute; TOP: 512px"
					runat="server" Width="64px" Height="20px" BackColor="Gold" ForeColor="Blue" AutoPostBack="True" Visible="False">
					<asp:ListItem Value="01">01</asp:ListItem>
					<asp:ListItem Value="02">02</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DReason" style="Z-INDEX: 125; LEFT: 240px; POSITION: absolute; TOP: 512px" runat="server"
					Width="352px" Height="20px" ReadOnly="True" BackColor="Gold" ForeColor="Blue" BorderStyle="Groove" Visible="False">DReason</asp:textbox><asp:textbox id="DReasonDesc" style="Z-INDEX: 126; LEFT: 168px; POSITION: absolute; TOP: 544px"
					runat="server" Width="424px" Height="59px" BackColor="Gold" ForeColor="Blue" BorderStyle="Groove" TextMode="MultiLine" Visible="False">DReasonDesc</asp:textbox><asp:textbox id="DBStartTime" style="Z-INDEX: 124; LEFT: 168px; POSITION: absolute; TOP: 404px"
					runat="server" Width="152px" Height="20px" ReadOnly="True" BackColor="Gold" ForeColor="Blue" BorderStyle="Groove" Font-Size="9pt">DBStartTime</asp:textbox><asp:textbox id="DAStartTime" style="Z-INDEX: 119; LEFT: 168px; POSITION: absolute; TOP: 438px"
					runat="server" Width="152px" Height="20px" ReadOnly="True" BackColor="Gold" ForeColor="Blue" BorderStyle="Groove" Font-Size="9pt">DAStartTime</asp:textbox><asp:textbox id="DAEndTime" style="Z-INDEX: 122; LEFT: 440px; POSITION: absolute; TOP: 438px"
					runat="server" Width="152px" Height="20px" ReadOnly="True" BackColor="Gold" ForeColor="Blue" BorderStyle="Groove" Font-Size="9pt">DAEndTime</asp:textbox><asp:textbox id="DBEndTime" style="Z-INDEX: 118; LEFT: 440px; POSITION: absolute; TOP: 404px"
					runat="server" Width="152px" Height="20px" ReadOnly="True" BackColor="Gold" ForeColor="Blue" BorderStyle="Groove" Font-Size="9pt">DBEndTime</asp:textbox><asp:hyperlink id="LBefOP" style="Z-INDEX: 123; LEFT: 168px; POSITION: absolute; TOP: 476px" runat="server"
					Width="144px" Height="8px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">工程履歷</asp:hyperlink><asp:textbox id="DDecideDesc" style="Z-INDEX: 121; LEFT: 56px; POSITION: absolute; TOP: 328px"
					runat="server" Width="536px" Height="56px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" TextMode="MultiLine">DDecideDesc</asp:textbox><asp:image id="DDelay" style="Z-INDEX: 120; LEFT: 8px; POSITION: absolute; TOP: 504px" runat="server"
					ImageUrl="Images\Sheet_Delay.jpg" Width="593px" Height="107px" Visible="False"></asp:image><asp:image id="DDelivery" style="Z-INDEX: 116; LEFT: 8px; POSITION: absolute; TOP: 392px" runat="server"
					ImageUrl="Images\Sheet_Delivery.jpg" Width="593px" Height="110px"></asp:image><asp:image id="DDescSheet" style="Z-INDEX: 114; LEFT: 8px; POSITION: absolute; TOP: 320px"
					runat="server" ImageUrl="Images\MapSheet_004.jpg" Width="593px" Height="75px"></asp:image><asp:textbox id="DOFormNo" style="Z-INDEX: 113; LEFT: 416px; POSITION: absolute; TOP: 154px"
					runat="server" Width="56px" Height="20px" ReadOnly="True" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" AutoPostBack="True">DOFormNo</asp:textbox><asp:textbox id="DOFormSno" style="Z-INDEX: 111; LEFT: 472px; POSITION: absolute; TOP: 154px"
					runat="server" Width="56px" Height="20px" ReadOnly="True" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" AutoPostBack="True">DOFormSno</asp:textbox><asp:hyperlink id="LOFormNo" style="Z-INDEX: 112; LEFT: 536px; POSITION: absolute; TOP: 156px"
					runat="server" Width="48px" Height="8px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">原委託</asp:hyperlink><asp:hyperlink id="LAttachFile" style="Z-INDEX: 110; LEFT: 120px; POSITION: absolute; TOP: 294px"
					runat="server" Width="32px" Height="8px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">附件</asp:hyperlink><INPUT id="DAttachFile" style="Z-INDEX: 109; LEFT: 120px; WIDTH: 368px; POSITION: absolute; TOP: 292px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					type="file" size="42" name="File1" runat="server"><INPUT id="DSliderFile" style="Z-INDEX: 108; LEFT: 120px; WIDTH: 368px; POSITION: absolute; TOP: 258px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					type="file" size="42" name="File1" runat="server"><asp:textbox id="DNo" style="Z-INDEX: 107; LEFT: 120px; POSITION: absolute; TOP: 86px" runat="server"
					Width="112px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DNo</asp:textbox><asp:button id="BOut" style="Z-INDEX: 106; LEFT: 240px; POSITION: absolute; TOP: 86px" runat="server"
					Width="24px" Height="20px" Text="外"></asp:button><asp:textbox id="DDate" style="Z-INDEX: 103; LEFT: 416px; POSITION: absolute; TOP: 86px" runat="server"
					Width="152px" Height="20px" ReadOnly="True" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DDate</asp:textbox><asp:button id="BDate" style="Z-INDEX: 104; LEFT: 568px; POSITION: absolute; TOP: 86px" runat="server"
					Width="20px" Height="20px" Text="....."></asp:button><asp:dropdownlist id="DDivision" style="Z-INDEX: 101; LEFT: 120px; POSITION: absolute; TOP: 120px"
					runat="server" Width="176px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DPerson" style="Z-INDEX: 102; LEFT: 416px; POSITION: absolute; TOP: 120px" runat="server"
					Width="176px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DSliderGRCode" style="Z-INDEX: 105; LEFT: 120px; POSITION: absolute; TOP: 154px"
					runat="server" Width="176px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DSliderGRCode</asp:textbox></FONT><INPUT id="BSAVE" style="Z-INDEX: 129; LEFT: 256px; WIDTH: 80px; POSITION: absolute; TOP: 616px; HEIGHT: 28px"
				onclick="Button('SAVE');" type="button" value="儲存" name="BSAVE" runat="server"><INPUT id="BNG2" style="Z-INDEX: 130; LEFT: 344px; WIDTH: 80px; POSITION: absolute; TOP: 616px; HEIGHT: 28px"
				onclick="Button('NG2');" type="button" value="NG2" name="Button1" runat="server"><INPUT id="BNG1" style="Z-INDEX: 131; LEFT: 432px; WIDTH: 80px; POSITION: absolute; TOP: 616px; HEIGHT: 28px"
				onclick="Button('NG1');" type="button" value="NG1" name="Button1" runat="server"><INPUT id="BOK" style="Z-INDEX: 132; LEFT: 520px; WIDTH: 80px; POSITION: absolute; TOP: 616px; HEIGHT: 28px"
				onclick="Button('OK');" type="button" value="OK" name="BOK" runat="server"></form>
	</BODY>
</HTML>
