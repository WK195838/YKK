<%@ Page Language="vb" AspCompat="true" AutoEventWireup="false" Codebehind="MapModSheet_01.aspx.vb" Inherits="SPD.MapSheetMod_01"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>圖面修改委託書</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<script language="javascript" type="text/javascript">
			function calendarPicker(strField)
			{
				window.open('DatePicker.aspx?field=' + strField,'calendarPopup','width=250,height=190,resizable=yes');
			}
			function MapModPicker(strField)
			{
				window.open('MapModPicker.aspx?field=' + strField,'MapPopup','width=168,height=360,resizable=yes');
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
				<asp:image id="DMapSheet1" style="Z-INDEX: 100; LEFT: 8px; POSITION: absolute; TOP: 8px" runat="server"
					Height="554px" Width="593px" ImageUrl="Images\MapSheetMod_002_B.jpg"></asp:image><asp:hyperlink id="LPdfFile" style="Z-INDEX: 357; LEFT: 440px; POSITION: absolute; TOP: 608px"
					runat="server" Height="16px" Width="41px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">圖檔</asp:hyperlink><INPUT id="DPdfFile" style="Z-INDEX: 156; LEFT: 440px; WIDTH: 152px; POSITION: absolute; TOP: 606px; HEIGHT: 18px; BACKGROUND-COLOR: #ffff00"
					type="file" size="6" name="LPdfFile" runat="server">
				<asp:dropdownlist id="DSuppiler" style="Z-INDEX: 356; LEFT: 304px; POSITION: absolute; TOP: 528px"
					runat="server" Height="20px" Width="146px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DOPReadyDesc" style="Z-INDEX: 154; LEFT: 248px; POSITION: absolute; TOP: 792px"
					runat="server" Height="20px" Width="96px" ForeColor="Red" BackColor="White" BorderStyle="None">需閱讀工程履歷</asp:textbox><asp:textbox id="DOPReady" style="Z-INDEX: 153; LEFT: 352px; POSITION: absolute; TOP: 792px"
					runat="server" Height="20px" Width="48px" ForeColor="Red" BackColor="GreenYellow" BorderStyle="Groove" ReadOnly="True"></asp:textbox><asp:dropdownlist id="DCpsc" style="Z-INDEX: 151; LEFT: 540px; POSITION: absolute; TOP: 528px" runat="server"
					Height="20px" Width="56px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DBuyer" style="Z-INDEX: 150; LEFT: 168px; POSITION: absolute; TOP: 212px" runat="server"
					Height="20px" Width="128px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">DBuyer</asp:textbox><asp:dropdownlist id="DManufType" style="Z-INDEX: 149; LEFT: 168px; POSITION: absolute; TOP: 528px"
					runat="server" Height="20px" Width="128px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:imagebutton id="BFlow" style="Z-INDEX: 148; LEFT: 8px; POSITION: absolute; TOP: 32px" runat="server"
					Height="16px" Width="16px" ImageUrl="Images\Flow-01.gif"></asp:imagebutton><asp:hyperlink id="LBefMap" style="Z-INDEX: 147; LEFT: 456px; POSITION: absolute; TOP: 172px" runat="server"
					Height="8px" Width="50px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">前圖號</asp:hyperlink><asp:hyperlink id="LOriMap" style="Z-INDEX: 146; LEFT: 168px; POSITION: absolute; TOP: 172px" runat="server"
					Height="8px" Width="50px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">原圖號</asp:hyperlink><asp:dropdownlist id="DLevel" style="Z-INDEX: 145; LEFT: 540px; POSITION: absolute; TOP: 570px" runat="server"
					Height="20px" Width="56px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:hyperlink id="LRefAttach" style="Z-INDEX: 143; LEFT: 168px; POSITION: absolute; TOP: 496px"
					runat="server" Height="8px" Width="40px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">附件</asp:hyperlink><asp:imagebutton id="BPrint" style="Z-INDEX: 142; LEFT: 8px; POSITION: absolute; TOP: 8px" runat="server"
					Height="16px" Width="16px" ImageUrl="Images\Print.gif"></asp:imagebutton><asp:dropdownlist id="DSample" style="Z-INDEX: 137; LEFT: 539px; POSITION: absolute; TOP: 358px" runat="server"
					Height="20px" Width="55px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DModContent" style="Z-INDEX: 136; LEFT: 64px; POSITION: absolute; TOP: 392px"
					runat="server" Height="92px" Width="528px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" TextMode="MultiLine">DModContent</asp:textbox><asp:dropdownlist id="DModReasonCode" style="Z-INDEX: 135; LEFT: 168px; POSITION: absolute; TOP: 290px"
					runat="server" Height="20px" Width="424px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="01">01</asp:ListItem>
					<asp:ListItem Value="02">02</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DModReasonDesc" style="Z-INDEX: 109; LEFT: 168px; POSITION: absolute; TOP: 324px"
					runat="server" Height="20px" Width="424px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">DModReasonDesc</asp:textbox><asp:textbox id="DBefFormSno" style="Z-INDEX: 134; LEFT: 520px; POSITION: absolute; TOP: 168px"
					runat="server" Height="20px" Width="48px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">DBefFormSno</asp:textbox><asp:textbox id="DBefFormNo" style="Z-INDEX: 133; LEFT: 456px; POSITION: absolute; TOP: 168px"
					runat="server" Height="20px" Width="64px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">DBefFormNo</asp:textbox><asp:textbox id="DOriFormSno" style="Z-INDEX: 132; LEFT: 224px; POSITION: absolute; TOP: 168px"
					runat="server" Height="20px" Width="48px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">DOriFormSno</asp:textbox><asp:textbox id="DOriFormNo" style="Z-INDEX: 131; LEFT: 168px; POSITION: absolute; TOP: 168px"
					runat="server" Height="20px" Width="56px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">DOriFormNo</asp:textbox><asp:button id="BBefMapNo" style="Z-INDEX: 130; LEFT: 568px; POSITION: absolute; TOP: 134px"
					runat="server" Height="20px" Width="20px" Text="....."></asp:button><asp:textbox id="DBefMapNo" style="Z-INDEX: 129; LEFT: 456px; POSITION: absolute; TOP: 134px"
					runat="server" Height="20px" Width="112px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DBefMapNo</asp:textbox><asp:textbox id="DOriMapNo" style="Z-INDEX: 128; LEFT: 168px; POSITION: absolute; TOP: 134px"
					runat="server" Height="20px" Width="112px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True" AutoPostBack="True">DOriMapNo</asp:textbox><asp:button id="BOriMapNo" style="Z-INDEX: 123; LEFT: 280px; POSITION: absolute; TOP: 134px"
					runat="server" Height="20px" Width="20px" Text="....."></asp:button><INPUT id="DRefAttach" style="Z-INDEX: 127; LEFT: 168px; WIDTH: 416px; POSITION: absolute; TOP: 496px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					type="file" size="50" name="File1" runat="server"><asp:dropdownlist id="DDivision" style="Z-INDEX: 126; LEFT: 168px; POSITION: absolute; TOP: 246px"
					runat="server" Height="20px" Width="128px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="3">IS</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DPerson" style="Z-INDEX: 125; LEFT: 456px; POSITION: absolute; TOP: 246px" runat="server"
					Height="20px" Width="136px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="3">徐滿</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DSellVendor" style="Z-INDEX: 107; LEFT: 456px; POSITION: absolute; TOP: 212px"
					runat="server" Height="20px" Width="136px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">DSellVendor</asp:textbox><asp:textbox id="DNo" style="Z-INDEX: 124; LEFT: 168px; POSITION: absolute; TOP: 90px" runat="server"
					Height="20px" Width="128px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">DNo</asp:textbox><asp:textbox id="DDate" style="Z-INDEX: 105; LEFT: 456px; POSITION: absolute; TOP: 90px" runat="server"
					Height="20px" Width="112px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DDate</asp:textbox><asp:button id="BDate" style="Z-INDEX: 117; LEFT: 568px; POSITION: absolute; TOP: 90px" runat="server"
					Height="20px" Width="20px" Text="....."></asp:button><asp:dropdownlist id="DMakeMap" style="Z-INDEX: 122; LEFT: 396px; POSITION: absolute; TOP: 570px"
					runat="server" Height="20px" Width="72px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><INPUT id="DMapFile" style="Z-INDEX: 118; LEFT: 168px; WIDTH: 152px; POSITION: absolute; TOP: 608px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					type="file" size="6" name="File1" runat="server">
				<asp:hyperlink id="LMapFile" style="Z-INDEX: 119; LEFT: 168px; POSITION: absolute; TOP: 609px"
					runat="server" Height="8px" Width="40px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">圖檔</asp:hyperlink><asp:textbox id="DMapNo" style="Z-INDEX: 121; LEFT: 168px; POSITION: absolute; TOP: 570px" runat="server"
					Height="20px" Width="152px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">DMapNo</asp:textbox><asp:textbox id="DDecideDesc" style="Z-INDEX: 120; LEFT: 64px; POSITION: absolute; TOP: 648px"
					runat="server" Height="56px" Width="528px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" TextMode="MultiLine">DDecideDesc</asp:textbox><asp:hyperlink id="LBefOP" style="Z-INDEX: 116; LEFT: 176px; POSITION: absolute; TOP: 794px" runat="server"
					Height="8px" Width="144px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">工程履歷</asp:hyperlink><asp:textbox id="DAStartTime" style="Z-INDEX: 114; LEFT: 168px; POSITION: absolute; TOP: 757px"
					runat="server" Height="20px" Width="152px" ForeColor="Blue" BackColor="Gold" BorderStyle="Groove" ReadOnly="True" Font-Size="9pt">DAStartTime</asp:textbox><asp:textbox id="DBStartTime" style="Z-INDEX: 112; LEFT: 168px; POSITION: absolute; TOP: 724px"
					runat="server" Height="20px" Width="152px" ForeColor="Blue" BackColor="Gold" BorderStyle="Groove" ReadOnly="True" Font-Size="9pt">DBStartTime</asp:textbox><asp:textbox id="DBEndTime" style="Z-INDEX: 113; LEFT: 440px; POSITION: absolute; TOP: 724px"
					runat="server" Height="20px" Width="152px" ForeColor="Blue" BackColor="Gold" BorderStyle="Groove" ReadOnly="True" Font-Size="9pt">DBEndTime</asp:textbox><asp:textbox id="DAEndTime" style="Z-INDEX: 115; LEFT: 440px; POSITION: absolute; TOP: 757px"
					runat="server" Height="20px" Width="152px" ForeColor="Blue" BackColor="Gold" BorderStyle="Groove" ReadOnly="True" Font-Size="9pt">DAEndTime</asp:textbox><asp:dropdownlist id="DReasonCode" style="Z-INDEX: 111; LEFT: 168px; POSITION: absolute; TOP: 834px"
					runat="server" Height="20px" Width="64px" ForeColor="Blue" BackColor="Gold" AutoPostBack="True" Visible="False">
					<asp:ListItem Value="01">01</asp:ListItem>
					<asp:ListItem Value="02">02</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DReason" style="Z-INDEX: 110; LEFT: 240px; POSITION: absolute; TOP: 834px" runat="server"
					Height="20px" Width="352px" ForeColor="Blue" BackColor="Gold" BorderStyle="Groove" ReadOnly="True" Visible="False">DReason</asp:textbox><asp:textbox id="DReasonDesc" style="Z-INDEX: 108; LEFT: 168px; POSITION: absolute; TOP: 865px"
					runat="server" Height="59px" Width="424px" ForeColor="Blue" BackColor="Gold" BorderStyle="Groove" TextMode="MultiLine" Visible="False">DReasonDesc</asp:textbox><asp:textbox id="DFormSno" style="Z-INDEX: 106; LEFT: 16px; POSITION: absolute; TOP: 936px" runat="server"
					Height="20px" Width="97px" ForeColor="Blue" BackColor="White" BorderStyle="None">單號：123</asp:textbox><asp:image id="DDelay" style="Z-INDEX: 104; LEFT: 8px; POSITION: absolute; TOP: 824px" runat="server"
					Height="107px" Width="593px" ImageUrl="Images\Sheet_Delay.jpg" Visible="False"></asp:image><asp:image id="DDelivery" style="Z-INDEX: 103; LEFT: 8px; POSITION: absolute; TOP: 712px" runat="server"
					Height="110px" Width="593px" ImageUrl="Images\Sheet_Delivery.jpg"></asp:image><asp:image id="DMapSheet4" style="Z-INDEX: 102; LEFT: 8px; POSITION: absolute; TOP: 640px"
					runat="server" Height="75px" Width="593px" ImageUrl="Images\MapSheet_004.jpg"></asp:image><asp:image id="DMapSheet3" style="Z-INDEX: 101; LEFT: 8px; POSITION: absolute; TOP: 560px"
					runat="server" Height="77px" Width="593px" ImageUrl="Images\MapSheet_003_A1.JPG"></asp:image></FONT><INPUT id="BSAVE" style="Z-INDEX: 138; LEFT: 256px; WIDTH: 80px; POSITION: absolute; TOP: 936px; HEIGHT: 28px"
				onclick="Button('SAVE');" type="button" value="儲存" name="Button1" runat="server"><INPUT id="BNG2" style="Z-INDEX: 139; LEFT: 344px; WIDTH: 80px; POSITION: absolute; TOP: 936px; HEIGHT: 28px"
				onclick="Button('NG2');" type="button" value="NG2" name="Button1" runat="server"><INPUT id="BNG1" style="Z-INDEX: 140; LEFT: 432px; WIDTH: 80px; POSITION: absolute; TOP: 936px; HEIGHT: 28px"
				onclick="Button('NG1');" type="button" value="NG1" name="Button1" runat="server"><INPUT id="BOK" style="Z-INDEX: 141; LEFT: 520px; WIDTH: 80px; POSITION: absolute; TOP: 936px; HEIGHT: 28px"
				onclick="Button('OK');" type="button" value="OK" name="Button2" runat="server"></form>
	</body>
</HTML>
