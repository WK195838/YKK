<%@ Page Language="vb" AutoEventWireup="false" Codebehind="ModifyDataSheet_01.aspx.vb" Inherits="SPD.ModifyDataSheet_01" aspCompat="True"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>ModifyDataSheet_01</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<style type="text/css">BODY { FONT-SIZE: 9pt; FONT-FAMILY: Verdana, Geneva, Sans-Serif }
	TABLE { FONT-SIZE: 9pt; FONT-FAMILY: Verdana, Geneva, Sans-Serif }
	TR { FONT-SIZE: 9pt; FONT-FAMILY: Verdana, Geneva, Sans-Serif }
	TD { FONT-SIZE: 9pt; FONT-FAMILY: Verdana, Geneva, Sans-Serif }
	UL { FONT-SIZE: 9pt; FONT-FAMILY: Verdana, Geneva, Sans-Serif }
	LI { FONT-SIZE: 9pt; FONT-FAMILY: Verdana, Geneva, Sans-Serif }
	.normal { FONT-SIZE: 9pt; FONT-FAMILY: Verdana, Geneva, Sans-Serif }
	H1 { FONT-WEIGHT: 900; FONT-SIZE: 10.5pt; COLOR: #666666; FONT-FAMILY: Verdana, Geneva, Sans-Serif }
	.small { FONT-SIZE: 7.5pt; FONT-FAMILY: Verdana, Geneva, Sans-Serif }
	.error { COLOR: #ff0033 }
	.required { FONT-WEIGHT: 900; COLOR: #ff0033; FONT-FAMILY: Verdana, Geneva, Sans-Serif }
	.success { FONT-WEIGHT: 900; FONT-SIZE: 11pt; MARGIN: 10px 0px; COLOR: #009933 }
		</style>
		<script language="javascript" type="text/javascript">
			var wPop;
			//--Calendar------------------------------------
			function calendarPicker(strField)
			{
				window.open('DatePicker.aspx?field=' + strField,'calendarPopup','width=250,height=190,resizable=yes');
			}
		    function SendToSpec() {
				wPop.document.SpecForm.DSpec.value=document.Form1.DSpec.value;
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
		<form id="Form1" method="post" encType="multipart/form-data" runat="server">
			<FONT face="新細明體">
				<asp:image id="DModifyDataSheet1" style="Z-INDEX: 104; LEFT: 8px; POSITION: absolute; TOP: 8px"
					runat="server" Height="684px" Width="593px" ImageUrl="Images\ModifyDatatSheet_001.jpg"></asp:image>
				<asp:hyperlink id="LSheet1" style="Z-INDEX: 142; LEFT: 528px; POSITION: absolute; TOP: 224px" runat="server"
					Width="65px" Height="8px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">連結表單</asp:hyperlink><asp:dropdownlist id="DIPerson" style="Z-INDEX: 141; LEFT: 440px; POSITION: absolute; TOP: 664px"
					runat="server" Height="20px" Width="152px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="宋繼明" Selected="True">宋繼明</asp:ListItem>
					<asp:ListItem Value="徐滿霖">徐滿霖</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DFDateTime" style="Z-INDEX: 140; LEFT: 144px; POSITION: absolute; TOP: 664px"
					runat="server" Height="20px" Width="152px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" MaxLength="20">DFDateTime</asp:textbox><asp:textbox id="DIContent" style="Z-INDEX: 139; LEFT: 144px; POSITION: absolute; TOP: 560px"
					runat="server" Height="90px" Width="448px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" TextMode="MultiLine" MaxLength="256">DIContent</asp:textbox><asp:textbox id="DMContent2" style="Z-INDEX: 138; LEFT: 144px; POSITION: absolute; TOP: 432px"
					runat="server" Height="80px" Width="448px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" TextMode="MultiLine" MaxLength="256">DMContent2</asp:textbox><asp:textbox id="DMContent1" style="Z-INDEX: 137; LEFT: 144px; POSITION: absolute; TOP: 352px"
					runat="server" Height="80px" Width="448px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" TextMode="MultiLine" MaxLength="256">DMContent1</asp:textbox><asp:dropdownlist id="DWPerson" style="Z-INDEX: 136; LEFT: 304px; POSITION: absolute; TOP: 188px"
					runat="server" Height="20px" Width="162px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="3">徐滿</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DWDivision" style="Z-INDEX: 135; LEFT: 144px; POSITION: absolute; TOP: 188px"
					runat="server" Height="20px" Width="152px" ForeColor="Blue" BackColor="Yellow" AutoPostBack="True">
					<asp:ListItem Value="3">徐滿</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DMPerson" style="Z-INDEX: 134; LEFT: 304px; POSITION: absolute; TOP: 154px"
					runat="server" Height="20px" Width="162px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="3">徐滿</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DMDivision" style="Z-INDEX: 133; LEFT: 144px; POSITION: absolute; TOP: 154px"
					runat="server" Height="20px" Width="152px" ForeColor="Blue" BackColor="Yellow" AutoPostBack="True">
					<asp:ListItem Value="3">徐滿</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DOPReadyDesc" style="Z-INDEX: 132; LEFT: 248px; POSITION: absolute; TOP: 848px"
					runat="server" Height="20px" Width="96px" ForeColor="Red" BackColor="White" BorderStyle="None">需閱讀工程履歷</asp:textbox><asp:textbox id="DOPReady" style="Z-INDEX: 131; LEFT: 352px; POSITION: absolute; TOP: 848px"
					runat="server" Height="20px" Width="48px" ForeColor="Red" BackColor="GreenYellow" BorderStyle="Groove" ReadOnly="True"></asp:textbox><asp:imagebutton id="BFlow" style="Z-INDEX: 130; LEFT: 8px; POSITION: absolute; TOP: 32px" runat="server"
					Height="16px" Width="16px" ImageUrl="Images\Flow-01.gif"></asp:imagebutton><INPUT id="BSAVE" style="Z-INDEX: 128; LEFT: 256px; WIDTH: 80px; POSITION: absolute; TOP: 992px; HEIGHT: 28px"
					onclick="Button('SAVE');" type="button" value="儲存" name="Button1" runat="server"><INPUT id="BNG2" style="Z-INDEX: 127; LEFT: 344px; WIDTH: 80px; POSITION: absolute; TOP: 992px; HEIGHT: 28px"
					onclick="Button('NG2');" type="button" value="NG2" name="Button1" runat="server"><INPUT id="BNG1" style="Z-INDEX: 126; LEFT: 432px; WIDTH: 80px; POSITION: absolute; TOP: 992px; HEIGHT: 28px"
					onclick="Button('NG1');" type="button" value="NG1" name="Button1" runat="server"><asp:dropdownlist id="DSheet" style="Z-INDEX: 124; LEFT: 144px; POSITION: absolute; TOP: 222px" runat="server"
					Height="20px" Width="152px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DMReasonType" style="Z-INDEX: 123; LEFT: 144px; POSITION: absolute; TOP: 288px"
					runat="server" Height="20px" Width="152px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="Custom" Selected="True">客戶要求</asp:ListItem>
					<asp:ListItem Value="YKK">內部需求</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DComNo" style="Z-INDEX: 122; LEFT: 432px; POSITION: absolute; TOP: 222px" runat="server"
					Height="20px" Width="162px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" MaxLength="30">DComNo</asp:textbox><asp:textbox id="DMReason" style="Z-INDEX: 121; LEFT: 16px; POSITION: absolute; TOP: 320px" runat="server"
					Height="20px" Width="577px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" MaxLength="256">DMReason</asp:textbox><asp:textbox id="DFormSno" style="Z-INDEX: 120; LEFT: 16px; POSITION: absolute; TOP: 992px" runat="server"
					Height="20px" Width="97px" ForeColor="Blue" BackColor="White" BorderStyle="None">單號：123</asp:textbox><asp:textbox id="DAEndTime" style="Z-INDEX: 117; LEFT: 440px; POSITION: absolute; TOP: 816px"
					runat="server" Height="20px" Width="152px" ForeColor="Blue" BackColor="Gold" BorderStyle="Groove" ReadOnly="True" Font-Size="9pt">DAEndTime</asp:textbox><asp:textbox id="DAStartTime" style="Z-INDEX: 116; LEFT: 168px; POSITION: absolute; TOP: 816px"
					runat="server" Height="20px" Width="152px" ForeColor="Blue" BackColor="Gold" BorderStyle="Groove" ReadOnly="True" Font-Size="9pt">DAStartTime</asp:textbox><asp:textbox id="DBEndTime" style="Z-INDEX: 115; LEFT: 440px; POSITION: absolute; TOP: 784px"
					runat="server" Height="20px" Width="152px" ForeColor="Blue" BackColor="Gold" BorderStyle="Groove" ReadOnly="True" Font-Size="9pt">DBEndTime</asp:textbox><asp:textbox id="DBStartTime" style="Z-INDEX: 114; LEFT: 168px; POSITION: absolute; TOP: 784px"
					runat="server" Height="20px" Width="152px" ForeColor="Blue" BackColor="Gold" BorderStyle="Groove" ReadOnly="True" Font-Size="9pt">DBStartTime</asp:textbox><asp:textbox id="DReason" style="Z-INDEX: 113; LEFT: 240px; POSITION: absolute; TOP: 888px" runat="server"
					Height="20px" Width="352px" ForeColor="Blue" BackColor="Gold" BorderStyle="Groove" ReadOnly="True" Visible="False">DReason</asp:textbox><asp:dropdownlist id="DReasonCode" style="Z-INDEX: 112; LEFT: 168px; POSITION: absolute; TOP: 888px"
					runat="server" Height="20px" Width="64px" ForeColor="Blue" BackColor="Gold" Visible="False" AutoPostBack="True">
					<asp:ListItem Value="01">01</asp:ListItem>
					<asp:ListItem Value="02">02</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DReasonDesc" style="Z-INDEX: 111; LEFT: 168px; POSITION: absolute; TOP: 920px"
					runat="server" Height="59px" Width="424px" ForeColor="Blue" BackColor="Gold" BorderStyle="Groove" TextMode="MultiLine"
					Visible="False" MaxLength="256">DReasonDesc</asp:textbox><asp:textbox id="DDecideDesc" style="Z-INDEX: 110; LEFT: 64px; POSITION: absolute; TOP: 704px"
					runat="server" Height="56px" Width="528px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" TextMode="MultiLine" MaxLength="256">DDecideDesc</asp:textbox><asp:hyperlink id="LBefOP" style="Z-INDEX: 109; LEFT: 168px; POSITION: absolute; TOP: 848px" runat="server"
					Height="8px" Width="144px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">工程履歷</asp:hyperlink><asp:textbox id="DNo" style="Z-INDEX: 105; LEFT: 144px; POSITION: absolute; TOP: 86px" runat="server"
					Height="20px" Width="152px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">DNo</asp:textbox></FONT><asp:textbox id="DDate" style="Z-INDEX: 106; LEFT: 432px; POSITION: absolute; TOP: 86px" runat="server"
				Height="20px" Width="136px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DDate</asp:textbox><asp:button id="BDate" style="Z-INDEX: 107; LEFT: 568px; POSITION: absolute; TOP: 86px" runat="server"
				Height="20px" Width="20px" Text="....."></asp:button><asp:dropdownlist id="DStatus" style="Z-INDEX: 108; LEFT: 144px; POSITION: absolute; TOP: 256px" runat="server"
				Height="20px" Width="152px" ForeColor="Blue" BackColor="Yellow">
				<asp:ListItem Value="開發中" Selected="True">開發中</asp:ListItem>
				<asp:ListItem Value="OK完成">OK完成</asp:ListItem>
				<asp:ListItem Value="NG完成">NG完成</asp:ListItem>
				<asp:ListItem Value="取消完成">取消完成</asp:ListItem>
			</asp:dropdownlist><asp:dropdownlist id="DDivision" style="Z-INDEX: 118; LEFT: 144px; POSITION: absolute; TOP: 120px"
				runat="server" Height="20px" Width="152px" ForeColor="Blue" BackColor="Yellow">
				<asp:ListItem Value="3">IS</asp:ListItem>
			</asp:dropdownlist><asp:dropdownlist id="DPerson" style="Z-INDEX: 119; LEFT: 432px; POSITION: absolute; TOP: 120px" runat="server"
				Height="20px" Width="162px" ForeColor="Blue" BackColor="Yellow">
				<asp:ListItem Value="3">徐滿</asp:ListItem>
			</asp:dropdownlist><asp:image id="DMapSheet3" style="Z-INDEX: 103; LEFT: 8px; POSITION: absolute; TOP: 608px"
				runat="server" Height="77px" Width="593px" ImageUrl="Images\MapSheet_003_A.jpg"></asp:image><asp:image id="DMapSheet4" style="Z-INDEX: 102; LEFT: 8px; POSITION: absolute; TOP: 696px"
				runat="server" Height="75px" Width="593px" ImageUrl="Images\MapSheet_004.jpg"></asp:image><asp:image id="DDelivery" style="Z-INDEX: 101; LEFT: 8px; POSITION: absolute; TOP: 768px" runat="server"
				Height="110px" Width="593px" ImageUrl="Images\Sheet_Delivery.jpg"></asp:image><asp:image id="DDelay" style="Z-INDEX: 100; LEFT: 8px; POSITION: absolute; TOP: 880px" runat="server"
				Height="107px" Width="593px" ImageUrl="Images\Sheet_Delay.jpg" Visible="False"></asp:image><INPUT id="BOK" style="Z-INDEX: 125; LEFT: 520px; WIDTH: 80px; POSITION: absolute; TOP: 992px; HEIGHT: 28px"
				onclick="Button('OK');" type="button" value="OK" name="BOK" runat="server">
			<asp:imagebutton id="BPrint" style="Z-INDEX: 129; LEFT: 8px; POSITION: absolute; TOP: 8px" runat="server"
				Height="16px" Width="16px" ImageUrl="Images\Print.gif"></asp:imagebutton></form>
	</body>
</HTML>
