<%@ Page Language="vb" AspCompat="true" AutoEventWireup="false" Codebehind="MapSheet_01.aspx.vb" Inherits="SPD.MapSheet_01"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>製圖委託書</title>
		<meta content="True" name="vs_showGrid">
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
			//--Spec------------------------------------
			function SpecPicker(strField, Fun)
			{
				wPop=window.open('SpecPicker.aspx?field=' + strField + '&fun=' + Fun, 'SpecPopup','width=330,height=250,resizable=yes');
				if (document.Form1.DSpec.value != "") {
					setTimeout("SendToSpec()",200);
				}
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
				<asp:image id="DMapSheet1" style="Z-INDEX: 104; LEFT: 8px; POSITION: absolute; TOP: 8px" runat="server"
					ImageUrl="Images\MapSheet_001_D.jpg" Width="593px" Height="597px"></asp:image><asp:dropdownlist id="DMakeMap" style="Z-INDEX: 158; LEFT: 400px; POSITION: absolute; TOP: 618px"
					runat="server" Width="64px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:hyperlink id="LPdfFile" style="Z-INDEX: 157; LEFT: 440px; POSITION: absolute; TOP: 660px"
					runat="server" Width="48px" Height="6px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">圖檔</asp:hyperlink><INPUT id="DPdfFile" style="Z-INDEX: 156; LEFT: 440px; WIDTH: 152px; POSITION: absolute; TOP: 656px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					type="file" size="6" name="LPdfFile" runat="server"><asp:dropdownlist id="DSuppiler" style="Z-INDEX: 155; LEFT: 304px; POSITION: absolute; TOP: 438px"
					runat="server" Width="146px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DOPReadyDesc" style="Z-INDEX: 154; LEFT: 248px; POSITION: absolute; TOP: 840px"
					runat="server" Width="96px" Height="20px" BackColor="White" ForeColor="Red" BorderStyle="None">需閱讀工程履歷</asp:textbox><asp:textbox id="DOPReady" style="Z-INDEX: 153; LEFT: 352px; POSITION: absolute; TOP: 840px"
					runat="server" Width="48px" Height="20px" BackColor="GreenYellow" ForeColor="Red" BorderStyle="Groove" ReadOnly="True"></asp:textbox><asp:dropdownlist id="DBuyer" style="Z-INDEX: 152; LEFT: 168px; POSITION: absolute; TOP: 132px" runat="server"
					Width="128px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">IS</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DManufType" style="Z-INDEX: 151; LEFT: 168px; POSITION: absolute; TOP: 438px"
					runat="server" Width="128px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="內製">內製</asp:ListItem>
					<asp:ListItem Value="外注">外注</asp:ListItem>
				</asp:dropdownlist><asp:imagebutton id="BFlow" style="Z-INDEX: 150; LEFT: 8px; POSITION: absolute; TOP: 32px" runat="server"
					ImageUrl="Images\Flow-01.gif" Width="16px" Height="16px"></asp:imagebutton><INPUT id="BSAVE" style="Z-INDEX: 148; LEFT: 256px; WIDTH: 80px; POSITION: absolute; TOP: 984px; HEIGHT: 28px"
					onclick="Button('SAVE');" type="button" value="儲存" name="Button1" runat="server"><INPUT id="BNG2" style="Z-INDEX: 147; LEFT: 344px; WIDTH: 80px; POSITION: absolute; TOP: 984px; HEIGHT: 28px"
					onclick="Button('NG2');" type="button" value="NG2" name="Button1" runat="server"><INPUT id="BNG1" style="Z-INDEX: 146; LEFT: 432px; WIDTH: 80px; POSITION: absolute; TOP: 984px; HEIGHT: 28px"
					onclick="Button('NG1');" type="button" value="NG1" name="Button1" runat="server"><asp:dropdownlist id="DCPSC" style="Z-INDEX: 144; LEFT: 168px; POSITION: absolute; TOP: 234px" runat="server"
					Width="128px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DSample" style="Z-INDEX: 143; LEFT: 542px; POSITION: absolute; TOP: 336px" runat="server"
					Width="51px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:button id="BSpec" style="Z-INDEX: 142; LEFT: 512px; POSITION: absolute; TOP: 336px" runat="server"
					Width="20px" Height="20px" Text="....."></asp:button><asp:textbox id="DSpec" style="Z-INDEX: 141; LEFT: 64px; POSITION: absolute; TOP: 336px" runat="server"
					Width="448px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True">DSpec</asp:textbox><asp:dropdownlist id="DLevel" style="Z-INDEX: 140; LEFT: 540px; POSITION: absolute; TOP: 618px" runat="server"
					Width="56px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DMaterialDetail_1" style="Z-INDEX: 139; LEFT: 168px; POSITION: absolute; TOP: 506px"
					runat="server" Width="424px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DMaterialDetail_1</asp:textbox><asp:dropdownlist id="Dhalffinish" style="Z-INDEX: 138; LEFT: 456px; POSITION: absolute; TOP: 404px"
					runat="server" Width="136px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DFormSno" style="Z-INDEX: 137; LEFT: 16px; POSITION: absolute; TOP: 984px" runat="server"
					Width="97px" Height="20px" BackColor="White" ForeColor="Blue" BorderStyle="None">單號：123</asp:textbox><asp:textbox id="DAEndTime" style="Z-INDEX: 134; LEFT: 440px; POSITION: absolute; TOP: 804px"
					runat="server" Width="152px" Height="20px" BackColor="Gold" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True" Font-Size="9pt">DAEndTime</asp:textbox><asp:textbox id="DAStartTime" style="Z-INDEX: 133; LEFT: 168px; POSITION: absolute; TOP: 804px"
					runat="server" Width="152px" Height="20px" BackColor="Gold" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True" Font-Size="9pt">DAStartTime</asp:textbox><asp:textbox id="DBEndTime" style="Z-INDEX: 132; LEFT: 440px; POSITION: absolute; TOP: 772px"
					runat="server" Width="152px" Height="20px" BackColor="Gold" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True" Font-Size="9pt">DBEndTime</asp:textbox><asp:textbox id="DBStartTime" style="Z-INDEX: 131; LEFT: 168px; POSITION: absolute; TOP: 772px"
					runat="server" Width="152px" Height="20px" BackColor="Gold" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True" Font-Size="9pt">DBStartTime</asp:textbox><asp:textbox id="DReason" style="Z-INDEX: 130; LEFT: 240px; POSITION: absolute; TOP: 882px" runat="server"
					Width="352px" Height="20px" BackColor="Gold" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True" Visible="False">DReason</asp:textbox><asp:dropdownlist id="DReasonCode" style="Z-INDEX: 129; LEFT: 168px; POSITION: absolute; TOP: 882px"
					runat="server" Width="64px" Height="20px" BackColor="Gold" ForeColor="Blue" Visible="False" AutoPostBack="True">
					<asp:ListItem Value="01">01</asp:ListItem>
					<asp:ListItem Value="02">02</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DReasonDesc" style="Z-INDEX: 128; LEFT: 168px; POSITION: absolute; TOP: 912px"
					runat="server" Width="424px" Height="59px" BackColor="Gold" ForeColor="Blue" BorderStyle="Groove" Visible="False"
					TextMode="MultiLine">DReasonDesc</asp:textbox><asp:textbox id="DDecideDesc" style="Z-INDEX: 127; LEFT: 64px; POSITION: absolute; TOP: 696px"
					runat="server" Width="528px" Height="56px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" TextMode="MultiLine">DDecideDesc</asp:textbox>
				<asp:hyperlink id="LMapFile" style="Z-INDEX: 126; LEFT: 168px; POSITION: absolute; TOP: 656px"
					runat="server" Width="40px" Height="8px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">圖檔</asp:hyperlink><asp:hyperlink id="LRefMapFile" style="Z-INDEX: 125; LEFT: 168px; POSITION: absolute; TOP: 542px"
					runat="server" Width="40px" Height="8px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">草圖</asp:hyperlink><asp:hyperlink id="LBefOP" style="Z-INDEX: 124; LEFT: 168px; POSITION: absolute; TOP: 842px" runat="server"
					Width="144px" Height="8px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">工程履歷</asp:hyperlink><INPUT id="DMapFile" style="Z-INDEX: 123; LEFT: 168px; WIDTH: 152px; POSITION: absolute; TOP: 654px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					type="file" size="6" name="File1" runat="server">
				<asp:textbox id="DMapNo" style="Z-INDEX: 120; LEFT: 168px; POSITION: absolute; TOP: 618px" runat="server"
					Width="152px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DMapNo</asp:textbox><INPUT id="DRefMapFile" style="Z-INDEX: 114; LEFT: 168px; WIDTH: 424px; POSITION: absolute; TOP: 542px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					type="file" size="50" name="File1" runat="server"><asp:textbox id="DSurface" style="Z-INDEX: 110; LEFT: 456px; POSITION: absolute; TOP: 370px"
					runat="server" Width="136px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DSurface</asp:textbox><asp:textbox id="DNo" style="Z-INDEX: 105; LEFT: 168px; POSITION: absolute; TOP: 88px" runat="server"
					Width="128px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DNo</asp:textbox></FONT><asp:textbox id="DDate" style="Z-INDEX: 106; LEFT: 456px; POSITION: absolute; TOP: 88px" runat="server"
				Width="112px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True">DDate</asp:textbox><asp:textbox id="DSellVendor" style="Z-INDEX: 107; LEFT: 456px; POSITION: absolute; TOP: 132px"
				runat="server" Width="136px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DSellVendor</asp:textbox><asp:textbox id="DBackground" style="Z-INDEX: 108; LEFT: 168px; POSITION: absolute; TOP: 200px"
				runat="server" Width="424px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DBackground</asp:textbox><asp:textbox id="DCramper" style="Z-INDEX: 109; LEFT: 168px; POSITION: absolute; TOP: 370px"
				runat="server" Width="128px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DCramper</asp:textbox><asp:dropdownlist id="DFrontBack" style="Z-INDEX: 111; LEFT: 168px; POSITION: absolute; TOP: 404px"
				runat="server" Width="128px" Height="30px" BackColor="Yellow" ForeColor="Blue">
				<asp:ListItem Value="3">相同</asp:ListItem>
			</asp:dropdownlist><asp:dropdownlist id="DMaterial" style="Z-INDEX: 112; LEFT: 168px; POSITION: absolute; TOP: 472px"
				runat="server" Width="128px" Height="20px" BackColor="Yellow" ForeColor="Blue" AutoPostBack="True">
				<asp:ListItem Value="3">3</asp:ListItem>
			</asp:dropdownlist><asp:dropdownlist id="DMaterialDetail" style="Z-INDEX: 113; LEFT: 296px; POSITION: absolute; TOP: 472px"
				runat="server" Width="296px" Height="20px" BackColor="Yellow" ForeColor="Blue">
				<asp:ListItem Value="3">CF</asp:ListItem>
			</asp:dropdownlist><asp:textbox id="DDescription" style="Z-INDEX: 115; LEFT: 168px; POSITION: absolute; TOP: 574px"
				runat="server" Width="424px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DDescription</asp:textbox><asp:button id="BDate" style="Z-INDEX: 116; LEFT: 568px; POSITION: absolute; TOP: 88px" runat="server"
				Width="20px" Height="20px" Text="....."></asp:button><asp:textbox id="DMapReqDelDate" style="Z-INDEX: 117; LEFT: 456px; POSITION: absolute; TOP: 234px"
				runat="server" Width="112px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True">DMapReqDelDate</asp:textbox><asp:button id="BMapReqDelDate" style="Z-INDEX: 118; LEFT: 568px; POSITION: absolute; TOP: 234px"
				runat="server" Width="20px" Height="20px" Text="....."></asp:button><asp:dropdownlist id="DLight" style="Z-INDEX: 119; LEFT: 168px; POSITION: absolute; TOP: 268px" runat="server"
				Width="128px" Height="20px" BackColor="Yellow" ForeColor="Blue">
				<asp:ListItem Value="Y">Yes</asp:ListItem>
				<asp:ListItem Value="N">No</asp:ListItem>
			</asp:dropdownlist><asp:dropdownlist id="DDivision" style="Z-INDEX: 135; LEFT: 168px; POSITION: absolute; TOP: 166px"
				runat="server" Width="128px" Height="20px" BackColor="Yellow" ForeColor="Blue">
				<asp:ListItem Value="3">IS</asp:ListItem>
			</asp:dropdownlist><asp:dropdownlist id="DPerson" style="Z-INDEX: 136; LEFT: 456px; POSITION: absolute; TOP: 166px" runat="server"
				Width="136px" Height="20px" BackColor="Yellow" ForeColor="Blue">
				<asp:ListItem Value="3">徐滿</asp:ListItem>
			</asp:dropdownlist><asp:image id="DMapSheet3" style="Z-INDEX: 103; LEFT: 8px; POSITION: absolute; TOP: 608px"
				runat="server" ImageUrl="Images\MapSheet_003_A1.JPG" Width="593px" Height="77px"></asp:image><asp:image id="DMapSheet4" style="Z-INDEX: 102; LEFT: 8px; POSITION: absolute; TOP: 688px"
				runat="server" ImageUrl="Images\MapSheet_004.jpg" Width="593px" Height="75px"></asp:image><asp:image id="DDelivery" style="Z-INDEX: 101; LEFT: 8px; POSITION: absolute; TOP: 760px" runat="server"
				ImageUrl="Images\Sheet_Delivery.jpg" Width="593px" Height="110px"></asp:image><asp:image id="DDelay" style="Z-INDEX: 100; LEFT: 8px; POSITION: absolute; TOP: 872px" runat="server"
				ImageUrl="Images\Sheet_Delay.jpg" Width="593px" Height="107px" Visible="False"></asp:image><INPUT id="BOK" style="Z-INDEX: 145; LEFT: 520px; WIDTH: 80px; POSITION: absolute; TOP: 984px; HEIGHT: 28px"
				onclick="Button('OK');" type="button" value="OK" runat="server">
			<asp:imagebutton id="BPrint" style="Z-INDEX: 149; LEFT: 8px; POSITION: absolute; TOP: 8px" runat="server"
				ImageUrl="Images\Print.gif" Width="16px" Height="16px"></asp:imagebutton></form>
	</body>
</HTML>
