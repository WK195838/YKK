<%@ Page Language="vb" AutoEventWireup="False" Codebehind="MapBefSheet_01.aspx.vb" Inherits="SPD.MapBefSheet_01" aspCompat="True"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>舊版製圖委託書</title>
		<meta content="True" name="vs_showGrid">
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<style type="text/css">BODY { FONT-FAMILY: Verdana, Geneva, Sans-Serif; FONT-SIZE: 9pt }
	TABLE { FONT-FAMILY: Verdana, Geneva, Sans-Serif; FONT-SIZE: 9pt }
	TR { FONT-FAMILY: Verdana, Geneva, Sans-Serif; FONT-SIZE: 9pt }
	TD { FONT-FAMILY: Verdana, Geneva, Sans-Serif; FONT-SIZE: 9pt }
	UL { FONT-FAMILY: Verdana, Geneva, Sans-Serif; FONT-SIZE: 9pt }
	LI { FONT-FAMILY: Verdana, Geneva, Sans-Serif; FONT-SIZE: 9pt }
	.normal { FONT-FAMILY: Verdana, Geneva, Sans-Serif; FONT-SIZE: 9pt }
	H1 { FONT-FAMILY: Verdana, Geneva, Sans-Serif; COLOR: #666666; FONT-SIZE: 10.5pt; FONT-WEIGHT: 900 }
	.small { FONT-FAMILY: Verdana, Geneva, Sans-Serif; FONT-SIZE: 7.5pt }
	.error { COLOR: #ff0033 }
	.required { FONT-FAMILY: Verdana, Geneva, Sans-Serif; COLOR: #ff0033; FONT-WEIGHT: 900 }
	.success { MARGIN: 10px 0px; COLOR: #009933; FONT-SIZE: 11pt; FONT-WEIGHT: 900 }
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
				<asp:image id="DMapSheet1" style="Z-INDEX: 104; POSITION: absolute; TOP: 8px; LEFT: 8px" runat="server"
					ImageUrl="Images\MapSheet_001_D.jpg" Width="593px" Height="597px"></asp:image>
				<asp:Label id="Lable27" style="Z-INDEX: 357; POSITION: absolute; TOP: 24px; LEFT: 152px" runat="server"
					Width="88px" ForeColor="Red" BorderColor="White" Font-Size="16pt" Font-Bold="True">〔舊版〕</asp:Label><asp:dropdownlist id="DSuppiler" style="Z-INDEX: 356; POSITION: absolute; TOP: 438px; LEFT: 304px"
					runat="server" Width="146px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DBuyer" style="Z-INDEX: 152; POSITION: absolute; TOP: 132px; LEFT: 168px" runat="server"
					Width="128px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">IS</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DManufType" style="Z-INDEX: 151; POSITION: absolute; TOP: 438px; LEFT: 168px"
					runat="server" Width="128px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="內製">內製</asp:ListItem>
					<asp:ListItem Value="外注">外注</asp:ListItem>
				</asp:dropdownlist><INPUT id="BSAVE" style="Z-INDEX: 148; POSITION: absolute; WIDTH: 80px; HEIGHT: 28px; TOP: 688px; LEFT: 256px"
					onclick="Button('SAVE');" type="button" value="儲存" name="Button1" runat="server"><INPUT id="BNG2" style="Z-INDEX: 147; POSITION: absolute; WIDTH: 80px; HEIGHT: 28px; TOP: 688px; LEFT: 344px"
					onclick="Button('NG2');" type="button" value="NG2" name="Button1" runat="server"><INPUT id="BNG1" style="Z-INDEX: 146; POSITION: absolute; WIDTH: 80px; HEIGHT: 28px; TOP: 688px; LEFT: 432px"
					onclick="Button('NG1');" type="button" value="NG1" name="Button1" runat="server"><asp:dropdownlist id="DCPSC" style="Z-INDEX: 144; POSITION: absolute; TOP: 234px; LEFT: 168px" runat="server"
					Width="128px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DSample" style="Z-INDEX: 143; POSITION: absolute; TOP: 336px; LEFT: 542px" runat="server"
					Width="51px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:button id="BSpec" style="Z-INDEX: 142; POSITION: absolute; TOP: 336px; LEFT: 512px" runat="server"
					Width="20px" Height="20px" Text="....."></asp:button><asp:textbox id="DSpec" style="Z-INDEX: 141; POSITION: absolute; TOP: 336px; LEFT: 64px" runat="server"
					Width="448px" Height="20px" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue" ReadOnly="True">DSpec</asp:textbox><asp:dropdownlist id="DMakeMap" style="Z-INDEX: 140; POSITION: absolute; TOP: 618px; LEFT: 394px"
					runat="server" Width="76px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DLevel" style="Z-INDEX: 139; POSITION: absolute; TOP: 618px; LEFT: 540px" runat="server"
					Width="56px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DMaterialDetail_1" style="Z-INDEX: 138; POSITION: absolute; TOP: 506px; LEFT: 168px"
					runat="server" Width="424px" Height="20px" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue">DMaterialDetail_1</asp:textbox><asp:dropdownlist id="Dhalffinish" style="Z-INDEX: 137; POSITION: absolute; TOP: 404px; LEFT: 456px"
					runat="server" Width="136px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><INPUT id="DMapFile" style="Z-INDEX: 121; POSITION: absolute; BACKGROUND-COLOR: #ffff00; WIDTH: 424px; HEIGHT: 20px; TOP: 654px; LEFT: 168px"
					type="file" size="50" name="File1" runat="server">
				<asp:textbox id="DMapNo" style="Z-INDEX: 120; POSITION: absolute; TOP: 618px; LEFT: 168px" runat="server"
					Width="152px" Height="20px" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue">DMapNo</asp:textbox><INPUT id="DRefMapFile" style="Z-INDEX: 114; POSITION: absolute; BACKGROUND-COLOR: #ffff00; WIDTH: 424px; HEIGHT: 20px; TOP: 542px; LEFT: 168px"
					type="file" size="50" name="File1" runat="server"><asp:textbox id="DSurface" style="Z-INDEX: 110; POSITION: absolute; TOP: 370px; LEFT: 456px"
					runat="server" Width="136px" Height="20px" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue">DSurface</asp:textbox><asp:textbox id="DNo" style="Z-INDEX: 105; POSITION: absolute; TOP: 88px; LEFT: 168px" runat="server"
					Width="128px" Height="20px" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue">DNo</asp:textbox></FONT><asp:textbox id="DDate" style="Z-INDEX: 106; POSITION: absolute; TOP: 88px; LEFT: 456px" runat="server"
				Width="112px" Height="20px" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue" ReadOnly="True">DDate</asp:textbox><asp:textbox id="DSellVendor" style="Z-INDEX: 107; POSITION: absolute; TOP: 132px; LEFT: 456px"
				runat="server" Width="136px" Height="20px" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue">DSellVendor</asp:textbox><asp:textbox id="DBackground" style="Z-INDEX: 108; POSITION: absolute; TOP: 200px; LEFT: 168px"
				runat="server" Width="424px" Height="20px" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue">DBackground</asp:textbox><asp:textbox id="DCramper" style="Z-INDEX: 109; POSITION: absolute; TOP: 370px; LEFT: 168px"
				runat="server" Width="128px" Height="20px" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue">DCramper</asp:textbox><asp:dropdownlist id="DFrontBack" style="Z-INDEX: 111; POSITION: absolute; TOP: 404px; LEFT: 168px"
				runat="server" Width="128px" Height="30px" BackColor="Yellow" ForeColor="Blue">
				<asp:ListItem Value="3">相同</asp:ListItem>
			</asp:dropdownlist><asp:dropdownlist id="DMaterial" style="Z-INDEX: 112; POSITION: absolute; TOP: 472px; LEFT: 168px"
				runat="server" Width="128px" Height="20px" BackColor="Yellow" ForeColor="Blue" AutoPostBack="True">
				<asp:ListItem Value="3">3</asp:ListItem>
			</asp:dropdownlist><asp:dropdownlist id="DMaterialDetail" style="Z-INDEX: 113; POSITION: absolute; TOP: 472px; LEFT: 296px"
				runat="server" Width="296px" Height="20px" BackColor="Yellow" ForeColor="Blue">
				<asp:ListItem Value="3">CF</asp:ListItem>
			</asp:dropdownlist><asp:textbox id="DDescription" style="Z-INDEX: 115; POSITION: absolute; TOP: 574px; LEFT: 168px"
				runat="server" Width="424px" Height="20px" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue">DDescription</asp:textbox><asp:button id="BDate" style="Z-INDEX: 116; POSITION: absolute; TOP: 88px; LEFT: 568px" runat="server"
				Width="20px" Height="20px" Text="....."></asp:button><asp:textbox id="DMapReqDelDate" style="Z-INDEX: 117; POSITION: absolute; TOP: 234px; LEFT: 456px"
				runat="server" Width="112px" Height="20px" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue" ReadOnly="True">DMapReqDelDate</asp:textbox><asp:button id="BMapReqDelDate" style="Z-INDEX: 118; POSITION: absolute; TOP: 234px; LEFT: 568px"
				runat="server" Width="20px" Height="20px" Text="....."></asp:button><asp:dropdownlist id="DLight" style="Z-INDEX: 119; POSITION: absolute; TOP: 268px; LEFT: 168px" runat="server"
				Width="128px" Height="20px" BackColor="Yellow" ForeColor="Blue">
				<asp:ListItem Value="Y">Yes</asp:ListItem>
				<asp:ListItem Value="N">No</asp:ListItem>
			</asp:dropdownlist><asp:dropdownlist id="DDivision" style="Z-INDEX: 133; POSITION: absolute; TOP: 166px; LEFT: 168px"
				runat="server" Width="128px" Height="20px" BackColor="Yellow" ForeColor="Blue">
				<asp:ListItem Value="3">IS</asp:ListItem>
			</asp:dropdownlist><asp:dropdownlist id="DPerson" style="Z-INDEX: 134; POSITION: absolute; TOP: 166px; LEFT: 456px" runat="server"
				Width="136px" Height="20px" BackColor="Yellow" ForeColor="Blue">
				<asp:ListItem Value="3">徐滿</asp:ListItem>
			</asp:dropdownlist><asp:image id="DMapSheet3" style="Z-INDEX: 103; POSITION: absolute; TOP: 608px; LEFT: 8px"
				runat="server" ImageUrl="Images\MapSheet_003_A.jpg" Width="593px" Height="77px"></asp:image><INPUT id="BOK" style="Z-INDEX: 145; POSITION: absolute; WIDTH: 80px; HEIGHT: 28px; TOP: 688px; LEFT: 520px"
				onclick="Button('OK');" type="button" value="OK" runat="server" NAME="BOK">
		</form>
	</body>
</HTML>
