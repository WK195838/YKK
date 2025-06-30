<%@ Page Language="vb" AutoEventWireup="false" Codebehind="ManufInBefSheet_01.aspx.vb" Inherits="SPD.ManufInBefSheet_01" aspCompat="True"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>內製委託書</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<script language="javascript" type="text/javascript">
			var wPop;
			var val;
			//--Calendar------------------------------------
			function CalendarPicker(strField) {
				window.open('DatePicker.aspx?field=' + strField,'CalendarPopup','width=250,height=190,resizable=yes');
			}
			//--圖號------------------------------------
			function MapPicker(strField) {
				window.open('MapPicker.aspx?field=' + strField,'MapPopup','width=168,height=360,resizable=yes');
			}
			//--Slider Code------------------------------------
			function SliderPicker(strField) {
			    val=0;
				wPop=window.open('SliderPicker.aspx?field=' + strField,'SliderPopup','width=300,height=250,resizable=yes');
				if ((document.Form1.DSliderGRCode.value != "") || (document.Form1.DSliderCode.value != "")) {
					setTimeout("SendToSlider()",200);
					//------------------------------------
					//while (val == 0) {
						//setTimeout("SendToSlider()",200);
						//if ((wPop.document.SliderForm.DSlider2.value!="") || (wPop.document.SliderForm.DContent.value!=""))  {
						//	val=1;
						//}
 					//} 
					//------------------------------------
				}
			}
		    function SendToSlider() {
				wPop.document.SliderForm.DSlider2.value=document.Form1.DSliderGRCode.value;
				wPop.document.SliderForm.DContent.value=document.Form1.DSliderCode.value;
			}
			//--Spec------------------------------------
			function SpecPicker(strField, Fun) {
			    val=0;
				wPop=window.open('SpecPicker.aspx?field=' + strField + '&fun=' + Fun, 'SpecPopup','width=330,height=250,resizable=yes');
				if (document.Form1.DSpec.value != "") {
					setTimeout("SendToSpec()",200);
					//------------------------------------
					//while (val == 0) {
						//setTimeout("SendToSpec()",200);
						//if (wPop.document.SpecForm.DSpec.value!="")  {
						//	val=1;
						//}
 					//} 
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
		<form id="Form1" method="post" encType="multipart/form-data" runat="server">
			<FONT face="新細明體">
				<asp:dropdownlist id="DQAD13" style="Z-INDEX: 308; POSITION: absolute; TOP: 750px; LEFT: 664px" runat="server"
					Width="48px" Height="346px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0點">0點</asp:ListItem>
					<asp:ListItem Value="2點">2點</asp:ListItem>
					<asp:ListItem Value="4點">4點</asp:ListItem>
					<asp:ListItem Value="6點">6點</asp:ListItem>
					<asp:ListItem Value="8點">8點</asp:ListItem>
					<asp:ListItem Value="10點">10點</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DOPReadyDesc" style="Z-INDEX: 313; POSITION: absolute; TOP: 1184px; LEFT: 248px"
					runat="server" Width="104px" Height="20px" BackColor="White" ForeColor="Red" BorderStyle="None">需閱讀工程履歷</asp:textbox><asp:textbox id="DOPReady" style="Z-INDEX: 153; POSITION: absolute; TOP: 1184px; LEFT: 352px"
					runat="server" Width="48px" Height="20px" BackColor="GreenYellow" ForeColor="Red" BorderStyle="Groove" ReadOnly="True"></asp:textbox><asp:dropdownlist id="DCpsc" style="Z-INDEX: 312; POSITION: absolute; TOP: 192px; LEFT: 708px" runat="server"
					Width="56px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DMakeCAM" style="Z-INDEX: 311; POSITION: absolute; TOP: 360px; LEFT: 584px"
					runat="server" Width="174px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAC13" style="Z-INDEX: 307; POSITION: absolute; TOP: 732px; LEFT: 664px" runat="server"
					Width="48px" Height="346px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0點">0點</asp:ListItem>
					<asp:ListItem Value="2點">2點</asp:ListItem>
					<asp:ListItem Value="4點">4點</asp:ListItem>
					<asp:ListItem Value="6點">6點</asp:ListItem>
					<asp:ListItem Value="8點">8點</asp:ListItem>
					<asp:ListItem Value="10點">10點</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAB13" style="Z-INDEX: 306; POSITION: absolute; TOP: 714px; LEFT: 664px" runat="server"
					Width="48px" Height="346px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0點">0點</asp:ListItem>
					<asp:ListItem Value="2點">2點</asp:ListItem>
					<asp:ListItem Value="4點">4點</asp:ListItem>
					<asp:ListItem Value="6點">6點</asp:ListItem>
					<asp:ListItem Value="8點">8點</asp:ListItem>
					<asp:ListItem Value="10點">10點</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAD12" style="Z-INDEX: 305; POSITION: absolute; TOP: 750px; LEFT: 616px" runat="server"
					Width="48px" Height="346px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0級">0級</asp:ListItem>
					<asp:ListItem Value="1級">1級</asp:ListItem>
					<asp:ListItem Value="2級">2級</asp:ListItem>
					<asp:ListItem Value="3級">3級</asp:ListItem>
					<asp:ListItem Value="4級">4級</asp:ListItem>
					<asp:ListItem Value="5級">5級</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAC12" style="Z-INDEX: 304; POSITION: absolute; TOP: 732px; LEFT: 616px" runat="server"
					Width="48px" Height="346px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0級">0級</asp:ListItem>
					<asp:ListItem Value="1級">1級</asp:ListItem>
					<asp:ListItem Value="2級">2級</asp:ListItem>
					<asp:ListItem Value="3級">3級</asp:ListItem>
					<asp:ListItem Value="4級">4級</asp:ListItem>
					<asp:ListItem Value="5級">5級</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAB12" style="Z-INDEX: 303; POSITION: absolute; TOP: 714px; LEFT: 616px" runat="server"
					Width="48px" Height="346px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0級">0級</asp:ListItem>
					<asp:ListItem Value="1級">1級</asp:ListItem>
					<asp:ListItem Value="2級">2級</asp:ListItem>
					<asp:ListItem Value="3級">3級</asp:ListItem>
					<asp:ListItem Value="4級">4級</asp:ListItem>
					<asp:ListItem Value="5級">5級</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAD11" style="Z-INDEX: 302; POSITION: absolute; TOP: 750px; LEFT: 568px" runat="server"
					Width="48px" Height="346px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="1級">1級</asp:ListItem>
					<asp:ListItem Value="1~2級">1~2級</asp:ListItem>
					<asp:ListItem Value="2級">2級</asp:ListItem>
					<asp:ListItem Value="2~3級">2~3級</asp:ListItem>
					<asp:ListItem Value="3級">3級</asp:ListItem>
					<asp:ListItem Value="3~4級">3~4級</asp:ListItem>
					<asp:ListItem Value="4級">4級</asp:ListItem>
					<asp:ListItem Value="4~5級">4~5級</asp:ListItem>
					<asp:ListItem Value="5級">5級</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAC11" style="Z-INDEX: 301; POSITION: absolute; TOP: 732px; LEFT: 568px" runat="server"
					Width="48px" Height="346px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="1級">1級</asp:ListItem>
					<asp:ListItem Value="1~2級">1~2級</asp:ListItem>
					<asp:ListItem Value="2級">2級</asp:ListItem>
					<asp:ListItem Value="2~3級">2~3級</asp:ListItem>
					<asp:ListItem Value="3級">3級</asp:ListItem>
					<asp:ListItem Value="3~4級">3~4級</asp:ListItem>
					<asp:ListItem Value="4級">4級</asp:ListItem>
					<asp:ListItem Value="4~5級">4~5級</asp:ListItem>
					<asp:ListItem Value="5級">5級</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAB11" style="Z-INDEX: 300; POSITION: absolute; TOP: 714px; LEFT: 568px" runat="server"
					Width="48px" Height="346px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="1級">1級</asp:ListItem>
					<asp:ListItem Value="1~2級">1~2級</asp:ListItem>
					<asp:ListItem Value="2級">2級</asp:ListItem>
					<asp:ListItem Value="2~3級">2~3級</asp:ListItem>
					<asp:ListItem Value="3級">3級</asp:ListItem>
					<asp:ListItem Value="3~4級">3~4級</asp:ListItem>
					<asp:ListItem Value="4級">4級</asp:ListItem>
					<asp:ListItem Value="4~5級">4~5級</asp:ListItem>
					<asp:ListItem Value="5級">5級</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAA13" style="Z-INDEX: 299; POSITION: absolute; TOP: 696px; LEFT: 664px" runat="server"
					Width="48px" Height="346px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0點">0點</asp:ListItem>
					<asp:ListItem Value="2點">2點</asp:ListItem>
					<asp:ListItem Value="4點">4點</asp:ListItem>
					<asp:ListItem Value="6點">6點</asp:ListItem>
					<asp:ListItem Value="8點">8點</asp:ListItem>
					<asp:ListItem Value="10點">10點</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAA12" style="Z-INDEX: 298; POSITION: absolute; TOP: 696px; LEFT: 616px" runat="server"
					Width="48px" Height="346px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0級">0級</asp:ListItem>
					<asp:ListItem Value="1級">1級</asp:ListItem>
					<asp:ListItem Value="2級">2級</asp:ListItem>
					<asp:ListItem Value="3級">3級</asp:ListItem>
					<asp:ListItem Value="4級">4級</asp:ListItem>
					<asp:ListItem Value="5級">5級</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAA11" style="Z-INDEX: 297; POSITION: absolute; TOP: 696px; LEFT: 568px" runat="server"
					Width="48px" Height="346px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="1級">1級</asp:ListItem>
					<asp:ListItem Value="1~2級">1~2級</asp:ListItem>
					<asp:ListItem Value="2級">2級</asp:ListItem>
					<asp:ListItem Value="2~3級">2~3級</asp:ListItem>
					<asp:ListItem Value="3級">3級</asp:ListItem>
					<asp:ListItem Value="3~4級">3~4級</asp:ListItem>
					<asp:ListItem Value="4級">4級</asp:ListItem>
					<asp:ListItem Value="4~5級">4~5級</asp:ListItem>
					<asp:ListItem Value="5級">5級</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAD14" style="Z-INDEX: 241; POSITION: absolute; TOP: 750px; LEFT: 712px" runat="server"
					Width="48px" Height="22px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAD10" style="Z-INDEX: 240; POSITION: absolute; TOP: 750px; LEFT: 520px" runat="server"
					Width="48px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAD9" style="Z-INDEX: 239; POSITION: absolute; TOP: 750px; LEFT: 472px" runat="server"
					Width="48px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAD8" style="Z-INDEX: 238; POSITION: absolute; TOP: 750px; LEFT: 424px" runat="server"
					Width="48px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="NCA">NCA</asp:ListItem>
					<asp:ListItem Value="NCB">NCB</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAD7" style="Z-INDEX: 237; POSITION: absolute; TOP: 750px; LEFT: 376px" runat="server"
					Width="48px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAD6" style="Z-INDEX: 236; POSITION: absolute; TOP: 750px; LEFT: 328px" runat="server"
					Width="48px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DQAF4" style="Z-INDEX: 235; POSITION: absolute; TOP: 750px; LEFT: 304px" runat="server"
					Width="22px" Height="16px" BackColor="White" ForeColor="Blue" Font-Size="8pt" BorderStyle="None">g/pc</asp:textbox><asp:textbox id="DQAD5" style="Z-INDEX: 234; POSITION: absolute; TOP: 750px; LEFT: 264px" runat="server"
					Width="40px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DQAD4" style="Z-INDEX: 233; POSITION: absolute; TOP: 750px; LEFT: 232px" runat="server"
					Width="32px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:dropdownlist id="DQAD3" style="Z-INDEX: 232; POSITION: absolute; TOP: 750px; LEFT: 192px" runat="server"
					Width="40px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Selected="True">無</asp:ListItem>
					<asp:ListItem Value="正">正</asp:ListItem>
					<asp:ListItem Value="反">反</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DQAD2" style="Z-INDEX: 231; POSITION: absolute; TOP: 750px; LEFT: 80px" runat="server"
					Width="112px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DQAD1" style="Z-INDEX: 230; POSITION: absolute; TOP: 750px; LEFT: 16px" runat="server"
					Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:dropdownlist id="DQAC14" style="Z-INDEX: 229; POSITION: absolute; TOP: 732px; LEFT: 712px" runat="server"
					Width="48px" Height="22px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAC10" style="Z-INDEX: 228; POSITION: absolute; TOP: 732px; LEFT: 520px" runat="server"
					Width="48px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAC9" style="Z-INDEX: 227; POSITION: absolute; TOP: 732px; LEFT: 472px" runat="server"
					Width="48px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAC8" style="Z-INDEX: 226; POSITION: absolute; TOP: 732px; LEFT: 424px" runat="server"
					Width="48px" Height="36px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="NCA">NCA</asp:ListItem>
					<asp:ListItem Value="NCB">NCB</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAC7" style="Z-INDEX: 225; POSITION: absolute; TOP: 732px; LEFT: 376px" runat="server"
					Width="48px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAC6" style="Z-INDEX: 224; POSITION: absolute; TOP: 732px; LEFT: 328px" runat="server"
					Width="48px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DQAF3" style="Z-INDEX: 223; POSITION: absolute; TOP: 732px; LEFT: 304px" runat="server"
					Width="22px" Height="16px" BackColor="White" ForeColor="Blue" Font-Size="8pt" BorderStyle="None">g/pc</asp:textbox><asp:textbox id="DQAC5" style="Z-INDEX: 222; POSITION: absolute; TOP: 732px; LEFT: 264px" runat="server"
					Width="40px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DQAC4" style="Z-INDEX: 221; POSITION: absolute; TOP: 732px; LEFT: 232px" runat="server"
					Width="32px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:dropdownlist id="DQAC3" style="Z-INDEX: 220; POSITION: absolute; TOP: 732px; LEFT: 192px" runat="server"
					Width="40px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Selected="True">無</asp:ListItem>
					<asp:ListItem Value="正">正</asp:ListItem>
					<asp:ListItem Value="反">反</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DQAC2" style="Z-INDEX: 219; POSITION: absolute; TOP: 732px; LEFT: 80px" runat="server"
					Width="112px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DQAC1" style="Z-INDEX: 218; POSITION: absolute; TOP: 732px; LEFT: 16px" runat="server"
					Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:dropdownlist id="DQAB14" style="Z-INDEX: 217; POSITION: absolute; TOP: 714px; LEFT: 712px" runat="server"
					Width="48px" Height="22px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAB10" style="Z-INDEX: 216; POSITION: absolute; TOP: 714px; LEFT: 520px" runat="server"
					Width="48px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAB9" style="Z-INDEX: 215; POSITION: absolute; TOP: 714px; LEFT: 472px" runat="server"
					Width="48px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DQAF2" style="Z-INDEX: 214; POSITION: absolute; TOP: 714px; LEFT: 304px" runat="server"
					Width="22px" Height="16px" BackColor="White" ForeColor="Blue" Font-Size="8pt" BorderStyle="None">g/pc</asp:textbox><asp:dropdownlist id="DQAB8" style="Z-INDEX: 213; POSITION: absolute; TOP: 714px; LEFT: 424px" runat="server"
					Width="48px" Height="36px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="NCA">NCA</asp:ListItem>
					<asp:ListItem Value="NCB">NCB</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAB7" style="Z-INDEX: 212; POSITION: absolute; TOP: 714px; LEFT: 376px" runat="server"
					Width="48px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAB6" style="Z-INDEX: 211; POSITION: absolute; TOP: 714px; LEFT: 328px" runat="server"
					Width="48px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DQAB5" style="Z-INDEX: 210; POSITION: absolute; TOP: 714px; LEFT: 264px" runat="server"
					Width="40px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DQAB4" style="Z-INDEX: 209; POSITION: absolute; TOP: 714px; LEFT: 232px" runat="server"
					Width="32px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:dropdownlist id="DQAB3" style="Z-INDEX: 208; POSITION: absolute; TOP: 714px; LEFT: 192px" runat="server"
					Width="40px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Selected="True">無</asp:ListItem>
					<asp:ListItem Value="正">正</asp:ListItem>
					<asp:ListItem Value="反">反</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DQAB2" style="Z-INDEX: 207; POSITION: absolute; TOP: 714px; LEFT: 80px" runat="server"
					Width="112px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DQAB1" style="Z-INDEX: 206; POSITION: absolute; TOP: 714px; LEFT: 16px" runat="server"
					Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:dropdownlist id="DQAA14" style="Z-INDEX: 162; POSITION: absolute; TOP: 696px; LEFT: 712px" runat="server"
					Width="48px" Height="22px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAA10" style="Z-INDEX: 161; POSITION: absolute; TOP: 696px; LEFT: 520px" runat="server"
					Width="48px" Height="346px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAA9" style="Z-INDEX: 160; POSITION: absolute; TOP: 696px; LEFT: 472px" runat="server"
					Width="48px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAA8" style="Z-INDEX: 159; POSITION: absolute; TOP: 696px; LEFT: 424px" runat="server"
					Width="48px" Height="36px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="NCA">NCA</asp:ListItem>
					<asp:ListItem Value="NCB">NCB</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAA7" style="Z-INDEX: 158; POSITION: absolute; TOP: 696px; LEFT: 376px" runat="server"
					Width="48px" Height="162px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAA6" style="Z-INDEX: 157; POSITION: absolute; TOP: 696px; LEFT: 328px" runat="server"
					Width="48px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DQAA5" style="Z-INDEX: 156; POSITION: absolute; TOP: 696px; LEFT: 264px" runat="server"
					Width="40px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DQAA4" style="Z-INDEX: 155; POSITION: absolute; TOP: 696px; LEFT: 232px" runat="server"
					Width="32px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:dropdownlist id="DQAA3" style="Z-INDEX: 154; POSITION: absolute; TOP: 696px; LEFT: 192px" runat="server"
					Width="40px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Selected="True">無</asp:ListItem>
					<asp:ListItem Value="正">正</asp:ListItem>
					<asp:ListItem Value="反">反</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DQAA1" style="Z-INDEX: 152; POSITION: absolute; TOP: 696px; LEFT: 16px" runat="server"
					Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DQAA2" style="Z-INDEX: 171; POSITION: absolute; TOP: 696px; LEFT: 80px" runat="server"
					Width="112px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DQAF1" style="Z-INDEX: 172; POSITION: absolute; TOP: 696px; LEFT: 304px" runat="server"
					Width="22px" Height="16px" BackColor="White" ForeColor="Blue" Font-Size="8pt" BorderStyle="None">g/pc</asp:textbox><asp:image id="DManuaInSheet1" style="Z-INDEX: 101; POSITION: absolute; TOP: 8px; LEFT: 8px"
					runat="server" Width="761px" Height="530px" ImageUrl="Images\ManufInSheet_003_A.jpg"></asp:image><asp:dropdownlist id="DSuppiler" style="Z-INDEX: 309; POSITION: absolute; TOP: 395px; LEFT: 400px"
					runat="server" Width="88px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:label id="Label25" style="Z-INDEX: 296; POSITION: absolute; TOP: 440px; LEFT: 704px" runat="server"
					Width="24px" Height="10px" Font-Size="8pt">數量</asp:label><asp:label id="Label24" style="Z-INDEX: 295; POSITION: absolute; TOP: 440px; LEFT: 648px" runat="server"
					Width="24px" Height="10px" Font-Size="8pt">顏色</asp:label><asp:label id="Label23" style="Z-INDEX: 294; POSITION: absolute; TOP: 440px; LEFT: 472px" runat="server"
					Width="48px" Height="10px" Font-Size="8pt">拉頭種類</asp:label><asp:label id="Label1" style="Z-INDEX: 293; POSITION: absolute; TOP: 684px; LEFT: 48px" runat="server"
					Width="24px" Height="10px" Font-Size="8pt">日期</asp:label><asp:label id="Label22" style="Z-INDEX: 292; POSITION: absolute; TOP: 556px; LEFT: 504px" runat="server"
					Width="24px" Height="10px" Font-Size="8pt">單價</asp:label><asp:label id="Label21" style="Z-INDEX: 291; POSITION: absolute; TOP: 556px; LEFT: 464px" runat="server"
					Width="24px" Height="10px" Font-Size="8pt">幣別</asp:label><asp:label id="Label20" style="Z-INDEX: 290; POSITION: absolute; TOP: 556px; LEFT: 296px" runat="server"
					Width="48px" Height="10px" Font-Size="8pt">拉頭種類</asp:label><asp:label id="Label19" style="Z-INDEX: 289; POSITION: absolute; TOP: 556px; LEFT: 224px" runat="server"
					Width="24px" Height="10px" Font-Size="8pt">單價</asp:label><asp:label id="Label18" style="Z-INDEX: 288; POSITION: absolute; TOP: 556px; LEFT: 184px" runat="server"
					Width="24px" Height="10px" Font-Size="8pt">幣別</asp:label><asp:label id="Label17" style="Z-INDEX: 287; POSITION: absolute; TOP: 556px; LEFT: 56px" runat="server"
					Width="48px" Height="10px" Font-Size="8pt">拉頭種類</asp:label><asp:label id="Label16" style="Z-INDEX: 286; POSITION: absolute; TOP: 812px; LEFT: 136px" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">備註</asp:label><asp:label id="Label15" style="Z-INDEX: 285; POSITION: absolute; TOP: 812px; LEFT: 80px" runat="server"
					Width="48px" Height="10px" Font-Size="8pt">檢測結果</asp:label><asp:label id="Label14" style="Z-INDEX: 284; POSITION: absolute; TOP: 812px; LEFT: 48px" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">日期</asp:label><asp:label id="Label13" style="Z-INDEX: 283; POSITION: absolute; TOP: 684px; LEFT: 712px" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">CPSC</asp:label><asp:label id="Label26" style="Z-INDEX: 310; POSITION: absolute; TOP: 684px; LEFT: 616px" runat="server"
					Width="48px" Height="10px" Font-Size="8pt">一次密著</asp:label><asp:label id="Label12" style="Z-INDEX: 282; POSITION: absolute; TOP: 684px; LEFT: 664px" runat="server"
					Width="48px" Height="10px" Font-Size="8pt">二次密著</asp:label><asp:label id="Label11" style="Z-INDEX: 281; POSITION: absolute; TOP: 684px; LEFT: 568px" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">黃變</asp:label><asp:label id="Label10" style="Z-INDEX: 280; POSITION: absolute; TOP: 684px; LEFT: 520px" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">乾洗</asp:label><asp:label id="Label9" style="Z-INDEX: 279; POSITION: absolute; TOP: 684px; LEFT: 472px" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">水洗</asp:label><asp:label id="Label8" style="Z-INDEX: 278; POSITION: absolute; TOP: 684px; LEFT: 424px" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">檢針</asp:label><asp:label id="Label7" style="Z-INDEX: 277; POSITION: absolute; TOP: 684px; LEFT: 376px" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">扭力</asp:label><asp:label id="Label6" style="Z-INDEX: 276; POSITION: absolute; TOP: 684px; LEFT: 328px" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">強度</asp:label><asp:label id="Label5" style="Z-INDEX: 275; POSITION: absolute; TOP: 684px; LEFT: 264px" runat="server"
					Width="40px" Height="10px" Font-Size="8pt">原單位</asp:label><asp:label id="Label4" style="Z-INDEX: 274; POSITION: absolute; TOP: 684px; LEFT: 232px" runat="server"
					Width="48px" Height="10px" Font-Size="8pt">表面</asp:label><asp:label id="Label3" style="Z-INDEX: 273; POSITION: absolute; TOP: 684px; LEFT: 192px" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">組立</asp:label><asp:label id="Label2" style="Z-INDEX: 272; POSITION: absolute; TOP: 684px; LEFT: 80px" runat="server"
					Width="48px" Height="10px" Font-Size="8pt">拉頭種類</asp:label><asp:dropdownlist id="DBuyer" style="Z-INDEX: 271; POSITION: absolute; TOP: 328px; LEFT: 312px" runat="server"
					Width="176px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:hyperlink id="LQAttachFile1" style="Z-INDEX: 270; POSITION: absolute; TOP: 772px; LEFT: 16px"
					runat="server" Width="64px" Height="8px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">其他附件</asp:hyperlink><asp:hyperlink id="LQAttachFile2" style="Z-INDEX: 269; POSITION: absolute; TOP: 900px; LEFT: 136px"
					runat="server" Width="120px" Height="8px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">品質分析依賴書</asp:hyperlink><INPUT id="DQAttachFile2" style="Z-INDEX: 268; POSITION: absolute; BACKGROUND-COLOR: #ffff00; WIDTH: 352px; HEIGHT: 20px; TOP: 900px; LEFT: 136px"
					type="file" size="39" name="File1" runat="server"><INPUT id="DQAttachFile1" style="Z-INDEX: 267; POSITION: absolute; BACKGROUND-COLOR: #ffff00; WIDTH: 744px; HEIGHT: 20px; TOP: 772px; LEFT: 16px"
					type="file" size="104" name="File1" runat="server">
				<asp:hyperlink id="LSAttachFile" style="Z-INDEX: 266; POSITION: absolute; TOP: 512px; LEFT: 440px"
					runat="server" Width="64px" Height="8px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">其他附件</asp:hyperlink><INPUT id="DSAttachFile" style="Z-INDEX: 265; POSITION: absolute; BACKGROUND-COLOR: #ffff00; WIDTH: 314px; HEIGHT: 20px; TOP: 508px; LEFT: 440px"
					type="file" size="31" name="File1" runat="server"><asp:imagebutton id="BFlow" style="Z-INDEX: 264; POSITION: absolute; TOP: 32px; LEFT: 8px" runat="server"
					Width="16px" Height="16px" ImageUrl="Images\Flow-01.gif"></asp:imagebutton>
				<asp:textbox id="DFQAD3" style="Z-INDEX: 262; POSITION: absolute; TOP: 878px; LEFT: 136px" runat="server"
					Width="620px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:dropdownlist id="DFQAD2" style="Z-INDEX: 261; POSITION: absolute; TOP: 878px; LEFT: 80px" runat="server"
					Width="56px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DFQAD1" style="Z-INDEX: 260; POSITION: absolute; TOP: 878px; LEFT: 16px" runat="server"
					Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DFQAC1" style="Z-INDEX: 259; POSITION: absolute; TOP: 860px; LEFT: 16px" runat="server"
					Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:dropdownlist id="DFQAC2" style="Z-INDEX: 258; POSITION: absolute; TOP: 860px; LEFT: 80px" runat="server"
					Width="56px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DFQAC3" style="Z-INDEX: 257; POSITION: absolute; TOP: 860px; LEFT: 136px" runat="server"
					Width="620px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DFQAB3" style="Z-INDEX: 256; POSITION: absolute; TOP: 842px; LEFT: 136px" runat="server"
					Width="620px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:dropdownlist id="DFQAB2" style="Z-INDEX: 255; POSITION: absolute; TOP: 842px; LEFT: 80px" runat="server"
					Width="56px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DFQAB1" style="Z-INDEX: 254; POSITION: absolute; TOP: 842px; LEFT: 16px" runat="server"
					Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DOFormSno" style="Z-INDEX: 249; POSITION: absolute; TOP: 293px; LEFT: 520px"
					runat="server" Width="32px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" AutoPostBack="True">123</asp:textbox><asp:textbox id="DOFormNo" style="Z-INDEX: 248; POSITION: absolute; TOP: 293px; LEFT: 472px"
					runat="server" Width="48px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" AutoPostBack="True">000001</asp:textbox><INPUT id="DMapFile" style="Z-INDEX: 247; POSITION: absolute; BACKGROUND-COLOR: #ffff00; WIDTH: 200px; HEIGHT: 20px; TOP: 328px; LEFT: 16px"
					type="file" size="14" name="File1" runat="server">
				<asp:hyperlink id="LMapNo" style="Z-INDEX: 245; POSITION: absolute; TOP: 296px; LEFT: 312px" runat="server"
					Width="112px" Height="8px" Target="_blank" NavigateUrl="BoardEdit.aspx">圖號文件</asp:hyperlink><asp:textbox id="DFQAA3" style="Z-INDEX: 244; POSITION: absolute; TOP: 824px; LEFT: 136px" runat="server"
					Width="620px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:dropdownlist id="DFQAA2" style="Z-INDEX: 243; POSITION: absolute; TOP: 824px; LEFT: 80px" runat="server"
					Width="56px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DFQAA1" style="Z-INDEX: 242; POSITION: absolute; TOP: 824px; LEFT: 16px" runat="server"
					Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DPriceJ1" style="Z-INDEX: 203; POSITION: absolute; TOP: 640px; LEFT: 296px"
					runat="server" Width="166px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DPriceI3" style="Z-INDEX: 202; POSITION: absolute; TOP: 640px; LEFT: 224px"
					runat="server" Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:dropdownlist id="DPriceI2" style="Z-INDEX: 201; POSITION: absolute; TOP: 640px; LEFT: 184px"
					runat="server" Width="40px" Height="12px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceI1" style="Z-INDEX: 200; POSITION: absolute; TOP: 640px; LEFT: 16px" runat="server"
					Width="166px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DPriceH3" style="Z-INDEX: 199; POSITION: absolute; TOP: 622px; LEFT: 504px"
					runat="server" Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:dropdownlist id="DPriceH2" style="Z-INDEX: 198; POSITION: absolute; TOP: 622px; LEFT: 464px"
					runat="server" Width="40px" Height="12px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceJ3" style="Z-INDEX: 205; POSITION: absolute; TOP: 640px; LEFT: 504px"
					runat="server" Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:dropdownlist id="DPriceJ2" style="Z-INDEX: 204; POSITION: absolute; TOP: 640px; LEFT: 464px"
					runat="server" Width="40px" Height="12px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceA3" style="Z-INDEX: 151; POSITION: absolute; TOP: 568px; LEFT: 224px"
					runat="server" Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:dropdownlist id="DPriceA2" style="Z-INDEX: 150; POSITION: absolute; TOP: 568px; LEFT: 184px"
					runat="server" Width="40px" Height="12px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceH1" style="Z-INDEX: 197; POSITION: absolute; TOP: 622px; LEFT: 296px"
					runat="server" Width="166px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DPriceG3" style="Z-INDEX: 196; POSITION: absolute; TOP: 622px; LEFT: 224px"
					runat="server" Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:dropdownlist id="DPriceG2" style="Z-INDEX: 195; POSITION: absolute; TOP: 622px; LEFT: 184px"
					runat="server" Width="40px" Height="12px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceG1" style="Z-INDEX: 194; POSITION: absolute; TOP: 622px; LEFT: 16px" runat="server"
					Width="166px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DPriceF3" style="Z-INDEX: 193; POSITION: absolute; TOP: 604px; LEFT: 504px"
					runat="server" Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:dropdownlist id="DPriceF2" style="Z-INDEX: 192; POSITION: absolute; TOP: 604px; LEFT: 464px"
					runat="server" Width="40px" Height="12px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceF1" style="Z-INDEX: 191; POSITION: absolute; TOP: 604px; LEFT: 296px"
					runat="server" Width="166px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DPriceE3" style="Z-INDEX: 190; POSITION: absolute; TOP: 604px; LEFT: 224px"
					runat="server" Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:dropdownlist id="DPriceE2" style="Z-INDEX: 189; POSITION: absolute; TOP: 604px; LEFT: 184px"
					runat="server" Width="40px" Height="12px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceE1" style="Z-INDEX: 188; POSITION: absolute; TOP: 604px; LEFT: 16px" runat="server"
					Width="166px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DPriceD3" style="Z-INDEX: 187; POSITION: absolute; TOP: 586px; LEFT: 504px"
					runat="server" Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:dropdownlist id="DPriceD2" style="Z-INDEX: 186; POSITION: absolute; TOP: 586px; LEFT: 464px"
					runat="server" Width="40px" Height="12px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceD1" style="Z-INDEX: 185; POSITION: absolute; TOP: 586px; LEFT: 296px"
					runat="server" Width="166px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DPriceC3" style="Z-INDEX: 184; POSITION: absolute; TOP: 586px; LEFT: 224px"
					runat="server" Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:dropdownlist id="DPriceC2" style="Z-INDEX: 183; POSITION: absolute; TOP: 586px; LEFT: 184px"
					runat="server" Width="40px" Height="12px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceC1" style="Z-INDEX: 182; POSITION: absolute; TOP: 586px; LEFT: 16px" runat="server"
					Width="166px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DPriceB3" style="Z-INDEX: 181; POSITION: absolute; TOP: 568px; LEFT: 504px"
					runat="server" Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:dropdownlist id="DPriceB2" style="Z-INDEX: 180; POSITION: absolute; TOP: 568px; LEFT: 464px"
					runat="server" Width="40px" Height="12px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceB1" style="Z-INDEX: 179; POSITION: absolute; TOP: 568px; LEFT: 296px"
					runat="server" Width="166px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DSampleC3" style="Z-INDEX: 178; POSITION: absolute; TOP: 488px; LEFT: 704px"
					runat="server" Width="52px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DSampleC2" style="Z-INDEX: 177; POSITION: absolute; TOP: 488px; LEFT: 648px"
					runat="server" Width="52px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DSampleC1" style="Z-INDEX: 176; POSITION: absolute; TOP: 488px; LEFT: 440px"
					runat="server" Width="204px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DSampleB3" style="Z-INDEX: 175; POSITION: absolute; TOP: 470px; LEFT: 704px"
					runat="server" Width="52px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DSampleB2" style="Z-INDEX: 174; POSITION: absolute; TOP: 470px; LEFT: 648px"
					runat="server" Width="52px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DSampleB1" style="Z-INDEX: 173; POSITION: absolute; TOP: 470px; LEFT: 440px"
					runat="server" Width="204px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DPriceA1" style="Z-INDEX: 170; POSITION: absolute; TOP: 568px; LEFT: 16px" runat="server"
					Width="166px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DSampleA1" style="Z-INDEX: 169; POSITION: absolute; TOP: 452px; LEFT: 440px"
					runat="server" Width="204px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:button id="BSliderCode" style="Z-INDEX: 168; POSITION: absolute; TOP: 120px; LEFT: 568px"
					runat="server" Width="20px" Height="20px" Text="....."></asp:button><asp:hyperlink id="LContactFile" style="Z-INDEX: 167; POSITION: absolute; TOP: 1000px; LEFT: 120px"
					runat="server" Width="50px" Height="8px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">切結書</asp:hyperlink><INPUT id="DContactFile" style="Z-INDEX: 166; POSITION: absolute; BACKGROUND-COLOR: #ffff00; WIDTH: 240px; HEIGHT: 20px; TOP: 1000px; LEFT: 120px"
					type="file" name="File1" runat="server">
				<asp:hyperlink id="LRefFile" style="Z-INDEX: 165; POSITION: absolute; TOP: 932px; LEFT: 120px"
					runat="server" Width="64px" Height="8px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">其他文件</asp:hyperlink><asp:hyperlink id="LQAFile" style="Z-INDEX: 164; POSITION: absolute; TOP: 932px; LEFT: 496px" runat="server"
					Width="64px" Height="8px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">測試報告</asp:hyperlink><INPUT id="DQAFile" style="Z-INDEX: 163; POSITION: absolute; BACKGROUND-COLOR: #ffff00; WIDTH: 260px; HEIGHT: 20px; TOP: 932px; LEFT: 496px"
					type="file" size="24" name="File1" runat="server"><asp:textbox id="DSampleA2" style="Z-INDEX: 105; POSITION: absolute; TOP: 452px; LEFT: 648px"
					runat="server" Width="52px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox>
				<asp:textbox id="DSampleA3" style="Z-INDEX: 107; POSITION: absolute; TOP: 452px; LEFT: 704px"
					runat="server" Width="52px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt"
					BorderStyle="Groove"></asp:textbox><asp:button id="BDate" style="Z-INDEX: 149; POSITION: absolute; TOP: 192px; LEFT: 616px" runat="server"
					Width="20px" Height="20px" Text="....."></asp:button><asp:textbox id="DDate" style="Z-INDEX: 148; POSITION: absolute; TOP: 192px; LEFT: 520px" runat="server"
					Width="96px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True">DDate</asp:textbox><asp:textbox id="DSliderCode" style="Z-INDEX: 147; POSITION: absolute; TOP: 120px; LEFT: 224px"
					runat="server" Width="342px" Height="56px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove" ReadOnly="True" TextMode="MultiLine">SliderCode</asp:textbox><asp:button id="BMMapNo" style="Z-INDEX: 146; POSITION: absolute; TOP: 293px; LEFT: 448px" runat="server"
					Width="20px" Height="20px" Text="修"></asp:button><asp:hyperlink id="LSampleFile" style="Z-INDEX: 145; POSITION: absolute; TOP: 1000px; LEFT: 496px"
					runat="server" Width="50px" Height="8px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">樣品圖</asp:hyperlink><INPUT id="DSampleFile" style="Z-INDEX: 144; POSITION: absolute; BACKGROUND-COLOR: #ffff00; WIDTH: 260px; HEIGHT: 20px; TOP: 1000px; LEFT: 496px"
					type="file" size="24" name="File1" runat="server"><asp:dropdownlist id="DPerson" style="Z-INDEX: 143; POSITION: absolute; TOP: 226px; LEFT: 520px" runat="server"
					Width="120px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DDivision" style="Z-INDEX: 142; POSITION: absolute; TOP: 226px; LEFT: 312px"
					runat="server" Width="112px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DReasonDesc" style="Z-INDEX: 141; POSITION: absolute; TOP: 1256px; LEFT: 168px"
					runat="server" Width="424px" Height="59px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" TextMode="MultiLine"
					Visible="False">DReasonDesc</asp:textbox><asp:textbox id="DReason" style="Z-INDEX: 140; POSITION: absolute; TOP: 1224px; LEFT: 240px"
					runat="server" Width="352px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True" Visible="False">DReason</asp:textbox><asp:dropdownlist id="DReasonCode" style="Z-INDEX: 139; POSITION: absolute; TOP: 1224px; LEFT: 168px"
					runat="server" Width="64px" Height="20px" BackColor="Yellow" ForeColor="Blue" AutoPostBack="True" Visible="False">
					<asp:ListItem Value="01">01</asp:ListItem>
					<asp:ListItem Value="02">02</asp:ListItem>
				</asp:dropdownlist><asp:hyperlink id="LBefOP" style="Z-INDEX: 138; POSITION: absolute; TOP: 1184px; LEFT: 168px" runat="server"
					Width="144px" Height="8px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">工程履歷</asp:hyperlink><asp:textbox id="DAEndTime" style="Z-INDEX: 137; POSITION: absolute; TOP: 1150px; LEFT: 440px"
					runat="server" Width="152px" Height="20px" BackColor="Gold" ForeColor="Blue" Font-Size="9pt" BorderStyle="Groove" ReadOnly="True">DAEndTime</asp:textbox><asp:textbox id="DAStartTime" style="Z-INDEX: 136; POSITION: absolute; TOP: 1150px; LEFT: 168px"
					runat="server" Width="152px" Height="20px" BackColor="Gold" ForeColor="Blue" Font-Size="9pt" BorderStyle="Groove" ReadOnly="True">DAStartTime</asp:textbox><asp:textbox id="DBEndTime" style="Z-INDEX: 135; POSITION: absolute; TOP: 1116px; LEFT: 440px"
					runat="server" Width="152px" Height="20px" BackColor="Gold" ForeColor="Blue" Font-Size="9pt" BorderStyle="Groove" ReadOnly="True">DBEndTime</asp:textbox><asp:textbox id="DBStartTime" style="Z-INDEX: 134; POSITION: absolute; TOP: 1116px; LEFT: 168px"
					runat="server" Width="152px" Height="20px" BackColor="Gold" ForeColor="Blue" Font-Size="9pt" BorderStyle="Groove" ReadOnly="True">DBStartTime</asp:textbox><asp:textbox id="DDecideDesc" style="Z-INDEX: 133; POSITION: absolute; TOP: 1040px; LEFT: 56px"
					runat="server" Width="536px" Height="56px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" TextMode="MultiLine">DDecideDesc</asp:textbox><asp:textbox id="DMoldPoint" style="Z-INDEX: 132; POSITION: absolute; TOP: 395px; LEFT: 672px"
					runat="server" Width="45px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">MoldPoint</asp:textbox><asp:textbox id="DMoldQty" style="Z-INDEX: 131; POSITION: absolute; TOP: 395px; LEFT: 584px"
					runat="server" Width="48px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DMoldQty</asp:textbox><asp:textbox id="DPullerPrice" style="Z-INDEX: 130; POSITION: absolute; TOP: 640px; LEFT: 688px"
					runat="server" Width="72px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">PullerPrice</asp:textbox><asp:textbox id="DPurMold" style="Z-INDEX: 129; POSITION: absolute; TOP: 608px; LEFT: 688px"
					runat="server" Width="72px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">PurMold</asp:textbox><asp:textbox id="DArMoldFee" style="Z-INDEX: 128; POSITION: absolute; TOP: 574px; LEFT: 688px"
					runat="server" Width="72px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">ArMoldFee</asp:textbox><asp:textbox id="DDevReason" style="Z-INDEX: 127; POSITION: absolute; TOP: 452px; LEFT: 16px"
					runat="server" Width="408px" Height="80px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" TextMode="MultiLine">DevReason</asp:textbox><asp:hyperlink id="LAuthorizeFile" style="Z-INDEX: 126; POSITION: absolute; TOP: 968px; LEFT: 496px"
					runat="server" Width="50px" Height="8px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">授權書</asp:hyperlink><INPUT id="DAuthorizeFile" style="Z-INDEX: 125; POSITION: absolute; BACKGROUND-COLOR: #ffff00; WIDTH: 260px; HEIGHT: 20px; TOP: 968px; LEFT: 496px"
					type="file" size="24" name="File1" runat="server">
				<asp:hyperlink id="LConfirmFile" style="Z-INDEX: 124; POSITION: absolute; TOP: 968px; LEFT: 120px"
					runat="server" Width="50px" Height="8px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">確認書</asp:hyperlink><INPUT id="DConfirmFile" style="Z-INDEX: 123; POSITION: absolute; BACKGROUND-COLOR: #ffff00; WIDTH: 240px; HEIGHT: 20px; TOP: 968px; LEFT: 120px"
					type="file" name="File1" runat="server"><asp:textbox id="DSellVendor" style="Z-INDEX: 122; POSITION: absolute; TOP: 328px; LEFT: 584px"
					runat="server" Width="174px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DSellVendor</asp:textbox><asp:dropdownlist id="DMaterial" style="Z-INDEX: 121; POSITION: absolute; TOP: 395px; LEFT: 104px"
					runat="server" Width="192px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DManufPlace" style="Z-INDEX: 120; POSITION: absolute; TOP: 362px; LEFT: 400px"
					runat="server" Width="88px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DSliderType2" style="Z-INDEX: 119; POSITION: absolute; TOP: 362px; LEFT: 200px"
					runat="server" Width="94px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DSliderType1" style="Z-INDEX: 118; POSITION: absolute; TOP: 362px; LEFT: 104px"
					runat="server" Width="94px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DAssembler" style="Z-INDEX: 117; POSITION: absolute; TOP: 293px; LEFT: 648px"
					runat="server" Width="110px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DAssembler</asp:textbox><asp:dropdownlist id="DLevel" style="Z-INDEX: 116; POSITION: absolute; TOP: 259px; LEFT: 648px" runat="server"
					Width="110px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><INPUT id="DRefFile" style="Z-INDEX: 115; POSITION: absolute; BACKGROUND-COLOR: #ffff00; WIDTH: 240px; HEIGHT: 20px; TOP: 932px; LEFT: 120px"
					type="file" name="File1" runat="server">
				<asp:button id="BOMapNo" style="Z-INDEX: 114; POSITION: absolute; TOP: 293px; LEFT: 424px" runat="server"
					Width="20px" Height="20px" Text="原"></asp:button><asp:textbox id="DMapNo" style="Z-INDEX: 113; POSITION: absolute; TOP: 293px; LEFT: 312px" runat="server"
					Width="112px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True" AutoPostBack="True">123456789012345</asp:textbox><asp:button id="BSpec" style="Z-INDEX: 112; POSITION: absolute; TOP: 259px; LEFT: 528px" runat="server"
					Width="20px" Height="20px" Text="....."></asp:button><asp:textbox id="DSpec" style="Z-INDEX: 111; POSITION: absolute; TOP: 259px; LEFT: 312px" runat="server"
					Width="216px" Height="20px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove" ReadOnly="True">DSpec</asp:textbox><asp:textbox id="DSliderGRCode" style="Z-INDEX: 110; POSITION: absolute; TOP: 158px; LEFT: 604px"
					runat="server" Width="154px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True">DSliderGRCode</asp:textbox><asp:textbox id="DNo" style="Z-INDEX: 109; POSITION: absolute; TOP: 192px; LEFT: 312px" runat="server"
					Width="112px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DNo</asp:textbox><asp:textbox id="DFormSno" style="Z-INDEX: 108; POSITION: absolute; TOP: 1320px; LEFT: 16px"
					runat="server" Width="97px" Height="20px" BackColor="White" ForeColor="Blue" BorderStyle="None">單號：123</asp:textbox><asp:image id="DDelivery" style="Z-INDEX: 106; POSITION: absolute; TOP: 1104px; LEFT: 8px"
					runat="server" Width="593px" Height="110px" ImageUrl="Images\Sheet_Delivery.jpg"></asp:image><asp:image id="DDelay" style="Z-INDEX: 104; POSITION: absolute; TOP: 1216px; LEFT: 8px" runat="server"
					Width="593px" Height="107px" ImageUrl="Images\Sheet_Delay.jpg" Visible="False"></asp:image><asp:image id="DDescSheet" style="Z-INDEX: 103; POSITION: absolute; TOP: 1032px; LEFT: 8px"
					runat="server" Width="593px" Height="75px" ImageUrl="Images\MapSheet_004.jpg"></asp:image><asp:image id="DManuaInSheet2" style="Z-INDEX: 102; POSITION: absolute; TOP: 536px; LEFT: 8px"
					runat="server" Width="761px" Height="493px" ImageUrl="Images\ManufInSheet_005_A.jpg"></asp:image><asp:image id="LMapFile" style="Z-INDEX: 246; POSITION: absolute; TOP: 120px; LEFT: 16px" runat="server"
					Width="200px" Height="230px" BorderStyle="Groove" ImageUrl="http://10.245.0.178//WorkFlow/Document/000003/1-Map2-2006124111750-2.jpg"></asp:image></FONT><INPUT id="BSAVE" style="Z-INDEX: 250; POSITION: absolute; WIDTH: 80px; HEIGHT: 28px; TOP: 1328px; LEFT: 256px"
				onclick="Button('SAVE');" type="button" value="儲存" name="Button1" runat="server"><INPUT id="BNG2" style="Z-INDEX: 251; POSITION: absolute; WIDTH: 80px; HEIGHT: 28px; TOP: 1328px; LEFT: 344px"
				onclick="Button('NG2');" type="button" value="NG2" name="Button1" runat="server"><INPUT id="BNG1" style="Z-INDEX: 252; POSITION: absolute; WIDTH: 80px; HEIGHT: 28px; TOP: 1328px; LEFT: 432px"
				onclick="Button('NG1');" type="button" value="NG1" name="Button1" runat="server"><INPUT id="BOK" style="Z-INDEX: 253; POSITION: absolute; WIDTH: 80px; HEIGHT: 28px; TOP: 1328px; LEFT: 520px"
				onclick="Button('OK');" type="button" value="OK" name="Button2" runat="server">
			<asp:imagebutton id="BPrint" style="Z-INDEX: 263; POSITION: absolute; TOP: 8px; LEFT: 8px" runat="server"
				Width="16px" Height="16px" ImageUrl="Images\Print.gif"></asp:imagebutton>
			<asp:Label id="Lable27" style="Z-INDEX: 314; POSITION: absolute; TOP: 24px; LEFT: 208px" runat="server"
				Font-Size="16pt" ForeColor="Red" Width="88px" BorderColor="White" Font-Bold="True">〔舊版〕</asp:Label></form>
		</FONT></FORM>
	</body>
</HTML>
