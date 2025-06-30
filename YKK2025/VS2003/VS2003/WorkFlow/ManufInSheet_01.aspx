<%@ Page Language="vb" AspCompat="true" AutoEventWireup="false" Codebehind="ManufInSheet_01.aspx.vb" Inherits="SPD.ManufInSheet_01"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>內製委託書</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
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
		<form id="Form1" encType="multipart/form-data" method="post" runat="server">
			<FONT face="新細明體">
				<asp:textbox style="Z-INDEX: 315; LEFT: 488px; POSITION: absolute; TOP: 328px" id="DCustomerCode"
					runat="server" BorderStyle="Groove" Width="104px" Height="20px" BackColor="Yellow" ForeColor="Blue">DCustomerCode</asp:textbox>
				<asp:hyperlink style="Z-INDEX: 387; LEFT: 400px; POSITION: absolute; TOP: 902px" id="LAssemblerFile"
					runat="server" Height="10px" Width="80px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">組立判定書</asp:hyperlink><INPUT style="Z-INDEX: 386; LEFT: 400px; WIDTH: 192px; POSITION: absolute; TOP: 900px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					id="DAssemblerFile" size="12" type="file" name="File1" runat="server">
				<asp:label style="Z-INDEX: 384; LEFT: 312px; POSITION: absolute; TOP: 900px" id="Label42" runat="server"
					Height="20px" Width="80px">組立判定書</asp:label><asp:dropdownlist style="Z-INDEX: 323; LEFT: 648px; POSITION: absolute; TOP: 256px" id="DLevel" runat="server"
					Width="112px" Height="30px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 322; LEFT: 688px; POSITION: absolute; TOP: 640px" id="DOriginalDep"
					runat="server" BorderStyle="Groove" Width="64px" Height="20px" BackColor="Yellow" ForeColor="Blue">OriginalDep</asp:textbox><asp:image style="Z-INDEX: 321; LEFT: 560px; POSITION: absolute; TOP: 8px" id="DMMSSts" runat="server"
					BorderStyle="None" Width="206px" Height="68px" ImageUrl="Images\mmsdelete.jpg"></asp:image><asp:dropdownlist style="Z-INDEX: 320; LEFT: 712px; POSITION: absolute; TOP: 750px" id="DQAD15" runat="server"
					Width="48px" Height="22px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 319; LEFT: 712px; POSITION: absolute; TOP: 732px" id="DQAC15" runat="server"
					Width="48px" Height="22px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 318; LEFT: 712px; POSITION: absolute; TOP: 714px" id="DQAB15" runat="server"
					Width="48px" Height="22px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:label style="Z-INDEX: 317; LEFT: 712px; POSITION: absolute; TOP: 684px" id="Label29" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">EDX</asp:label><asp:dropdownlist style="Z-INDEX: 316; LEFT: 712px; POSITION: absolute; TOP: 696px" id="DQAA15" runat="server"
					Width="48px" Height="22px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 314; LEFT: 686px; POSITION: absolute; TOP: 224px" id="DLogo" runat="server"
					Width="78px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 313; LEFT: 56px; POSITION: absolute; TOP: 480px" id="DManufFlow"
					runat="server" BorderStyle="Groove" Width="369px" Height="48px" BackColor="Yellow" ForeColor="Blue" TextMode="MultiLine"
					MaxLength="240">ManufFlow</asp:textbox><asp:textbox style="Z-INDEX: 127; LEFT: 56px; POSITION: absolute; TOP: 424px" id="DDevReason"
					runat="server" BorderStyle="Groove" Width="369px" Height="48px" BackColor="Yellow" ForeColor="Blue" TextMode="MultiLine" MaxLength="240">DevReason</asp:textbox><asp:dropdownlist style="Z-INDEX: 304; LEFT: 616px; POSITION: absolute; TOP: 750px" id="DQAD13" runat="server"
					Width="48px" Height="346px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0級">0級</asp:ListItem>
					<asp:ListItem Value="1級">1級</asp:ListItem>
					<asp:ListItem Value="2級">2級</asp:ListItem>
					<asp:ListItem Value="3級">3級</asp:ListItem>
					<asp:ListItem Value="4級">4級</asp:ListItem>
					<asp:ListItem Value="5級">5級</asp:ListItem>
				</asp:dropdownlist><asp:label style="Z-INDEX: 312; LEFT: 744px; POSITION: absolute; TOP: 902px" id="Label28" runat="server"
					Width="17px" Height="10px">分</asp:label><asp:label style="Z-INDEX: 311; LEFT: 600px; POSITION: absolute; TOP: 902px" id="Label27" runat="server"
					Width="49px" Height="10px">QC-L/T</asp:label><asp:textbox style="Z-INDEX: 310; LEFT: 656px; POSITION: absolute; TOP: 900px" id="DQCLT" runat="server"
					BorderStyle="Groove" Width="81px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt"></asp:textbox><asp:textbox style="Z-INDEX: 309; LEFT: 248px; POSITION: absolute; TOP: 1184px" id="DOPReadyDesc"
					runat="server" BorderStyle="None" Width="104px" Height="20px" BackColor="White" ForeColor="Red">需閱讀工程履歷</asp:textbox><asp:textbox style="Z-INDEX: 153; LEFT: 352px; POSITION: absolute; TOP: 1184px" id="DOPReady"
					runat="server" BorderStyle="Groove" Width="48px" Height="20px" BackColor="GreenYellow" ForeColor="Red" ReadOnly="True"></asp:textbox><asp:dropdownlist style="Z-INDEX: 308; LEFT: 708px; POSITION: absolute; TOP: 192px" id="DCpsc" runat="server"
					Width="56px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 307; LEFT: 584px; POSITION: absolute; TOP: 360px" id="DMakeCAM"
					runat="server" Width="176px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 303; LEFT: 616px; POSITION: absolute; TOP: 732px" id="DQAC13" runat="server"
					Width="48px" Height="346px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0級">0級</asp:ListItem>
					<asp:ListItem Value="1級">1級</asp:ListItem>
					<asp:ListItem Value="2級">2級</asp:ListItem>
					<asp:ListItem Value="3級">3級</asp:ListItem>
					<asp:ListItem Value="4級">4級</asp:ListItem>
					<asp:ListItem Value="5級">5級</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 302; LEFT: 616px; POSITION: absolute; TOP: 714px" id="DQAB13" runat="server"
					Width="48px" Height="346px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0級">0級</asp:ListItem>
					<asp:ListItem Value="1級">1級</asp:ListItem>
					<asp:ListItem Value="2級">2級</asp:ListItem>
					<asp:ListItem Value="3級">3級</asp:ListItem>
					<asp:ListItem Value="4級">4級</asp:ListItem>
					<asp:ListItem Value="5級">5級</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 301; LEFT: 568px; POSITION: absolute; TOP: 750px" id="DQAD12" runat="server"
					Width="48px" Height="346px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0級">0級</asp:ListItem>
					<asp:ListItem Value="1級">1級</asp:ListItem>
					<asp:ListItem Value="2級">2級</asp:ListItem>
					<asp:ListItem Value="3級">3級</asp:ListItem>
					<asp:ListItem Value="4級">4級</asp:ListItem>
					<asp:ListItem Value="5級">5級</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 300; LEFT: 568px; POSITION: absolute; TOP: 732px" id="DQAC12" runat="server"
					Width="48px" Height="346px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0級">0級</asp:ListItem>
					<asp:ListItem Value="1級">1級</asp:ListItem>
					<asp:ListItem Value="2級">2級</asp:ListItem>
					<asp:ListItem Value="3級">3級</asp:ListItem>
					<asp:ListItem Value="4級">4級</asp:ListItem>
					<asp:ListItem Value="5級">5級</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 299; LEFT: 568px; POSITION: absolute; TOP: 714px" id="DQAB12" runat="server"
					Width="48px" Height="346px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0級">0級</asp:ListItem>
					<asp:ListItem Value="1級">1級</asp:ListItem>
					<asp:ListItem Value="2級">2級</asp:ListItem>
					<asp:ListItem Value="3級">3級</asp:ListItem>
					<asp:ListItem Value="4級">4級</asp:ListItem>
					<asp:ListItem Value="5級">5級</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 298; LEFT: 520px; POSITION: absolute; TOP: 750px" id="DQAD11" runat="server"
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
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 297; LEFT: 520px; POSITION: absolute; TOP: 732px" id="DQAC11" runat="server"
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
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 296; LEFT: 520px; POSITION: absolute; TOP: 714px" id="DQAB11" runat="server"
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
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 295; LEFT: 616px; POSITION: absolute; TOP: 696px" id="DQAA13" runat="server"
					Width="48px" Height="346px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0級">0級</asp:ListItem>
					<asp:ListItem Value="1級">1級</asp:ListItem>
					<asp:ListItem Value="2級">2級</asp:ListItem>
					<asp:ListItem Value="3級">3級</asp:ListItem>
					<asp:ListItem Value="4級">4級</asp:ListItem>
					<asp:ListItem Value="5級">5級</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 294; LEFT: 568px; POSITION: absolute; TOP: 696px" id="DQAA12" runat="server"
					Width="48px" Height="346px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0級">0級</asp:ListItem>
					<asp:ListItem Value="1級">1級</asp:ListItem>
					<asp:ListItem Value="2級">2級</asp:ListItem>
					<asp:ListItem Value="3級">3級</asp:ListItem>
					<asp:ListItem Value="4級">4級</asp:ListItem>
					<asp:ListItem Value="5級">5級</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 293; LEFT: 520px; POSITION: absolute; TOP: 696px" id="DQAA11" runat="server"
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
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 237; LEFT: 664px; POSITION: absolute; TOP: 750px" id="DQAD14" runat="server"
					Width="48px" Height="22px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 236; LEFT: 472px; POSITION: absolute; TOP: 750px" id="DQAD10" runat="server"
					Width="48px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 235; LEFT: 424px; POSITION: absolute; TOP: 750px" id="DQAD9" runat="server"
					Width="48px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 234; LEFT: 376px; POSITION: absolute; TOP: 750px" id="DQAD8" runat="server"
					Width="48px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="NCB">NCB</asp:ListItem>
					<asp:ListItem Value="KENSIN">KENSIN</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 233; LEFT: 328px; POSITION: absolute; TOP: 750px" id="DQAD7" runat="server"
					Width="48px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 232; LEFT: 280px; POSITION: absolute; TOP: 750px" id="DQAD6" runat="server"
					Width="48px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 231; LEFT: 232px; POSITION: absolute; TOP: 750px" id="DQAD5" runat="server"
					BorderStyle="Groove" Width="47px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt"></asp:textbox><asp:textbox style="Z-INDEX: 230; LEFT: 204px; POSITION: absolute; TOP: 750px" id="DQAD4" runat="server"
					BorderStyle="Groove" Width="28px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt"></asp:textbox><asp:dropdownlist style="Z-INDEX: 229; LEFT: 170px; POSITION: absolute; TOP: 750px" id="DQAD3" runat="server"
					Width="35px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 228; LEFT: 72px; POSITION: absolute; TOP: 750px" id="DQAD2" runat="server"
					BorderStyle="Groove" Width="98px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt"></asp:textbox><asp:textbox style="Z-INDEX: 227; LEFT: 16px; POSITION: absolute; TOP: 750px" id="DQAD1" runat="server"
					BorderStyle="Groove" Width="56px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt"></asp:textbox><asp:dropdownlist style="Z-INDEX: 226; LEFT: 664px; POSITION: absolute; TOP: 732px" id="DQAC14" runat="server"
					Width="48px" Height="22px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 225; LEFT: 472px; POSITION: absolute; TOP: 732px" id="DQAC10" runat="server"
					Width="48px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 224; LEFT: 424px; POSITION: absolute; TOP: 732px" id="DQAC9" runat="server"
					Width="48px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 223; LEFT: 376px; POSITION: absolute; TOP: 732px" id="DQAC8" runat="server"
					Width="48px" Height="36px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="NCB">NCB</asp:ListItem>
					<asp:ListItem Value="KENSIN">KENSIN</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 222; LEFT: 328px; POSITION: absolute; TOP: 732px" id="DQAC7" runat="server"
					Width="48px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 221; LEFT: 280px; POSITION: absolute; TOP: 732px" id="DQAC6" runat="server"
					Width="48px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 220; LEFT: 232px; POSITION: absolute; TOP: 732px" id="DQAC5" runat="server"
					BorderStyle="Groove" Width="47px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt"></asp:textbox><asp:textbox style="Z-INDEX: 219; LEFT: 204px; POSITION: absolute; TOP: 732px" id="DQAC4" runat="server"
					BorderStyle="Groove" Width="28px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt"></asp:textbox><asp:dropdownlist style="Z-INDEX: 218; LEFT: 170px; POSITION: absolute; TOP: 732px" id="DQAC3" runat="server"
					Width="35px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 217; LEFT: 72px; POSITION: absolute; TOP: 732px" id="DQAC2" runat="server"
					BorderStyle="Groove" Width="98px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt"></asp:textbox><asp:textbox style="Z-INDEX: 216; LEFT: 16px; POSITION: absolute; TOP: 732px" id="DQAC1" runat="server"
					BorderStyle="Groove" Width="56px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt"></asp:textbox><asp:dropdownlist style="Z-INDEX: 215; LEFT: 664px; POSITION: absolute; TOP: 714px" id="DQAB14" runat="server"
					Width="48px" Height="22px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 214; LEFT: 472px; POSITION: absolute; TOP: 714px" id="DQAB10" runat="server"
					Width="48px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 213; LEFT: 424px; POSITION: absolute; TOP: 714px" id="DQAB9" runat="server"
					Width="48px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 212; LEFT: 376px; POSITION: absolute; TOP: 714px" id="DQAB8" runat="server"
					Width="48px" Height="36px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="NCB">NCB</asp:ListItem>
					<asp:ListItem Value="KENSIN">KENSIN</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 211; LEFT: 328px; POSITION: absolute; TOP: 714px" id="DQAB7" runat="server"
					Width="48px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 210; LEFT: 280px; POSITION: absolute; TOP: 714px" id="DQAB6" runat="server"
					Width="48px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 209; LEFT: 232px; POSITION: absolute; TOP: 714px" id="DQAB5" runat="server"
					BorderStyle="Groove" Width="47px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt"></asp:textbox><asp:textbox style="Z-INDEX: 208; LEFT: 204px; POSITION: absolute; TOP: 714px" id="DQAB4" runat="server"
					BorderStyle="Groove" Width="28px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt"></asp:textbox><asp:dropdownlist style="Z-INDEX: 207; LEFT: 170px; POSITION: absolute; TOP: 714px" id="DQAB3" runat="server"
					Width="35px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 206; LEFT: 72px; POSITION: absolute; TOP: 714px" id="DQAB2" runat="server"
					BorderStyle="Groove" Width="98px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt"></asp:textbox><asp:textbox style="Z-INDEX: 205; LEFT: 16px; POSITION: absolute; TOP: 714px" id="DQAB1" runat="server"
					BorderStyle="Groove" Width="56px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt"></asp:textbox><asp:dropdownlist style="Z-INDEX: 162; LEFT: 664px; POSITION: absolute; TOP: 696px" id="DQAA14" runat="server"
					Width="48px" Height="22px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 161; LEFT: 472px; POSITION: absolute; TOP: 696px" id="DQAA10" runat="server"
					Width="48px" Height="346px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 160; LEFT: 424px; POSITION: absolute; TOP: 696px" id="DQAA9" runat="server"
					Width="48px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 159; LEFT: 376px; POSITION: absolute; TOP: 696px" id="DQAA8" runat="server"
					Width="48px" Height="36px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="NCB">NCB</asp:ListItem>
					<asp:ListItem Value="KENSIN">KENSIN</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 158; LEFT: 328px; POSITION: absolute; TOP: 696px" id="DQAA7" runat="server"
					Width="48px" Height="162px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 157; LEFT: 280px; POSITION: absolute; TOP: 696px" id="DQAA6" runat="server"
					Width="48px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 156; LEFT: 232px; POSITION: absolute; TOP: 696px" id="DQAA5" runat="server"
					BorderStyle="Groove" Width="47px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt"></asp:textbox><asp:textbox style="Z-INDEX: 155; LEFT: 204px; POSITION: absolute; TOP: 696px" id="DQAA4" runat="server"
					BorderStyle="Groove" Width="28px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt"></asp:textbox><asp:dropdownlist style="Z-INDEX: 154; LEFT: 170px; POSITION: absolute; TOP: 696px" id="DQAA3" runat="server"
					Width="35px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 152; LEFT: 16px; POSITION: absolute; TOP: 696px" id="DQAA1" runat="server"
					BorderStyle="Groove" Width="56px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt"></asp:textbox><asp:textbox style="Z-INDEX: 171; LEFT: 72px; POSITION: absolute; TOP: 696px" id="DQAA2" runat="server"
					BorderStyle="Groove" Width="98px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt"></asp:textbox><asp:image style="Z-INDEX: 100; LEFT: 8px; POSITION: absolute; TOP: 8px" id="DManuaInSheet1"
					runat="server" Width="761px" Height="530px" ImageUrl="Images\ManufInSheet_003_E.jpg"></asp:image><asp:dropdownlist style="Z-INDEX: 305; LEFT: 400px; POSITION: absolute; TOP: 395px" id="DSuppiler"
					runat="server" Width="88px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:label style="Z-INDEX: 292; LEFT: 704px; POSITION: absolute; TOP: 440px" id="Label25" runat="server"
					Width="24px" Height="10px" Font-Size="8pt">數量</asp:label><asp:label style="Z-INDEX: 291; LEFT: 648px; POSITION: absolute; TOP: 440px" id="Label24" runat="server"
					Width="24px" Height="10px" Font-Size="8pt">顏色</asp:label><asp:label style="Z-INDEX: 290; LEFT: 472px; POSITION: absolute; TOP: 440px" id="Label23" runat="server"
					Width="48px" Height="10px" Font-Size="8pt">拉頭種類</asp:label><asp:label style="Z-INDEX: 289; LEFT: 48px; POSITION: absolute; TOP: 684px" id="Label1" runat="server"
					Width="24px" Height="10px" Font-Size="8pt">日期</asp:label><asp:label style="Z-INDEX: 288; LEFT: 504px; POSITION: absolute; TOP: 556px" id="Label22" runat="server"
					Width="24px" Height="10px" Font-Size="8pt">單價</asp:label><asp:label style="Z-INDEX: 287; LEFT: 464px; POSITION: absolute; TOP: 556px" id="Label21" runat="server"
					Width="24px" Height="10px" Font-Size="8pt">幣別</asp:label><asp:label style="Z-INDEX: 286; LEFT: 296px; POSITION: absolute; TOP: 556px" id="Label20" runat="server"
					Width="48px" Height="10px" Font-Size="8pt">拉頭種類</asp:label><asp:label style="Z-INDEX: 285; LEFT: 224px; POSITION: absolute; TOP: 556px" id="Label19" runat="server"
					Width="24px" Height="10px" Font-Size="8pt">單價</asp:label><asp:label style="Z-INDEX: 284; LEFT: 184px; POSITION: absolute; TOP: 556px" id="Label18" runat="server"
					Width="24px" Height="10px" Font-Size="8pt">幣別</asp:label><asp:label style="Z-INDEX: 283; LEFT: 56px; POSITION: absolute; TOP: 556px" id="Label17" runat="server"
					Width="48px" Height="10px" Font-Size="8pt">拉頭種類</asp:label><asp:label style="Z-INDEX: 282; LEFT: 136px; POSITION: absolute; TOP: 812px" id="Label16" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">備註</asp:label><asp:label style="Z-INDEX: 281; LEFT: 80px; POSITION: absolute; TOP: 812px" id="Label15" runat="server"
					Width="48px" Height="10px" Font-Size="8pt">檢測結果</asp:label><asp:label style="Z-INDEX: 280; LEFT: 48px; POSITION: absolute; TOP: 812px" id="Label14" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">日期</asp:label><asp:label style="Z-INDEX: 279; LEFT: 664px; POSITION: absolute; TOP: 684px" id="Label13" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">CPSC</asp:label><asp:label style="Z-INDEX: 306; LEFT: 568px; POSITION: absolute; TOP: 684px" id="Label26" runat="server"
					Width="48px" Height="10px" Font-Size="8pt">一次密著</asp:label><asp:label style="Z-INDEX: 278; LEFT: 616px; POSITION: absolute; TOP: 684px" id="Label12" runat="server"
					Width="48px" Height="10px" Font-Size="8pt">二次密著</asp:label><asp:label style="Z-INDEX: 277; LEFT: 520px; POSITION: absolute; TOP: 684px" id="Label11" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">黃變</asp:label><asp:label style="Z-INDEX: 276; LEFT: 472px; POSITION: absolute; TOP: 684px" id="Label10" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">乾洗</asp:label><asp:label style="Z-INDEX: 275; LEFT: 424px; POSITION: absolute; TOP: 684px" id="Label9" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">水洗</asp:label><asp:label style="Z-INDEX: 274; LEFT: 376px; POSITION: absolute; TOP: 684px" id="Label8" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">檢針</asp:label><asp:label style="Z-INDEX: 273; LEFT: 328px; POSITION: absolute; TOP: 684px" id="Label7" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">扭力</asp:label><asp:label style="Z-INDEX: 272; LEFT: 280px; POSITION: absolute; TOP: 684px" id="Label6" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">強度</asp:label><asp:label style="Z-INDEX: 271; LEFT: 240px; POSITION: absolute; TOP: 684px" id="Label5" runat="server"
					Width="40px" Height="10px" Font-Size="8pt">原單位</asp:label><asp:label style="Z-INDEX: 270; LEFT: 208px; POSITION: absolute; TOP: 684px" id="Label4" runat="server"
					Width="48px" Height="10px" Font-Size="8pt">表面</asp:label><asp:label style="Z-INDEX: 269; LEFT: 168px; POSITION: absolute; TOP: 684px" id="Label3" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">口厚</asp:label><asp:label style="Z-INDEX: 268; LEFT: 80px; POSITION: absolute; TOP: 684px" id="Label2" runat="server"
					Width="48px" Height="10px" Font-Size="8pt">拉頭種類</asp:label><asp:dropdownlist style="Z-INDEX: 267; LEFT: 312px; POSITION: absolute; TOP: 328px" id="DBuyer" runat="server"
					Width="176px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:hyperlink style="Z-INDEX: 266; LEFT: 16px; POSITION: absolute; TOP: 772px" id="LQAttachFile1"
					runat="server" Width="64px" Height="8px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">其他附件</asp:hyperlink><asp:hyperlink style="Z-INDEX: 265; LEFT: 136px; POSITION: absolute; TOP: 900px" id="LQAttachFile2"
					runat="server" Width="120px" Height="8px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">品質分析依賴書</asp:hyperlink><INPUT style="Z-INDEX: 264; LEFT: 136px; WIDTH: 176px; POSITION: absolute; TOP: 900px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					id="DQAttachFile2" size="10" type="file" name="File1" runat="server"><INPUT style="Z-INDEX: 263; LEFT: 16px; WIDTH: 744px; POSITION: absolute; TOP: 772px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					id="DQAttachFile1" size="104" type="file" name="File1" runat="server">
				<asp:hyperlink style="Z-INDEX: 262; LEFT: 440px; POSITION: absolute; TOP: 512px" id="LSAttachFile"
					runat="server" Width="64px" Height="8px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">其他附件</asp:hyperlink><INPUT style="Z-INDEX: 261; LEFT: 440px; WIDTH: 314px; POSITION: absolute; TOP: 508px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					id="DSAttachFile" size="31" type="file" name="File1" runat="server"><asp:imagebutton style="Z-INDEX: 260; LEFT: 8px; POSITION: absolute; TOP: 32px" id="BFlow" runat="server"
					Width="16px" Height="16px" ImageUrl="Images\Flow-01.gif"></asp:imagebutton>
				<asp:textbox style="Z-INDEX: 258; LEFT: 136px; POSITION: absolute; TOP: 878px" id="DFQAD3" runat="server"
					BorderStyle="Groove" Width="620px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt"></asp:textbox><asp:dropdownlist style="Z-INDEX: 257; LEFT: 80px; POSITION: absolute; TOP: 878px" id="DFQAD2" runat="server"
					Width="56px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 256; LEFT: 16px; POSITION: absolute; TOP: 878px" id="DFQAD1" runat="server"
					BorderStyle="Groove" Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt"></asp:textbox><asp:textbox style="Z-INDEX: 255; LEFT: 16px; POSITION: absolute; TOP: 860px" id="DFQAC1" runat="server"
					BorderStyle="Groove" Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt"></asp:textbox><asp:dropdownlist style="Z-INDEX: 254; LEFT: 80px; POSITION: absolute; TOP: 860px" id="DFQAC2" runat="server"
					Width="56px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 253; LEFT: 136px; POSITION: absolute; TOP: 860px" id="DFQAC3" runat="server"
					BorderStyle="Groove" Width="620px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt"></asp:textbox><asp:textbox style="Z-INDEX: 252; LEFT: 136px; POSITION: absolute; TOP: 842px" id="DFQAB3" runat="server"
					BorderStyle="Groove" Width="620px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt"></asp:textbox><asp:dropdownlist style="Z-INDEX: 251; LEFT: 80px; POSITION: absolute; TOP: 842px" id="DFQAB2" runat="server"
					Width="56px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 250; LEFT: 16px; POSITION: absolute; TOP: 842px" id="DFQAB1" runat="server"
					BorderStyle="Groove" Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt"></asp:textbox><asp:textbox style="Z-INDEX: 245; LEFT: 480px; POSITION: absolute; TOP: 293px" id="DOFormSno"
					runat="server" BorderStyle="Groove" Width="32px" Height="20px" BackColor="Yellow" ForeColor="Blue" AutoPostBack="True">123</asp:textbox><asp:textbox style="Z-INDEX: 244; LEFT: 432px; POSITION: absolute; TOP: 293px" id="DOFormNo"
					runat="server" BorderStyle="Groove" Width="48px" Height="20px" BackColor="Yellow" ForeColor="Blue" AutoPostBack="True">000001</asp:textbox><INPUT style="Z-INDEX: 243; LEFT: 16px; WIDTH: 200px; POSITION: absolute; TOP: 328px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					id="DMapFile" size="14" type="file" name="File1" runat="server">
				<asp:hyperlink style="Z-INDEX: 241; LEFT: 312px; POSITION: absolute; TOP: 296px" id="LMapNo" runat="server"
					Width="88px" Height="8px" Target="_blank" NavigateUrl="BoardEdit.aspx">圖號文件</asp:hyperlink><asp:textbox style="Z-INDEX: 240; LEFT: 136px; POSITION: absolute; TOP: 824px" id="DFQAA3" runat="server"
					BorderStyle="Groove" Width="620px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt"></asp:textbox><asp:dropdownlist style="Z-INDEX: 239; LEFT: 80px; POSITION: absolute; TOP: 824px" id="DFQAA2" runat="server"
					Width="56px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 238; LEFT: 16px; POSITION: absolute; TOP: 824px" id="DFQAA1" runat="server"
					BorderStyle="Groove" Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt"></asp:textbox><asp:textbox style="Z-INDEX: 202; LEFT: 296px; POSITION: absolute; TOP: 640px" id="DPriceJ1"
					runat="server" BorderStyle="Groove" Width="166px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt"></asp:textbox><asp:textbox style="Z-INDEX: 201; LEFT: 224px; POSITION: absolute; TOP: 640px" id="DPriceI3"
					runat="server" BorderStyle="Groove" Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt"></asp:textbox><asp:dropdownlist style="Z-INDEX: 200; LEFT: 184px; POSITION: absolute; TOP: 640px" id="DPriceI2"
					runat="server" Width="40px" Height="12px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 199; LEFT: 16px; POSITION: absolute; TOP: 640px" id="DPriceI1" runat="server"
					BorderStyle="Groove" Width="166px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt"></asp:textbox><asp:textbox style="Z-INDEX: 198; LEFT: 504px; POSITION: absolute; TOP: 622px" id="DPriceH3"
					runat="server" BorderStyle="Groove" Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt"></asp:textbox><asp:dropdownlist style="Z-INDEX: 197; LEFT: 464px; POSITION: absolute; TOP: 622px" id="DPriceH2"
					runat="server" Width="40px" Height="12px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 204; LEFT: 504px; POSITION: absolute; TOP: 640px" id="DPriceJ3"
					runat="server" BorderStyle="Groove" Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt"></asp:textbox><asp:dropdownlist style="Z-INDEX: 203; LEFT: 464px; POSITION: absolute; TOP: 640px" id="DPriceJ2"
					runat="server" Width="40px" Height="12px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 151; LEFT: 224px; POSITION: absolute; TOP: 568px" id="DPriceA3"
					runat="server" BorderStyle="Groove" Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt"></asp:textbox><asp:dropdownlist style="Z-INDEX: 150; LEFT: 184px; POSITION: absolute; TOP: 568px" id="DPriceA2"
					runat="server" Width="40px" Height="12px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 196; LEFT: 296px; POSITION: absolute; TOP: 622px" id="DPriceH1"
					runat="server" BorderStyle="Groove" Width="166px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt"></asp:textbox><asp:textbox style="Z-INDEX: 195; LEFT: 224px; POSITION: absolute; TOP: 622px" id="DPriceG3"
					runat="server" BorderStyle="Groove" Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt"></asp:textbox><asp:dropdownlist style="Z-INDEX: 194; LEFT: 184px; POSITION: absolute; TOP: 622px" id="DPriceG2"
					runat="server" Width="40px" Height="12px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 193; LEFT: 16px; POSITION: absolute; TOP: 622px" id="DPriceG1" runat="server"
					BorderStyle="Groove" Width="166px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt"></asp:textbox><asp:textbox style="Z-INDEX: 192; LEFT: 504px; POSITION: absolute; TOP: 604px" id="DPriceF3"
					runat="server" BorderStyle="Groove" Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt"></asp:textbox><asp:dropdownlist style="Z-INDEX: 191; LEFT: 464px; POSITION: absolute; TOP: 604px" id="DPriceF2"
					runat="server" Width="40px" Height="12px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 190; LEFT: 296px; POSITION: absolute; TOP: 604px" id="DPriceF1"
					runat="server" BorderStyle="Groove" Width="166px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt"></asp:textbox><asp:textbox style="Z-INDEX: 189; LEFT: 224px; POSITION: absolute; TOP: 604px" id="DPriceE3"
					runat="server" BorderStyle="Groove" Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt"></asp:textbox><asp:dropdownlist style="Z-INDEX: 188; LEFT: 184px; POSITION: absolute; TOP: 604px" id="DPriceE2"
					runat="server" Width="40px" Height="12px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 187; LEFT: 16px; POSITION: absolute; TOP: 604px" id="DPriceE1" runat="server"
					BorderStyle="Groove" Width="166px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt"></asp:textbox><asp:textbox style="Z-INDEX: 186; LEFT: 504px; POSITION: absolute; TOP: 586px" id="DPriceD3"
					runat="server" BorderStyle="Groove" Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt"></asp:textbox><asp:dropdownlist style="Z-INDEX: 185; LEFT: 464px; POSITION: absolute; TOP: 586px" id="DPriceD2"
					runat="server" Width="40px" Height="12px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 184; LEFT: 296px; POSITION: absolute; TOP: 586px" id="DPriceD1"
					runat="server" BorderStyle="Groove" Width="166px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt"></asp:textbox><asp:textbox style="Z-INDEX: 183; LEFT: 224px; POSITION: absolute; TOP: 586px" id="DPriceC3"
					runat="server" BorderStyle="Groove" Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt"></asp:textbox><asp:dropdownlist style="Z-INDEX: 182; LEFT: 184px; POSITION: absolute; TOP: 586px" id="DPriceC2"
					runat="server" Width="40px" Height="12px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 181; LEFT: 16px; POSITION: absolute; TOP: 586px" id="DPriceC1" runat="server"
					BorderStyle="Groove" Width="166px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt"></asp:textbox><asp:textbox style="Z-INDEX: 180; LEFT: 504px; POSITION: absolute; TOP: 568px" id="DPriceB3"
					runat="server" BorderStyle="Groove" Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt"></asp:textbox><asp:dropdownlist style="Z-INDEX: 179; LEFT: 464px; POSITION: absolute; TOP: 568px" id="DPriceB2"
					runat="server" Width="40px" Height="12px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 178; LEFT: 296px; POSITION: absolute; TOP: 568px" id="DPriceB1"
					runat="server" BorderStyle="Groove" Width="166px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt"></asp:textbox><asp:textbox style="Z-INDEX: 177; LEFT: 704px; POSITION: absolute; TOP: 488px" id="DSampleC3"
					runat="server" BorderStyle="Groove" Width="52px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt"></asp:textbox><asp:textbox style="Z-INDEX: 176; LEFT: 648px; POSITION: absolute; TOP: 488px" id="DSampleC2"
					runat="server" BorderStyle="Groove" Width="52px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt"></asp:textbox><asp:textbox style="Z-INDEX: 175; LEFT: 440px; POSITION: absolute; TOP: 488px" id="DSampleC1"
					runat="server" BorderStyle="Groove" Width="204px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt"></asp:textbox><asp:textbox style="Z-INDEX: 174; LEFT: 704px; POSITION: absolute; TOP: 470px" id="DSampleB3"
					runat="server" BorderStyle="Groove" Width="52px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt"></asp:textbox><asp:textbox style="Z-INDEX: 173; LEFT: 648px; POSITION: absolute; TOP: 470px" id="DSampleB2"
					runat="server" BorderStyle="Groove" Width="52px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt"></asp:textbox><asp:textbox style="Z-INDEX: 172; LEFT: 440px; POSITION: absolute; TOP: 470px" id="DSampleB1"
					runat="server" BorderStyle="Groove" Width="204px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt"></asp:textbox><asp:textbox style="Z-INDEX: 170; LEFT: 16px; POSITION: absolute; TOP: 568px" id="DPriceA1" runat="server"
					BorderStyle="Groove" Width="166px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt"></asp:textbox><asp:textbox style="Z-INDEX: 169; LEFT: 440px; POSITION: absolute; TOP: 452px" id="DSampleA1"
					runat="server" BorderStyle="Groove" Width="204px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt"></asp:textbox><asp:button style="Z-INDEX: 168; LEFT: 568px; POSITION: absolute; TOP: 120px" id="BSliderCode"
					runat="server" Width="20px" Height="20px" Text="....."></asp:button><asp:hyperlink style="Z-INDEX: 167; LEFT: 120px; POSITION: absolute; TOP: 1000px" id="LContactFile"
					runat="server" Width="50px" Height="8px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">切結書</asp:hyperlink><INPUT style="Z-INDEX: 166; LEFT: 120px; WIDTH: 240px; POSITION: absolute; TOP: 1000px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					id="DContactFile" type="file" name="File1" runat="server">
				<asp:hyperlink style="Z-INDEX: 165; LEFT: 120px; POSITION: absolute; TOP: 932px" id="LRefFile"
					runat="server" Width="64px" Height="8px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">其他文件</asp:hyperlink><asp:hyperlink style="Z-INDEX: 164; LEFT: 496px; POSITION: absolute; TOP: 932px" id="LQAFile" runat="server"
					Width="64px" Height="8px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">測試報告</asp:hyperlink><INPUT style="Z-INDEX: 163; LEFT: 496px; WIDTH: 260px; POSITION: absolute; TOP: 932px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					id="DQAFile" size="24" type="file" name="File1" runat="server"><asp:textbox style="Z-INDEX: 104; LEFT: 648px; POSITION: absolute; TOP: 452px" id="DSampleA2"
					runat="server" BorderStyle="Groove" Width="52px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt"></asp:textbox>
				<asp:textbox style="Z-INDEX: 106; LEFT: 704px; POSITION: absolute; TOP: 452px" id="DSampleA3"
					runat="server" BorderStyle="Groove" Width="52px" Height="18px" BackColor="Yellow" ForeColor="Blue"
					Font-Size="8pt"></asp:textbox><asp:button style="Z-INDEX: 149; LEFT: 616px; POSITION: absolute; TOP: 192px" id="BDate" runat="server"
					Width="20px" Height="20px" Text="....."></asp:button><asp:textbox style="Z-INDEX: 148; LEFT: 520px; POSITION: absolute; TOP: 192px" id="DDate" runat="server"
					BorderStyle="Groove" Width="96px" Height="20px" BackColor="Yellow" ForeColor="Blue" ReadOnly="True">DDate</asp:textbox><asp:textbox style="Z-INDEX: 147; LEFT: 224px; POSITION: absolute; TOP: 120px" id="DSliderCode"
					runat="server" BorderStyle="Groove" Width="342px" Height="56px" BackColor="Yellow" ForeColor="Blue" Font-Size="12pt" TextMode="MultiLine" ReadOnly="True">SliderCode</asp:textbox><asp:button style="Z-INDEX: 146; LEFT: 416px; POSITION: absolute; TOP: 293px" id="BMMapNo" runat="server"
					Width="18px" Height="20px" Text="修"></asp:button><asp:hyperlink style="Z-INDEX: 145; LEFT: 496px; POSITION: absolute; TOP: 1000px" id="LSampleFile"
					runat="server" Width="50px" Height="8px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">樣品圖</asp:hyperlink><INPUT style="Z-INDEX: 144; LEFT: 496px; WIDTH: 260px; POSITION: absolute; TOP: 1000px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					id="DSampleFile" size="24" type="file" name="File1" runat="server"><asp:dropdownlist style="Z-INDEX: 143; LEFT: 520px; POSITION: absolute; TOP: 226px" id="DPerson" runat="server"
					Width="96px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist style="Z-INDEX: 142; LEFT: 312px; POSITION: absolute; TOP: 226px" id="DDivision"
					runat="server" Width="112px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 141; LEFT: 168px; POSITION: absolute; TOP: 1256px" id="DReasonDesc"
					runat="server" BorderStyle="Groove" Width="424px" Height="59px" BackColor="Yellow" ForeColor="Blue" TextMode="MultiLine"
					Visible="False">DReasonDesc</asp:textbox><asp:textbox style="Z-INDEX: 140; LEFT: 240px; POSITION: absolute; TOP: 1224px" id="DReason"
					runat="server" BorderStyle="Groove" Width="352px" Height="20px" BackColor="Yellow" ForeColor="Blue" ReadOnly="True" Visible="False">DReason</asp:textbox><asp:dropdownlist style="Z-INDEX: 139; LEFT: 168px; POSITION: absolute; TOP: 1224px" id="DReasonCode"
					runat="server" Width="64px" Height="20px" BackColor="Yellow" ForeColor="Blue" AutoPostBack="True" Visible="False">
					<asp:ListItem Value="01">01</asp:ListItem>
					<asp:ListItem Value="02">02</asp:ListItem>
				</asp:dropdownlist><asp:hyperlink style="Z-INDEX: 138; LEFT: 168px; POSITION: absolute; TOP: 1184px" id="LBefOP" runat="server"
					Width="144px" Height="8px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">工程履歷</asp:hyperlink><asp:textbox style="Z-INDEX: 137; LEFT: 440px; POSITION: absolute; TOP: 1150px" id="DAEndTime"
					runat="server" BorderStyle="Groove" Width="152px" Height="20px" BackColor="Gold" ForeColor="Blue" Font-Size="9pt" ReadOnly="True">DAEndTime</asp:textbox><asp:textbox style="Z-INDEX: 136; LEFT: 168px; POSITION: absolute; TOP: 1150px" id="DAStartTime"
					runat="server" BorderStyle="Groove" Width="152px" Height="20px" BackColor="Gold" ForeColor="Blue" Font-Size="9pt" ReadOnly="True">DAStartTime</asp:textbox><asp:textbox style="Z-INDEX: 135; LEFT: 440px; POSITION: absolute; TOP: 1116px" id="DBEndTime"
					runat="server" BorderStyle="Groove" Width="152px" Height="20px" BackColor="Gold" ForeColor="Blue" Font-Size="9pt" ReadOnly="True">DBEndTime</asp:textbox><asp:textbox style="Z-INDEX: 134; LEFT: 168px; POSITION: absolute; TOP: 1116px" id="DBStartTime"
					runat="server" BorderStyle="Groove" Width="152px" Height="20px" BackColor="Gold" ForeColor="Blue" Font-Size="9pt" ReadOnly="True">DBStartTime</asp:textbox><asp:textbox style="Z-INDEX: 133; LEFT: 56px; POSITION: absolute; TOP: 1040px" id="DDecideDesc"
					runat="server" BorderStyle="Groove" Width="536px" Height="56px" BackColor="Yellow" ForeColor="Blue" TextMode="MultiLine">DDecideDesc</asp:textbox><asp:textbox style="Z-INDEX: 132; LEFT: 672px; POSITION: absolute; TOP: 395px" id="DMoldPoint"
					runat="server" BorderStyle="Groove" Width="45px" Height="20px" BackColor="Yellow" ForeColor="Blue">MoldPoint</asp:textbox><asp:textbox style="Z-INDEX: 131; LEFT: 584px; POSITION: absolute; TOP: 395px" id="DMoldQty"
					runat="server" BorderStyle="Groove" Width="48px" Height="20px" BackColor="Yellow" ForeColor="Blue">DMoldQty</asp:textbox><asp:textbox style="Z-INDEX: 130; LEFT: 688px; POSITION: absolute; TOP: 640px" id="DPullerPrice"
					runat="server" BorderStyle="Groove" Width="48px" Height="20px" BackColor="Yellow" ForeColor="Blue" Visible="False">PullerPrice</asp:textbox><asp:textbox style="Z-INDEX: 129; LEFT: 688px; POSITION: absolute; TOP: 608px" id="DPurMold"
					runat="server" BorderStyle="Groove" Width="72px" Height="20px" BackColor="Yellow" ForeColor="Blue">PurMold</asp:textbox><asp:textbox style="Z-INDEX: 128; LEFT: 688px; POSITION: absolute; TOP: 574px" id="DArMoldFee"
					runat="server" BorderStyle="Groove" Width="72px" Height="20px" BackColor="Yellow" ForeColor="Blue">ArMoldFee</asp:textbox><asp:hyperlink style="Z-INDEX: 126; LEFT: 496px; POSITION: absolute; TOP: 968px" id="LAuthorizeFile"
					runat="server" Width="50px" Height="8px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">授權書</asp:hyperlink><INPUT style="Z-INDEX: 125; LEFT: 496px; WIDTH: 260px; POSITION: absolute; TOP: 968px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					id="DAuthorizeFile" size="24" type="file" name="File1" runat="server">
				<asp:hyperlink style="Z-INDEX: 124; LEFT: 120px; POSITION: absolute; TOP: 968px" id="LConfirmFile"
					runat="server" Width="50px" Height="8px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">確認書</asp:hyperlink><INPUT style="Z-INDEX: 123; LEFT: 120px; WIDTH: 240px; POSITION: absolute; TOP: 968px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					id="DConfirmFile" type="file" name="File1" runat="server"><asp:textbox style="Z-INDEX: 122; LEFT: 688px; POSITION: absolute; TOP: 328px" id="DSellVendor"
					runat="server" BorderStyle="Groove" Width="73px" Height="20px" BackColor="Yellow" ForeColor="Blue">DSellVendor</asp:textbox><asp:dropdownlist style="Z-INDEX: 121; LEFT: 104px; POSITION: absolute; TOP: 395px" id="DMaterial"
					runat="server" Width="192px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 120; LEFT: 400px; POSITION: absolute; TOP: 362px" id="DManufPlace"
					runat="server" Width="88px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 119; LEFT: 200px; POSITION: absolute; TOP: 362px" id="DSliderType2"
					runat="server" Width="94px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 118; LEFT: 104px; POSITION: absolute; TOP: 362px" id="DSliderType1"
					runat="server" Width="94px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 117; LEFT: 600px; POSITION: absolute; TOP: 293px" id="DAssembler"
					runat="server" BorderStyle="Groove" Width="48px" Height="20px" BackColor="Yellow" ForeColor="Blue">DAssembler</asp:textbox><asp:dropdownlist style="Z-INDEX: 115; LEFT: 648px; POSITION: absolute; TOP: 293px" id="DAssembler1"
					runat="server" Width="118px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><INPUT style="Z-INDEX: 114; LEFT: 120px; WIDTH: 240px; POSITION: absolute; TOP: 932px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					id="DRefFile" type="file" name="File1" runat="server">
				<asp:button style="Z-INDEX: 113; LEFT: 398px; POSITION: absolute; TOP: 293px" id="BOMapNo" runat="server"
					Width="18px" Height="20px" Text="原"></asp:button><asp:textbox style="Z-INDEX: 112; LEFT: 312px; POSITION: absolute; TOP: 293px" id="DMapNo" runat="server"
					BorderStyle="Groove" Width="80px" Height="20px" BackColor="Yellow" ForeColor="Blue" ReadOnly="True" AutoPostBack="True">123456789012345</asp:textbox><asp:button style="Z-INDEX: 111; LEFT: 528px; POSITION: absolute; TOP: 259px" id="BSpec" runat="server"
					Width="20px" Height="20px" Text="....."></asp:button><asp:textbox style="Z-INDEX: 110; LEFT: 312px; POSITION: absolute; TOP: 259px" id="DSpec" runat="server"
					BorderStyle="Groove" Width="216px" Height="20px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" ReadOnly="True">DSpec</asp:textbox><asp:textbox style="Z-INDEX: 109; LEFT: 604px; POSITION: absolute; TOP: 158px" id="DSliderGRCode"
					runat="server" BorderStyle="Groove" Width="154px" Height="20px" BackColor="Yellow" ForeColor="Blue" ReadOnly="True">DSliderGRCode</asp:textbox><asp:textbox style="Z-INDEX: 108; LEFT: 312px; POSITION: absolute; TOP: 192px" id="DNo" runat="server"
					BorderStyle="Groove" Width="112px" Height="20px" BackColor="Yellow" ForeColor="Blue">DNo</asp:textbox><asp:textbox style="Z-INDEX: 107; LEFT: 16px; POSITION: absolute; TOP: 1320px" id="DFormSno"
					runat="server" BorderStyle="None" Width="97px" Height="20px" BackColor="White" ForeColor="Blue">單號：123</asp:textbox><asp:image style="Z-INDEX: 105; LEFT: 8px; POSITION: absolute; TOP: 1104px" id="DDelivery"
					runat="server" Width="593px" Height="110px" ImageUrl="Images\Sheet_Delivery.jpg"></asp:image><asp:image style="Z-INDEX: 103; LEFT: 8px; POSITION: absolute; TOP: 1216px" id="DDelay" runat="server"
					Width="593px" Height="107px" ImageUrl="Images\Sheet_Delay.jpg" Visible="False"></asp:image><asp:image style="Z-INDEX: 102; LEFT: 8px; POSITION: absolute; TOP: 1032px" id="DDescSheet"
					runat="server" Width="593px" Height="75px" ImageUrl="Images\MapSheet_004.jpg"></asp:image><asp:image style="Z-INDEX: 101; LEFT: 8px; POSITION: absolute; TOP: 536px" id="DManuaInSheet2"
					runat="server" Width="761px" Height="493px" ImageUrl="Images\ManufInSheet_005_C1.jpg"></asp:image><asp:image style="Z-INDEX: 242; LEFT: 16px; POSITION: absolute; TOP: 120px" id="LMapFile" runat="server"
					BorderStyle="Groove" Width="200px" Height="230px" ImageUrl="http://10.245.0.178//WorkFlow/Document/000003/1-Map2-2006124111750-2.jpg"></asp:image></FONT><INPUT style="Z-INDEX: 246; LEFT: 256px; WIDTH: 80px; POSITION: absolute; TOP: 1328px; HEIGHT: 28px"
				id="BSAVE" onclick="Button('SAVE');" value="儲存" type="button" name="Button1" runat="server"><INPUT style="Z-INDEX: 247; LEFT: 344px; WIDTH: 80px; POSITION: absolute; TOP: 1328px; HEIGHT: 28px"
				id="BNG2" onclick="Button('NG2');" value="NG2" type="button" name="Button1" runat="server"><INPUT style="Z-INDEX: 248; LEFT: 432px; WIDTH: 80px; POSITION: absolute; TOP: 1328px; HEIGHT: 28px"
				id="BNG1" onclick="Button('NG1');" value="NG1" type="button" name="Button1" runat="server"><INPUT style="Z-INDEX: 249; LEFT: 520px; WIDTH: 80px; POSITION: absolute; TOP: 1328px; HEIGHT: 28px"
				id="BOK" onclick="Button('OK');" value="OK" type="button" name="Button2" runat="server">
			<asp:imagebutton style="Z-INDEX: 259; LEFT: 8px; POSITION: absolute; TOP: 8px" id="BPrint" runat="server"
				Width="16px" Height="16px" ImageUrl="Images\Print.gif"></asp:imagebutton></form>
		</FONT></FORM>
	</body>
</HTML>
