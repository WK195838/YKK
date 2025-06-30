<%@ Page Language="vb" AutoEventWireup="false" Codebehind="ManufOutSheet_01.aspx.vb" Inherits="SPD.ManufOutSheet_01" aspCompat="True"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>外注委託書</title>
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
					//while (val == 0) {
					//	setTimeout("SendToSlider()",200);
					//	if ((wPop.document.SliderForm.DSlider2.value!="") || (wPop.document.SliderForm.DContent.value!=""))  {
					//		val=1;
					//	}
 					//} 
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
					//while (val == 0) {
					//	setTimeout("SendToSpec()",200);
					//	if (wPop.document.SpecForm.DSpec.value!="")  {
					//		val=1;
					//	}
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
		<form id="Form1" method="post" encType="multipart/form-data" runat="server">
			<FONT face="新細明體">
				<asp:textbox id="DCustomerCode" style="Z-INDEX: 320; POSITION: absolute; TOP: 328px; LEFT: 464px"
					runat="server" Width="88px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow">DCustomerCode</asp:textbox><asp:label id="Label42" style="Z-INDEX: 384; POSITION: absolute; TOP: 1044px; LEFT: 328px"
					runat="server" Width="80px" Height="12px">組立判定書</asp:label><INPUT id="DAssemblerFile" style="Z-INDEX: 347; POSITION: absolute; BACKGROUND-COLOR: #ffff00; WIDTH: 168px; HEIGHT: 20px; TOP: 1044px; LEFT: 416px"
					type="file" size="8" name="File1" runat="server">
				<asp:hyperlink id="LAssemblerFile" style="Z-INDEX: 383; POSITION: absolute; TOP: 1042px; LEFT: 416px"
					runat="server" Width="88px" Height="14px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">組立判定書</asp:hyperlink><asp:dropdownlist id="DAssembler1" style="Z-INDEX: 382; POSITION: absolute; TOP: 360px; LEFT: 644px"
					runat="server" Width="114px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:label id="Label41" style="Z-INDEX: 381; POSITION: absolute; TOP: 490px; LEFT: 688px" runat="server"
					Width="65px" Height="8px" Font-Size="8pt">稼動時間/日</asp:label><asp:label id="Label40" style="Z-INDEX: 380; POSITION: absolute; TOP: 448px; LEFT: 688px" runat="server"
					Width="64px" Height="10px" Font-Size="8pt">CYCLE TIME</asp:label><asp:textbox id="DWorkTime" style="Z-INDEX: 379; POSITION: absolute; TOP: 504px; LEFT: 688px"
					runat="server" Width="72px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DCYCLETIME" style="Z-INDEX: 378; POSITION: absolute; TOP: 464px; LEFT: 688px"
					runat="server" Width="73px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DLOSS" style="Z-INDEX: 377; POSITION: absolute; TOP: 540px; LEFT: 656px" runat="server"
					Width="103px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" MaxLength="10">DLOSS</asp:textbox><asp:label id="Label39" style="Z-INDEX: 376; POSITION: absolute; TOP: 544px; LEFT: 576px" runat="server"
					Width="80px" Height="10px">LOSS率(%)</asp:label><asp:hyperlink id="LQAttachFile1" style="Z-INDEX: 375; POSITION: absolute; TOP: 912px; LEFT: 16px"
					runat="server" Width="64px" Height="8px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">其他附件</asp:hyperlink><asp:dropdownlist id="DQAD15" style="Z-INDEX: 374; POSITION: absolute; TOP: 888px; LEFT: 712px" runat="server"
					Width="48px" Height="22px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAC15" style="Z-INDEX: 317; POSITION: absolute; TOP: 872px; LEFT: 712px" runat="server"
					Width="48px" Height="22px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAB15" style="Z-INDEX: 315; POSITION: absolute; TOP: 856px; LEFT: 712px" runat="server"
					Width="48px" Height="22px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:label id="Label38" style="Z-INDEX: 313; POSITION: absolute; TOP: 824px; LEFT: 712px" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">EDX</asp:label><asp:dropdownlist id="DQAA15" style="Z-INDEX: 311; POSITION: absolute; TOP: 840px; LEFT: 712px" runat="server"
					Width="48px" Height="22px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAD13" style="Z-INDEX: 301; POSITION: absolute; TOP: 888px; LEFT: 616px" runat="server"
					Width="48px" Height="346px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0級">0級</asp:ListItem>
					<asp:ListItem Value="1級">1級</asp:ListItem>
					<asp:ListItem Value="2級">2級</asp:ListItem>
					<asp:ListItem Value="3級">3級</asp:ListItem>
					<asp:ListItem Value="4級">4級</asp:ListItem>
					<asp:ListItem Value="5級">5級</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAC13" style="Z-INDEX: 299; POSITION: absolute; TOP: 872px; LEFT: 616px" runat="server"
					Width="48px" Height="346px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0級">0級</asp:ListItem>
					<asp:ListItem Value="1級">1級</asp:ListItem>
					<asp:ListItem Value="2級">2級</asp:ListItem>
					<asp:ListItem Value="3級">3級</asp:ListItem>
					<asp:ListItem Value="4級">4級</asp:ListItem>
					<asp:ListItem Value="5級">5級</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAB13" style="Z-INDEX: 298; POSITION: absolute; TOP: 856px; LEFT: 616px" runat="server"
					Width="48px" Height="346px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0級">0級</asp:ListItem>
					<asp:ListItem Value="1級">1級</asp:ListItem>
					<asp:ListItem Value="2級">2級</asp:ListItem>
					<asp:ListItem Value="3級">3級</asp:ListItem>
					<asp:ListItem Value="4級">4級</asp:ListItem>
					<asp:ListItem Value="5級">5級</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAD12" style="Z-INDEX: 297; POSITION: absolute; TOP: 888px; LEFT: 568px" runat="server"
					Width="48px" Height="346px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0級">0級</asp:ListItem>
					<asp:ListItem Value="1級">1級</asp:ListItem>
					<asp:ListItem Value="2級">2級</asp:ListItem>
					<asp:ListItem Value="3級">3級</asp:ListItem>
					<asp:ListItem Value="4級">4級</asp:ListItem>
					<asp:ListItem Value="5級">5級</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAC12" style="Z-INDEX: 296; POSITION: absolute; TOP: 872px; LEFT: 568px" runat="server"
					Width="48px" Height="346px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0級">0級</asp:ListItem>
					<asp:ListItem Value="1級">1級</asp:ListItem>
					<asp:ListItem Value="2級">2級</asp:ListItem>
					<asp:ListItem Value="3級">3級</asp:ListItem>
					<asp:ListItem Value="4級">4級</asp:ListItem>
					<asp:ListItem Value="5級">5級</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAB12" style="Z-INDEX: 295; POSITION: absolute; TOP: 856px; LEFT: 568px" runat="server"
					Width="48px" Height="346px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0級">0級</asp:ListItem>
					<asp:ListItem Value="1級">1級</asp:ListItem>
					<asp:ListItem Value="2級">2級</asp:ListItem>
					<asp:ListItem Value="3級">3級</asp:ListItem>
					<asp:ListItem Value="4級">4級</asp:ListItem>
					<asp:ListItem Value="5級">5級</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAD11" style="Z-INDEX: 294; POSITION: absolute; TOP: 888px; LEFT: 520px" runat="server"
					Width="48px" Height="346px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
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
				</asp:dropdownlist><asp:dropdownlist id="DQAC11" style="Z-INDEX: 292; POSITION: absolute; TOP: 872px; LEFT: 520px" runat="server"
					Width="48px" Height="346px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
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
				</asp:dropdownlist><asp:dropdownlist id="DQAB11" style="Z-INDEX: 291; POSITION: absolute; TOP: 856px; LEFT: 520px" runat="server"
					Width="48px" Height="346px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
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
				</asp:dropdownlist><asp:dropdownlist id="DQAA13" style="Z-INDEX: 289; POSITION: absolute; TOP: 840px; LEFT: 616px" runat="server"
					Width="48px" Height="346px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0級">0級</asp:ListItem>
					<asp:ListItem Value="1級">1級</asp:ListItem>
					<asp:ListItem Value="2級">2級</asp:ListItem>
					<asp:ListItem Value="3級">3級</asp:ListItem>
					<asp:ListItem Value="4級">4級</asp:ListItem>
					<asp:ListItem Value="5級">5級</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAA12" style="Z-INDEX: 288; POSITION: absolute; TOP: 840px; LEFT: 568px" runat="server"
					Width="48px" Height="346px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0級">0級</asp:ListItem>
					<asp:ListItem Value="1級">1級</asp:ListItem>
					<asp:ListItem Value="2級">2級</asp:ListItem>
					<asp:ListItem Value="3級">3級</asp:ListItem>
					<asp:ListItem Value="4級">4級</asp:ListItem>
					<asp:ListItem Value="5級">5級</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAA11" style="Z-INDEX: 287; POSITION: absolute; TOP: 840px; LEFT: 520px" runat="server"
					Width="48px" Height="346px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
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
				</asp:dropdownlist><asp:dropdownlist id="DQAD14" style="Z-INDEX: 237; POSITION: absolute; TOP: 888px; LEFT: 664px" runat="server"
					Width="48px" Height="22px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAD10" style="Z-INDEX: 236; POSITION: absolute; TOP: 888px; LEFT: 472px" runat="server"
					Width="48px" Height="18px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAD9" style="Z-INDEX: 235; POSITION: absolute; TOP: 888px; LEFT: 424px" runat="server"
					Width="48px" Height="18px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAD8" style="Z-INDEX: 234; POSITION: absolute; TOP: 888px; LEFT: 376px" runat="server"
					Width="48px" Height="18px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="NCB">NCB</asp:ListItem>
					<asp:ListItem Value="KENSIN">KENSIN</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAD7" style="Z-INDEX: 233; POSITION: absolute; TOP: 888px; LEFT: 328px" runat="server"
					Width="48px" Height="18px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAD6" style="Z-INDEX: 232; POSITION: absolute; TOP: 888px; LEFT: 280px" runat="server"
					Width="48px" Height="18px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DQAD5" style="Z-INDEX: 231; POSITION: absolute; TOP: 888px; LEFT: 232px" runat="server"
					Width="47px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt"></asp:textbox><asp:textbox id="DQAD4" style="Z-INDEX: 230; POSITION: absolute; TOP: 888px; LEFT: 204px" runat="server"
					Width="28px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt"></asp:textbox><asp:dropdownlist id="DQAD3" style="Z-INDEX: 229; POSITION: absolute; TOP: 888px; LEFT: 170px" runat="server"
					Width="35px" Height="18px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DQAD2" style="Z-INDEX: 228; POSITION: absolute; TOP: 888px; LEFT: 72px" runat="server"
					Width="98px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt"></asp:textbox><asp:textbox id="DQAD1" style="Z-INDEX: 227; POSITION: absolute; TOP: 888px; LEFT: 16px" runat="server"
					Width="56px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt"></asp:textbox><asp:dropdownlist id="DQAC14" style="Z-INDEX: 226; POSITION: absolute; TOP: 872px; LEFT: 664px" runat="server"
					Width="48px" Height="22px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAC10" style="Z-INDEX: 225; POSITION: absolute; TOP: 872px; LEFT: 472px" runat="server"
					Width="48px" Height="18px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAC9" style="Z-INDEX: 224; POSITION: absolute; TOP: 872px; LEFT: 424px" runat="server"
					Width="48px" Height="18px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAC8" style="Z-INDEX: 223; POSITION: absolute; TOP: 872px; LEFT: 376px" runat="server"
					Width="48px" Height="36px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="NCB">NCB</asp:ListItem>
					<asp:ListItem Value="KENSIN">KENSIN</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAC7" style="Z-INDEX: 222; POSITION: absolute; TOP: 872px; LEFT: 328px" runat="server"
					Width="48px" Height="18px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAC6" style="Z-INDEX: 221; POSITION: absolute; TOP: 872px; LEFT: 280px" runat="server"
					Width="48px" Height="18px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DQAC5" style="Z-INDEX: 220; POSITION: absolute; TOP: 872px; LEFT: 232px" runat="server"
					Width="47px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt"></asp:textbox><asp:textbox id="DQAC4" style="Z-INDEX: 219; POSITION: absolute; TOP: 872px; LEFT: 204px" runat="server"
					Width="28px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt"></asp:textbox><asp:dropdownlist id="DQAC3" style="Z-INDEX: 218; POSITION: absolute; TOP: 872px; LEFT: 170px" runat="server"
					Width="35px" Height="18px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DQAC2" style="Z-INDEX: 217; POSITION: absolute; TOP: 872px; LEFT: 72px" runat="server"
					Width="98px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt"></asp:textbox><asp:textbox id="DQAC1" style="Z-INDEX: 216; POSITION: absolute; TOP: 872px; LEFT: 16px" runat="server"
					Width="56px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt"></asp:textbox><asp:dropdownlist id="DQAB14" style="Z-INDEX: 215; POSITION: absolute; TOP: 856px; LEFT: 664px" runat="server"
					Width="48px" Height="22px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAB10" style="Z-INDEX: 214; POSITION: absolute; TOP: 856px; LEFT: 472px" runat="server"
					Width="48px" Height="18px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAB9" style="Z-INDEX: 213; POSITION: absolute; TOP: 856px; LEFT: 424px" runat="server"
					Width="48px" Height="18px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAB8" style="Z-INDEX: 211; POSITION: absolute; TOP: 856px; LEFT: 376px" runat="server"
					Width="48px" Height="36px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="NCB">NCB</asp:ListItem>
					<asp:ListItem Value="KENSIN">KENSIN</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAB7" style="Z-INDEX: 210; POSITION: absolute; TOP: 856px; LEFT: 328px" runat="server"
					Width="48px" Height="18px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAB6" style="Z-INDEX: 208; POSITION: absolute; TOP: 856px; LEFT: 280px" runat="server"
					Width="48px" Height="18px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DQAB5" style="Z-INDEX: 207; POSITION: absolute; TOP: 856px; LEFT: 232px" runat="server"
					Width="47px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt"></asp:textbox><asp:textbox id="DQAB4" style="Z-INDEX: 206; POSITION: absolute; TOP: 856px; LEFT: 204px" runat="server"
					Width="28px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt"></asp:textbox><asp:dropdownlist id="DQAB3" style="Z-INDEX: 204; POSITION: absolute; TOP: 856px; LEFT: 170px" runat="server"
					Width="35px" Height="18px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DQAB2" style="Z-INDEX: 203; POSITION: absolute; TOP: 856px; LEFT: 72px" runat="server"
					Width="98px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt"></asp:textbox><asp:textbox id="DQAB1" style="Z-INDEX: 201; POSITION: absolute; TOP: 856px; LEFT: 16px" runat="server"
					Width="56px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt"></asp:textbox><asp:dropdownlist id="DQAA14" style="Z-INDEX: 162; POSITION: absolute; TOP: 840px; LEFT: 664px" runat="server"
					Width="48px" Height="22px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAA10" style="Z-INDEX: 161; POSITION: absolute; TOP: 840px; LEFT: 472px" runat="server"
					Width="48px" Height="346px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAA9" style="Z-INDEX: 160; POSITION: absolute; TOP: 840px; LEFT: 424px" runat="server"
					Width="48px" Height="18px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAA8" style="Z-INDEX: 159; POSITION: absolute; TOP: 840px; LEFT: 376px" runat="server"
					Width="48px" Height="26px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="NCB">NCB</asp:ListItem>
					<asp:ListItem Value="KENSIN">KENSIN</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAA7" style="Z-INDEX: 158; POSITION: absolute; TOP: 840px; LEFT: 328px" runat="server"
					Width="48px" Height="162px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQAA6" style="Z-INDEX: 157; POSITION: absolute; TOP: 840px; LEFT: 280px" runat="server"
					Width="48px" Height="40px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DQAA5" style="Z-INDEX: 156; POSITION: absolute; TOP: 840px; LEFT: 232px" runat="server"
					Width="47px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt"></asp:textbox><asp:textbox id="DQAA4" style="Z-INDEX: 154; POSITION: absolute; TOP: 840px; LEFT: 204px" runat="server"
					Width="28px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt"></asp:textbox><asp:dropdownlist id="DQAA3" style="Z-INDEX: 153; POSITION: absolute; TOP: 840px; LEFT: 170px" runat="server"
					Width="35px" Height="18px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DQAA1" style="Z-INDEX: 150; POSITION: absolute; TOP: 840px; LEFT: 16px" runat="server"
					Width="56px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt"></asp:textbox><asp:textbox id="DQAA2" style="Z-INDEX: 169; POSITION: absolute; TOP: 840px; LEFT: 72px" runat="server"
					Width="98px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt"></asp:textbox><asp:label id="Label1" style="Z-INDEX: 284; POSITION: absolute; TOP: 824px; LEFT: 48px" runat="server"
					Width="24px" Height="10px" Font-Size="8pt">日期</asp:label><asp:label id="Label13" style="Z-INDEX: 281; POSITION: absolute; TOP: 824px; LEFT: 664px" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">CPSC</asp:label><asp:label id="Label35" style="Z-INDEX: 302; POSITION: absolute; TOP: 824px; LEFT: 568px" runat="server"
					Width="48px" Height="10px" Font-Size="8pt">一次密著</asp:label><asp:label id="Label12" style="Z-INDEX: 280; POSITION: absolute; TOP: 824px; LEFT: 616px" runat="server"
					Width="48px" Height="10px" Font-Size="8pt">二次密著</asp:label><asp:label id="Label11" style="Z-INDEX: 278; POSITION: absolute; TOP: 824px; LEFT: 520px" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">黃變</asp:label><asp:label id="Label10" style="Z-INDEX: 276; POSITION: absolute; TOP: 824px; LEFT: 472px" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">乾洗</asp:label><asp:label id="Label9" style="Z-INDEX: 275; POSITION: absolute; TOP: 824px; LEFT: 424px" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">水洗</asp:label><asp:label id="Label8" style="Z-INDEX: 273; POSITION: absolute; TOP: 824px; LEFT: 376px" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">檢針</asp:label><asp:label id="Label7" style="Z-INDEX: 272; POSITION: absolute; TOP: 824px; LEFT: 328px" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">扭力</asp:label><asp:label id="Label6" style="Z-INDEX: 269; POSITION: absolute; TOP: 824px; LEFT: 280px" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">強度</asp:label><asp:label id="Label5" style="Z-INDEX: 268; POSITION: absolute; TOP: 824px; LEFT: 240px" runat="server"
					Width="40px" Height="10px" Font-Size="8pt">原單位</asp:label><asp:label id="Label4" style="Z-INDEX: 265; POSITION: absolute; TOP: 824px; LEFT: 208px" runat="server"
					Width="48px" Height="10px" Font-Size="8pt">表面</asp:label><asp:label id="Label3" style="Z-INDEX: 264; POSITION: absolute; TOP: 824px; LEFT: 168px" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">口厚</asp:label><asp:label id="Label2" style="Z-INDEX: 261; POSITION: absolute; TOP: 824px; LEFT: 80px" runat="server"
					Width="48px" Height="10px" Font-Size="8pt">拉頭種類</asp:label><asp:dropdownlist id="DLogo" style="Z-INDEX: 318; POSITION: absolute; TOP: 224px; LEFT: 686px" runat="server"
					Width="78px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:hyperlink id="LForCastFile" style="Z-INDEX: 373; POSITION: absolute; TOP: 1112px; LEFT: 608px"
					runat="server" Width="56px" Height="8px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">報價單</asp:hyperlink><INPUT id="DForCastFile" style="Z-INDEX: 372; POSITION: absolute; BACKGROUND-COLOR: #ffff00; WIDTH: 164px; HEIGHT: 20px; TOP: 1112px; LEFT: 600px"
					type="file" size="8" name="File1" runat="server">
				<asp:textbox id="DPullerPrice" style="Z-INDEX: 130; POSITION: absolute; TOP: 752px; LEFT: 688px"
					runat="server" Width="72px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow">PullerPrice</asp:textbox><asp:textbox id="DPurMold" style="Z-INDEX: 128; POSITION: absolute; TOP: 718px; LEFT: 688px"
					runat="server" Width="72px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow">PurMold</asp:textbox><asp:textbox id="DPullerPrice1" style="Z-INDEX: 129; POSITION: absolute; TOP: 786px; LEFT: 688px"
					runat="server" Width="72px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow">PullerPrice1</asp:textbox><asp:textbox id="DArMoldFee" style="Z-INDEX: 127; POSITION: absolute; TOP: 686px; LEFT: 688px"
					runat="server" Width="72px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow">ArMoldFee</asp:textbox><asp:label id="Label37" style="Z-INDEX: 314; POSITION: absolute; TOP: 1044px; LEFT: 744px"
					runat="server" Width="1px" Height="10px">分</asp:label><asp:textbox id="DManufFlow" style="Z-INDEX: 371; POSITION: absolute; TOP: 624px; LEFT: 56px"
					runat="server" Width="369px" Height="50px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" MaxLength="240" TextMode="MultiLine">DManufFlow</asp:textbox><asp:textbox id="DOFormSno1" style="Z-INDEX: 370; POSITION: absolute; TOP: 293px; LEFT: 736px"
					runat="server" Width="28px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt" AutoPostBack="True">1234</asp:textbox><asp:textbox id="DOFormNo1" style="Z-INDEX: 369; POSITION: absolute; TOP: 293px; LEFT: 696px"
					runat="server" Width="38px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt" AutoPostBack="True">000001</asp:textbox><asp:button id="BMMapNo1" style="Z-INDEX: 368; POSITION: absolute; TOP: 293px; LEFT: 672px"
					runat="server" Width="20px" Height="20px" Text="修"></asp:button><asp:button id="BOMapNo1" style="Z-INDEX: 367; POSITION: absolute; TOP: 293px; LEFT: 648px"
					runat="server" Width="20px" Height="20px" Text="原"></asp:button><asp:hyperlink id="LMapNo1" style="Z-INDEX: 366; POSITION: absolute; TOP: 293px; LEFT: 544px" runat="server"
					Width="100px" Height="8px" Target="_blank" NavigateUrl="BoardEdit.aspx">圖號文件</asp:hyperlink><asp:textbox id="DMapNo1" style="Z-INDEX: 365; POSITION: absolute; TOP: 293px; LEFT: 544px" runat="server"
					Width="104px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" AutoPostBack="True" ReadOnly="True">123456789012345</asp:textbox><asp:label id="Label36" style="Z-INDEX: 363; POSITION: absolute; TOP: 1044px; LEFT: 600px"
					runat="server" Width="57px" Height="10px">QC-L/T</asp:label><asp:textbox id="DQCLT" style="Z-INDEX: 316; POSITION: absolute; TOP: 1044px; LEFT: 664px" runat="server"
					Width="81px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DOPReadyDesc" style="Z-INDEX: 362; POSITION: absolute; TOP: 1328px; LEFT: 240px"
					runat="server" Width="104px" Height="20px" BorderStyle="None" ForeColor="Red" BackColor="White">需閱讀工程履歷</asp:textbox><asp:textbox id="DOPReady" style="Z-INDEX: 155; POSITION: absolute; TOP: 1328px; LEFT: 344px"
					runat="server" Width="48px" Height="20px" BorderStyle="Groove" ForeColor="Red" BackColor="GreenYellow" ReadOnly="True"></asp:textbox><asp:dropdownlist id="DCpsc" style="Z-INDEX: 361; POSITION: absolute; TOP: 190px; LEFT: 708px" runat="server"
					Width="56px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:label id="Label25" style="Z-INDEX: 337; POSITION: absolute; TOP: 580px; LEFT: 704px" runat="server"
					Width="24px" Height="10px" Font-Size="8pt">數量</asp:label><asp:dropdownlist id="DSuppiler" style="Z-INDEX: 360; POSITION: absolute; TOP: 395px; LEFT: 400px"
					runat="server" Width="146px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:label id="Label32" style="Z-INDEX: 359; POSITION: absolute; TOP: 440px; LEFT: 136px" runat="server"
					Width="32px" Height="10px" Font-Size="8pt" Visible="False">品名-4</asp:label><asp:label id="Label31" style="Z-INDEX: 358; POSITION: absolute; TOP: 440px; LEFT: 96px" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">品名-3</asp:label><asp:label id="Label30" style="Z-INDEX: 357; POSITION: absolute; TOP: 440px; LEFT: 56px" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">品名-2</asp:label><asp:label id="Label34" style="Z-INDEX: 356; POSITION: absolute; TOP: 440px; LEFT: 568px" runat="server"
					Width="24px" Height="10px" Font-Size="8pt">備註</asp:label><asp:label id="Label33" style="Z-INDEX: 355; POSITION: absolute; TOP: 440px; LEFT: 472px" runat="server"
					Width="64px" Height="10px" Font-Size="8pt">預定完成日</asp:label><asp:label id="Label29" style="Z-INDEX: 354; POSITION: absolute; TOP: 440px; LEFT: 376px" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">OR-NO</asp:label><asp:label id="Label28" style="Z-INDEX: 353; POSITION: absolute; TOP: 440px; LEFT: 280px" runat="server"
					Width="56px" Height="10px" Font-Size="8pt">基準日程</asp:label><asp:label id="Label27" style="Z-INDEX: 352; POSITION: absolute; TOP: 440px; LEFT: 184px" runat="server"
					Width="64px" Height="10px" Font-Size="8pt">一日產能</asp:label><asp:label id="Label26" style="Z-INDEX: 351; POSITION: absolute; TOP: 440px; LEFT: 16px" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">品名-1</asp:label><asp:label id="Label24" style="Z-INDEX: 336; POSITION: absolute; TOP: 580px; LEFT: 648px" runat="server"
					Width="24px" Height="10px" Font-Size="8pt">顏色</asp:label><asp:label id="Label23" style="Z-INDEX: 333; POSITION: absolute; TOP: 580px; LEFT: 472px" runat="server"
					Width="48px" Height="10px" Font-Size="8pt">拉頭種類</asp:label><asp:label id="Label22" style="Z-INDEX: 330; POSITION: absolute; TOP: 700px; LEFT: 504px" runat="server"
					Width="24px" Height="10px" Font-Size="8pt">單價</asp:label><asp:label id="Label21" style="Z-INDEX: 328; POSITION: absolute; TOP: 700px; LEFT: 464px" runat="server"
					Width="24px" Height="10px" Font-Size="8pt">幣別</asp:label><asp:label id="Label20" style="Z-INDEX: 326; POSITION: absolute; TOP: 700px; LEFT: 296px" runat="server"
					Width="48px" Height="10px" Font-Size="8pt">拉頭種類</asp:label><asp:label id="Label19" style="Z-INDEX: 325; POSITION: absolute; TOP: 700px; LEFT: 224px" runat="server"
					Width="24px" Height="10px" Font-Size="8pt">單價</asp:label><asp:label id="Label18" style="Z-INDEX: 322; POSITION: absolute; TOP: 700px; LEFT: 184px" runat="server"
					Width="24px" Height="10px" Font-Size="8pt">幣別</asp:label><asp:label id="Label17" style="Z-INDEX: 319; POSITION: absolute; TOP: 700px; LEFT: 56px" runat="server"
					Width="48px" Height="10px" Font-Size="8pt">拉頭種類</asp:label><asp:label id="Label16" style="Z-INDEX: 312; POSITION: absolute; TOP: 956px; LEFT: 136px" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">備註</asp:label><asp:label id="Label15" style="Z-INDEX: 308; POSITION: absolute; TOP: 956px; LEFT: 80px" runat="server"
					Width="48px" Height="10px" Font-Size="8pt">檢測結果</asp:label><asp:label id="Label14" style="Z-INDEX: 306; POSITION: absolute; TOP: 956px; LEFT: 48px" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">日期</asp:label><asp:image id="DManuaInSheet1" style="Z-INDEX: 100; POSITION: absolute; TOP: 8px; LEFT: 8px"
					runat="server" Width="761px" Height="674px" ImageUrl="Images\ManufOutSheet_003_f.jpg"></asp:image><asp:dropdownlist id="DBuyer" style="Z-INDEX: 350; POSITION: absolute; TOP: 328px; LEFT: 312px" runat="server"
					Width="152px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:hyperlink id="LQAttachFile2" style="Z-INDEX: 349; POSITION: absolute; TOP: 1044px; LEFT: 136px"
					runat="server" Width="120px" Height="8px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">品質分析依賴書</asp:hyperlink><INPUT id="DQAttachFile2" style="Z-INDEX: 348; POSITION: absolute; BACKGROUND-COLOR: #ffff00; WIDTH: 176px; HEIGHT: 20px; TOP: 1044px; LEFT: 136px"
					type="file" size="10" name="File1" runat="server"><INPUT id="DQAttachFile1" style="Z-INDEX: 346; POSITION: absolute; BACKGROUND-COLOR: #ffff00; WIDTH: 744px; HEIGHT: 20px; TOP: 916px; LEFT: 16px"
					type="file" name="File1" runat="server">
				<asp:hyperlink id="LSAttachFile" style="Z-INDEX: 345; POSITION: absolute; TOP: 652px; LEFT: 440px"
					runat="server" Width="64px" Height="8px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">其他附件</asp:hyperlink><INPUT id="DSAttachFile" style="Z-INDEX: 344; POSITION: absolute; BACKGROUND-COLOR: #ffff00; WIDTH: 314px; HEIGHT: 20px; TOP: 652px; LEFT: 440px"
					type="file" size="24" name="File1" runat="server"><asp:imagebutton id="BFlow" style="Z-INDEX: 343; POSITION: absolute; TOP: 32px; LEFT: 8px" runat="server"
					Width="16px" Height="16px" ImageUrl="Images\Flow-01.gif"></asp:imagebutton><asp:textbox id="DHAD9" style="Z-INDEX: 342; POSITION: absolute; TOP: 506px; LEFT: 568px" runat="server"
					Width="110px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DHAD8" style="Z-INDEX: 341; POSITION: absolute; TOP: 506px; LEFT: 472px" runat="server"
					Width="88px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DHAD7" style="Z-INDEX: 340; POSITION: absolute; TOP: 506px; LEFT: 376px" runat="server"
					Width="88px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DHAD6" style="Z-INDEX: 339; POSITION: absolute; TOP: 506px; LEFT: 280px" runat="server"
					Width="88px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DHAD5" style="Z-INDEX: 338; POSITION: absolute; TOP: 506px; LEFT: 184px" runat="server"
					Width="88px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DHAD4" style="Z-INDEX: 335; POSITION: absolute; TOP: 506px; LEFT: 136px" runat="server"
					Width="40px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DHAD3" style="Z-INDEX: 334; POSITION: absolute; TOP: 506px; LEFT: 96px" runat="server"
					Width="80px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DHAD2" style="Z-INDEX: 332; POSITION: absolute; TOP: 506px; LEFT: 56px" runat="server"
					Width="40px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DHAD1" style="Z-INDEX: 331; POSITION: absolute; TOP: 506px; LEFT: 16px" runat="server"
					Width="40px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DHAC9" style="Z-INDEX: 329; POSITION: absolute; TOP: 488px; LEFT: 568px" runat="server"
					Width="110px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DHAC8" style="Z-INDEX: 327; POSITION: absolute; TOP: 488px; LEFT: 472px" runat="server"
					Width="88px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DHAC7" style="Z-INDEX: 324; POSITION: absolute; TOP: 488px; LEFT: 376px" runat="server"
					Width="88px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DHAC6" style="Z-INDEX: 323; POSITION: absolute; TOP: 488px; LEFT: 280px" runat="server"
					Width="88px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DHAC5" style="Z-INDEX: 321; POSITION: absolute; TOP: 488px; LEFT: 184px" runat="server"
					Width="88px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DHAC4" style="Z-INDEX: 310; POSITION: absolute; TOP: 488px; LEFT: 136px" runat="server"
					Width="40px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DHAC3" style="Z-INDEX: 309; POSITION: absolute; TOP: 488px; LEFT: 96px" runat="server"
					Width="80px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DHAC2" style="Z-INDEX: 307; POSITION: absolute; TOP: 488px; LEFT: 56px" runat="server"
					Width="40px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DHAC1" style="Z-INDEX: 305; POSITION: absolute; TOP: 488px; LEFT: 16px" runat="server"
					Width="40px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DHAB9" style="Z-INDEX: 304; POSITION: absolute; TOP: 470px; LEFT: 568px" runat="server"
					Width="110px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DHAB8" style="Z-INDEX: 303; POSITION: absolute; TOP: 470px; LEFT: 472px" runat="server"
					Width="88px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DHAB7" style="Z-INDEX: 300; POSITION: absolute; TOP: 470px; LEFT: 376px" runat="server"
					Width="88px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DHAB6" style="Z-INDEX: 293; POSITION: absolute; TOP: 470px; LEFT: 280px" runat="server"
					Width="88px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DHAB5" style="Z-INDEX: 290; POSITION: absolute; TOP: 470px; LEFT: 184px" runat="server"
					Width="88px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DHAB4" style="Z-INDEX: 286; POSITION: absolute; TOP: 470px; LEFT: 136px" runat="server"
					Width="40px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DHAB3" style="Z-INDEX: 285; POSITION: absolute; TOP: 470px; LEFT: 96px" runat="server"
					Width="80px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DHAB2" style="Z-INDEX: 283; POSITION: absolute; TOP: 470px; LEFT: 56px" runat="server"
					Width="40px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DHAB1" style="Z-INDEX: 282; POSITION: absolute; TOP: 470px; LEFT: 16px" runat="server"
					Width="40px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DHAA9" style="Z-INDEX: 279; POSITION: absolute; TOP: 452px; LEFT: 568px" runat="server"
					Width="110px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DHAA8" style="Z-INDEX: 277; POSITION: absolute; TOP: 452px; LEFT: 472px" runat="server"
					Width="88px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DHAA7" style="Z-INDEX: 274; POSITION: absolute; TOP: 452px; LEFT: 376px" runat="server"
					Width="88px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DHAA6" style="Z-INDEX: 271; POSITION: absolute; TOP: 452px; LEFT: 280px" runat="server"
					Width="88px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DHAA5" style="Z-INDEX: 270; POSITION: absolute; TOP: 452px; LEFT: 184px" runat="server"
					Width="88px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DHAA4" style="Z-INDEX: 267; POSITION: absolute; TOP: 452px; LEFT: 136px" runat="server"
					Width="40px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DHAA3" style="Z-INDEX: 266; POSITION: absolute; TOP: 452px; LEFT: 96px" runat="server"
					Width="80px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DHAA2" style="Z-INDEX: 263; POSITION: absolute; TOP: 452px; LEFT: 56px" runat="server"
					Width="40px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DHAA1" style="Z-INDEX: 262; POSITION: absolute; TOP: 452px; LEFT: 16px" runat="server"
					Width="40px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DHADesc" style="Z-INDEX: 260; POSITION: absolute; TOP: 534px; LEFT: 104px" runat="server"
					Width="464px" Height="28px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" TextMode="MultiLine">DHDesc</asp:textbox><asp:textbox id="DFQAD3" style="Z-INDEX: 258; POSITION: absolute; TOP: 1022px; LEFT: 136px" runat="server"
					Width="620px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:dropdownlist id="DFQAD2" style="Z-INDEX: 257; POSITION: absolute; TOP: 1022px; LEFT: 80px" runat="server"
					Width="56px" Height="18px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DFQAD1" style="Z-INDEX: 256; POSITION: absolute; TOP: 1022px; LEFT: 16px" runat="server"
					Width="64px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DFQAC1" style="Z-INDEX: 255; POSITION: absolute; TOP: 1004px; LEFT: 16px" runat="server"
					Width="64px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:dropdownlist id="DFQAC2" style="Z-INDEX: 254; POSITION: absolute; TOP: 1004px; LEFT: 80px" runat="server"
					Width="56px" Height="18px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DFQAC3" style="Z-INDEX: 253; POSITION: absolute; TOP: 1004px; LEFT: 136px" runat="server"
					Width="620px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DFQAB3" style="Z-INDEX: 252; POSITION: absolute; TOP: 986px; LEFT: 136px" runat="server"
					Width="620px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:dropdownlist id="DFQAB2" style="Z-INDEX: 251; POSITION: absolute; TOP: 986px; LEFT: 80px" runat="server"
					Width="56px" Height="18px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DFQAB1" style="Z-INDEX: 250; POSITION: absolute; TOP: 986px; LEFT: 16px" runat="server"
					Width="64px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DOFormSno" style="Z-INDEX: 245; POSITION: absolute; TOP: 293px; LEFT: 504px"
					runat="server" Width="28px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt" AutoPostBack="True">1234</asp:textbox><asp:textbox id="DOFormNo" style="Z-INDEX: 244; POSITION: absolute; TOP: 293px; LEFT: 464px"
					runat="server" Width="38px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt" AutoPostBack="True">000001</asp:textbox><INPUT id="DMapFile" style="Z-INDEX: 243; POSITION: absolute; BACKGROUND-COLOR: #ffff00; WIDTH: 200px; HEIGHT: 20px; TOP: 328px; LEFT: 16px"
					type="file" size="14" name="File1" runat="server">
				<asp:hyperlink id="LMapNo" style="Z-INDEX: 241; POSITION: absolute; TOP: 293px; LEFT: 312px" runat="server"
					Width="100px" Height="8px" Target="_blank" NavigateUrl="BoardEdit.aspx">圖號文件</asp:hyperlink><asp:textbox id="DFQAA3" style="Z-INDEX: 240; POSITION: absolute; TOP: 968px; LEFT: 136px" runat="server"
					Width="620px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:dropdownlist id="DFQAA2" style="Z-INDEX: 239; POSITION: absolute; TOP: 968px; LEFT: 80px" runat="server"
					Width="56px" Height="18px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DFQAA1" style="Z-INDEX: 238; POSITION: absolute; TOP: 968px; LEFT: 16px" runat="server"
					Width="64px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DPriceJ3" style="Z-INDEX: 212; POSITION: absolute; TOP: 784px; LEFT: 504px"
					runat="server" Width="64px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:dropdownlist id="DPriceJ2" style="Z-INDEX: 209; POSITION: absolute; TOP: 784px; LEFT: 464px"
					runat="server" Width="40px" Height="12px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceJ1" style="Z-INDEX: 205; POSITION: absolute; TOP: 784px; LEFT: 296px"
					runat="server" Width="166px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DPriceI3" style="Z-INDEX: 202; POSITION: absolute; TOP: 784px; LEFT: 224px"
					runat="server" Width="64px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:dropdownlist id="DPriceI2" style="Z-INDEX: 200; POSITION: absolute; TOP: 784px; LEFT: 184px"
					runat="server" Width="40px" Height="12px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceI1" style="Z-INDEX: 199; POSITION: absolute; TOP: 784px; LEFT: 16px" runat="server"
					Width="166px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DPriceH3" style="Z-INDEX: 198; POSITION: absolute; TOP: 766px; LEFT: 504px"
					runat="server" Width="64px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:dropdownlist id="DPriceH2" style="Z-INDEX: 197; POSITION: absolute; TOP: 766px; LEFT: 464px"
					runat="server" Width="40px" Height="12px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceH1" style="Z-INDEX: 196; POSITION: absolute; TOP: 766px; LEFT: 296px"
					runat="server" Width="166px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DPriceG3" style="Z-INDEX: 195; POSITION: absolute; TOP: 766px; LEFT: 224px"
					runat="server" Width="64px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:dropdownlist id="DPriceG2" style="Z-INDEX: 194; POSITION: absolute; TOP: 766px; LEFT: 184px"
					runat="server" Width="40px" Height="12px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceG1" style="Z-INDEX: 193; POSITION: absolute; TOP: 766px; LEFT: 16px" runat="server"
					Width="166px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DPriceF3" style="Z-INDEX: 192; POSITION: absolute; TOP: 748px; LEFT: 504px"
					runat="server" Width="64px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:dropdownlist id="DPriceF2" style="Z-INDEX: 191; POSITION: absolute; TOP: 748px; LEFT: 464px"
					runat="server" Width="40px" Height="12px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceF1" style="Z-INDEX: 190; POSITION: absolute; TOP: 748px; LEFT: 296px"
					runat="server" Width="166px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DPriceE3" style="Z-INDEX: 189; POSITION: absolute; TOP: 748px; LEFT: 224px"
					runat="server" Width="64px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:dropdownlist id="DPriceE2" style="Z-INDEX: 188; POSITION: absolute; TOP: 748px; LEFT: 184px"
					runat="server" Width="40px" Height="12px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceE1" style="Z-INDEX: 187; POSITION: absolute; TOP: 748px; LEFT: 16px" runat="server"
					Width="166px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DPriceD3" style="Z-INDEX: 186; POSITION: absolute; TOP: 730px; LEFT: 504px"
					runat="server" Width="64px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:dropdownlist id="DPriceD2" style="Z-INDEX: 185; POSITION: absolute; TOP: 730px; LEFT: 464px"
					runat="server" Width="40px" Height="12px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceD1" style="Z-INDEX: 184; POSITION: absolute; TOP: 730px; LEFT: 296px"
					runat="server" Width="166px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DPriceC3" style="Z-INDEX: 183; POSITION: absolute; TOP: 730px; LEFT: 224px"
					runat="server" Width="64px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:dropdownlist id="DPriceC2" style="Z-INDEX: 182; POSITION: absolute; TOP: 730px; LEFT: 184px"
					runat="server" Width="40px" Height="12px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceC1" style="Z-INDEX: 181; POSITION: absolute; TOP: 730px; LEFT: 16px" runat="server"
					Width="166px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DPriceB3" style="Z-INDEX: 180; POSITION: absolute; TOP: 712px; LEFT: 504px"
					runat="server" Width="64px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:dropdownlist id="DPriceB2" style="Z-INDEX: 179; POSITION: absolute; TOP: 712px; LEFT: 464px"
					runat="server" Width="40px" Height="12px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceB1" style="Z-INDEX: 178; POSITION: absolute; TOP: 712px; LEFT: 296px"
					runat="server" Width="166px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DSampleC3" style="Z-INDEX: 177; POSITION: absolute; TOP: 630px; LEFT: 704px"
					runat="server" Width="52px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DSampleC2" style="Z-INDEX: 176; POSITION: absolute; TOP: 630px; LEFT: 648px"
					runat="server" Width="52px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DSampleC1" style="Z-INDEX: 175; POSITION: absolute; TOP: 630px; LEFT: 440px"
					runat="server" Width="204px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DSampleB3" style="Z-INDEX: 174; POSITION: absolute; TOP: 612px; LEFT: 704px"
					runat="server" Width="52px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DSampleB2" style="Z-INDEX: 173; POSITION: absolute; TOP: 612px; LEFT: 648px"
					runat="server" Width="52px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DSampleB1" style="Z-INDEX: 172; POSITION: absolute; TOP: 612px; LEFT: 440px"
					runat="server" Width="204px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DPriceA1" style="Z-INDEX: 171; POSITION: absolute; TOP: 712px; LEFT: 16px" runat="server"
					Width="166px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DSampleA1" style="Z-INDEX: 170; POSITION: absolute; TOP: 594px; LEFT: 440px"
					runat="server" Width="204px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:button id="BSliderCode" style="Z-INDEX: 168; POSITION: absolute; TOP: 120px; LEFT: 568px"
					runat="server" Width="20px" Height="20px" Text="....."></asp:button><asp:hyperlink id="LContactFile" style="Z-INDEX: 167; POSITION: absolute; TOP: 1144px; LEFT: 104px"
					runat="server" Width="50px" Height="8px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">切結書</asp:hyperlink><INPUT id="DContactFile" style="Z-INDEX: 166; POSITION: absolute; BACKGROUND-COLOR: #ffff00; WIDTH: 264px; HEIGHT: 20px; TOP: 1144px; LEFT: 100px"
					type="file" size="24" name="File1" runat="server">
				<asp:hyperlink id="LRefFile" style="Z-INDEX: 165; POSITION: absolute; TOP: 1078px; LEFT: 104px"
					runat="server" Width="88px" Height="8px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">生產條件表</asp:hyperlink><asp:hyperlink id="LQAFile" style="Z-INDEX: 164; POSITION: absolute; TOP: 1078px; LEFT: 480px"
					runat="server" Width="64px" Height="8px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">測試報告</asp:hyperlink><INPUT id="DQAFile" style="Z-INDEX: 163; POSITION: absolute; BACKGROUND-COLOR: #ffff00; WIDTH: 288px; HEIGHT: 20px; TOP: 1078px; LEFT: 476px"
					type="file" size="28" name="File1" runat="server"><asp:textbox id="DPriceA3" style="Z-INDEX: 152; POSITION: absolute; TOP: 712px; LEFT: 224px"
					runat="server" Width="64px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:dropdownlist id="DPriceA2" style="Z-INDEX: 151; POSITION: absolute; TOP: 712px; LEFT: 184px"
					runat="server" Width="40px" Height="12px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DSampleA2" style="Z-INDEX: 104; POSITION: absolute; TOP: 594px; LEFT: 648px"
					runat="server" Width="52px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:textbox id="DSampleA3" style="Z-INDEX: 106; POSITION: absolute; TOP: 594px; LEFT: 704px"
					runat="server" Width="52px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:button id="BDate" style="Z-INDEX: 149; POSITION: absolute; TOP: 190px; LEFT: 616px" runat="server"
					Width="20px" Height="20px" Text="....."></asp:button><asp:textbox id="DDate" style="Z-INDEX: 148; POSITION: absolute; TOP: 190px; LEFT: 520px" runat="server"
					Width="96px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" ReadOnly="True">DDate</asp:textbox><asp:textbox id="DSliderCode" style="Z-INDEX: 147; POSITION: absolute; TOP: 120px; LEFT: 224px"
					runat="server" Width="342px" Height="56px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="12pt" TextMode="MultiLine" ReadOnly="True">SliderCode</asp:textbox><asp:button id="BMMapNo" style="Z-INDEX: 146; POSITION: absolute; TOP: 293px; LEFT: 440px" runat="server"
					Width="20px" Height="20px" Text="修"></asp:button><asp:hyperlink id="LSampleFile" style="Z-INDEX: 145; POSITION: absolute; TOP: 1144px; LEFT: 480px"
					runat="server" Width="50px" Height="8px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">樣品圖</asp:hyperlink><INPUT id="DSampleFile" style="Z-INDEX: 144; POSITION: absolute; BACKGROUND-COLOR: #ffff00; WIDTH: 288px; HEIGHT: 20px; TOP: 1144px; LEFT: 476px"
					type="file" size="28" name="File1" runat="server"><asp:dropdownlist id="DPerson" style="Z-INDEX: 143; POSITION: absolute; TOP: 224px; LEFT: 520px" runat="server"
					Width="96px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DDivision" style="Z-INDEX: 142; POSITION: absolute; TOP: 224px; LEFT: 312px"
					runat="server" Width="112px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DReasonDesc" style="Z-INDEX: 141; POSITION: absolute; TOP: 1400px; LEFT: 168px"
					runat="server" Width="424px" Height="59px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" TextMode="MultiLine"
					Visible="False">DReasonDesc</asp:textbox><asp:textbox id="DReason" style="Z-INDEX: 140; POSITION: absolute; TOP: 1368px; LEFT: 240px"
					runat="server" Width="352px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" ReadOnly="True" Visible="False">DReason</asp:textbox><asp:dropdownlist id="DReasonCode" style="Z-INDEX: 139; POSITION: absolute; TOP: 1368px; LEFT: 168px"
					runat="server" Width="64px" Height="20px" ForeColor="Blue" BackColor="Yellow" AutoPostBack="True" Visible="False">
					<asp:ListItem Value="01">01</asp:ListItem>
					<asp:ListItem Value="02">02</asp:ListItem>
				</asp:dropdownlist><asp:hyperlink id="LBefOP" style="Z-INDEX: 138; POSITION: absolute; TOP: 1330px; LEFT: 168px" runat="server"
					Width="144px" Height="8px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">工程履歷</asp:hyperlink><asp:textbox id="DAEndTime" style="Z-INDEX: 137; POSITION: absolute; TOP: 1294px; LEFT: 440px"
					runat="server" Width="152px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Gold" Font-Size="9pt" ReadOnly="True">DAEndTime</asp:textbox><asp:textbox id="DAStartTime" style="Z-INDEX: 136; POSITION: absolute; TOP: 1294px; LEFT: 168px"
					runat="server" Width="152px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Gold" Font-Size="9pt" ReadOnly="True">DAStartTime</asp:textbox><asp:textbox id="DBEndTime" style="Z-INDEX: 135; POSITION: absolute; TOP: 1262px; LEFT: 440px"
					runat="server" Width="152px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Gold" Font-Size="9pt" ReadOnly="True">DBEndTime</asp:textbox><asp:textbox id="DBStartTime" style="Z-INDEX: 134; POSITION: absolute; TOP: 1262px; LEFT: 168px"
					runat="server" Width="152px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Gold" Font-Size="9pt" ReadOnly="True">DBStartTime</asp:textbox><asp:textbox id="DDecideDesc" style="Z-INDEX: 133; POSITION: absolute; TOP: 1184px; LEFT: 56px"
					runat="server" Width="536px" Height="56px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" TextMode="MultiLine">DDecideDesc</asp:textbox><asp:textbox id="DMoldPoint" style="Z-INDEX: 132; POSITION: absolute; TOP: 395px; LEFT: 672px"
					runat="server" Width="45px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow">MoldPoint</asp:textbox><asp:textbox id="DMoldQty" style="Z-INDEX: 131; POSITION: absolute; TOP: 395px; LEFT: 568px"
					runat="server" Width="45px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow">DMoldQty</asp:textbox>
				<asp:textbox id="DDevReason" style="Z-INDEX: 126; POSITION: absolute; TOP: 568px; LEFT: 56px"
					runat="server" Width="369px" Height="48px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow"
					MaxLength="240" TextMode="MultiLine">DevReason</asp:textbox><asp:hyperlink id="LAuthorizeFile" style="Z-INDEX: 125; POSITION: absolute; TOP: 1112px; LEFT: 352px"
					runat="server" Width="56px" Height="8px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">授權書</asp:hyperlink><INPUT id="DAuthorizeFile" style="Z-INDEX: 124; POSITION: absolute; BACKGROUND-COLOR: #ffff00; WIDTH: 164px; HEIGHT: 20px; TOP: 1112px; LEFT: 348px"
					type="file" size="8" name="File1" runat="server">
				<asp:hyperlink id="LConfirmFile" style="Z-INDEX: 123; POSITION: absolute; TOP: 1112px; LEFT: 104px"
					runat="server" Width="50px" Height="8px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">確認書</asp:hyperlink><INPUT id="DConfirmFile" style="Z-INDEX: 122; POSITION: absolute; BACKGROUND-COLOR: #ffff00; WIDTH: 160px; HEIGHT: 20px; TOP: 1112px; LEFT: 100px"
					type="file" size="7" name="File1" runat="server"><asp:textbox id="DSellVendor" style="Z-INDEX: 121; POSITION: absolute; TOP: 326px; LEFT: 648px"
					runat="server" Width="111px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow">DSellVendor</asp:textbox><asp:dropdownlist id="DMaterial" style="Z-INDEX: 120; POSITION: absolute; TOP: 395px; LEFT: 104px"
					runat="server" Width="192px" Height="20px" ForeColor="Blue" BackColor="Yellow" Font-Size="7.5pt">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DManufPlace" style="Z-INDEX: 119; POSITION: absolute; TOP: 360px; LEFT: 400px"
					runat="server" Width="80px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DSliderType2" style="Z-INDEX: 118; POSITION: absolute; TOP: 360px; LEFT: 200px"
					runat="server" Width="94px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DSliderType1" style="Z-INDEX: 117; POSITION: absolute; TOP: 360px; LEFT: 104px"
					runat="server" Width="94px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DAssembler" style="Z-INDEX: 116; POSITION: absolute; TOP: 360px; LEFT: 568px"
					runat="server" Width="74px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow">DAssembler</asp:textbox><asp:dropdownlist id="DLevel" style="Z-INDEX: 115; POSITION: absolute; TOP: 259px; LEFT: 648px" runat="server"
					Width="110px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><INPUT id="DRefFile" style="Z-INDEX: 114; POSITION: absolute; BACKGROUND-COLOR: #ffff00; WIDTH: 264px; HEIGHT: 20px; TOP: 1078px; LEFT: 100px"
					type="file" size="24" name="File1" runat="server">
				<asp:button id="BOMapNo" style="Z-INDEX: 113; POSITION: absolute; TOP: 293px; LEFT: 416px" runat="server"
					Width="20px" Height="20px" Text="原"></asp:button><asp:textbox id="DMapNo" style="Z-INDEX: 112; POSITION: absolute; TOP: 293px; LEFT: 312px" runat="server"
					Width="104px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" AutoPostBack="True" ReadOnly="True">123456789012345</asp:textbox><asp:button id="BSpec" style="Z-INDEX: 111; POSITION: absolute; TOP: 259px; LEFT: 528px" runat="server"
					Width="20px" Height="20px" Text="....."></asp:button><asp:textbox id="DSpec" style="Z-INDEX: 110; POSITION: absolute; TOP: 259px; LEFT: 312px" runat="server"
					Width="216px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True">DSpec</asp:textbox><asp:textbox id="DSliderGRCode" style="Z-INDEX: 109; POSITION: absolute; TOP: 156px; LEFT: 604px"
					runat="server" Width="154px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" ReadOnly="True">DSliderGRCode</asp:textbox><asp:textbox id="DNo" style="Z-INDEX: 108; POSITION: absolute; TOP: 190px; LEFT: 312px" runat="server"
					Width="112px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow">DNo</asp:textbox><asp:textbox id="DFormSno" style="Z-INDEX: 107; POSITION: absolute; TOP: 1472px; LEFT: 16px"
					runat="server" Width="97px" Height="20px" BorderStyle="None" ForeColor="Blue" BackColor="White">單號：123</asp:textbox><asp:image id="DDelivery" style="Z-INDEX: 105; POSITION: absolute; TOP: 1248px; LEFT: 8px"
					runat="server" Width="593px" Height="110px" ImageUrl="Images\Sheet_Delivery.jpg"></asp:image><asp:image id="DDelay" style="Z-INDEX: 103; POSITION: absolute; TOP: 1360px; LEFT: 8px" runat="server"
					Width="593px" Height="107px" Visible="False" ImageUrl="Images\Sheet_Delay.jpg"></asp:image><asp:image id="DDescSheet" style="Z-INDEX: 102; POSITION: absolute; TOP: 1176px; LEFT: 8px"
					runat="server" Width="593px" Height="75px" ImageUrl="Images\MapSheet_004.jpg"></asp:image><asp:image id="DManuaInSheet2" style="Z-INDEX: 101; POSITION: absolute; TOP: 680px; LEFT: 8px"
					runat="server" Width="761px" Height="493px" ImageUrl="Images\ManufOutSheet_006_c.jpg"></asp:image><asp:image id="LMapFile" style="Z-INDEX: 242; POSITION: absolute; TOP: 120px; LEFT: 16px" runat="server"
					Width="200px" Height="230px" BorderStyle="Groove" ImageUrl="http://10.245.0.178//WorkFlow/Document/000003/1-Map2-2006124111750-2.jpg"></asp:image></FONT><INPUT id="BSAVE" style="Z-INDEX: 246; POSITION: absolute; WIDTH: 80px; HEIGHT: 28px; TOP: 1472px; LEFT: 256px"
				onclick="Button('SAVE');" type="button" value="儲存" name="Button1" runat="server"><INPUT id="BNG2" style="Z-INDEX: 247; POSITION: absolute; WIDTH: 80px; HEIGHT: 28px; TOP: 1472px; LEFT: 344px"
				onclick="Button('NG2');" type="button" value="NG2" name="Button1" runat="server"><INPUT id="BNG1" style="Z-INDEX: 248; POSITION: absolute; WIDTH: 80px; HEIGHT: 28px; TOP: 1472px; LEFT: 432px"
				onclick="Button('NG1');" type="button" value="NG1" name="Button1" runat="server"><INPUT id="BOK" style="Z-INDEX: 249; POSITION: absolute; WIDTH: 80px; HEIGHT: 28px; TOP: 1472px; LEFT: 520px"
				onclick="Button('OK');" type="button" value="OK" name="Button2" runat="server">
			<asp:imagebutton id="BPrint" style="Z-INDEX: 259; POSITION: absolute; TOP: 8px; LEFT: 8px" runat="server"
				Width="16px" Height="16px" ImageUrl="Images\Print.gif"></asp:imagebutton></form>
	</body>
</HTML>
