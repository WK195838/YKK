<%@ Page Language="vb" AutoEventWireup="false" Codebehind="ManufOutSheet_02.aspx.vb" Inherits="SPD.ManufOutSheet_021"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>外注委託書</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<script language="javascript" type="text/javascript">
			var wPop;
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
				wPop=window.open('SliderPicker.aspx?field=' + strField,'SliderPopup','width=300,height=250,resizable=yes');
				if ((document.Form1.DSliderGRCode.value != "") || (document.Form1.DSliderCode.value != "")) {
					setTimeout("SendToSlider()",200);
				}
			}
		    function SendToSlider() {
				wPop.document.SliderForm.DSlider2.value=document.Form1.DSliderGRCode.value;
				wPop.document.SliderForm.DContent.value=document.Form1.DSliderCode.value;
			}
			//--Spec------------------------------------
			function SpecPicker(strField) {
				wPop=window.open('SpecPicker.aspx?field=' + strField,'SpecPopup','width=330,height=250,resizable=yes');
				if (document.Form1.DSpec.value != "") {
					setTimeout("SendToSpec()",200);
				}
			}
		    function SendToSpec() {
				wPop.document.SpecForm.DSpec.value=document.Form1.DSpec.value;
			}
		    function Button(F) {
				//alert(F);
				document.cookie="RunBOK=False";
				document.cookie="RunBNG1=False";
				document.cookie="RunBNG2=False";
				document.cookie="RunBSAVE=False";
				if (F=="OK")   answer = confirm("是否執行[申請/完工/OK]作業嗎？ 請確認....");
				if (F=="NG1")  answer = confirm("是否執行[跳前工程/NG]作業嗎？ 請確認....");
				if (F=="NG2")  answer = confirm("是否執行[跳前工程/NG/結束]作業嗎？ 請確認....");
				if (F=="SAVE") answer = confirm("是否執行[暫存資料]作業嗎？ 請確認....");
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
		</script>
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" encType="multipart/form-data" method="post" runat="server">
			<FONT face="新細明體">
				<asp:textbox style="Z-INDEX: 321; LEFT: 464px; POSITION: absolute; TOP: 328px" id="DCustomerCode"
					runat="server" Width="88px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow">DCustomerCode</asp:textbox><asp:hyperlink style="Z-INDEX: 385; LEFT: 400px; POSITION: absolute; TOP: 1044px" id="LAssemblerFile"
					runat="server" Width="129px" Height="20px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">組立判定書</asp:hyperlink><asp:label style="Z-INDEX: 384; LEFT: 304px; POSITION: absolute; TOP: 1044px" id="Label42"
					runat="server" Width="81px" Height="20px">組立判定書</asp:label><asp:label style="Z-INDEX: 346; LEFT: 608px; POSITION: absolute; TOP: 1048px" id="Label36"
					runat="server" Width="48px" Height="10px">QC-L/T</asp:label><asp:textbox style="Z-INDEX: 317; LEFT: 664px; POSITION: absolute; TOP: 1044px" id="DQCLT" runat="server"
					Width="81px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:label style="Z-INDEX: 315; LEFT: 744px; POSITION: absolute; TOP: 1044px" id="Label37"
					runat="server" Width="1px" Height="10px">分</asp:label><asp:hyperlink style="Z-INDEX: 345; LEFT: 152px; POSITION: absolute; TOP: 32px" id="LYKKGroupCopy"
					runat="server" Width="144px" Height="16px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">姊妹社圖面複製履歷</asp:hyperlink><asp:textbox style="Z-INDEX: 344; LEFT: 644px; POSITION: absolute; TOP: 360px" id="DAssembler1"
					runat="server" Width="116px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" ReadOnly="True">DAssembler</asp:textbox><asp:label style="Z-INDEX: 343; LEFT: 688px; POSITION: absolute; TOP: 448px" id="Label40" runat="server"
					Width="73px" Height="10px" Font-Size="8pt">CYCLE TIME</asp:label><asp:textbox style="Z-INDEX: 340; LEFT: 688px; POSITION: absolute; TOP: 464px" id="DCYCLETIME"
					runat="server" Width="73px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:label style="Z-INDEX: 342; LEFT: 688px; POSITION: absolute; TOP: 490px" id="Label41" runat="server"
					Width="65px" Height="8px" Font-Size="8pt">稼動時間/日</asp:label><asp:textbox style="Z-INDEX: 341; LEFT: 688px; POSITION: absolute; TOP: 504px" id="DWorkTime"
					runat="server" Width="72px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:image style="Z-INDEX: 339; LEFT: 624px; POSITION: absolute; TOP: 8px" id="DMMSSts" runat="server"
					Width="145px" Height="48px" BorderStyle="None" ImageUrl="Images\mmsdelete.jpg"></asp:image><asp:label style="Z-INDEX: 338; LEFT: 576px; POSITION: absolute; TOP: 544px" id="Label39" runat="server"
					Width="80px" Height="10px">LOSS率(%)</asp:label><asp:textbox style="Z-INDEX: 337; LEFT: 656px; POSITION: absolute; TOP: 540px" id="DLOSS" runat="server"
					Width="103px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" MaxLength="10">DLOSS</asp:textbox><asp:dropdownlist style="Z-INDEX: 336; LEFT: 712px; POSITION: absolute; TOP: 888px" id="DQAD15" runat="server"
					Width="48px" Height="22px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 316; LEFT: 712px; POSITION: absolute; TOP: 872px" id="DQAC15" runat="server"
					Width="48px" Height="22px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 313; LEFT: 712px; POSITION: absolute; TOP: 856px" id="DQAB15" runat="server"
					Width="48px" Height="22px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:label style="Z-INDEX: 310; LEFT: 712px; POSITION: absolute; TOP: 824px" id="Label38" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">EDX</asp:label><asp:dropdownlist style="Z-INDEX: 307; LEFT: 712px; POSITION: absolute; TOP: 840px" id="DQAA15" runat="server"
					Width="48px" Height="22px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 295; LEFT: 616px; POSITION: absolute; TOP: 888px" id="DQAD13" runat="server"
					Width="48px" Height="346px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0級">0級</asp:ListItem>
					<asp:ListItem Value="1級">1級</asp:ListItem>
					<asp:ListItem Value="2級">2級</asp:ListItem>
					<asp:ListItem Value="3級">3級</asp:ListItem>
					<asp:ListItem Value="4級">4級</asp:ListItem>
					<asp:ListItem Value="5級">5級</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 294; LEFT: 616px; POSITION: absolute; TOP: 872px" id="DQAC13" runat="server"
					Width="48px" Height="346px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0級">0級</asp:ListItem>
					<asp:ListItem Value="1級">1級</asp:ListItem>
					<asp:ListItem Value="2級">2級</asp:ListItem>
					<asp:ListItem Value="3級">3級</asp:ListItem>
					<asp:ListItem Value="4級">4級</asp:ListItem>
					<asp:ListItem Value="5級">5級</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 293; LEFT: 616px; POSITION: absolute; TOP: 856px" id="DQAB13" runat="server"
					Width="48px" Height="346px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0級">0級</asp:ListItem>
					<asp:ListItem Value="1級">1級</asp:ListItem>
					<asp:ListItem Value="2級">2級</asp:ListItem>
					<asp:ListItem Value="3級">3級</asp:ListItem>
					<asp:ListItem Value="4級">4級</asp:ListItem>
					<asp:ListItem Value="5級">5級</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 292; LEFT: 568px; POSITION: absolute; TOP: 888px" id="DQAD12" runat="server"
					Width="48px" Height="346px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0級">0級</asp:ListItem>
					<asp:ListItem Value="1級">1級</asp:ListItem>
					<asp:ListItem Value="2級">2級</asp:ListItem>
					<asp:ListItem Value="3級">3級</asp:ListItem>
					<asp:ListItem Value="4級">4級</asp:ListItem>
					<asp:ListItem Value="5級">5級</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 291; LEFT: 568px; POSITION: absolute; TOP: 872px" id="DQAC12" runat="server"
					Width="48px" Height="346px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0級">0級</asp:ListItem>
					<asp:ListItem Value="1級">1級</asp:ListItem>
					<asp:ListItem Value="2級">2級</asp:ListItem>
					<asp:ListItem Value="3級">3級</asp:ListItem>
					<asp:ListItem Value="4級">4級</asp:ListItem>
					<asp:ListItem Value="5級">5級</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 290; LEFT: 568px; POSITION: absolute; TOP: 856px" id="DQAB12" runat="server"
					Width="48px" Height="346px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0級">0級</asp:ListItem>
					<asp:ListItem Value="1級">1級</asp:ListItem>
					<asp:ListItem Value="2級">2級</asp:ListItem>
					<asp:ListItem Value="3級">3級</asp:ListItem>
					<asp:ListItem Value="4級">4級</asp:ListItem>
					<asp:ListItem Value="5級">5級</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 289; LEFT: 520px; POSITION: absolute; TOP: 888px" id="DQAD11" runat="server"
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
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 288; LEFT: 520px; POSITION: absolute; TOP: 872px" id="DQAC11" runat="server"
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
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 287; LEFT: 520px; POSITION: absolute; TOP: 856px" id="DQAB11" runat="server"
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
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 286; LEFT: 616px; POSITION: absolute; TOP: 840px" id="DQAA13" runat="server"
					Width="48px" Height="346px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0級">0級</asp:ListItem>
					<asp:ListItem Value="1級">1級</asp:ListItem>
					<asp:ListItem Value="2級">2級</asp:ListItem>
					<asp:ListItem Value="3級">3級</asp:ListItem>
					<asp:ListItem Value="4級">4級</asp:ListItem>
					<asp:ListItem Value="5級">5級</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 285; LEFT: 568px; POSITION: absolute; TOP: 840px" id="DQAA12" runat="server"
					Width="48px" Height="346px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0級">0級</asp:ListItem>
					<asp:ListItem Value="1級">1級</asp:ListItem>
					<asp:ListItem Value="2級">2級</asp:ListItem>
					<asp:ListItem Value="3級">3級</asp:ListItem>
					<asp:ListItem Value="4級">4級</asp:ListItem>
					<asp:ListItem Value="5級">5級</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 284; LEFT: 520px; POSITION: absolute; TOP: 840px" id="DQAA11" runat="server"
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
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 232; LEFT: 664px; POSITION: absolute; TOP: 888px" id="DQAD14" runat="server"
					Width="48px" Height="22px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 230; LEFT: 472px; POSITION: absolute; TOP: 888px" id="DQAD10" runat="server"
					Width="48px" Height="18px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 228; LEFT: 424px; POSITION: absolute; TOP: 888px" id="DQAD9" runat="server"
					Width="48px" Height="18px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 227; LEFT: 376px; POSITION: absolute; TOP: 888px" id="DQAD8" runat="server"
					Width="48px" Height="18px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="F06">F06</asp:ListItem>
					<asp:ListItem Value="F07">F07</asp:ListItem>
					<asp:ListItem Value="F08">F08</asp:ListItem>
					<asp:ListItem Value="F09">F09</asp:ListItem>
					<asp:ListItem Value="NM-12">NM-12</asp:ListItem>
					<asp:ListItem Value="NCA">NCA</asp:ListItem>
					<asp:ListItem Value="NCB">NCB</asp:ListItem>
					<asp:ListItem Value="NG">NG</asp:ListItem>
					<asp:ListItem Value="KENSIN">KENSIN</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 224; LEFT: 328px; POSITION: absolute; TOP: 888px" id="DQAD7" runat="server"
					Width="48px" Height="18px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 222; LEFT: 280px; POSITION: absolute; TOP: 888px" id="DQAD6" runat="server"
					Width="48px" Height="18px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 220; LEFT: 232px; POSITION: absolute; TOP: 888px" id="DQAD5" runat="server"
					Width="47px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt"></asp:textbox><asp:textbox style="Z-INDEX: 219; LEFT: 204px; POSITION: absolute; TOP: 888px" id="DQAD4" runat="server"
					Width="28px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt"></asp:textbox><asp:dropdownlist style="Z-INDEX: 217; LEFT: 170px; POSITION: absolute; TOP: 888px" id="DQAD3" runat="server"
					Width="35px" Height="18px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 215; LEFT: 72px; POSITION: absolute; TOP: 888px" id="DQAD2" runat="server"
					Width="98px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt"></asp:textbox><asp:textbox style="Z-INDEX: 212; LEFT: 16px; POSITION: absolute; TOP: 888px" id="DQAD1" runat="server"
					Width="56px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt"></asp:textbox><asp:dropdownlist style="Z-INDEX: 211; LEFT: 664px; POSITION: absolute; TOP: 872px" id="DQAC14" runat="server"
					Width="48px" Height="22px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 208; LEFT: 472px; POSITION: absolute; TOP: 872px" id="DQAC10" runat="server"
					Width="48px" Height="18px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 207; LEFT: 424px; POSITION: absolute; TOP: 872px" id="DQAC9" runat="server"
					Width="48px" Height="18px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 205; LEFT: 376px; POSITION: absolute; TOP: 872px" id="DQAC8" runat="server"
					Width="48px" Height="36px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="F06">F06</asp:ListItem>
					<asp:ListItem Value="F07">F07</asp:ListItem>
					<asp:ListItem Value="F08">F08</asp:ListItem>
					<asp:ListItem Value="F09">F09</asp:ListItem>
					<asp:ListItem Value="NM-12">NM-12</asp:ListItem>
					<asp:ListItem Value="NCA">NCA</asp:ListItem>
					<asp:ListItem Value="NCB">NCB</asp:ListItem>
					<asp:ListItem Value="NG">NG</asp:ListItem>
					<asp:ListItem Value="KENSIN">KENSIN</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 203; LEFT: 328px; POSITION: absolute; TOP: 872px" id="DQAC7" runat="server"
					Width="48px" Height="18px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 201; LEFT: 280px; POSITION: absolute; TOP: 872px" id="DQAC6" runat="server"
					Width="48px" Height="18px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 199; LEFT: 232px; POSITION: absolute; TOP: 872px" id="DQAC5" runat="server"
					Width="47px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt"></asp:textbox><asp:textbox style="Z-INDEX: 198; LEFT: 204px; POSITION: absolute; TOP: 872px" id="DQAC4" runat="server"
					Width="28px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt"></asp:textbox><asp:dropdownlist style="Z-INDEX: 196; LEFT: 170px; POSITION: absolute; TOP: 872px" id="DQAC3" runat="server"
					Width="35px" Height="18px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 195; LEFT: 72px; POSITION: absolute; TOP: 872px" id="DQAC2" runat="server"
					Width="98px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt"></asp:textbox><asp:textbox style="Z-INDEX: 193; LEFT: 16px; POSITION: absolute; TOP: 872px" id="DQAC1" runat="server"
					Width="56px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt"></asp:textbox><asp:dropdownlist style="Z-INDEX: 192; LEFT: 664px; POSITION: absolute; TOP: 856px" id="DQAB14" runat="server"
					Width="48px" Height="22px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 191; LEFT: 472px; POSITION: absolute; TOP: 856px" id="DQAB10" runat="server"
					Width="48px" Height="18px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 189; LEFT: 424px; POSITION: absolute; TOP: 856px" id="DQAB9" runat="server"
					Width="48px" Height="18px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 188; LEFT: 376px; POSITION: absolute; TOP: 856px" id="DQAB8" runat="server"
					Width="48px" Height="36px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="F06">F06</asp:ListItem>
					<asp:ListItem Value="F07">F07</asp:ListItem>
					<asp:ListItem Value="F08">F08</asp:ListItem>
					<asp:ListItem Value="F09">F09</asp:ListItem>
					<asp:ListItem Value="NM-12">NM-12</asp:ListItem>
					<asp:ListItem Value="NCA">NCA</asp:ListItem>
					<asp:ListItem Value="NCB">NCB</asp:ListItem>
					<asp:ListItem Value="NG">NG</asp:ListItem>
					<asp:ListItem Value="KENSIN">KENSIN</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 187; LEFT: 328px; POSITION: absolute; TOP: 856px" id="DQAB7" runat="server"
					Width="48px" Height="18px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 186; LEFT: 280px; POSITION: absolute; TOP: 856px" id="DQAB6" runat="server"
					Width="48px" Height="18px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 185; LEFT: 232px; POSITION: absolute; TOP: 856px" id="DQAB5" runat="server"
					Width="47px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt"></asp:textbox><asp:textbox style="Z-INDEX: 184; LEFT: 204px; POSITION: absolute; TOP: 856px" id="DQAB4" runat="server"
					Width="28px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt"></asp:textbox><asp:dropdownlist style="Z-INDEX: 183; LEFT: 170px; POSITION: absolute; TOP: 856px" id="DQAB3" runat="server"
					Width="35px" Height="18px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 182; LEFT: 72px; POSITION: absolute; TOP: 856px" id="DQAB2" runat="server"
					Width="98px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt"></asp:textbox><asp:textbox style="Z-INDEX: 181; LEFT: 16px; POSITION: absolute; TOP: 856px" id="DQAB1" runat="server"
					Width="56px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt"></asp:textbox><asp:dropdownlist style="Z-INDEX: 162; LEFT: 664px; POSITION: absolute; TOP: 840px" id="DQAA14" runat="server"
					Width="48px" Height="22px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 160; LEFT: 472px; POSITION: absolute; TOP: 840px" id="DQAA10" runat="server"
					Width="48px" Height="346px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 158; LEFT: 424px; POSITION: absolute; TOP: 840px" id="DQAA9" runat="server"
					Width="48px" Height="18px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 156; LEFT: 376px; POSITION: absolute; TOP: 840px" id="DQAA8" runat="server"
					Width="48px" Height="36px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="F06">F06</asp:ListItem>
					<asp:ListItem Value="F07">F07</asp:ListItem>
					<asp:ListItem Value="F08">F08</asp:ListItem>
					<asp:ListItem Value="F09">F09</asp:ListItem>
					<asp:ListItem Value="NM-12">NM-12</asp:ListItem>
					<asp:ListItem Value="NCA">NCA</asp:ListItem>
					<asp:ListItem Value="NCB">NCB</asp:ListItem>
					<asp:ListItem Value="NG">NG</asp:ListItem>
					<asp:ListItem Value="KENSIN">KENSIN</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 153; LEFT: 328px; POSITION: absolute; TOP: 840px" id="DQAA7" runat="server"
					Width="48px" Height="162px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 151; LEFT: 280px; POSITION: absolute; TOP: 840px" id="DQAA6" runat="server"
					Width="48px" Height="18px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 149; LEFT: 232px; POSITION: absolute; TOP: 840px" id="DQAA5" runat="server"
					Width="47px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt"></asp:textbox><asp:textbox style="Z-INDEX: 148; LEFT: 204px; POSITION: absolute; TOP: 840px" id="DQAA4" runat="server"
					Width="28px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt"></asp:textbox><asp:dropdownlist style="Z-INDEX: 146; LEFT: 170px; POSITION: absolute; TOP: 840px" id="DQAA3" runat="server"
					Width="35px" Height="18px" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 144; LEFT: 16px; POSITION: absolute; TOP: 840px" id="DQAA1" runat="server"
					Width="56px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt"></asp:textbox><asp:textbox style="Z-INDEX: 168; LEFT: 72px; POSITION: absolute; TOP: 840px" id="DQAA2" runat="server"
					Width="98px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="7pt"></asp:textbox><asp:label style="Z-INDEX: 283; LEFT: 48px; POSITION: absolute; TOP: 824px" id="Label1" runat="server"
					Width="24px" Height="10px" Font-Size="8pt">日期</asp:label><asp:label style="Z-INDEX: 282; LEFT: 664px; POSITION: absolute; TOP: 824px" id="Label13" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">CPSC</asp:label><asp:label style="Z-INDEX: 297; LEFT: 568px; POSITION: absolute; TOP: 824px" id="Label35" runat="server"
					Width="48px" Height="10px" Font-Size="8pt">一次密著</asp:label><asp:label style="Z-INDEX: 281; LEFT: 616px; POSITION: absolute; TOP: 824px" id="Label12" runat="server"
					Width="48px" Height="10px" Font-Size="8pt">二次密著</asp:label><asp:label style="Z-INDEX: 280; LEFT: 520px; POSITION: absolute; TOP: 824px" id="Label11" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">黃變</asp:label><asp:label style="Z-INDEX: 278; LEFT: 472px; POSITION: absolute; TOP: 824px" id="Label10" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">乾洗</asp:label><asp:label style="Z-INDEX: 275; LEFT: 424px; POSITION: absolute; TOP: 824px" id="Label9" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">水洗</asp:label><asp:label style="Z-INDEX: 274; LEFT: 376px; POSITION: absolute; TOP: 824px" id="Label8" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">檢針</asp:label><asp:label style="Z-INDEX: 272; LEFT: 328px; POSITION: absolute; TOP: 824px" id="Label7" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">扭力</asp:label><asp:label style="Z-INDEX: 269; LEFT: 280px; POSITION: absolute; TOP: 824px" id="Label6" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">強度</asp:label><asp:label style="Z-INDEX: 266; LEFT: 240px; POSITION: absolute; TOP: 824px" id="Label5" runat="server"
					Width="40px" Height="10px" Font-Size="8pt">原單位</asp:label><asp:label style="Z-INDEX: 263; LEFT: 208px; POSITION: absolute; TOP: 824px" id="Label4" runat="server"
					Width="48px" Height="10px" Font-Size="8pt">表面</asp:label><asp:label style="Z-INDEX: 261; LEFT: 168px; POSITION: absolute; TOP: 824px" id="Label3" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">口厚</asp:label><asp:label style="Z-INDEX: 257; LEFT: 80px; POSITION: absolute; TOP: 824px" id="Label2" runat="server"
					Width="48px" Height="10px" Font-Size="8pt">拉頭種類</asp:label><asp:hyperlink style="Z-INDEX: 335; LEFT: 152px; POSITION: absolute; TOP: 56px" id="LSurface" runat="server"
					Width="80px" Height="8px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">有表面處理</asp:hyperlink><asp:hyperlink style="Z-INDEX: 334; LEFT: 240px; POSITION: absolute; TOP: 56px" id="LColorAppend"
					runat="server" Width="112px" Height="8px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">有外注色番追加</asp:hyperlink><asp:dropdownlist style="Z-INDEX: 319; LEFT: 686px; POSITION: absolute; TOP: 224px" id="DLogo" runat="server"
					Width="78px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 131; LEFT: 688px; POSITION: absolute; TOP: 786px" id="DPullerPrice1"
					runat="server" Width="72px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow">PullerPrice1</asp:textbox><asp:hyperlink style="Z-INDEX: 333; LEFT: 608px; POSITION: absolute; TOP: 1112px" id="LForCastFile"
					runat="server" Width="112px" Height="8px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">報價單</asp:hyperlink><asp:textbox style="Z-INDEX: 129; LEFT: 688px; POSITION: absolute; TOP: 752px" id="DPullerPrice"
					runat="server" Width="72px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow">PullerPrice</asp:textbox><asp:textbox style="Z-INDEX: 127; LEFT: 688px; POSITION: absolute; TOP: 718px" id="DPurMold"
					runat="server" Width="72px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow">PurMold</asp:textbox><asp:textbox style="Z-INDEX: 125; LEFT: 688px; POSITION: absolute; TOP: 686px" id="DArMoldFee"
					runat="server" Width="72px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow">ArMoldFee</asp:textbox><asp:textbox style="Z-INDEX: 332; LEFT: 56px; POSITION: absolute; TOP: 624px" id="DManufFlow"
					runat="server" Width="369px" Height="50px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" MaxLength="240" TextMode="MultiLine">DManufFlow</asp:textbox><asp:hyperlink style="Z-INDEX: 331; LEFT: 464px; POSITION: absolute; TOP: 296px" id="LMapNo1" runat="server"
					Width="112px" Height="8px" NavigateUrl="BoardEdit.aspx" Target="_blank">圖號文件</asp:hyperlink><asp:dropdownlist style="Z-INDEX: 330; LEFT: 708px; POSITION: absolute; TOP: 190px" id="DCpsc" runat="server"
					Width="56px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 329; LEFT: 560px; POSITION: absolute; TOP: 16px" id="DStatus" runat="server"
					Width="212px" Height="32px" BorderStyle="Groove" ForeColor="White" BackColor="Red" Font-Size="10pt" ReadOnly="True">修改工程進行中(xxxxxxxxxxxxxxxxxx)</asp:textbox><asp:hyperlink style="Z-INDEX: 328; LEFT: 416px; POSITION: absolute; TOP: 56px" id="LOPContact"
					runat="server" Width="112px" Height="8px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">有型別胴體追加</asp:hyperlink><asp:hyperlink style="Z-INDEX: 327; LEFT: 536px; POSITION: absolute; TOP: 56px" id="LContact" runat="server"
					Width="112px" Height="8px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">有追加業務連絡</asp:hyperlink><asp:hyperlink style="Z-INDEX: 326; LEFT: 656px; POSITION: absolute; TOP: 56px" id="LSliderDetail"
					runat="server" Width="112px" Height="16px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">有追加拉頭細目</asp:hyperlink><asp:label style="Z-INDEX: 309; LEFT: 704px; POSITION: absolute; TOP: 580px" id="Label25" runat="server"
					Width="24px" Height="10px" Font-Size="8pt">數量</asp:label><asp:label style="Z-INDEX: 325; LEFT: 136px; POSITION: absolute; TOP: 440px" id="Label32" runat="server"
					Width="32px" Height="10px" Font-Size="8pt" Visible="False">品名-4</asp:label><asp:label style="Z-INDEX: 324; LEFT: 96px; POSITION: absolute; TOP: 440px" id="Label31" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">品名-3</asp:label><asp:label style="Z-INDEX: 323; LEFT: 56px; POSITION: absolute; TOP: 440px" id="Label30" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">品名-2</asp:label><asp:label style="Z-INDEX: 322; LEFT: 568px; POSITION: absolute; TOP: 440px" id="Label34" runat="server"
					Width="24px" Height="10px" Font-Size="8pt">備註</asp:label><asp:label style="Z-INDEX: 320; LEFT: 472px; POSITION: absolute; TOP: 440px" id="Label33" runat="server"
					Width="64px" Height="10px" Font-Size="8pt">預定完成日</asp:label><asp:label style="Z-INDEX: 318; LEFT: 376px; POSITION: absolute; TOP: 440px" id="Label29" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">OR-NO</asp:label><asp:label style="Z-INDEX: 314; LEFT: 280px; POSITION: absolute; TOP: 440px" id="Label28" runat="server"
					Width="56px" Height="10px" Font-Size="8pt">基準日程</asp:label><asp:label style="Z-INDEX: 312; LEFT: 184px; POSITION: absolute; TOP: 440px" id="Label27" runat="server"
					Width="64px" Height="10px" Font-Size="8pt">一日產能</asp:label><asp:label style="Z-INDEX: 311; LEFT: 16px; POSITION: absolute; TOP: 440px" id="Label26" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">品名-1</asp:label><asp:label style="Z-INDEX: 308; LEFT: 648px; POSITION: absolute; TOP: 580px" id="Label24" runat="server"
					Width="24px" Height="10px" Font-Size="8pt">顏色</asp:label><asp:label style="Z-INDEX: 306; LEFT: 472px; POSITION: absolute; TOP: 580px" id="Label23" runat="server"
					Width="48px" Height="10px" Font-Size="8pt">拉頭種類</asp:label><asp:label style="Z-INDEX: 305; LEFT: 504px; POSITION: absolute; TOP: 700px" id="Label22" runat="server"
					Width="24px" Height="10px" Font-Size="8pt">單價</asp:label><asp:label style="Z-INDEX: 304; LEFT: 464px; POSITION: absolute; TOP: 700px" id="Label21" runat="server"
					Width="24px" Height="10px" Font-Size="8pt">幣別</asp:label><asp:label style="Z-INDEX: 303; LEFT: 296px; POSITION: absolute; TOP: 700px" id="Label20" runat="server"
					Width="48px" Height="10px" Font-Size="8pt">拉頭種類</asp:label><asp:label style="Z-INDEX: 302; LEFT: 224px; POSITION: absolute; TOP: 700px" id="Label19" runat="server"
					Width="24px" Height="10px" Font-Size="8pt">單價</asp:label><asp:label style="Z-INDEX: 301; LEFT: 184px; POSITION: absolute; TOP: 700px" id="Label18" runat="server"
					Width="24px" Height="10px" Font-Size="8pt">幣別</asp:label><asp:label style="Z-INDEX: 300; LEFT: 56px; POSITION: absolute; TOP: 700px" id="Label17" runat="server"
					Width="48px" Height="10px" Font-Size="8pt">拉頭種類</asp:label><asp:label style="Z-INDEX: 299; LEFT: 136px; POSITION: absolute; TOP: 956px" id="Label16" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">備註</asp:label><asp:label style="Z-INDEX: 298; LEFT: 80px; POSITION: absolute; TOP: 956px" id="Label15" runat="server"
					Width="48px" Height="10px" Font-Size="8pt">檢測結果</asp:label><asp:label style="Z-INDEX: 296; LEFT: 48px; POSITION: absolute; TOP: 956px" id="Label14" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">日期</asp:label><asp:hyperlink style="Z-INDEX: 276; LEFT: 440px; POSITION: absolute; TOP: 652px" id="LSAttachFile"
					runat="server" Width="64px" Height="8px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">其他附件</asp:hyperlink><asp:hyperlink style="Z-INDEX: 279; LEFT: 136px; POSITION: absolute; TOP: 1044px" id="LQAttachFile2"
					runat="server" Width="128px" Height="8px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">品質分析依賴書</asp:hyperlink><asp:hyperlink style="Z-INDEX: 277; LEFT: 16px; POSITION: absolute; TOP: 916px" id="LQAttachFile1"
					runat="server" Width="64px" Height="8px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">其他附件</asp:hyperlink><asp:image style="Z-INDEX: 100; LEFT: 8px; POSITION: absolute; TOP: 8px" id="DManuaInSheet1"
					runat="server" Width="761px" Height="674px" ImageUrl="Images\ManufOutSheet_003_f.jpg"></asp:image><asp:textbox style="Z-INDEX: 273; LEFT: 8px; POSITION: absolute; TOP: 1176px" id="DFormSno" runat="server"
					Width="97px" Height="20px" BorderStyle="None" ForeColor="Blue" BackColor="White">單號：123</asp:textbox><asp:image style="Z-INDEX: 270; LEFT: 8px; POSITION: absolute; TOP: 1176px" id="Image1" runat="server"
					Width="755px" Height="29px" ImageUrl="Images\ManufOutSheet_NEW_004.jpg"></asp:image><asp:textbox style="Z-INDEX: 268; LEFT: 568px; POSITION: absolute; TOP: 506px" id="DHAD9" runat="server"
					Width="110px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 267; LEFT: 472px; POSITION: absolute; TOP: 506px" id="DHAD8" runat="server"
					Width="88px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 265; LEFT: 376px; POSITION: absolute; TOP: 506px" id="DHAD7" runat="server"
					Width="88px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 264; LEFT: 280px; POSITION: absolute; TOP: 506px" id="DHAD6" runat="server"
					Width="88px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 262; LEFT: 184px; POSITION: absolute; TOP: 506px" id="DHAD5" runat="server"
					Width="88px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 260; LEFT: 136px; POSITION: absolute; TOP: 506px" id="DHAD4" runat="server"
					Width="40px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True" Visible="False"></asp:textbox><asp:textbox style="Z-INDEX: 259; LEFT: 96px; POSITION: absolute; TOP: 506px" id="DHAD3" runat="server"
					Width="80px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 258; LEFT: 56px; POSITION: absolute; TOP: 506px" id="DHAD2" runat="server"
					Width="40px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 256; LEFT: 16px; POSITION: absolute; TOP: 506px" id="DHAD1" runat="server"
					Width="40px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 255; LEFT: 568px; POSITION: absolute; TOP: 488px" id="DHAC9" runat="server"
					Width="110px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 254; LEFT: 472px; POSITION: absolute; TOP: 488px" id="DHAC8" runat="server"
					Width="88px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 253; LEFT: 376px; POSITION: absolute; TOP: 488px" id="DHAC7" runat="server"
					Width="88px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 252; LEFT: 280px; POSITION: absolute; TOP: 488px" id="DHAC6" runat="server"
					Width="88px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 251; LEFT: 184px; POSITION: absolute; TOP: 488px" id="DHAC5" runat="server"
					Width="88px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 250; LEFT: 136px; POSITION: absolute; TOP: 488px" id="DHAC4" runat="server"
					Width="40px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True" Visible="False"></asp:textbox><asp:textbox style="Z-INDEX: 249; LEFT: 96px; POSITION: absolute; TOP: 488px" id="DHAC3" runat="server"
					Width="80px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 248; LEFT: 56px; POSITION: absolute; TOP: 488px" id="DHAC2" runat="server"
					Width="40px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 247; LEFT: 16px; POSITION: absolute; TOP: 488px" id="DHAC1" runat="server"
					Width="40px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 246; LEFT: 568px; POSITION: absolute; TOP: 470px" id="DHAB9" runat="server"
					Width="110px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 245; LEFT: 472px; POSITION: absolute; TOP: 470px" id="DHAB8" runat="server"
					Width="88px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 244; LEFT: 376px; POSITION: absolute; TOP: 470px" id="DHAB7" runat="server"
					Width="88px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 243; LEFT: 280px; POSITION: absolute; TOP: 470px" id="DHAB6" runat="server"
					Width="88px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 242; LEFT: 184px; POSITION: absolute; TOP: 470px" id="DHAB5" runat="server"
					Width="88px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 241; LEFT: 136px; POSITION: absolute; TOP: 470px" id="DHAB4" runat="server"
					Width="40px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True" Visible="False"></asp:textbox><asp:textbox style="Z-INDEX: 240; LEFT: 96px; POSITION: absolute; TOP: 470px" id="DHAB3" runat="server"
					Width="80px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 239; LEFT: 56px; POSITION: absolute; TOP: 470px" id="DHAB2" runat="server"
					Width="40px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 238; LEFT: 16px; POSITION: absolute; TOP: 470px" id="DHAB1" runat="server"
					Width="40px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 237; LEFT: 568px; POSITION: absolute; TOP: 452px" id="DHAA9" runat="server"
					Width="110px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 236; LEFT: 472px; POSITION: absolute; TOP: 452px" id="DHAA8" runat="server"
					Width="88px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 235; LEFT: 376px; POSITION: absolute; TOP: 452px" id="DHAA7" runat="server"
					Width="88px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 234; LEFT: 280px; POSITION: absolute; TOP: 452px" id="DHAA6" runat="server"
					Width="88px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 233; LEFT: 184px; POSITION: absolute; TOP: 452px" id="DHAA5" runat="server"
					Width="88px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 231; LEFT: 136px; POSITION: absolute; TOP: 452px" id="DHAA4" runat="server"
					Width="40px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True" Visible="False"></asp:textbox><asp:textbox style="Z-INDEX: 229; LEFT: 96px; POSITION: absolute; TOP: 452px" id="DHAA3" runat="server"
					Width="80px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 226; LEFT: 56px; POSITION: absolute; TOP: 452px" id="DHAA2" runat="server"
					Width="40px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 225; LEFT: 16px; POSITION: absolute; TOP: 452px" id="DHAA1" runat="server"
					Width="40px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 223; LEFT: 104px; POSITION: absolute; TOP: 534px" id="DHADesc" runat="server"
					Width="464px" Height="28px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" ReadOnly="True" TextMode="MultiLine">DHDesc</asp:textbox><asp:textbox style="Z-INDEX: 221; LEFT: 136px; POSITION: absolute; TOP: 1022px" id="DFQAD3" runat="server"
					Width="620px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:dropdownlist style="Z-INDEX: 218; LEFT: 80px; POSITION: absolute; TOP: 1022px" id="DFQAD2" runat="server"
					Width="56px" Height="18px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 216; LEFT: 16px; POSITION: absolute; TOP: 1022px" id="DFQAD1" runat="server"
					Width="64px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 214; LEFT: 16px; POSITION: absolute; TOP: 1004px" id="DFQAC1" runat="server"
					Width="64px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:dropdownlist style="Z-INDEX: 213; LEFT: 80px; POSITION: absolute; TOP: 1004px" id="DFQAC2" runat="server"
					Width="56px" Height="18px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 210; LEFT: 136px; POSITION: absolute; TOP: 1004px" id="DFQAC3" runat="server"
					Width="620px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 209; LEFT: 136px; POSITION: absolute; TOP: 986px" id="DFQAB3" runat="server"
					Width="620px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:dropdownlist style="Z-INDEX: 206; LEFT: 80px; POSITION: absolute; TOP: 986px" id="DFQAB2" runat="server"
					Width="56px" Height="18px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 204; LEFT: 16px; POSITION: absolute; TOP: 986px" id="DFQAB1" runat="server"
					Width="64px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:hyperlink style="Z-INDEX: 200; LEFT: 312px; POSITION: absolute; TOP: 296px" id="LMapNo" runat="server"
					Width="112px" Height="8px" NavigateUrl="BoardEdit.aspx" Target="_blank">圖號文件</asp:hyperlink><asp:textbox style="Z-INDEX: 197; LEFT: 136px; POSITION: absolute; TOP: 968px" id="DFQAA3" runat="server"
					Width="620px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:dropdownlist style="Z-INDEX: 194; LEFT: 80px; POSITION: absolute; TOP: 968px" id="DFQAA2" runat="server"
					Width="56px" Height="18px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 190; LEFT: 16px; POSITION: absolute; TOP: 968px" id="DFQAA1" runat="server"
					Width="64px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 180; LEFT: 504px; POSITION: absolute; TOP: 784px" id="DPriceJ3"
					runat="server" Width="64px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:dropdownlist style="Z-INDEX: 179; LEFT: 464px; POSITION: absolute; TOP: 784px" id="DPriceJ2"
					runat="server" Width="40px" Height="12px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 178; LEFT: 296px; POSITION: absolute; TOP: 784px" id="DPriceJ1"
					runat="server" Width="166px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"
					ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 177; LEFT: 224px; POSITION: absolute; TOP: 784px" id="DPriceI3"
					runat="server" Width="64px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:dropdownlist style="Z-INDEX: 176; LEFT: 184px; POSITION: absolute; TOP: 784px" id="DPriceI2"
					runat="server" Width="40px" Height="12px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 175; LEFT: 16px; POSITION: absolute; TOP: 784px" id="DPriceI1" runat="server"
					Width="166px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 174; LEFT: 504px; POSITION: absolute; TOP: 766px" id="DPriceH3"
					runat="server" Width="64px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:dropdownlist style="Z-INDEX: 173; LEFT: 464px; POSITION: absolute; TOP: 766px" id="DPriceH2"
					runat="server" Width="40px" Height="12px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 172; LEFT: 296px; POSITION: absolute; TOP: 766px" id="DPriceH1"
					runat="server" Width="166px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"
					ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 171; LEFT: 224px; POSITION: absolute; TOP: 766px" id="DPriceG3"
					runat="server" Width="64px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:dropdownlist style="Z-INDEX: 170; LEFT: 184px; POSITION: absolute; TOP: 766px" id="DPriceG2"
					runat="server" Width="40px" Height="12px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 169; LEFT: 16px; POSITION: absolute; TOP: 766px" id="DPriceG1" runat="server"
					Width="166px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 167; LEFT: 504px; POSITION: absolute; TOP: 748px" id="DPriceF3"
					runat="server" Width="64px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:dropdownlist style="Z-INDEX: 166; LEFT: 464px; POSITION: absolute; TOP: 748px" id="DPriceF2"
					runat="server" Width="40px" Height="12px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 165; LEFT: 296px; POSITION: absolute; TOP: 748px" id="DPriceF1"
					runat="server" Width="166px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"
					ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 164; LEFT: 224px; POSITION: absolute; TOP: 748px" id="DPriceE3"
					runat="server" Width="64px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:dropdownlist style="Z-INDEX: 163; LEFT: 184px; POSITION: absolute; TOP: 748px" id="DPriceE2"
					runat="server" Width="40px" Height="12px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 161; LEFT: 16px; POSITION: absolute; TOP: 748px" id="DPriceE1" runat="server"
					Width="166px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 159; LEFT: 504px; POSITION: absolute; TOP: 730px" id="DPriceD3"
					runat="server" Width="64px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:dropdownlist style="Z-INDEX: 157; LEFT: 464px; POSITION: absolute; TOP: 730px" id="DPriceD2"
					runat="server" Width="40px" Height="12px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 155; LEFT: 296px; POSITION: absolute; TOP: 730px" id="DPriceD1"
					runat="server" Width="166px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"
					ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 154; LEFT: 224px; POSITION: absolute; TOP: 730px" id="DPriceC3"
					runat="server" Width="64px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:dropdownlist style="Z-INDEX: 152; LEFT: 184px; POSITION: absolute; TOP: 730px" id="DPriceC2"
					runat="server" Width="40px" Height="12px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 150; LEFT: 16px; POSITION: absolute; TOP: 730px" id="DPriceC1" runat="server"
					Width="166px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 147; LEFT: 504px; POSITION: absolute; TOP: 712px" id="DPriceB3"
					runat="server" Width="64px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:dropdownlist style="Z-INDEX: 145; LEFT: 464px; POSITION: absolute; TOP: 712px" id="DPriceB2"
					runat="server" Width="40px" Height="12px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 143; LEFT: 296px; POSITION: absolute; TOP: 712px" id="DPriceB1"
					runat="server" Width="166px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"
					ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 142; LEFT: 704px; POSITION: absolute; TOP: 630px" id="DSampleC3"
					runat="server" Width="52px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 141; LEFT: 648px; POSITION: absolute; TOP: 630px" id="DSampleC2"
					runat="server" Width="52px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 140; LEFT: 440px; POSITION: absolute; TOP: 630px" id="DSampleC1"
					runat="server" Width="204px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 139; LEFT: 704px; POSITION: absolute; TOP: 612px" id="DSampleB3"
					runat="server" Width="52px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 138; LEFT: 648px; POSITION: absolute; TOP: 612px" id="DSampleB2"
					runat="server" Width="52px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 137; LEFT: 440px; POSITION: absolute; TOP: 612px" id="DSampleB1"
					runat="server" Width="204px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 136; LEFT: 16px; POSITION: absolute; TOP: 712px" id="DPriceA1" runat="server"
					Width="166px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 135; LEFT: 440px; POSITION: absolute; TOP: 594px" id="DSampleA1"
					runat="server" Width="204px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:hyperlink style="Z-INDEX: 134; LEFT: 112px; POSITION: absolute; TOP: 1144px" id="LContactFile"
					runat="server" Width="50px" Height="8px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">切結書</asp:hyperlink><asp:hyperlink style="Z-INDEX: 133; LEFT: 112px; POSITION: absolute; TOP: 1080px" id="LRefFile"
					runat="server" Width="88px" Height="8px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">生產條件表</asp:hyperlink><asp:hyperlink style="Z-INDEX: 132; LEFT: 488px; POSITION: absolute; TOP: 1080px" id="LQAFile"
					runat="server" Width="64px" Height="8px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">測試報告</asp:hyperlink><asp:textbox style="Z-INDEX: 130; LEFT: 224px; POSITION: absolute; TOP: 712px" id="DPriceA3"
					runat="server" Width="64px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"></asp:textbox><asp:dropdownlist style="Z-INDEX: 128; LEFT: 184px; POSITION: absolute; TOP: 712px" id="DPriceA2"
					runat="server" Width="40px" Height="12px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 102; LEFT: 648px; POSITION: absolute; TOP: 594px" id="DSampleA2"
					runat="server" Width="52px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"
					ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 103; LEFT: 704px; POSITION: absolute; TOP: 594px" id="DSampleA3"
					runat="server" Width="52px" Height="18px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox style="Z-INDEX: 126; LEFT: 520px; POSITION: absolute; TOP: 190px" id="DDate" runat="server"
					Width="118px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" ReadOnly="True">DDate</asp:textbox><asp:textbox style="Z-INDEX: 124; LEFT: 224px; POSITION: absolute; TOP: 120px" id="DSliderCode"
					runat="server" Width="365px" Height="56px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="12pt" ReadOnly="True" TextMode="MultiLine">SliderCode</asp:textbox><asp:hyperlink style="Z-INDEX: 123; LEFT: 488px; POSITION: absolute; TOP: 1144px" id="LSampleFile"
					runat="server" Width="50px" Height="8px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">樣品圖</asp:hyperlink><asp:dropdownlist style="Z-INDEX: 122; LEFT: 520px; POSITION: absolute; TOP: 224px" id="DPerson" runat="server"
					Width="96px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 121; LEFT: 312px; POSITION: absolute; TOP: 224px" id="DDivision"
					runat="server" Width="112px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 120; LEFT: 664px; POSITION: absolute; TOP: 395px" id="DMoldPoint"
					runat="server" Width="45px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" ReadOnly="True">MoldPoint</asp:textbox><asp:textbox style="Z-INDEX: 119; LEFT: 560px; POSITION: absolute; TOP: 395px" id="DMoldQty"
					runat="server" Width="45px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" ReadOnly="True">DMoldQty</asp:textbox><asp:textbox style="Z-INDEX: 118; LEFT: 400px; POSITION: absolute; TOP: 395px" id="DSuppiler"
					runat="server" Width="146px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" ReadOnly="True">Suppiler</asp:textbox><asp:textbox style="Z-INDEX: 117; LEFT: 56px; POSITION: absolute; TOP: 568px" id="DDevReason"
					runat="server" Width="369px" Height="48px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" ReadOnly="True" MaxLength="240" TextMode="MultiLine">DevReason</asp:textbox><asp:hyperlink style="Z-INDEX: 116; LEFT: 360px; POSITION: absolute; TOP: 1112px" id="LAuthorizeFile"
					runat="server" Width="112px" Height="8px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">授權書</asp:hyperlink><asp:hyperlink style="Z-INDEX: 115; LEFT: 112px; POSITION: absolute; TOP: 1112px" id="LConfirmFile"
					runat="server" Width="50px" Height="8px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">確認書</asp:hyperlink><asp:textbox style="Z-INDEX: 114; LEFT: 312px; POSITION: absolute; TOP: 326px" id="DBuyer" runat="server"
					Width="152px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" ReadOnly="True">DBuyer</asp:textbox><asp:textbox style="Z-INDEX: 113; LEFT: 648px; POSITION: absolute; TOP: 326px" id="DSellVendor"
					runat="server" Width="111px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" ReadOnly="True">DSellVendor</asp:textbox><asp:dropdownlist style="Z-INDEX: 112; LEFT: 104px; POSITION: absolute; TOP: 395px" id="DMaterial"
					runat="server" Width="192px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 111; LEFT: 400px; POSITION: absolute; TOP: 360px" id="DManufPlace"
					runat="server" Width="80px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 110; LEFT: 200px; POSITION: absolute; TOP: 360px" id="DSliderType2"
					runat="server" Width="94px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 109; LEFT: 104px; POSITION: absolute; TOP: 360px" id="DSliderType1"
					runat="server" Width="94px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 108; LEFT: 564px; POSITION: absolute; TOP: 360px" id="DAssembler"
					runat="server" Width="76px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" ReadOnly="True">DAssembler</asp:textbox><asp:dropdownlist style="Z-INDEX: 107; LEFT: 648px; POSITION: absolute; TOP: 259px" id="DLevel" runat="server"
					Width="110px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:textbox style="Z-INDEX: 106; LEFT: 312px; POSITION: absolute; TOP: 259px" id="DSpec" runat="server"
					Width="234px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True">DSpec</asp:textbox><asp:textbox style="Z-INDEX: 105; LEFT: 604px; POSITION: absolute; TOP: 156px" id="DSliderGRCode"
					runat="server" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" Height="20px" Width="154px" ReadOnly="True">DSliderGRCode</asp:textbox>
				<asp:textbox id="DNo" style="Z-INDEX: 104; LEFT: 312px; POSITION: absolute; TOP: 190px" runat="server"
					Width="112px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True">DNo</asp:textbox>
				<asp:image id="DManuaInSheet2" style="Z-INDEX: 101; LEFT: 8px; POSITION: absolute; TOP: 680px"
					runat="server" ImageUrl="Images\ManufOutSheet_006_c.JPG" Width="761px" Height="493px"></asp:image>
				<asp:image id="LMapFile" style="Z-INDEX: 202; LEFT: 16px; POSITION: absolute; TOP: 120px" runat="server"
					ImageUrl="http://10.245.0.178//WorkFlow/Document/000003/1-Map2-2006124111750-2.jpg" Width="200px"
					Height="230px" BorderStyle="Groove"></asp:image>
				<asp:Panel id="Panel1" style="Z-INDEX: 271; LEFT: 560px; POSITION: absolute; TOP: 8px" runat="server"
					BackColor="White" Width="210px" Height="64px"></asp:Panel></FONT></form>
		</FONT></FORM>
	</body>
</HTML>
