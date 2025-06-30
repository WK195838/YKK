<%@ Page Language="vb" AutoEventWireup="false" Codebehind="ManufOutModSheet_02.aspx.vb" Inherits="SPD.ManufOutModSheet_02"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>外注委託書(修改版)</title>
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
					while (val == 0) {
						setTimeout("SendToSlider()",200);
						if ((wPop.document.SliderForm.DSlider2.value!="") || (wPop.document.SliderForm.DContent.value!=""))  {
							val=1;
						}
 					} 
				}
			}
		    function SendToSlider() {
				wPop.document.SliderForm.DSlider2.value=document.Form1.DSliderGRCode.value;
				wPop.document.SliderForm.DContent.value=document.Form1.DSliderCode.value;
			}
			//--Spec------------------------------------
			function SpecPicker(strField) {
			    val=0;
				wPop=window.open('SpecPicker.aspx?field=' + strField,'SpecPopup','width=330,height=250,resizable=yes');
				if (document.Form1.DSpec.value != "") {
					while (val == 0) {
						setTimeout("SendToSpec()",200);
						if (wPop.document.SpecForm.DSpec.value!="")  {
							val=1;
						}
 					} 
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
				if (F=="SAVE") answer = confirm("是否執行[儲存資料]作業嗎？ 請確認....");
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
		<form id="Form1" method="post" encType="multipart/form-data" runat="server">
			<FONT face="新細明體">
				<asp:textbox id="DCustomerCode" style="Z-INDEX: 319; LEFT: 464px; POSITION: absolute; TOP: 326px"
					runat="server" Width="88px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow">DCustomerCode</asp:textbox>
				<asp:dropdownlist id="DQAD15" style="Z-INDEX: 331; LEFT: 712px; POSITION: absolute; TOP: 888px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Height="22px" Width="48px" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAC15" style="Z-INDEX: 315; LEFT: 712px; POSITION: absolute; TOP: 872px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Height="22px" Width="48px" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAB15" style="Z-INDEX: 310; LEFT: 712px; POSITION: absolute; TOP: 856px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Height="22px" Width="48px" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:label id="Label38" style="Z-INDEX: 307; LEFT: 712px; POSITION: absolute; TOP: 824px" runat="server"
					Height="10px" Width="32px" Font-Size="8pt">EDX</asp:label>
				<asp:dropdownlist id="DQAA15" style="Z-INDEX: 305; LEFT: 712px; POSITION: absolute; TOP: 840px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Height="22px" Width="48px" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAD13" style="Z-INDEX: 296; LEFT: 616px; POSITION: absolute; TOP: 888px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Height="346px" Width="48px" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0點">0點</asp:ListItem>
					<asp:ListItem Value="2點">2點</asp:ListItem>
					<asp:ListItem Value="4點">4點</asp:ListItem>
					<asp:ListItem Value="6點">6點</asp:ListItem>
					<asp:ListItem Value="8點">8點</asp:ListItem>
					<asp:ListItem Value="10點">10點</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAC13" style="Z-INDEX: 293; LEFT: 616px; POSITION: absolute; TOP: 872px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Height="346px" Width="48px" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0點">0點</asp:ListItem>
					<asp:ListItem Value="2點">2點</asp:ListItem>
					<asp:ListItem Value="4點">4點</asp:ListItem>
					<asp:ListItem Value="6點">6點</asp:ListItem>
					<asp:ListItem Value="8點">8點</asp:ListItem>
					<asp:ListItem Value="10點">10點</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAB13" style="Z-INDEX: 292; LEFT: 616px; POSITION: absolute; TOP: 856px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Height="346px" Width="48px" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0點">0點</asp:ListItem>
					<asp:ListItem Value="2點">2點</asp:ListItem>
					<asp:ListItem Value="4點">4點</asp:ListItem>
					<asp:ListItem Value="6點">6點</asp:ListItem>
					<asp:ListItem Value="8點">8點</asp:ListItem>
					<asp:ListItem Value="10點">10點</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAD12" style="Z-INDEX: 291; LEFT: 568px; POSITION: absolute; TOP: 888px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Height="346px" Width="48px" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0級">0級</asp:ListItem>
					<asp:ListItem Value="1級">1級</asp:ListItem>
					<asp:ListItem Value="2級">2級</asp:ListItem>
					<asp:ListItem Value="3級">3級</asp:ListItem>
					<asp:ListItem Value="4級">4級</asp:ListItem>
					<asp:ListItem Value="5級">5級</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAC12" style="Z-INDEX: 290; LEFT: 568px; POSITION: absolute; TOP: 872px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Height="346px" Width="48px" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0級">0級</asp:ListItem>
					<asp:ListItem Value="1級">1級</asp:ListItem>
					<asp:ListItem Value="2級">2級</asp:ListItem>
					<asp:ListItem Value="3級">3級</asp:ListItem>
					<asp:ListItem Value="4級">4級</asp:ListItem>
					<asp:ListItem Value="5級">5級</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAB12" style="Z-INDEX: 289; LEFT: 568px; POSITION: absolute; TOP: 856px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Height="346px" Width="48px" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0級">0級</asp:ListItem>
					<asp:ListItem Value="1級">1級</asp:ListItem>
					<asp:ListItem Value="2級">2級</asp:ListItem>
					<asp:ListItem Value="3級">3級</asp:ListItem>
					<asp:ListItem Value="4級">4級</asp:ListItem>
					<asp:ListItem Value="5級">5級</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAD11" style="Z-INDEX: 287; LEFT: 520px; POSITION: absolute; TOP: 888px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Height="346px" Width="48px" Font-Size="7pt">
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
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAC11" style="Z-INDEX: 285; LEFT: 520px; POSITION: absolute; TOP: 872px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Height="346px" Width="48px" Font-Size="7pt">
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
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAB11" style="Z-INDEX: 283; LEFT: 520px; POSITION: absolute; TOP: 856px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Height="346px" Width="48px" Font-Size="7pt">
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
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAA13" style="Z-INDEX: 282; LEFT: 616px; POSITION: absolute; TOP: 840px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Height="346px" Width="48px" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0點">0點</asp:ListItem>
					<asp:ListItem Value="2點">2點</asp:ListItem>
					<asp:ListItem Value="4點">4點</asp:ListItem>
					<asp:ListItem Value="6點">6點</asp:ListItem>
					<asp:ListItem Value="8點">8點</asp:ListItem>
					<asp:ListItem Value="10點">10點</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAA12" style="Z-INDEX: 281; LEFT: 568px; POSITION: absolute; TOP: 840px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Height="346px" Width="48px" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0級">0級</asp:ListItem>
					<asp:ListItem Value="1級">1級</asp:ListItem>
					<asp:ListItem Value="2級">2級</asp:ListItem>
					<asp:ListItem Value="3級">3級</asp:ListItem>
					<asp:ListItem Value="4級">4級</asp:ListItem>
					<asp:ListItem Value="5級">5級</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAA11" style="Z-INDEX: 279; LEFT: 520px; POSITION: absolute; TOP: 840px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Height="346px" Width="48px" Font-Size="7pt">
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
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAD14" style="Z-INDEX: 222; LEFT: 664px; POSITION: absolute; TOP: 888px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Height="22px" Width="48px" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAD10" style="Z-INDEX: 219; LEFT: 472px; POSITION: absolute; TOP: 888px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Height="18px" Width="48px" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAD9" style="Z-INDEX: 215; LEFT: 424px; POSITION: absolute; TOP: 888px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Height="18px" Width="48px" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAD8" style="Z-INDEX: 214; LEFT: 376px; POSITION: absolute; TOP: 888px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Height="18px" Width="48px" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="NCA">NCA</asp:ListItem>
					<asp:ListItem Value="NCB">NCB</asp:ListItem>
					<asp:ListItem Value="NG">NG</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAD7" style="Z-INDEX: 212; LEFT: 328px; POSITION: absolute; TOP: 888px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Height="18px" Width="48px" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAD6" style="Z-INDEX: 211; LEFT: 280px; POSITION: absolute; TOP: 888px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Height="18px" Width="48px" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:textbox id="DQAD5" style="Z-INDEX: 210; LEFT: 232px; POSITION: absolute; TOP: 888px" runat="server"
					BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" Height="18px" Width="47px" Font-Size="7pt"></asp:textbox>
				<asp:textbox id="DQAD4" style="Z-INDEX: 209; LEFT: 204px; POSITION: absolute; TOP: 888px" runat="server"
					BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" Height="18px" Width="28px" Font-Size="7pt"></asp:textbox>
				<asp:dropdownlist id="DQAD3" style="Z-INDEX: 208; LEFT: 170px; POSITION: absolute; TOP: 888px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Height="18px" Width="35px" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:textbox id="DQAD2" style="Z-INDEX: 207; LEFT: 72px; POSITION: absolute; TOP: 888px" runat="server"
					BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" Height="18px" Width="98px" Font-Size="7pt"></asp:textbox>
				<asp:textbox id="DQAD1" style="Z-INDEX: 206; LEFT: 16px; POSITION: absolute; TOP: 888px" runat="server"
					BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" Height="18px" Width="56px" Font-Size="7pt"></asp:textbox>
				<asp:dropdownlist id="DQAC14" style="Z-INDEX: 205; LEFT: 664px; POSITION: absolute; TOP: 872px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Height="22px" Width="48px" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAC10" style="Z-INDEX: 204; LEFT: 472px; POSITION: absolute; TOP: 872px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Height="18px" Width="48px" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAC9" style="Z-INDEX: 203; LEFT: 424px; POSITION: absolute; TOP: 872px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Height="18px" Width="48px" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAC8" style="Z-INDEX: 202; LEFT: 376px; POSITION: absolute; TOP: 872px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Height="36px" Width="48px" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="NCA">NCA</asp:ListItem>
					<asp:ListItem Value="NCB">NCB</asp:ListItem>
					<asp:ListItem Value="NG">NG</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAC7" style="Z-INDEX: 201; LEFT: 328px; POSITION: absolute; TOP: 872px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Height="18px" Width="48px" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAC6" style="Z-INDEX: 200; LEFT: 280px; POSITION: absolute; TOP: 872px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Height="18px" Width="48px" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:textbox id="DQAC5" style="Z-INDEX: 199; LEFT: 232px; POSITION: absolute; TOP: 872px" runat="server"
					BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" Height="18px" Width="47px" Font-Size="7pt"></asp:textbox>
				<asp:textbox id="DQAC4" style="Z-INDEX: 198; LEFT: 204px; POSITION: absolute; TOP: 872px" runat="server"
					BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" Height="18px" Width="28px" Font-Size="7pt"></asp:textbox>
				<asp:dropdownlist id="DQAC3" style="Z-INDEX: 197; LEFT: 170px; POSITION: absolute; TOP: 872px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Height="18px" Width="35px" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:textbox id="DQAC2" style="Z-INDEX: 196; LEFT: 72px; POSITION: absolute; TOP: 872px" runat="server"
					BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" Height="18px" Width="98px" Font-Size="7pt"></asp:textbox>
				<asp:textbox id="DQAC1" style="Z-INDEX: 195; LEFT: 16px; POSITION: absolute; TOP: 872px" runat="server"
					BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" Height="18px" Width="56px" Font-Size="7pt"></asp:textbox>
				<asp:dropdownlist id="DQAB14" style="Z-INDEX: 194; LEFT: 664px; POSITION: absolute; TOP: 856px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Height="22px" Width="48px" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAB10" style="Z-INDEX: 193; LEFT: 472px; POSITION: absolute; TOP: 856px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Height="18px" Width="48px" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAB9" style="Z-INDEX: 192; LEFT: 424px; POSITION: absolute; TOP: 856px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Height="18px" Width="48px" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAB8" style="Z-INDEX: 191; LEFT: 376px; POSITION: absolute; TOP: 856px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Height="36px" Width="48px" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="NCA">NCA</asp:ListItem>
					<asp:ListItem Value="NCB">NCB</asp:ListItem>
					<asp:ListItem Value="NG">NG</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAB7" style="Z-INDEX: 190; LEFT: 328px; POSITION: absolute; TOP: 856px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Height="18px" Width="48px" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAB6" style="Z-INDEX: 189; LEFT: 280px; POSITION: absolute; TOP: 856px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Height="18px" Width="48px" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:textbox id="DQAB5" style="Z-INDEX: 188; LEFT: 232px; POSITION: absolute; TOP: 856px" runat="server"
					BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" Height="18px" Width="47px" Font-Size="7pt"></asp:textbox>
				<asp:textbox id="DQAB4" style="Z-INDEX: 187; LEFT: 204px; POSITION: absolute; TOP: 856px" runat="server"
					BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" Height="18px" Width="28px" Font-Size="7pt"></asp:textbox>
				<asp:dropdownlist id="DQAB3" style="Z-INDEX: 186; LEFT: 170px; POSITION: absolute; TOP: 856px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Height="18px" Width="35px" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:textbox id="DQAB2" style="Z-INDEX: 184; LEFT: 72px; POSITION: absolute; TOP: 856px" runat="server"
					BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" Height="18px" Width="98px" Font-Size="7pt"></asp:textbox>
				<asp:textbox id="DQAB1" style="Z-INDEX: 179; LEFT: 16px; POSITION: absolute; TOP: 856px" runat="server"
					BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" Height="18px" Width="56px" Font-Size="7pt"></asp:textbox>
				<asp:dropdownlist id="DQAA14" style="Z-INDEX: 160; LEFT: 664px; POSITION: absolute; TOP: 840px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Height="22px" Width="48px" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAA10" style="Z-INDEX: 155; LEFT: 472px; POSITION: absolute; TOP: 840px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Height="346px" Width="48px" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAA9" style="Z-INDEX: 152; LEFT: 424px; POSITION: absolute; TOP: 840px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Height="18px" Width="48px" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAA8" style="Z-INDEX: 151; LEFT: 376px; POSITION: absolute; TOP: 840px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Height="36px" Width="48px" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="NCA">NCA</asp:ListItem>
					<asp:ListItem Value="NCB">NCB</asp:ListItem>
					<asp:ListItem Value="NG">NG</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAA7" style="Z-INDEX: 146; LEFT: 328px; POSITION: absolute; TOP: 840px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Height="162px" Width="48px" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAA6" style="Z-INDEX: 144; LEFT: 280px; POSITION: absolute; TOP: 840px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Height="18px" Width="48px" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:textbox id="DQAA5" style="Z-INDEX: 143; LEFT: 232px; POSITION: absolute; TOP: 840px" runat="server"
					BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" Height="18px" Width="47px" Font-Size="7pt"></asp:textbox>
				<asp:textbox id="DQAA4" style="Z-INDEX: 141; LEFT: 204px; POSITION: absolute; TOP: 840px" runat="server"
					BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" Height="18px" Width="28px" Font-Size="7pt"></asp:textbox>
				<asp:dropdownlist id="DQAA3" style="Z-INDEX: 139; LEFT: 170px; POSITION: absolute; TOP: 840px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Height="18px" Width="35px" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:textbox id="DQAA1" style="Z-INDEX: 137; LEFT: 16px; POSITION: absolute; TOP: 840px" runat="server"
					BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" Height="18px" Width="56px" Font-Size="7pt"></asp:textbox>
				<asp:textbox id="DQAA2" style="Z-INDEX: 165; LEFT: 72px; POSITION: absolute; TOP: 840px" runat="server"
					BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" Height="18px" Width="98px" Font-Size="7pt"></asp:textbox>
				<asp:label id="Label1" style="Z-INDEX: 278; LEFT: 48px; POSITION: absolute; TOP: 824px" runat="server"
					Height="10px" Width="24px" Font-Size="8pt">日期</asp:label>
				<asp:label id="Label13" style="Z-INDEX: 276; LEFT: 664px; POSITION: absolute; TOP: 824px" runat="server"
					Height="10px" Width="32px" Font-Size="8pt">CPSC</asp:label>
				<asp:label id="Label35" style="Z-INDEX: 300; LEFT: 568px; POSITION: absolute; TOP: 824px" runat="server"
					Height="10px" Width="48px" Font-Size="8pt">一次密著</asp:label>
				<asp:label id="Label12" style="Z-INDEX: 274; LEFT: 616px; POSITION: absolute; TOP: 824px" runat="server"
					Height="10px" Width="48px" Font-Size="8pt">二次密著</asp:label>
				<asp:label id="Label11" style="Z-INDEX: 272; LEFT: 520px; POSITION: absolute; TOP: 824px" runat="server"
					Height="10px" Width="32px" Font-Size="8pt">黃變</asp:label>
				<asp:label id="Label10" style="Z-INDEX: 271; LEFT: 472px; POSITION: absolute; TOP: 824px" runat="server"
					Height="10px" Width="32px" Font-Size="8pt">乾洗</asp:label>
				<asp:label id="Label9" style="Z-INDEX: 269; LEFT: 424px; POSITION: absolute; TOP: 824px" runat="server"
					Height="10px" Width="32px" Font-Size="8pt">水洗</asp:label>
				<asp:label id="Label8" style="Z-INDEX: 267; LEFT: 376px; POSITION: absolute; TOP: 824px" runat="server"
					Height="10px" Width="32px" Font-Size="8pt">檢針</asp:label>
				<asp:label id="Label7" style="Z-INDEX: 265; LEFT: 328px; POSITION: absolute; TOP: 824px" runat="server"
					Height="10px" Width="32px" Font-Size="8pt">扭力</asp:label>
				<asp:label id="Label6" style="Z-INDEX: 263; LEFT: 280px; POSITION: absolute; TOP: 824px" runat="server"
					Height="10px" Width="32px" Font-Size="8pt">強度</asp:label>
				<asp:label id="Label5" style="Z-INDEX: 261; LEFT: 240px; POSITION: absolute; TOP: 824px" runat="server"
					Height="10px" Width="40px" Font-Size="8pt">原單位</asp:label>
				<asp:label id="Label4" style="Z-INDEX: 258; LEFT: 208px; POSITION: absolute; TOP: 824px" runat="server"
					Height="10px" Width="48px" Font-Size="8pt">表面</asp:label>
				<asp:label id="Label3" style="Z-INDEX: 253; LEFT: 168px; POSITION: absolute; TOP: 824px" runat="server"
					Height="10px" Width="32px" Font-Size="8pt">口厚</asp:label>
				<asp:label id="Label2" style="Z-INDEX: 249; LEFT: 80px; POSITION: absolute; TOP: 824px" runat="server"
					Height="10px" Width="48px" Font-Size="8pt">拉頭種類</asp:label>
				<asp:hyperlink id="Hyperlink1" style="Z-INDEX: 326; LEFT: 16px; POSITION: absolute; TOP: 916px"
					runat="server" Height="8px" Width="64px" Target="_blank" NavigateUrl="BoardEdit.aspx" Font-Size="12pt">其他附件</asp:hyperlink><asp:dropdownlist id="DLogo" style="Z-INDEX: 318; LEFT: 686px; POSITION: absolute; TOP: 224px" runat="server"
					Width="78px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist>
				<asp:textbox id="DPullerPrice1" style="Z-INDEX: 129; LEFT: 688px; POSITION: absolute; TOP: 786px"
					runat="server" Width="72px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">PullerPrice1</asp:textbox>
				<asp:hyperlink id="LForCastFile" style="Z-INDEX: 330; LEFT: 616px; POSITION: absolute; TOP: 1112px"
					runat="server" Height="8px" Width="112px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">報價單</asp:hyperlink>
				<asp:textbox id="DPullerPrice" style="Z-INDEX: 131; LEFT: 688px; POSITION: absolute; TOP: 752px"
					runat="server" Width="72px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">PullerPrice</asp:textbox>
				<asp:textbox id="DPurMold" style="Z-INDEX: 127; LEFT: 688px; POSITION: absolute; TOP: 718px"
					runat="server" Width="72px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">PurMold</asp:textbox>
				<asp:textbox id="DArMoldFee" style="Z-INDEX: 126; LEFT: 688px; POSITION: absolute; TOP: 686px"
					runat="server" Width="72px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">ArMoldFee</asp:textbox>
				<asp:textbox id="DManufFlow" style="Z-INDEX: 329; LEFT: 56px; POSITION: absolute; TOP: 624px"
					runat="server" Height="50px" Width="369px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow"
					TextMode="MultiLine" MaxLength="240">DManufFlow</asp:textbox>
				<asp:textbox id="DDevReason" style="Z-INDEX: 124; LEFT: 56px; POSITION: absolute; TOP: 568px"
					runat="server" Width="369px" Height="48px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove"
					TextMode="MultiLine" MaxLength="240">DevReason</asp:textbox>
				<asp:hyperlink id="LMapNo1" style="Z-INDEX: 328; LEFT: 464px; POSITION: absolute; TOP: 296px" runat="server"
					Height="8px" Width="112px" Target="_blank" NavigateUrl="BoardEdit.aspx">圖號文件</asp:hyperlink>
				<asp:dropdownlist id="DCpsc" style="Z-INDEX: 327; LEFT: 708px; POSITION: absolute; TOP: 190px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="20px" Width="56px">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist>
				<asp:Label id="Label25" style="Z-INDEX: 313; LEFT: 704px; POSITION: absolute; TOP: 580px" runat="server"
					Width="24px" Height="10px" Font-Size="8pt">數量</asp:Label>
				<asp:Label id="Label32" style="Z-INDEX: 325; LEFT: 136px; POSITION: absolute; TOP: 440px" runat="server"
					Height="10px" Width="32px" Font-Size="8pt">品名-4</asp:Label>
				<asp:Label id="Label31" style="Z-INDEX: 324; LEFT: 96px; POSITION: absolute; TOP: 440px" runat="server"
					Height="10px" Width="32px" Font-Size="8pt">品名-3</asp:Label>
				<asp:Label id="Label30" style="Z-INDEX: 323; LEFT: 56px; POSITION: absolute; TOP: 440px" runat="server"
					Height="10px" Width="32px" Font-Size="8pt">品名-2</asp:Label>
				<asp:Label id="Label34" style="Z-INDEX: 322; LEFT: 568px; POSITION: absolute; TOP: 440px" runat="server"
					Height="10px" Width="24px" Font-Size="8pt">備註</asp:Label>
				<asp:Label id="Label33" style="Z-INDEX: 321; LEFT: 472px; POSITION: absolute; TOP: 440px" runat="server"
					Height="10px" Width="64px" Font-Size="8pt">預定完成日</asp:Label>
				<asp:Label id="Label29" style="Z-INDEX: 320; LEFT: 376px; POSITION: absolute; TOP: 440px" runat="server"
					Height="10px" Width="32px" Font-Size="8pt">OR-NO</asp:Label>
				<asp:Label id="Label28" style="Z-INDEX: 317; LEFT: 280px; POSITION: absolute; TOP: 440px" runat="server"
					Height="10px" Width="56px" Font-Size="8pt">基準日程</asp:Label>
				<asp:Label id="Label27" style="Z-INDEX: 316; LEFT: 184px; POSITION: absolute; TOP: 440px" runat="server"
					Height="10px" Width="64px" Font-Size="8pt">一日產能</asp:Label>
				<asp:Label id="Label26" style="Z-INDEX: 314; LEFT: 16px; POSITION: absolute; TOP: 440px" runat="server"
					Height="10px" Width="32px" Font-Size="8pt">品名-1</asp:Label>
				<asp:Label id="Label24" style="Z-INDEX: 312; LEFT: 648px; POSITION: absolute; TOP: 580px" runat="server"
					Width="24px" Height="10px" Font-Size="8pt">顏色</asp:Label>
				<asp:Label id="Label23" style="Z-INDEX: 311; LEFT: 472px; POSITION: absolute; TOP: 580px" runat="server"
					Width="48px" Height="10px" Font-Size="8pt">拉頭種類</asp:Label>
				<asp:Label id="Label22" style="Z-INDEX: 309; LEFT: 504px; POSITION: absolute; TOP: 700px" runat="server"
					Width="24px" Height="10px" Font-Size="8pt">單價</asp:Label>
				<asp:Label id="Label21" style="Z-INDEX: 308; LEFT: 464px; POSITION: absolute; TOP: 700px" runat="server"
					Width="24px" Height="10px" Font-Size="8pt">幣別</asp:Label>
				<asp:Label id="Label20" style="Z-INDEX: 304; LEFT: 296px; POSITION: absolute; TOP: 700px" runat="server"
					Width="48px" Height="10px" Font-Size="8pt">拉頭種類</asp:Label>
				<asp:Label id="Label19" style="Z-INDEX: 302; LEFT: 224px; POSITION: absolute; TOP: 700px" runat="server"
					Width="24px" Height="10px" Font-Size="8pt">單價</asp:Label>
				<asp:Label id="Label18" style="Z-INDEX: 301; LEFT: 184px; POSITION: absolute; TOP: 700px" runat="server"
					Width="24px" Height="10px" Font-Size="8pt">幣別</asp:Label>
				<asp:Label id="Label17" style="Z-INDEX: 298; LEFT: 56px; POSITION: absolute; TOP: 700px" runat="server"
					Width="48px" Height="10px" Font-Size="8pt">拉頭種類</asp:Label>
				<asp:Label id="Label16" style="Z-INDEX: 297; LEFT: 136px; POSITION: absolute; TOP: 956px" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">備註</asp:Label>
				<asp:Label id="Label15" style="Z-INDEX: 295; LEFT: 80px; POSITION: absolute; TOP: 956px" runat="server"
					Width="48px" Height="10px" Font-Size="8pt">檢測結果</asp:Label>
				<asp:Label id="Label14" style="Z-INDEX: 294; LEFT: 48px; POSITION: absolute; TOP: 956px" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">日期</asp:Label>
				<asp:hyperlink id="LSAttachFile" style="Z-INDEX: 299; LEFT: 440px; POSITION: absolute; TOP: 652px"
					runat="server" Height="8px" Width="64px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">其他附件</asp:hyperlink>
				<asp:hyperlink id="LQAttachFile2" style="Z-INDEX: 306; LEFT: 16px; POSITION: absolute; TOP: 1044px"
					runat="server" Height="8px" Width="136px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">品質分析依賴書</asp:hyperlink>
				<asp:hyperlink id="LQAttachFile1" style="Z-INDEX: 303; LEFT: 16px; POSITION: absolute; TOP: 916px"
					runat="server" Height="8px" Width="64px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">其他附件</asp:hyperlink>
				<asp:image id="DManuaInSheet1" style="Z-INDEX: 100; LEFT: 8px; POSITION: absolute; TOP: 8px"
					runat="server" Height="674px" Width="761px" ImageUrl="Images\ManufOutSheet_003_e.jpg"></asp:image>
				<asp:textbox id="DFormSno" style="Z-INDEX: 288; LEFT: 8px; POSITION: absolute; TOP: 1176px" runat="server"
					Width="97px" Height="20px" BackColor="White" ForeColor="Blue" BorderStyle="None">單號：123</asp:textbox>
				<asp:image id="Image1" style="Z-INDEX: 286; LEFT: 16px; POSITION: absolute; TOP: 1176px" runat="server"
					ImageUrl="Images\ManufInSheet_New_003.jpg" Width="755px" Height="29px"></asp:image><asp:textbox id="DHAD9" style="Z-INDEX: 284; LEFT: 568px; POSITION: absolute; TOP: 506px" runat="server"
					Height="18px" Width="192px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DHAD8" style="Z-INDEX: 280; LEFT: 472px; POSITION: absolute; TOP: 506px" runat="server"
					Height="18px" Width="88px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DHAD7" style="Z-INDEX: 277; LEFT: 376px; POSITION: absolute; TOP: 506px" runat="server"
					Height="18px" Width="88px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DHAD6" style="Z-INDEX: 275; LEFT: 280px; POSITION: absolute; TOP: 506px" runat="server"
					Height="18px" Width="88px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DHAD5" style="Z-INDEX: 273; LEFT: 184px; POSITION: absolute; TOP: 506px" runat="server"
					Height="18px" Width="88px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DHAD4" style="Z-INDEX: 270; LEFT: 136px; POSITION: absolute; TOP: 506px" runat="server"
					Height="18px" Width="40px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DHAD3" style="Z-INDEX: 268; LEFT: 96px; POSITION: absolute; TOP: 506px" runat="server"
					Height="18px" Width="40px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DHAD2" style="Z-INDEX: 266; LEFT: 56px; POSITION: absolute; TOP: 506px" runat="server"
					Height="18px" Width="40px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DHAD1" style="Z-INDEX: 264; LEFT: 16px; POSITION: absolute; TOP: 506px" runat="server"
					Height="18px" Width="40px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DHAC9" style="Z-INDEX: 262; LEFT: 568px; POSITION: absolute; TOP: 488px" runat="server"
					Height="18px" Width="192px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DHAC8" style="Z-INDEX: 260; LEFT: 472px; POSITION: absolute; TOP: 488px" runat="server"
					Height="18px" Width="88px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DHAC7" style="Z-INDEX: 259; LEFT: 376px; POSITION: absolute; TOP: 488px" runat="server"
					Height="18px" Width="88px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DHAC6" style="Z-INDEX: 257; LEFT: 280px; POSITION: absolute; TOP: 488px" runat="server"
					Height="18px" Width="88px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DHAC5" style="Z-INDEX: 256; LEFT: 184px; POSITION: absolute; TOP: 488px" runat="server"
					Height="18px" Width="88px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DHAC4" style="Z-INDEX: 255; LEFT: 136px; POSITION: absolute; TOP: 488px" runat="server"
					Height="18px" Width="40px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DHAC3" style="Z-INDEX: 254; LEFT: 96px; POSITION: absolute; TOP: 488px" runat="server"
					Height="18px" Width="40px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DHAC2" style="Z-INDEX: 252; LEFT: 56px; POSITION: absolute; TOP: 488px" runat="server"
					Height="18px" Width="40px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DHAC1" style="Z-INDEX: 251; LEFT: 16px; POSITION: absolute; TOP: 488px" runat="server"
					Height="18px" Width="40px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DHAB9" style="Z-INDEX: 250; LEFT: 568px; POSITION: absolute; TOP: 470px" runat="server"
					Height="18px" Width="192px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DHAB8" style="Z-INDEX: 248; LEFT: 472px; POSITION: absolute; TOP: 470px" runat="server"
					Height="18px" Width="88px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DHAB7" style="Z-INDEX: 247; LEFT: 376px; POSITION: absolute; TOP: 470px" runat="server"
					Height="18px" Width="88px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DHAB6" style="Z-INDEX: 246; LEFT: 280px; POSITION: absolute; TOP: 470px" runat="server"
					Height="18px" Width="88px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DHAB5" style="Z-INDEX: 245; LEFT: 184px; POSITION: absolute; TOP: 470px" runat="server"
					Height="18px" Width="88px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DHAB4" style="Z-INDEX: 244; LEFT: 136px; POSITION: absolute; TOP: 470px" runat="server"
					Height="18px" Width="40px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DHAB3" style="Z-INDEX: 243; LEFT: 96px; POSITION: absolute; TOP: 470px" runat="server"
					Height="18px" Width="40px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DHAB2" style="Z-INDEX: 242; LEFT: 56px; POSITION: absolute; TOP: 470px" runat="server"
					Height="18px" Width="40px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DHAB1" style="Z-INDEX: 241; LEFT: 16px; POSITION: absolute; TOP: 470px" runat="server"
					Height="18px" Width="40px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DHAA9" style="Z-INDEX: 240; LEFT: 568px; POSITION: absolute; TOP: 452px" runat="server"
					Height="18px" Width="192px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DHAA8" style="Z-INDEX: 239; LEFT: 472px; POSITION: absolute; TOP: 452px" runat="server"
					Height="18px" Width="88px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DHAA7" style="Z-INDEX: 238; LEFT: 376px; POSITION: absolute; TOP: 452px" runat="server"
					Height="18px" Width="88px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DHAA6" style="Z-INDEX: 237; LEFT: 280px; POSITION: absolute; TOP: 452px" runat="server"
					Height="18px" Width="88px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DHAA5" style="Z-INDEX: 236; LEFT: 184px; POSITION: absolute; TOP: 452px" runat="server"
					Height="18px" Width="88px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DHAA4" style="Z-INDEX: 235; LEFT: 136px; POSITION: absolute; TOP: 452px" runat="server"
					Height="18px" Width="40px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DHAA3" style="Z-INDEX: 234; LEFT: 96px; POSITION: absolute; TOP: 452px" runat="server"
					Height="18px" Width="40px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DHAA2" style="Z-INDEX: 233; LEFT: 56px; POSITION: absolute; TOP: 452px" runat="server"
					Height="18px" Width="40px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DHAA1" style="Z-INDEX: 232; LEFT: 16px; POSITION: absolute; TOP: 452px" runat="server"
					Height="18px" Width="40px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DHADesc" style="Z-INDEX: 231; LEFT: 104px; POSITION: absolute; TOP: 534px" runat="server"
					Height="28px" Width="656px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" TextMode="MultiLine" ReadOnly="True">DHDesc</asp:textbox><asp:textbox id="DFQAD3" style="Z-INDEX: 230; LEFT: 136px; POSITION: absolute; TOP: 1022px" runat="server"
					Height="18px" Width="620px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:dropdownlist id="DFQAD2" style="Z-INDEX: 229; LEFT: 80px; POSITION: absolute; TOP: 1022px" runat="server"
					Height="18px" Width="56px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DFQAD1" style="Z-INDEX: 228; LEFT: 16px; POSITION: absolute; TOP: 1022px" runat="server"
					Height="18px" Width="64px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DFQAC1" style="Z-INDEX: 227; LEFT: 16px; POSITION: absolute; TOP: 1004px" runat="server"
					Height="18px" Width="64px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:dropdownlist id="DFQAC2" style="Z-INDEX: 226; LEFT: 80px; POSITION: absolute; TOP: 1004px" runat="server"
					Height="18px" Width="56px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DFQAC3" style="Z-INDEX: 225; LEFT: 136px; POSITION: absolute; TOP: 1004px" runat="server"
					Height="18px" Width="620px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DFQAB3" style="Z-INDEX: 224; LEFT: 136px; POSITION: absolute; TOP: 986px" runat="server"
					Height="18px" Width="620px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:dropdownlist id="DFQAB2" style="Z-INDEX: 223; LEFT: 80px; POSITION: absolute; TOP: 986px" runat="server"
					Height="18px" Width="56px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DFQAB1" style="Z-INDEX: 221; LEFT: 16px; POSITION: absolute; TOP: 986px" runat="server"
					Height="18px" Width="64px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox>
				<asp:hyperlink id="LMapNo" style="Z-INDEX: 218; LEFT: 312px; POSITION: absolute; TOP: 296px" runat="server"
					Height="8px" Width="112px" NavigateUrl="BoardEdit.aspx" Target="_blank">圖號文件</asp:hyperlink><asp:textbox id="DFQAA3" style="Z-INDEX: 217; LEFT: 136px; POSITION: absolute; TOP: 968px" runat="server"
					Height="18px" Width="620px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:dropdownlist id="DFQAA2" style="Z-INDEX: 216; LEFT: 80px; POSITION: absolute; TOP: 968px" runat="server"
					Height="18px" Width="56px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DFQAA1" style="Z-INDEX: 213; LEFT: 16px; POSITION: absolute; TOP: 968px" runat="server"
					Height="18px" Width="64px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DPriceJ3" style="Z-INDEX: 185; LEFT: 504px; POSITION: absolute; TOP: 784px"
					runat="server" Height="18px" Width="64px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:dropdownlist id="DPriceJ2" style="Z-INDEX: 183; LEFT: 464px; POSITION: absolute; TOP: 784px"
					runat="server" Height="12px" Width="40px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceJ1" style="Z-INDEX: 182; LEFT: 296px; POSITION: absolute; TOP: 784px"
					runat="server" Height="18px" Width="166px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"
					ReadOnly="True"></asp:textbox><asp:textbox id="DPriceI3" style="Z-INDEX: 181; LEFT: 224px; POSITION: absolute; TOP: 784px"
					runat="server" Height="18px" Width="64px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:dropdownlist id="DPriceI2" style="Z-INDEX: 180; LEFT: 184px; POSITION: absolute; TOP: 784px"
					runat="server" Height="12px" Width="40px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceI1" style="Z-INDEX: 178; LEFT: 16px; POSITION: absolute; TOP: 784px" runat="server"
					Height="18px" Width="166px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DPriceH3" style="Z-INDEX: 177; LEFT: 504px; POSITION: absolute; TOP: 766px"
					runat="server" Height="18px" Width="64px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:dropdownlist id="DPriceH2" style="Z-INDEX: 176; LEFT: 464px; POSITION: absolute; TOP: 766px"
					runat="server" Height="12px" Width="40px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceH1" style="Z-INDEX: 175; LEFT: 296px; POSITION: absolute; TOP: 766px"
					runat="server" Height="18px" Width="166px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"
					ReadOnly="True"></asp:textbox><asp:textbox id="DPriceG3" style="Z-INDEX: 174; LEFT: 224px; POSITION: absolute; TOP: 766px"
					runat="server" Height="18px" Width="64px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:dropdownlist id="DPriceG2" style="Z-INDEX: 173; LEFT: 184px; POSITION: absolute; TOP: 766px"
					runat="server" Height="12px" Width="40px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceG1" style="Z-INDEX: 172; LEFT: 16px; POSITION: absolute; TOP: 766px" runat="server"
					Height="18px" Width="166px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DPriceF3" style="Z-INDEX: 171; LEFT: 504px; POSITION: absolute; TOP: 748px"
					runat="server" Height="18px" Width="64px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:dropdownlist id="DPriceF2" style="Z-INDEX: 170; LEFT: 464px; POSITION: absolute; TOP: 748px"
					runat="server" Height="12px" Width="40px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceF1" style="Z-INDEX: 169; LEFT: 296px; POSITION: absolute; TOP: 748px"
					runat="server" Height="18px" Width="166px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"
					ReadOnly="True"></asp:textbox><asp:textbox id="DPriceE3" style="Z-INDEX: 168; LEFT: 224px; POSITION: absolute; TOP: 748px"
					runat="server" Height="18px" Width="64px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:dropdownlist id="DPriceE2" style="Z-INDEX: 167; LEFT: 184px; POSITION: absolute; TOP: 748px"
					runat="server" Height="12px" Width="40px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceE1" style="Z-INDEX: 166; LEFT: 16px; POSITION: absolute; TOP: 748px" runat="server"
					Height="18px" Width="166px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DPriceD3" style="Z-INDEX: 164; LEFT: 504px; POSITION: absolute; TOP: 730px"
					runat="server" Height="18px" Width="64px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:dropdownlist id="DPriceD2" style="Z-INDEX: 163; LEFT: 464px; POSITION: absolute; TOP: 730px"
					runat="server" Height="12px" Width="40px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceD1" style="Z-INDEX: 162; LEFT: 296px; POSITION: absolute; TOP: 730px"
					runat="server" Height="18px" Width="166px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"
					ReadOnly="True"></asp:textbox><asp:textbox id="DPriceC3" style="Z-INDEX: 161; LEFT: 224px; POSITION: absolute; TOP: 730px"
					runat="server" Height="18px" Width="64px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:dropdownlist id="DPriceC2" style="Z-INDEX: 159; LEFT: 184px; POSITION: absolute; TOP: 730px"
					runat="server" Height="12px" Width="40px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceC1" style="Z-INDEX: 158; LEFT: 16px; POSITION: absolute; TOP: 730px" runat="server"
					Height="18px" Width="166px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DPriceB3" style="Z-INDEX: 157; LEFT: 504px; POSITION: absolute; TOP: 712px"
					runat="server" Height="18px" Width="64px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:dropdownlist id="DPriceB2" style="Z-INDEX: 156; LEFT: 464px; POSITION: absolute; TOP: 712px"
					runat="server" Height="12px" Width="40px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceB1" style="Z-INDEX: 154; LEFT: 296px; POSITION: absolute; TOP: 712px"
					runat="server" Height="18px" Width="166px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"
					ReadOnly="True"></asp:textbox><asp:textbox id="DSampleC3" style="Z-INDEX: 153; LEFT: 704px; POSITION: absolute; TOP: 630px"
					runat="server" Height="18px" Width="52px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DSampleC2" style="Z-INDEX: 150; LEFT: 648px; POSITION: absolute; TOP: 630px"
					runat="server" Height="18px" Width="52px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DSampleC1" style="Z-INDEX: 149; LEFT: 440px; POSITION: absolute; TOP: 630px"
					runat="server" Height="18px" Width="204px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DSampleB3" style="Z-INDEX: 148; LEFT: 704px; POSITION: absolute; TOP: 612px"
					runat="server" Height="18px" Width="52px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DSampleB2" style="Z-INDEX: 147; LEFT: 648px; POSITION: absolute; TOP: 612px"
					runat="server" Height="18px" Width="52px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DSampleB1" style="Z-INDEX: 145; LEFT: 440px; POSITION: absolute; TOP: 612px"
					runat="server" Height="18px" Width="204px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DPriceA1" style="Z-INDEX: 142; LEFT: 16px; POSITION: absolute; TOP: 712px" runat="server"
					Height="18px" Width="166px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DSampleA1" style="Z-INDEX: 140; LEFT: 440px; POSITION: absolute; TOP: 594px"
					runat="server" Height="18px" Width="204px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:button id="BSliderCode" style="Z-INDEX: 138; LEFT: 568px; POSITION: absolute; TOP: 120px"
					runat="server" Height="20px" Width="20px" Text="....."></asp:button><asp:hyperlink id="LContactFile" style="Z-INDEX: 136; LEFT: 104px; POSITION: absolute; TOP: 1144px"
					runat="server" Height="8px" Width="50px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">切結書</asp:hyperlink>
				<asp:hyperlink id="LRefFile" style="Z-INDEX: 135; LEFT: 104px; POSITION: absolute; TOP: 1076px"
					runat="server" Height="8px" Width="88px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">生產條件表</asp:hyperlink><asp:hyperlink id="LQAFile" style="Z-INDEX: 134; LEFT: 488px; POSITION: absolute; TOP: 1076px"
					runat="server" Height="8px" Width="64px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">測試報告</asp:hyperlink><asp:textbox id="DPriceA3" style="Z-INDEX: 133; LEFT: 224px; POSITION: absolute; TOP: 712px"
					runat="server" Height="18px" Width="64px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:dropdownlist id="DPriceA2" style="Z-INDEX: 132; LEFT: 184px; POSITION: absolute; TOP: 712px"
					runat="server" Height="12px" Width="40px" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DSampleA2" style="Z-INDEX: 102; LEFT: 648px; POSITION: absolute; TOP: 594px"
					runat="server" Height="18px" Width="52px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt"
					ReadOnly="True"></asp:textbox><asp:textbox id="DSampleA3" style="Z-INDEX: 103; LEFT: 704px; POSITION: absolute; TOP: 594px"
					runat="server" Height="18px" Width="52px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:button id="BDate" style="Z-INDEX: 130; LEFT: 616px; POSITION: absolute; TOP: 190px" runat="server"
					Height="20px" Width="20px" Text="....."></asp:button><asp:textbox id="DDate" style="Z-INDEX: 128; LEFT: 520px; POSITION: absolute; TOP: 190px" runat="server"
					Height="20px" Width="96px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" ReadOnly="True">DDate</asp:textbox><asp:textbox id="DSliderCode" style="Z-INDEX: 125; LEFT: 224px; POSITION: absolute; TOP: 120px"
					runat="server" Height="56px" Width="342px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" ReadOnly="True" Font-Size="12pt" TextMode="MultiLine">SliderCode</asp:textbox><asp:hyperlink id="LSampleFile" style="Z-INDEX: 123; LEFT: 488px; POSITION: absolute; TOP: 1144px"
					runat="server" Height="8px" Width="50px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">樣品圖</asp:hyperlink><asp:dropdownlist id="DPerson" style="Z-INDEX: 122; LEFT: 520px; POSITION: absolute; TOP: 224px" runat="server"
					Height="20px" Width="96px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DDivision" style="Z-INDEX: 121; LEFT: 312px; POSITION: absolute; TOP: 224px"
					runat="server" Height="20px" Width="112px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DMoldPoint" style="Z-INDEX: 120; LEFT: 672px; POSITION: absolute; TOP: 395px"
					runat="server" Height="20px" Width="45px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" ReadOnly="True">MoldPoint</asp:textbox><asp:textbox id="DMoldQty" style="Z-INDEX: 119; LEFT: 568px; POSITION: absolute; TOP: 395px"
					runat="server" Height="20px" Width="45px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" ReadOnly="True">DMoldQty</asp:textbox><asp:textbox id="DSuppiler" style="Z-INDEX: 118; LEFT: 400px; POSITION: absolute; TOP: 395px"
					runat="server" Height="20px" Width="146px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" ReadOnly="True">Suppiler</asp:textbox>
				<asp:hyperlink id="LAuthorizeFile" style="Z-INDEX: 117; LEFT: 360px; POSITION: absolute; TOP: 1112px"
					runat="server" Height="8px" Width="112px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">授權書</asp:hyperlink>
				<asp:hyperlink id="LConfirmFile" style="Z-INDEX: 116; LEFT: 104px; POSITION: absolute; TOP: 1112px"
					runat="server" Height="8px" Width="50px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">確認書</asp:hyperlink>
				<asp:textbox id="DBuyer" style="Z-INDEX: 115; LEFT: 312px; POSITION: absolute; TOP: 326px" runat="server"
					Height="20px" Width="152px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" ReadOnly="True">DBuyer</asp:textbox><asp:textbox id="DSellVendor" style="Z-INDEX: 114; LEFT: 648px; POSITION: absolute; TOP: 326px"
					runat="server" Height="20px" Width="111px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" ReadOnly="True">DSellVendor</asp:textbox><asp:dropdownlist id="DMaterial" style="Z-INDEX: 113; LEFT: 104px; POSITION: absolute; TOP: 395px"
					runat="server" Height="20px" Width="192px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DManufPlace" style="Z-INDEX: 112; LEFT: 400px; POSITION: absolute; TOP: 360px"
					runat="server" Height="20px" Width="146px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DSliderType2" style="Z-INDEX: 111; LEFT: 200px; POSITION: absolute; TOP: 360px"
					runat="server" Height="20px" Width="94px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DSliderType1" style="Z-INDEX: 110; LEFT: 104px; POSITION: absolute; TOP: 360px"
					runat="server" Height="20px" Width="94px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DAssembler" style="Z-INDEX: 109; LEFT: 648px; POSITION: absolute; TOP: 360px"
					runat="server" Height="20px" Width="110px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" ReadOnly="True">DAssembler</asp:textbox><asp:dropdownlist id="DLevel" style="Z-INDEX: 108; LEFT: 648px; POSITION: absolute; TOP: 259px" runat="server"
					Height="20px" Width="110px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:button id="BSpec" style="Z-INDEX: 107; LEFT: 528px; POSITION: absolute; TOP: 259px" runat="server"
					Height="20px" Width="20px" Text="....."></asp:button><asp:textbox id="DSpec" style="Z-INDEX: 106; LEFT: 312px; POSITION: absolute; TOP: 259px" runat="server"
					Height="20px" Width="216px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" ReadOnly="True" Font-Size="8pt">DSpec</asp:textbox><asp:textbox id="DSliderGRCode" style="Z-INDEX: 105; LEFT: 604px; POSITION: absolute; TOP: 156px"
					runat="server" Height="20px" Width="154px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" ReadOnly="True">DSliderGRCode</asp:textbox><asp:textbox id="DNo" style="Z-INDEX: 104; LEFT: 312px; POSITION: absolute; TOP: 190px" runat="server"
					Height="20px" Width="112px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" ReadOnly="True">DNo</asp:textbox><asp:image id="DManuaInSheet2" style="Z-INDEX: 101; LEFT: 8px; POSITION: absolute; TOP: 680px"
					runat="server" Height="493px" Width="761px" ImageUrl="Images\ManufOutSheet_006_c.jpg"></asp:image><asp:image id="LMapFile" style="Z-INDEX: 220; LEFT: 16px; POSITION: absolute; TOP: 120px" runat="server"
					Height="230px" Width="200px" ImageUrl="http://10.245.0.178//WorkFlow/Document/000003/1-Map2-2006124111750-2.jpg" BorderStyle="Groove"></asp:image></FONT></form>
	</body>
</HTML>
