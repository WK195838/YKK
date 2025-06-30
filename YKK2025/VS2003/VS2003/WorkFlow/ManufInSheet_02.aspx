<%@ Page Language="vb" AutoEventWireup="false" Codebehind="ManufInSheet_02.aspx.vb" Inherits="SPD.ManufInSheet_02"%>
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
		<form id="Form1" method="post" encType="multipart/form-data" runat="server">
			<FONT face="新細明體">
				<asp:textbox id="DCustomerCode" style="Z-INDEX: 265; LEFT: 488px; POSITION: absolute; TOP: 328px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="104px" BorderStyle="Groove">DCustomerCode</asp:textbox>
				<asp:label style="Z-INDEX: 384; LEFT: 288px; POSITION: absolute; TOP: 902px" id="Label42" runat="server"
					Width="96px" Height="10px">組立判定書</asp:label>
				<asp:hyperlink style="Z-INDEX: 295; LEFT: 400px; POSITION: absolute; TOP: 902px" id="LAssemblerFile"
					runat="server" Width="80px" Height="10px" Target="_blank" NavigateUrl="BoardEdit.aspx" Font-Size="12pt">組立判定書</asp:hyperlink>
				<asp:label style="Z-INDEX: 291; LEFT: 600px; POSITION: absolute; TOP: 902px" id="Label27" runat="server"
					Width="49px" Height="10px">QC-L/T</asp:label>
				<asp:textbox style="Z-INDEX: 290; LEFT: 656px; POSITION: absolute; TOP: 900px" id="DQCLT" runat="server"
					BorderStyle="Groove" Width="81px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt"></asp:textbox>
				<asp:label style="Z-INDEX: 292; LEFT: 744px; POSITION: absolute; TOP: 902px" id="Label28" runat="server"
					Width="17px" Height="10px">分</asp:label>
				<asp:hyperlink style="Z-INDEX: 293; LEFT: 152px; POSITION: absolute; TOP: 32px" id="LYKKGroupCopy"
					runat="server" Width="144px" Height="16px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">姊妹社圖面複製履歷</asp:hyperlink>
				<asp:textbox style="Z-INDEX: 289; LEFT: 648px; POSITION: absolute; TOP: 293px" id="DAssembler1"
					runat="server" BorderStyle="Groove" Width="113px" Height="20px" BackColor="Yellow" ForeColor="Blue"
					ReadOnly="True">DAssembler</asp:textbox>
				<asp:textbox style="Z-INDEX: 288; LEFT: 688px; POSITION: absolute; TOP: 640px" id="DOriginalDep"
					runat="server" BorderStyle="Groove" Width="64px" Height="20px" BackColor="Yellow" ForeColor="Blue"
					ReadOnly="True">OriginalDep</asp:textbox>
				<asp:image style="Z-INDEX: 287; LEFT: 624px; POSITION: absolute; TOP: 8px" id="DMMSSts" runat="server"
					BorderStyle="None" Width="145px" Height="48px" ImageUrl="Images\mmsdelete.jpg"></asp:image>
				<asp:dropdownlist id="DQAD15" style="Z-INDEX: 286; LEFT: 712px; POSITION: absolute; TOP: 750px" runat="server"
					Width="48px" Height="22px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAC15" style="Z-INDEX: 285; LEFT: 712px; POSITION: absolute; TOP: 732px" runat="server"
					Width="48px" Height="22px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAB15" style="Z-INDEX: 284; LEFT: 712px; POSITION: absolute; TOP: 714px" runat="server"
					Width="48px" Height="22px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:label id="Label29" style="Z-INDEX: 283; LEFT: 712px; POSITION: absolute; TOP: 684px" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">EDX</asp:label>
				<asp:dropdownlist id="DQAA15" style="Z-INDEX: 282; LEFT: 712px; POSITION: absolute; TOP: 696px" runat="server"
					Width="48px" Height="22px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAD13" style="Z-INDEX: 280; LEFT: 616px; POSITION: absolute; TOP: 750px" runat="server"
					Width="48px" Height="346px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0級">0級</asp:ListItem>
					<asp:ListItem Value="1級">1級</asp:ListItem>
					<asp:ListItem Value="2級">2級</asp:ListItem>
					<asp:ListItem Value="3級">3級</asp:ListItem>
					<asp:ListItem Value="4級">4級</asp:ListItem>
					<asp:ListItem Value="5級">5級</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAC13" style="Z-INDEX: 279; LEFT: 616px; POSITION: absolute; TOP: 732px" runat="server"
					Width="48px" Height="346px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0級">0級</asp:ListItem>
					<asp:ListItem Value="1級">1級</asp:ListItem>
					<asp:ListItem Value="2級">2級</asp:ListItem>
					<asp:ListItem Value="3級">3級</asp:ListItem>
					<asp:ListItem Value="4級">4級</asp:ListItem>
					<asp:ListItem Value="5級">5級</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAB13" style="Z-INDEX: 278; LEFT: 616px; POSITION: absolute; TOP: 714px" runat="server"
					Width="48px" Height="346px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0級">0級</asp:ListItem>
					<asp:ListItem Value="1級">1級</asp:ListItem>
					<asp:ListItem Value="2級">2級</asp:ListItem>
					<asp:ListItem Value="3級">3級</asp:ListItem>
					<asp:ListItem Value="4級">4級</asp:ListItem>
					<asp:ListItem Value="5級">5級</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAD12" style="Z-INDEX: 277; LEFT: 568px; POSITION: absolute; TOP: 750px" runat="server"
					Width="48px" Height="346px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0級">0級</asp:ListItem>
					<asp:ListItem Value="1級">1級</asp:ListItem>
					<asp:ListItem Value="2級">2級</asp:ListItem>
					<asp:ListItem Value="3級">3級</asp:ListItem>
					<asp:ListItem Value="4級">4級</asp:ListItem>
					<asp:ListItem Value="5級">5級</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAC12" style="Z-INDEX: 276; LEFT: 568px; POSITION: absolute; TOP: 732px" runat="server"
					Width="48px" Height="346px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0級">0級</asp:ListItem>
					<asp:ListItem Value="1級">1級</asp:ListItem>
					<asp:ListItem Value="2級">2級</asp:ListItem>
					<asp:ListItem Value="3級">3級</asp:ListItem>
					<asp:ListItem Value="4級">4級</asp:ListItem>
					<asp:ListItem Value="5級">5級</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAB12" style="Z-INDEX: 275; LEFT: 568px; POSITION: absolute; TOP: 714px" runat="server"
					Width="48px" Height="346px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0級">0級</asp:ListItem>
					<asp:ListItem Value="1級">1級</asp:ListItem>
					<asp:ListItem Value="2級">2級</asp:ListItem>
					<asp:ListItem Value="3級">3級</asp:ListItem>
					<asp:ListItem Value="4級">4級</asp:ListItem>
					<asp:ListItem Value="5級">5級</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAD11" style="Z-INDEX: 274; LEFT: 520px; POSITION: absolute; TOP: 750px" runat="server"
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
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAC11" style="Z-INDEX: 273; LEFT: 520px; POSITION: absolute; TOP: 732px" runat="server"
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
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAB11" style="Z-INDEX: 272; LEFT: 520px; POSITION: absolute; TOP: 714px" runat="server"
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
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAA13" style="Z-INDEX: 271; LEFT: 616px; POSITION: absolute; TOP: 696px" runat="server"
					Width="48px" Height="346px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0級">0級</asp:ListItem>
					<asp:ListItem Value="1級">1級</asp:ListItem>
					<asp:ListItem Value="2級">2級</asp:ListItem>
					<asp:ListItem Value="3級">3級</asp:ListItem>
					<asp:ListItem Value="4級">4級</asp:ListItem>
					<asp:ListItem Value="5級">5級</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAA12" style="Z-INDEX: 270; LEFT: 568px; POSITION: absolute; TOP: 696px" runat="server"
					Width="48px" Height="346px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="0級">0級</asp:ListItem>
					<asp:ListItem Value="1級">1級</asp:ListItem>
					<asp:ListItem Value="2級">2級</asp:ListItem>
					<asp:ListItem Value="3級">3級</asp:ListItem>
					<asp:ListItem Value="4級">4級</asp:ListItem>
					<asp:ListItem Value="5級">5級</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAA11" style="Z-INDEX: 269; LEFT: 520px; POSITION: absolute; TOP: 696px" runat="server"
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
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAD14" style="Z-INDEX: 231; LEFT: 664px; POSITION: absolute; TOP: 750px" runat="server"
					Width="48px" Height="22px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAD10" style="Z-INDEX: 229; LEFT: 472px; POSITION: absolute; TOP: 750px" runat="server"
					Width="48px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAD9" style="Z-INDEX: 228; LEFT: 424px; POSITION: absolute; TOP: 750px" runat="server"
					Width="48px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAD8" style="Z-INDEX: 225; LEFT: 376px; POSITION: absolute; TOP: 750px" runat="server"
					Width="48px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
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
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAD7" style="Z-INDEX: 224; LEFT: 328px; POSITION: absolute; TOP: 750px" runat="server"
					Width="48px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAD6" style="Z-INDEX: 221; LEFT: 280px; POSITION: absolute; TOP: 750px" runat="server"
					Width="48px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:textbox id="DQAD5" style="Z-INDEX: 220; LEFT: 232px; POSITION: absolute; TOP: 750px" runat="server"
					BorderStyle="Groove" Width="47px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt"></asp:textbox>
				<asp:textbox id="DQAD4" style="Z-INDEX: 217; LEFT: 204px; POSITION: absolute; TOP: 750px" runat="server"
					BorderStyle="Groove" Width="28px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt"></asp:textbox>
				<asp:dropdownlist id="DQAD3" style="Z-INDEX: 216; LEFT: 170px; POSITION: absolute; TOP: 750px" runat="server"
					Width="35px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:textbox id="DQAD2" style="Z-INDEX: 213; LEFT: 72px; POSITION: absolute; TOP: 750px" runat="server"
					BorderStyle="Groove" Width="98px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt"></asp:textbox>
				<asp:textbox id="DQAD1" style="Z-INDEX: 212; LEFT: 16px; POSITION: absolute; TOP: 750px" runat="server"
					BorderStyle="Groove" Width="56px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt"></asp:textbox>
				<asp:dropdownlist id="DQAC14" style="Z-INDEX: 209; LEFT: 664px; POSITION: absolute; TOP: 732px" runat="server"
					Width="48px" Height="22px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAC10" style="Z-INDEX: 208; LEFT: 472px; POSITION: absolute; TOP: 732px" runat="server"
					Width="48px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAC9" style="Z-INDEX: 206; LEFT: 424px; POSITION: absolute; TOP: 732px" runat="server"
					Width="48px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAC8" style="Z-INDEX: 204; LEFT: 376px; POSITION: absolute; TOP: 732px" runat="server"
					Width="48px" Height="36px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
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
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAC7" style="Z-INDEX: 201; LEFT: 328px; POSITION: absolute; TOP: 732px" runat="server"
					Width="48px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAC6" style="Z-INDEX: 200; LEFT: 280px; POSITION: absolute; TOP: 732px" runat="server"
					Width="48px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:textbox id="DQAC5" style="Z-INDEX: 198; LEFT: 232px; POSITION: absolute; TOP: 732px" runat="server"
					BorderStyle="Groove" Width="47px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt"></asp:textbox>
				<asp:textbox id="DQAC4" style="Z-INDEX: 196; LEFT: 204px; POSITION: absolute; TOP: 732px" runat="server"
					BorderStyle="Groove" Width="28px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt"></asp:textbox>
				<asp:dropdownlist id="DQAC3" style="Z-INDEX: 195; LEFT: 170px; POSITION: absolute; TOP: 732px" runat="server"
					Width="35px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:textbox id="DQAC2" style="Z-INDEX: 194; LEFT: 72px; POSITION: absolute; TOP: 732px" runat="server"
					BorderStyle="Groove" Width="98px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt"></asp:textbox>
				<asp:textbox id="DQAC1" style="Z-INDEX: 192; LEFT: 16px; POSITION: absolute; TOP: 732px" runat="server"
					BorderStyle="Groove" Width="56px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt"></asp:textbox>
				<asp:dropdownlist id="DQAB14" style="Z-INDEX: 191; LEFT: 664px; POSITION: absolute; TOP: 714px" runat="server"
					Width="48px" Height="22px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAB10" style="Z-INDEX: 189; LEFT: 472px; POSITION: absolute; TOP: 714px" runat="server"
					Width="48px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAB9" style="Z-INDEX: 188; LEFT: 424px; POSITION: absolute; TOP: 714px" runat="server"
					Width="48px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAB8" style="Z-INDEX: 186; LEFT: 376px; POSITION: absolute; TOP: 714px" runat="server"
					Width="48px" Height="36px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
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
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAB7" style="Z-INDEX: 185; LEFT: 328px; POSITION: absolute; TOP: 714px" runat="server"
					Width="48px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAB6" style="Z-INDEX: 184; LEFT: 280px; POSITION: absolute; TOP: 714px" runat="server"
					Width="48px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:textbox id="DQAB5" style="Z-INDEX: 183; LEFT: 232px; POSITION: absolute; TOP: 714px" runat="server"
					BorderStyle="Groove" Width="47px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt"></asp:textbox>
				<asp:textbox id="DQAB4" style="Z-INDEX: 182; LEFT: 204px; POSITION: absolute; TOP: 714px" runat="server"
					BorderStyle="Groove" Width="28px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt"></asp:textbox>
				<asp:dropdownlist id="DQAB3" style="Z-INDEX: 181; LEFT: 170px; POSITION: absolute; TOP: 714px" runat="server"
					Width="35px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:textbox id="DQAB2" style="Z-INDEX: 180; LEFT: 72px; POSITION: absolute; TOP: 714px" runat="server"
					BorderStyle="Groove" Width="98px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt"></asp:textbox>
				<asp:textbox id="DQAB1" style="Z-INDEX: 179; LEFT: 16px; POSITION: absolute; TOP: 714px" runat="server"
					BorderStyle="Groove" Width="56px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt"></asp:textbox>
				<asp:dropdownlist id="DQAA14" style="Z-INDEX: 160; LEFT: 664px; POSITION: absolute; TOP: 696px" runat="server"
					Width="48px" Height="22px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAA10" style="Z-INDEX: 159; LEFT: 472px; POSITION: absolute; TOP: 696px" runat="server"
					Width="48px" Height="346px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAA9" style="Z-INDEX: 156; LEFT: 424px; POSITION: absolute; TOP: 696px" runat="server"
					Width="48px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAA8" style="Z-INDEX: 154; LEFT: 376px; POSITION: absolute; TOP: 696px" runat="server"
					Width="48px" Height="36px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
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
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAA7" style="Z-INDEX: 152; LEFT: 328px; POSITION: absolute; TOP: 696px" runat="server"
					Width="48px" Height="162px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQAA6" style="Z-INDEX: 151; LEFT: 280px; POSITION: absolute; TOP: 696px" runat="server"
					Width="48px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:textbox id="DQAA5" style="Z-INDEX: 149; LEFT: 232px; POSITION: absolute; TOP: 696px" runat="server"
					BorderStyle="Groove" Width="47px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt"></asp:textbox>
				<asp:textbox id="DQAA4" style="Z-INDEX: 147; LEFT: 204px; POSITION: absolute; TOP: 696px" runat="server"
					BorderStyle="Groove" Width="28px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt"></asp:textbox>
				<asp:dropdownlist id="DQAA3" style="Z-INDEX: 145; LEFT: 170px; POSITION: absolute; TOP: 696px" runat="server"
					Width="35px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist>
				<asp:textbox id="DQAA1" style="Z-INDEX: 143; LEFT: 16px; POSITION: absolute; TOP: 696px" runat="server"
					BorderStyle="Groove" Width="56px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt"></asp:textbox>
				<asp:textbox id="DQAA2" style="Z-INDEX: 168; LEFT: 72px; POSITION: absolute; TOP: 696px" runat="server"
					BorderStyle="Groove" Width="98px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="7pt"></asp:textbox>
				<asp:label id="Label1" style="Z-INDEX: 268; LEFT: 48px; POSITION: absolute; TOP: 684px" runat="server"
					Width="24px" Height="10px" Font-Size="8pt">日期</asp:label>
				<asp:label id="Label13" style="Z-INDEX: 259; LEFT: 664px; POSITION: absolute; TOP: 684px" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">CPSC</asp:label>
				<asp:label id="Label26" style="Z-INDEX: 281; LEFT: 568px; POSITION: absolute; TOP: 684px" runat="server"
					Width="48px" Height="10px" Font-Size="8pt">一次密著</asp:label>
				<asp:label id="Label12" style="Z-INDEX: 257; LEFT: 616px; POSITION: absolute; TOP: 684px" runat="server"
					Width="48px" Height="10px" Font-Size="8pt">二次密著</asp:label>
				<asp:label id="Label11" style="Z-INDEX: 256; LEFT: 520px; POSITION: absolute; TOP: 684px" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">黃變</asp:label>
				<asp:label id="Label10" style="Z-INDEX: 254; LEFT: 472px; POSITION: absolute; TOP: 684px" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">乾洗</asp:label>
				<asp:label id="Label9" style="Z-INDEX: 252; LEFT: 424px; POSITION: absolute; TOP: 684px" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">水洗</asp:label>
				<asp:label id="Label8" style="Z-INDEX: 251; LEFT: 376px; POSITION: absolute; TOP: 684px" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">檢針</asp:label>
				<asp:label id="Label7" style="Z-INDEX: 250; LEFT: 328px; POSITION: absolute; TOP: 684px" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">扭力</asp:label>
				<asp:label id="Label6" style="Z-INDEX: 249; LEFT: 280px; POSITION: absolute; TOP: 684px" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">強度</asp:label>
				<asp:label id="Label5" style="Z-INDEX: 248; LEFT: 240px; POSITION: absolute; TOP: 684px" runat="server"
					Width="40px" Height="10px" Font-Size="8pt">原單位</asp:label>
				<asp:label id="Label4" style="Z-INDEX: 247; LEFT: 208px; POSITION: absolute; TOP: 684px" runat="server"
					Width="48px" Height="10px" Font-Size="8pt">表面</asp:label>
				<asp:label id="Label3" style="Z-INDEX: 246; LEFT: 168px; POSITION: absolute; TOP: 684px" runat="server"
					Width="32px" Height="10px" Font-Size="8pt">口厚</asp:label>
				<asp:label id="Label2" style="Z-INDEX: 245; LEFT: 80px; POSITION: absolute; TOP: 684px" runat="server"
					Width="48px" Height="10px" Font-Size="8pt">拉頭種類</asp:label>
				<asp:hyperlink id="LSurface" style="Z-INDEX: 267; LEFT: 152px; POSITION: absolute; TOP: 56px" runat="server"
					Width="80px" Height="8px" Target="_blank" NavigateUrl="BoardEdit.aspx" Font-Size="12pt">有表面處理</asp:hyperlink>
				<asp:hyperlink id="LColorAppend" style="Z-INDEX: 266; LEFT: 240px; POSITION: absolute; TOP: 56px"
					runat="server" Height="8px" Width="112px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">有外注色番追加</asp:hyperlink><asp:dropdownlist id="DLogo" style="Z-INDEX: 264; LEFT: 686px; POSITION: absolute; TOP: 224px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="20px" Width="78px">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DManufFlow" style="Z-INDEX: 263; LEFT: 56px; POSITION: absolute; TOP: 480px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="48px" Width="369px" BorderStyle="Groove" MaxLength="240"
					TextMode="MultiLine">ManufFlow</asp:textbox><asp:textbox id="DDevReason" style="Z-INDEX: 125; LEFT: 56px; POSITION: absolute; TOP: 424px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="48px" Width="369px" BorderStyle="Groove" MaxLength="240" TextMode="MultiLine">DevReason</asp:textbox><asp:dropdownlist id="DCPSC" style="Z-INDEX: 262; LEFT: 708px; POSITION: absolute; TOP: 192px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="20px" Width="56px">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:hyperlink id="LSliderDetail" style="Z-INDEX: 261; LEFT: 656px; POSITION: absolute; TOP: 56px"
					runat="server" Height="8px" Width="112px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">有追加拉頭細目</asp:hyperlink><asp:hyperlink id="LContact" style="Z-INDEX: 260; LEFT: 536px; POSITION: absolute; TOP: 56px" runat="server"
					Height="8px" Width="112px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">有追加業務連絡</asp:hyperlink><asp:hyperlink id="LOPContact" style="Z-INDEX: 258; LEFT: 416px; POSITION: absolute; TOP: 56px"
					runat="server" Height="8px" Width="121px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">有型別胴體追加</asp:hyperlink><asp:textbox id="DStatus" style="Z-INDEX: 255; LEFT: 560px; POSITION: absolute; TOP: 16px" runat="server"
					ForeColor="White" BackColor="Red" Height="32px" Width="212px" BorderStyle="Groove" Font-Size="10pt" ReadOnly="True">修改工程進行中(xxxxxxxxxxxxxxxxxx)</asp:textbox><asp:label id="Label25" style="Z-INDEX: 243; LEFT: 704px; POSITION: absolute; TOP: 436px" runat="server"
					Height="10px" Width="24px" Font-Size="8pt">數量</asp:label><asp:dropdownlist id="DMakeCAM" style="Z-INDEX: 244; LEFT: 584px; POSITION: absolute; TOP: 360px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="174px">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:label id="Label24" style="Z-INDEX: 242; LEFT: 648px; POSITION: absolute; TOP: 436px" runat="server"
					Height="10px" Width="24px" Font-Size="8pt">顏色</asp:label><asp:label id="Label23" style="Z-INDEX: 241; LEFT: 472px; POSITION: absolute; TOP: 436px" runat="server"
					Height="10px" Width="48px" Font-Size="8pt">拉頭種類</asp:label><asp:label id="Label22" style="Z-INDEX: 240; LEFT: 504px; POSITION: absolute; TOP: 556px" runat="server"
					Height="10px" Width="24px" Font-Size="8pt">單價</asp:label><asp:label id="Label21" style="Z-INDEX: 239; LEFT: 464px; POSITION: absolute; TOP: 556px" runat="server"
					Height="10px" Width="24px" Font-Size="8pt">幣別</asp:label><asp:label id="Label20" style="Z-INDEX: 238; LEFT: 296px; POSITION: absolute; TOP: 556px" runat="server"
					Height="10px" Width="48px" Font-Size="8pt">拉頭種類</asp:label><asp:label id="Label19" style="Z-INDEX: 237; LEFT: 224px; POSITION: absolute; TOP: 556px" runat="server"
					Height="10px" Width="24px" Font-Size="8pt">單價</asp:label><asp:label id="Label18" style="Z-INDEX: 236; LEFT: 184px; POSITION: absolute; TOP: 556px" runat="server"
					Height="10px" Width="24px" Font-Size="8pt">幣別</asp:label><asp:label id="Label17" style="Z-INDEX: 235; LEFT: 56px; POSITION: absolute; TOP: 556px" runat="server"
					Height="10px" Width="48px" Font-Size="8pt">拉頭種類</asp:label><asp:label id="Label16" style="Z-INDEX: 234; LEFT: 136px; POSITION: absolute; TOP: 812px" runat="server"
					Height="10px" Width="32px" Font-Size="8pt">備註</asp:label><asp:label id="Label15" style="Z-INDEX: 233; LEFT: 80px; POSITION: absolute; TOP: 812px" runat="server"
					Height="10px" Width="48px" Font-Size="8pt">檢測結果</asp:label><asp:label id="Label14" style="Z-INDEX: 232; LEFT: 48px; POSITION: absolute; TOP: 812px" runat="server"
					Height="10px" Width="32px" Font-Size="8pt">日期</asp:label><asp:image id="DManuaInSheet1" style="Z-INDEX: 100; LEFT: 8px; POSITION: absolute; TOP: 8px"
					runat="server" Height="530px" Width="761px" ImageUrl="Images\ManufInSheet_003_E.jpg"></asp:image><asp:hyperlink id="LQAttachFile2" style="Z-INDEX: 230; LEFT: 136px; POSITION: absolute; TOP: 900px"
					runat="server" Height="8px" Width="136px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">品質分析依賴書</asp:hyperlink><asp:hyperlink id="LQAttachFile1" style="Z-INDEX: 227; LEFT: 16px; POSITION: absolute; TOP: 772px"
					runat="server" Height="8px" Width="64px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">其他附件</asp:hyperlink><asp:dropdownlist id="DPriceI2" style="Z-INDEX: 226; LEFT: 184px; POSITION: absolute; TOP: 640px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="12px" Width="40px" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:hyperlink id="LSAttachFile" style="Z-INDEX: 223; LEFT: 440px; POSITION: absolute; TOP: 508px"
					runat="server" Height="8px" Width="64px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">其他附件</asp:hyperlink><asp:textbox id="DFormSno" style="Z-INDEX: 222; LEFT: 8px; POSITION: absolute; TOP: 1032px" runat="server"
					ForeColor="Blue" BackColor="White" Height="20px" Width="97px" BorderStyle="None">單號：123</asp:textbox><asp:image id="Image1" style="Z-INDEX: 219; LEFT: 8px; POSITION: absolute; TOP: 1032px" runat="server"
					Height="29px" Width="755px" ImageUrl="Images\ManufInSheet_New_004.jpg"></asp:image><asp:textbox id="DFQAD3" style="Z-INDEX: 218; LEFT: 136px; POSITION: absolute; TOP: 878px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="18px" Width="620px" BorderStyle="Groove" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:dropdownlist id="DFQAD2" style="Z-INDEX: 215; LEFT: 80px; POSITION: absolute; TOP: 878px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="18px" Width="56px" Font-Size="8pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DFQAD1" style="Z-INDEX: 214; LEFT: 16px; POSITION: absolute; TOP: 878px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="18px" Width="64px" BorderStyle="Groove" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DFQAC1" style="Z-INDEX: 211; LEFT: 16px; POSITION: absolute; TOP: 860px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="18px" Width="64px" BorderStyle="Groove" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:dropdownlist id="DFQAC2" style="Z-INDEX: 210; LEFT: 80px; POSITION: absolute; TOP: 860px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="18px" Width="56px" Font-Size="8pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DFQAC3" style="Z-INDEX: 207; LEFT: 136px; POSITION: absolute; TOP: 860px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="18px" Width="620px" BorderStyle="Groove" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DFQAB3" style="Z-INDEX: 205; LEFT: 136px; POSITION: absolute; TOP: 842px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="18px" Width="620px" BorderStyle="Groove" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:dropdownlist id="DFQAB2" style="Z-INDEX: 203; LEFT: 80px; POSITION: absolute; TOP: 842px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="18px" Width="56px" Font-Size="8pt">
					<asp:ListItem Value="無" Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DFQAB1" style="Z-INDEX: 202; LEFT: 16px; POSITION: absolute; TOP: 842px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="18px" Width="64px" BorderStyle="Groove" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:hyperlink id="LMapNo" style="Z-INDEX: 197; LEFT: 320px; POSITION: absolute; TOP: 296px" runat="server"
					Height="8px" Width="112px" NavigateUrl="BoardEdit.aspx" Target="_blank">圖號文件</asp:hyperlink><asp:textbox id="DFQAA3" style="Z-INDEX: 193; LEFT: 136px; POSITION: absolute; TOP: 824px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="18px" Width="620px" BorderStyle="Groove" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:dropdownlist id="DFQAA2" style="Z-INDEX: 190; LEFT: 80px; POSITION: absolute; TOP: 824px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="18px" Width="56px" Font-Size="8pt">
					<asp:ListItem Selected="True">無</asp:ListItem>
					<asp:ListItem Value="PASS">PASS</asp:ListItem>
					<asp:ListItem Value="FAIL">FAIL</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DFQAA1" style="Z-INDEX: 187; LEFT: 16px; POSITION: absolute; TOP: 824px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="18px" Width="64px" BorderStyle="Groove" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DPriceJ3" style="Z-INDEX: 178; LEFT: 504px; POSITION: absolute; TOP: 640px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="18px" Width="64px" BorderStyle="Groove" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:dropdownlist id="DPriceJ2" style="Z-INDEX: 177; LEFT: 464px; POSITION: absolute; TOP: 640px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="12px" Width="40px" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceJ1" style="Z-INDEX: 176; LEFT: 296px; POSITION: absolute; TOP: 640px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="18px" Width="166px" BorderStyle="Groove" Font-Size="8pt"
					ReadOnly="True"></asp:textbox><asp:textbox id="DPriceI3" style="Z-INDEX: 175; LEFT: 224px; POSITION: absolute; TOP: 640px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="18px" Width="64px" BorderStyle="Groove" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:dropdownlist id="DPriceG2" style="Z-INDEX: 174; LEFT: 184px; POSITION: absolute; TOP: 622px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="12px" Width="40px" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceI1" style="Z-INDEX: 173; LEFT: 16px; POSITION: absolute; TOP: 640px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="18px" Width="166px" BorderStyle="Groove" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DPriceH3" style="Z-INDEX: 172; LEFT: 504px; POSITION: absolute; TOP: 622px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="18px" Width="64px" BorderStyle="Groove" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:dropdownlist id="DPriceH2" style="Z-INDEX: 171; LEFT: 464px; POSITION: absolute; TOP: 622px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="12px" Width="40px" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceH1" style="Z-INDEX: 170; LEFT: 296px; POSITION: absolute; TOP: 622px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="18px" Width="166px" BorderStyle="Groove" Font-Size="8pt"
					ReadOnly="True"></asp:textbox><asp:textbox id="DPriceG3" style="Z-INDEX: 169; LEFT: 224px; POSITION: absolute; TOP: 622px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="18px" Width="64px" BorderStyle="Groove" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:dropdownlist id="DPriceE2" style="Z-INDEX: 167; LEFT: 184px; POSITION: absolute; TOP: 604px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="12px" Width="40px" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceG1" style="Z-INDEX: 166; LEFT: 16px; POSITION: absolute; TOP: 622px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="18px" Width="166px" BorderStyle="Groove" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DPriceF3" style="Z-INDEX: 165; LEFT: 504px; POSITION: absolute; TOP: 604px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="18px" Width="64px" BorderStyle="Groove" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:dropdownlist id="DPriceF2" style="Z-INDEX: 164; LEFT: 464px; POSITION: absolute; TOP: 604px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="12px" Width="40px" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceF1" style="Z-INDEX: 163; LEFT: 296px; POSITION: absolute; TOP: 604px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="18px" Width="166px" BorderStyle="Groove" Font-Size="8pt"
					ReadOnly="True"></asp:textbox><asp:textbox id="DPriceE3" style="Z-INDEX: 162; LEFT: 224px; POSITION: absolute; TOP: 604px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="18px" Width="64px" BorderStyle="Groove" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:dropdownlist id="DPriceC2" style="Z-INDEX: 161; LEFT: 184px; POSITION: absolute; TOP: 586px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="12px" Width="40px" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceE1" style="Z-INDEX: 158; LEFT: 16px; POSITION: absolute; TOP: 604px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="18px" Width="166px" BorderStyle="Groove" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DPriceD3" style="Z-INDEX: 157; LEFT: 504px; POSITION: absolute; TOP: 586px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="18px" Width="64px" BorderStyle="Groove" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:dropdownlist id="DPriceD2" style="Z-INDEX: 155; LEFT: 464px; POSITION: absolute; TOP: 586px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="12px" Width="40px" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceD1" style="Z-INDEX: 153; LEFT: 296px; POSITION: absolute; TOP: 586px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="18px" Width="166px" BorderStyle="Groove" Font-Size="8pt"
					ReadOnly="True"></asp:textbox><asp:textbox id="DPriceC3" style="Z-INDEX: 150; LEFT: 224px; POSITION: absolute; TOP: 586px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="18px" Width="64px" BorderStyle="Groove" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:dropdownlist id="DPriceA2" style="Z-INDEX: 148; LEFT: 184px; POSITION: absolute; TOP: 568px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="12px" Width="40px" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceC1" style="Z-INDEX: 146; LEFT: 16px; POSITION: absolute; TOP: 586px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="18px" Width="166px" BorderStyle="Groove" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DPriceB3" style="Z-INDEX: 144; LEFT: 504px; POSITION: absolute; TOP: 568px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="18px" Width="64px" BorderStyle="Groove" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:dropdownlist id="DPriceB2" style="Z-INDEX: 142; LEFT: 464px; POSITION: absolute; TOP: 568px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="12px" Width="40px" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceB1" style="Z-INDEX: 141; LEFT: 296px; POSITION: absolute; TOP: 568px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="18px" Width="166px" BorderStyle="Groove" Font-Size="8pt"
					ReadOnly="True"></asp:textbox><asp:textbox id="DSampleC3" style="Z-INDEX: 140; LEFT: 704px; POSITION: absolute; TOP: 486px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="18px" Width="52px" BorderStyle="Groove" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DSampleC2" style="Z-INDEX: 139; LEFT: 648px; POSITION: absolute; TOP: 486px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="18px" Width="52px" BorderStyle="Groove" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DSampleC1" style="Z-INDEX: 138; LEFT: 440px; POSITION: absolute; TOP: 486px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="18px" Width="204px" BorderStyle="Groove" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DSampleB3" style="Z-INDEX: 137; LEFT: 704px; POSITION: absolute; TOP: 468px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="18px" Width="52px" BorderStyle="Groove" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DSampleB2" style="Z-INDEX: 136; LEFT: 648px; POSITION: absolute; TOP: 468px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="18px" Width="52px" BorderStyle="Groove" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DSampleB1" style="Z-INDEX: 135; LEFT: 440px; POSITION: absolute; TOP: 468px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="18px" Width="204px" BorderStyle="Groove" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DPriceA1" style="Z-INDEX: 134; LEFT: 16px; POSITION: absolute; TOP: 568px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="18px" Width="166px" BorderStyle="Groove" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DSampleA1" style="Z-INDEX: 133; LEFT: 440px; POSITION: absolute; TOP: 450px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="18px" Width="204px" BorderStyle="Groove" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:hyperlink id="LContactFile" style="Z-INDEX: 132; LEFT: 120px; POSITION: absolute; TOP: 1000px"
					runat="server" Height="8px" Width="50px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">切結書</asp:hyperlink><asp:hyperlink id="LRefFile" style="Z-INDEX: 131; LEFT: 120px; POSITION: absolute; TOP: 932px"
					runat="server" Height="8px" Width="64px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">其他文件</asp:hyperlink><asp:hyperlink id="LQAFile" style="Z-INDEX: 130; LEFT: 496px; POSITION: absolute; TOP: 932px" runat="server"
					Height="8px" Width="64px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">測試報告</asp:hyperlink><asp:textbox id="DPriceA3" style="Z-INDEX: 129; LEFT: 224px; POSITION: absolute; TOP: 568px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="18px" Width="64px" BorderStyle="Groove" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DSampleA2" style="Z-INDEX: 102; LEFT: 648px; POSITION: absolute; TOP: 450px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="18px" Width="52px" BorderStyle="Groove" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DSampleA3" style="Z-INDEX: 103; LEFT: 704px; POSITION: absolute; TOP: 450px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="18px" Width="52px" BorderStyle="Groove" Font-Size="8pt" ReadOnly="True"></asp:textbox><asp:textbox id="DDate" style="Z-INDEX: 128; LEFT: 520px; POSITION: absolute; TOP: 192px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="20px" Width="112px" BorderStyle="Groove" ReadOnly="True">DDate</asp:textbox><asp:textbox id="DSliderCode" style="Z-INDEX: 127; LEFT: 224px; POSITION: absolute; TOP: 120px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="56px" Width="368px" BorderStyle="Groove" TextMode="MultiLine" Font-Size="12pt" ReadOnly="True">SliderCode</asp:textbox><asp:hyperlink id="LSampleFile" style="Z-INDEX: 126; LEFT: 496px; POSITION: absolute; TOP: 1000px"
					runat="server" Height="8px" Width="50px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">樣品圖</asp:hyperlink><asp:dropdownlist id="DPerson" style="Z-INDEX: 124; LEFT: 520px; POSITION: absolute; TOP: 226px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="20px" Width="96px" AutoPostBack="True">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DDivision" style="Z-INDEX: 123; LEFT: 312px; POSITION: absolute; TOP: 226px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="112px" AutoPostBack="True">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DMoldPoint" style="Z-INDEX: 122; LEFT: 672px; POSITION: absolute; TOP: 395px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="45px" BorderStyle="Groove" ReadOnly="True">MoldPoint</asp:textbox><asp:textbox id="DMoldQty" style="Z-INDEX: 121; LEFT: 584px; POSITION: absolute; TOP: 395px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="45px" BorderStyle="Groove" ReadOnly="True">DMoldQty</asp:textbox><asp:textbox id="DSuppiler" style="Z-INDEX: 120; LEFT: 392px; POSITION: absolute; TOP: 395px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="92px" BorderStyle="Groove" ReadOnly="True">Suppiler</asp:textbox><asp:textbox id="DPullerPrice" style="Z-INDEX: 119; LEFT: 688px; POSITION: absolute; TOP: 640px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="56px" BorderStyle="Groove" ReadOnly="True" Visible="False">PullerPrice</asp:textbox><asp:textbox id="DPurMold" style="Z-INDEX: 118; LEFT: 688px; POSITION: absolute; TOP: 608px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="76px" BorderStyle="Groove" ReadOnly="True">PurMold</asp:textbox><asp:textbox id="DArMoldFee" style="Z-INDEX: 117; LEFT: 688px; POSITION: absolute; TOP: 574px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="76px" BorderStyle="Groove" ReadOnly="True">ArMoldFee</asp:textbox><asp:hyperlink id="LAuthorizeFile" style="Z-INDEX: 116; LEFT: 496px; POSITION: absolute; TOP: 968px"
					runat="server" Height="8px" Width="50px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">授權書</asp:hyperlink><asp:hyperlink id="LConfirmFile" style="Z-INDEX: 115; LEFT: 120px; POSITION: absolute; TOP: 968px"
					runat="server" Height="8px" Width="50px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">確認書</asp:hyperlink><asp:textbox id="DBuyer" style="Z-INDEX: 114; LEFT: 312px; POSITION: absolute; TOP: 328px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="20px" Width="168px" BorderStyle="Groove" ReadOnly="True">DBuyer</asp:textbox><asp:textbox id="DSellVendor" style="Z-INDEX: 113; LEFT: 688px; POSITION: absolute; TOP: 328px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="71px" BorderStyle="Groove" ReadOnly="True">DSellVendor</asp:textbox><asp:dropdownlist id="DMaterial" style="Z-INDEX: 112; LEFT: 104px; POSITION: absolute; TOP: 395px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="192px" AutoPostBack="True">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DManufPlace" style="Z-INDEX: 111; LEFT: 392px; POSITION: absolute; TOP: 362px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="92px">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DSliderType2" style="Z-INDEX: 110; LEFT: 200px; POSITION: absolute; TOP: 362px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="94px">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DSliderType1" style="Z-INDEX: 109; LEFT: 104px; POSITION: absolute; TOP: 362px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="94px">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DAssembler" style="Z-INDEX: 108; LEFT: 600px; POSITION: absolute; TOP: 293px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="40px" BorderStyle="Groove" ReadOnly="True">DAssembler</asp:textbox><asp:dropdownlist id="DLevel" style="Z-INDEX: 107; LEFT: 648px; POSITION: absolute; TOP: 261px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="14px" Width="116px">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DSpec" style="Z-INDEX: 106; LEFT: 312px; POSITION: absolute; TOP: 259px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="20px" Width="232px" BorderStyle="Groove" Font-Size="8pt" ReadOnly="True">DSpec</asp:textbox><asp:textbox id="DSliderGRCode" style="Z-INDEX: 105; LEFT: 604px; POSITION: absolute; TOP: 158px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="154px" BorderStyle="Groove" ReadOnly="True">DSliderGRCode</asp:textbox><asp:textbox id="DNo" style="Z-INDEX: 104; LEFT: 312px; POSITION: absolute; TOP: 192px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="20px" Width="112px" BorderStyle="Groove" ReadOnly="True">DNo</asp:textbox><asp:image id="DManuaInSheet2" style="Z-INDEX: 101; LEFT: 8px; POSITION: absolute; TOP: 536px"
					runat="server" Height="493px" Width="761px" ImageUrl="Images\ManufInSheet_005_C1.jpg"></asp:image><asp:image id="LMapFile" style="Z-INDEX: 199; LEFT: 16px; POSITION: absolute; TOP: 120px" runat="server"
					Height="230px" Width="200px" BorderStyle="Groove" ImageUrl="http://10.245.0.178//WorkFlow/Document/000003/1-Map2-2006124111750-2.jpg"></asp:image><asp:panel id="Panel1" style="Z-INDEX: 253; LEFT: 560px; POSITION: absolute; TOP: 8px" runat="server"
					BackColor="White" Height="64px" Width="210px"></asp:panel></FONT></form>
	</body>
</HTML>
