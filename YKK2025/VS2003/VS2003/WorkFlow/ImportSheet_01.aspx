<%@ Page Language="vb" AutoEventWireup="false" Codebehind="ImportSheet_01.aspx.vb" Inherits="SPD.ImportSheet" aspCompat="True"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>ImportSheet</title>
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
		<form id="Form1" method="post" encType="multipart/form-data" runat="server">
			<FONT face="新細明體">
				<asp:textbox id="DOPReadyDesc" style="Z-INDEX: 312; LEFT: 248px; POSITION: absolute; TOP: 808px"
					runat="server" Width="104px" Height="20px" BackColor="White" ForeColor="Red" BorderStyle="None">需閱讀工程履歷</asp:textbox><asp:textbox id="DOPReady" style="Z-INDEX: 152; LEFT: 352px; POSITION: absolute; TOP: 808px"
					runat="server" Width="48px" Height="20px" BackColor="GreenYellow" ForeColor="Red" BorderStyle="Groove" ReadOnly="True"></asp:textbox><asp:image id="DImportSheet" style="Z-INDEX: 100; LEFT: 8px; POSITION: absolute; TOP: 8px"
					runat="server" Width="761px" Height="644px" ImageUrl="Images\ImportSheet_002.jpg"></asp:image><asp:label id="Label22" style="Z-INDEX: 291; LEFT: 504px; POSITION: absolute; TOP: 506px" runat="server"
					Width="24px" Height="10px" Font-Size="8pt">單價</asp:label><asp:label id="Label21" style="Z-INDEX: 290; LEFT: 464px; POSITION: absolute; TOP: 506px" runat="server"
					Width="24px" Height="10px" Font-Size="8pt">幣別</asp:label><asp:label id="Label20" style="Z-INDEX: 289; LEFT: 296px; POSITION: absolute; TOP: 506px" runat="server"
					Width="48px" Height="10px" Font-Size="8pt">拉頭種類</asp:label><asp:label id="Label19" style="Z-INDEX: 288; LEFT: 224px; POSITION: absolute; TOP: 506px" runat="server"
					Width="24px" Height="10px" Font-Size="8pt">單價</asp:label><asp:label id="Label18" style="Z-INDEX: 287; LEFT: 184px; POSITION: absolute; TOP: 506px" runat="server"
					Width="24px" Height="10px" Font-Size="8pt">幣別</asp:label><asp:label id="Label17" style="Z-INDEX: 286; LEFT: 56px; POSITION: absolute; TOP: 506px" runat="server"
					Width="48px" Height="10px" Font-Size="8pt">拉頭種類</asp:label><asp:dropdownlist id="DBuyer" style="Z-INDEX: 270; LEFT: 312px; POSITION: absolute; TOP: 296px" runat="server"
					Width="176px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:imagebutton id="BFlow" style="Z-INDEX: 263; LEFT: 8px; POSITION: absolute; TOP: 32px" runat="server"
					Width="16px" Height="16px" ImageUrl="Images\Flow-01.gif"></asp:imagebutton><INPUT id="DMapFile" style="Z-INDEX: 246; LEFT: 16px; WIDTH: 200px; POSITION: absolute; TOP: 328px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					type="file" size="14" name="File1" runat="server"><asp:textbox id="DPriceJ1" style="Z-INDEX: 202; LEFT: 296px; POSITION: absolute; TOP: 590px"
					runat="server" Width="166px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DPriceI3" style="Z-INDEX: 201; LEFT: 224px; POSITION: absolute; TOP: 590px"
					runat="server" Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:dropdownlist id="DPriceI2" style="Z-INDEX: 200; LEFT: 184px; POSITION: absolute; TOP: 590px"
					runat="server" Width="40px" Height="12px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceI1" style="Z-INDEX: 199; LEFT: 16px; POSITION: absolute; TOP: 590px" runat="server"
					Width="166px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DPriceH3" style="Z-INDEX: 198; LEFT: 504px; POSITION: absolute; TOP: 572px"
					runat="server" Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:dropdownlist id="DPriceH2" style="Z-INDEX: 197; LEFT: 464px; POSITION: absolute; TOP: 572px"
					runat="server" Width="40px" Height="12px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceJ3" style="Z-INDEX: 204; LEFT: 504px; POSITION: absolute; TOP: 590px"
					runat="server" Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:dropdownlist id="DPriceJ2" style="Z-INDEX: 203; LEFT: 464px; POSITION: absolute; TOP: 590px"
					runat="server" Width="40px" Height="12px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceA3" style="Z-INDEX: 150; LEFT: 224px; POSITION: absolute; TOP: 518px"
					runat="server" Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:dropdownlist id="DPriceA2" style="Z-INDEX: 149; LEFT: 184px; POSITION: absolute; TOP: 518px"
					runat="server" Width="40px" Height="12px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceH1" style="Z-INDEX: 196; LEFT: 296px; POSITION: absolute; TOP: 572px"
					runat="server" Width="166px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DPriceG3" style="Z-INDEX: 195; LEFT: 224px; POSITION: absolute; TOP: 572px"
					runat="server" Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:dropdownlist id="DPriceG2" style="Z-INDEX: 194; LEFT: 184px; POSITION: absolute; TOP: 572px"
					runat="server" Width="40px" Height="12px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceG1" style="Z-INDEX: 193; LEFT: 16px; POSITION: absolute; TOP: 572px" runat="server"
					Width="166px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DPriceF3" style="Z-INDEX: 192; LEFT: 504px; POSITION: absolute; TOP: 554px"
					runat="server" Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:dropdownlist id="DPriceF2" style="Z-INDEX: 191; LEFT: 464px; POSITION: absolute; TOP: 554px"
					runat="server" Width="40px" Height="12px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceF1" style="Z-INDEX: 190; LEFT: 296px; POSITION: absolute; TOP: 554px"
					runat="server" Width="166px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DPriceE3" style="Z-INDEX: 189; LEFT: 224px; POSITION: absolute; TOP: 554px"
					runat="server" Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:dropdownlist id="DPriceE2" style="Z-INDEX: 188; LEFT: 184px; POSITION: absolute; TOP: 554px"
					runat="server" Width="40px" Height="12px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceE1" style="Z-INDEX: 187; LEFT: 16px; POSITION: absolute; TOP: 554px" runat="server"
					Width="166px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DPriceD3" style="Z-INDEX: 186; LEFT: 504px; POSITION: absolute; TOP: 536px"
					runat="server" Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:dropdownlist id="DPriceD2" style="Z-INDEX: 185; LEFT: 464px; POSITION: absolute; TOP: 536px"
					runat="server" Width="40px" Height="12px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceD1" style="Z-INDEX: 184; LEFT: 296px; POSITION: absolute; TOP: 536px"
					runat="server" Width="166px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DPriceC3" style="Z-INDEX: 183; LEFT: 224px; POSITION: absolute; TOP: 536px"
					runat="server" Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:dropdownlist id="DPriceC2" style="Z-INDEX: 182; LEFT: 184px; POSITION: absolute; TOP: 536px"
					runat="server" Width="40px" Height="12px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceC1" style="Z-INDEX: 181; LEFT: 16px; POSITION: absolute; TOP: 536px" runat="server"
					Width="166px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DPriceB3" style="Z-INDEX: 180; LEFT: 504px; POSITION: absolute; TOP: 518px"
					runat="server" Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:dropdownlist id="DPriceB2" style="Z-INDEX: 179; LEFT: 464px; POSITION: absolute; TOP: 518px"
					runat="server" Width="40px" Height="12px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceB1" style="Z-INDEX: 178; LEFT: 296px; POSITION: absolute; TOP: 518px"
					runat="server" Width="166px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DPriceA1" style="Z-INDEX: 169; LEFT: 16px; POSITION: absolute; TOP: 518px" runat="server"
					Width="166px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:button id="BSliderCode" style="Z-INDEX: 167; LEFT: 568px; POSITION: absolute; TOP: 120px"
					runat="server" Width="20px" Height="20px" Text="....."></asp:button>
				<asp:hyperlink id="LRefFile" style="Z-INDEX: 164; LEFT: 120px; POSITION: absolute; TOP: 624px"
					runat="server" Width="64px" Height="8px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">其他文件</asp:hyperlink><asp:button id="BDate" style="Z-INDEX: 148; LEFT: 736px; POSITION: absolute; TOP: 192px" runat="server"
					Width="20px" Height="20px" Text="....."></asp:button><asp:textbox id="DDate" style="Z-INDEX: 147; LEFT: 584px; POSITION: absolute; TOP: 192px" runat="server"
					Width="152px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True">DDate</asp:textbox><asp:textbox id="DSliderCode" style="Z-INDEX: 146; LEFT: 224px; POSITION: absolute; TOP: 120px"
					runat="server" Width="342px" Height="56px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove" ReadOnly="True" TextMode="MultiLine">SliderCode</asp:textbox><asp:dropdownlist id="DPerson" style="Z-INDEX: 142; LEFT: 584px; POSITION: absolute; TOP: 226px" runat="server"
					Width="176px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DDivision" style="Z-INDEX: 141; LEFT: 312px; POSITION: absolute; TOP: 226px"
					runat="server" Width="176px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DReasonDesc" style="Z-INDEX: 140; LEFT: 168px; POSITION: absolute; TOP: 880px"
					runat="server" Width="424px" Height="59px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" TextMode="MultiLine"
					Visible="False">DReasonDesc</asp:textbox><asp:textbox id="DReason" style="Z-INDEX: 139; LEFT: 240px; POSITION: absolute; TOP: 848px" runat="server"
					Width="352px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True" Visible="False">DReason</asp:textbox><asp:dropdownlist id="DReasonCode" style="Z-INDEX: 138; LEFT: 168px; POSITION: absolute; TOP: 848px"
					runat="server" Width="64px" Height="20px" BackColor="Yellow" ForeColor="Blue" AutoPostBack="True" Visible="False">
					<asp:ListItem Value="01">01</asp:ListItem>
					<asp:ListItem Value="02">02</asp:ListItem>
				</asp:dropdownlist><asp:hyperlink id="LBefOP" style="Z-INDEX: 137; LEFT: 168px; POSITION: absolute; TOP: 808px" runat="server"
					Width="144px" Height="8px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">工程履歷</asp:hyperlink><asp:textbox id="DAEndTime" style="Z-INDEX: 136; LEFT: 440px; POSITION: absolute; TOP: 776px"
					runat="server" Width="152px" Height="20px" BackColor="Gold" ForeColor="Blue" Font-Size="9pt" BorderStyle="Groove" ReadOnly="True">DAEndTime</asp:textbox><asp:textbox id="DAStartTime" style="Z-INDEX: 135; LEFT: 168px; POSITION: absolute; TOP: 776px"
					runat="server" Width="152px" Height="20px" BackColor="Gold" ForeColor="Blue" Font-Size="9pt" BorderStyle="Groove" ReadOnly="True">DAStartTime</asp:textbox><asp:textbox id="DBEndTime" style="Z-INDEX: 134; LEFT: 440px; POSITION: absolute; TOP: 744px"
					runat="server" Width="152px" Height="20px" BackColor="Gold" ForeColor="Blue" Font-Size="9pt" BorderStyle="Groove" ReadOnly="True">DBEndTime</asp:textbox><asp:textbox id="DBStartTime" style="Z-INDEX: 133; LEFT: 168px; POSITION: absolute; TOP: 744px"
					runat="server" Width="152px" Height="20px" BackColor="Gold" ForeColor="Blue" Font-Size="9pt" BorderStyle="Groove" ReadOnly="True">DBStartTime</asp:textbox><asp:textbox id="DDecideDesc" style="Z-INDEX: 132; LEFT: 56px; POSITION: absolute; TOP: 664px"
					runat="server" Width="536px" Height="56px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" TextMode="MultiLine">DDecideDesc</asp:textbox><asp:textbox id="DSliderPrice" style="Z-INDEX: 129; LEFT: 648px; POSITION: absolute; TOP: 472px"
					runat="server" Width="112px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">SliderPrice</asp:textbox><asp:textbox id="DDevReason" style="Z-INDEX: 126; LEFT: 16px; POSITION: absolute; TOP: 384px"
					runat="server" Width="744px" Height="80px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" TextMode="MultiLine">DevReason</asp:textbox><asp:textbox id="DSellVendor" style="Z-INDEX: 121; LEFT: 584px; POSITION: absolute; TOP: 296px"
					runat="server" Width="174px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DSellVendor</asp:textbox><asp:dropdownlist id="DManufPlace" style="Z-INDEX: 119; LEFT: 584px; POSITION: absolute; TOP: 328px"
					runat="server" Width="176px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DSliderType" style="Z-INDEX: 117; LEFT: 312px; POSITION: absolute; TOP: 328px"
					runat="server" Width="176px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><INPUT id="DRefFile" style="Z-INDEX: 114; LEFT: 120px; WIDTH: 640px; POSITION: absolute; TOP: 624px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					type="file" size="87" name="File1" runat="server"><asp:button id="BSpec" style="Z-INDEX: 111; LEFT: 736px; POSITION: absolute; TOP: 259px" runat="server"
					Width="20px" Height="20px" Text="....."></asp:button><asp:textbox id="DSpec" style="Z-INDEX: 110; LEFT: 312px; POSITION: absolute; TOP: 259px" runat="server"
					Width="424px" Height="20px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove" ReadOnly="True">DSpec</asp:textbox><asp:textbox id="DSliderGRCode" style="Z-INDEX: 109; LEFT: 604px; POSITION: absolute; TOP: 158px"
					runat="server" Width="154px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True">DSliderGRCode</asp:textbox><asp:textbox id="DNo" style="Z-INDEX: 108; LEFT: 312px; POSITION: absolute; TOP: 192px" runat="server"
					Width="176px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DNo</asp:textbox><asp:textbox id="DFormSno" style="Z-INDEX: 107; LEFT: 16px; POSITION: absolute; TOP: 944px" runat="server"
					Width="97px" Height="20px" BackColor="White" ForeColor="Blue" BorderStyle="None">單號：123</asp:textbox><asp:image id="DDelivery" style="Z-INDEX: 105; LEFT: 8px; POSITION: absolute; TOP: 728px" runat="server"
					Width="593px" Height="110px" ImageUrl="Images\Sheet_Delivery.jpg"></asp:image><asp:image id="DDelay" style="Z-INDEX: 103; LEFT: 8px; POSITION: absolute; TOP: 840px" runat="server"
					Width="593px" Height="107px" ImageUrl="Images\Sheet_Delay.jpg" Visible="False"></asp:image><asp:image id="DDescSheet" style="Z-INDEX: 102; LEFT: 8px; POSITION: absolute; TOP: 656px"
					runat="server" Width="593px" Height="75px" ImageUrl="Images\MapSheet_004.jpg"></asp:image><asp:image id="LMapFile" style="Z-INDEX: 245; LEFT: 16px; POSITION: absolute; TOP: 120px" runat="server"
					Width="200px" Height="230px" BorderStyle="Groove" ImageUrl="http://10.245.0.178//WorkFlow/Document/000003/1-Map2-2006124111750-2.jpg"></asp:image></FONT><INPUT id="BSAVE" style="Z-INDEX: 249; LEFT: 256px; WIDTH: 80px; POSITION: absolute; TOP: 952px; HEIGHT: 28px"
				onclick="Button('SAVE');" type="button" value="儲存" name="Button1" runat="server"><INPUT id="BNG2" style="Z-INDEX: 250; LEFT: 344px; WIDTH: 80px; POSITION: absolute; TOP: 952px; HEIGHT: 28px"
				onclick="Button('NG2');" type="button" value="NG2" name="Button1" runat="server"><INPUT id="BNG1" style="Z-INDEX: 251; LEFT: 432px; WIDTH: 80px; POSITION: absolute; TOP: 952px; HEIGHT: 28px"
				onclick="Button('NG1');" type="button" value="NG1" name="Button1" runat="server"><INPUT id="BOK" style="Z-INDEX: 252; LEFT: 520px; WIDTH: 80px; POSITION: absolute; TOP: 952px; HEIGHT: 28px"
				onclick="Button('OK');" type="button" value="OK" name="Button2" runat="server">
			<asp:imagebutton id="BPrint" style="Z-INDEX: 262; LEFT: 8px; POSITION: absolute; TOP: 8px" runat="server"
				Width="16px" Height="16px" ImageUrl="Images\Print.gif"></asp:imagebutton></form>
		</FONT></FORM>
	</body>
</HTML>
