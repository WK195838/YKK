<%@ Page Language="vb" AutoEventWireup="false" Codebehind="ImportModSheet_01.aspx.vb" Inherits="SPD.ImportModSheet_01" aspCompat="True"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>ImportModSheet_01</title>
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
			function SpecPicker(strField) {
			    val=0;
				wPop=window.open('SpecPicker.aspx?field=' + strField,'SpecPopup','width=330,height=250,resizable=yes');
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
				else{alert("No");}
			}
		    function OpenPrintSheet(URL) {
				window.open(URL,'PrintSheet','width=630,height=580,resizable=yes,scrollbars=yes, menubar=yes');
			}			
		</script>
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" encType="multipart/form-data" runat="server">
			<FONT face="新細明體">
				<asp:image id="DImportSheet" style="Z-INDEX: 100; POSITION: absolute; TOP: 8px; LEFT: 8px"
					runat="server" Height="644px" Width="761px" ImageUrl="Images\ImportSheet_001.jpg"></asp:image><asp:label id="Label22" style="Z-INDEX: 291; POSITION: absolute; TOP: 506px; LEFT: 504px" runat="server"
					Height="10px" Width="24px" Font-Size="8pt">單價</asp:label><asp:label id="Label21" style="Z-INDEX: 290; POSITION: absolute; TOP: 506px; LEFT: 464px" runat="server"
					Height="10px" Width="24px" Font-Size="8pt">幣別</asp:label><asp:label id="Label20" style="Z-INDEX: 289; POSITION: absolute; TOP: 506px; LEFT: 296px" runat="server"
					Height="10px" Width="48px" Font-Size="8pt">拉頭種類</asp:label><asp:label id="Label19" style="Z-INDEX: 288; POSITION: absolute; TOP: 506px; LEFT: 224px" runat="server"
					Height="10px" Width="24px" Font-Size="8pt">單價</asp:label><asp:label id="Label18" style="Z-INDEX: 287; POSITION: absolute; TOP: 506px; LEFT: 184px" runat="server"
					Height="10px" Width="24px" Font-Size="8pt">幣別</asp:label><asp:label id="Label17" style="Z-INDEX: 286; POSITION: absolute; TOP: 506px; LEFT: 56px" runat="server"
					Height="10px" Width="48px" Font-Size="8pt">拉頭種類</asp:label><asp:dropdownlist id="DBuyer" style="Z-INDEX: 270; POSITION: absolute; TOP: 296px; LEFT: 312px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="20px" Width="176px">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:imagebutton id="BFlow" style="Z-INDEX: 263; POSITION: absolute; TOP: 32px; LEFT: 8px" runat="server"
					Height="16px" Width="16px" ImageUrl="Images\Flow-01.gif"></asp:imagebutton><INPUT id="DMapFile" style="Z-INDEX: 246; POSITION: absolute; BACKGROUND-COLOR: #ffff00; WIDTH: 200px; HEIGHT: 20px; TOP: 328px; LEFT: 16px"
					type="file" size="14" name="File1" runat="server"><asp:textbox id="DPriceJ1" style="Z-INDEX: 202; POSITION: absolute; TOP: 590px; LEFT: 296px"
					runat="server" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Height="18px" Width="166px" Font-Size="8pt"></asp:textbox><asp:textbox id="DPriceI3" style="Z-INDEX: 201; POSITION: absolute; TOP: 590px; LEFT: 224px"
					runat="server" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Height="18px" Width="64px" Font-Size="8pt"></asp:textbox><asp:dropdownlist id="DPriceI2" style="Z-INDEX: 200; POSITION: absolute; TOP: 590px; LEFT: 184px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="12px" Width="40px" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceI1" style="Z-INDEX: 199; POSITION: absolute; TOP: 590px; LEFT: 16px" runat="server"
					BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Height="18px" Width="166px" Font-Size="8pt"></asp:textbox><asp:textbox id="DPriceH3" style="Z-INDEX: 198; POSITION: absolute; TOP: 572px; LEFT: 504px"
					runat="server" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Height="18px" Width="64px" Font-Size="8pt"></asp:textbox><asp:dropdownlist id="DPriceH2" style="Z-INDEX: 197; POSITION: absolute; TOP: 572px; LEFT: 464px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="12px" Width="40px" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceJ3" style="Z-INDEX: 204; POSITION: absolute; TOP: 590px; LEFT: 504px"
					runat="server" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Height="18px" Width="64px" Font-Size="8pt"></asp:textbox><asp:dropdownlist id="DPriceJ2" style="Z-INDEX: 203; POSITION: absolute; TOP: 590px; LEFT: 464px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="12px" Width="40px" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceA3" style="Z-INDEX: 150; POSITION: absolute; TOP: 518px; LEFT: 224px"
					runat="server" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Height="18px" Width="64px" Font-Size="8pt"></asp:textbox><asp:dropdownlist id="DPriceA2" style="Z-INDEX: 149; POSITION: absolute; TOP: 518px; LEFT: 184px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="12px" Width="40px" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceH1" style="Z-INDEX: 196; POSITION: absolute; TOP: 572px; LEFT: 296px"
					runat="server" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Height="18px" Width="166px" Font-Size="8pt"></asp:textbox><asp:textbox id="DPriceG3" style="Z-INDEX: 195; POSITION: absolute; TOP: 572px; LEFT: 224px"
					runat="server" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Height="18px" Width="64px" Font-Size="8pt"></asp:textbox><asp:dropdownlist id="DPriceG2" style="Z-INDEX: 194; POSITION: absolute; TOP: 572px; LEFT: 184px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="12px" Width="40px" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceG1" style="Z-INDEX: 193; POSITION: absolute; TOP: 572px; LEFT: 16px" runat="server"
					BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Height="18px" Width="166px" Font-Size="8pt"></asp:textbox><asp:textbox id="DPriceF3" style="Z-INDEX: 192; POSITION: absolute; TOP: 554px; LEFT: 504px"
					runat="server" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Height="18px" Width="64px" Font-Size="8pt"></asp:textbox><asp:dropdownlist id="DPriceF2" style="Z-INDEX: 191; POSITION: absolute; TOP: 554px; LEFT: 464px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="12px" Width="40px" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceF1" style="Z-INDEX: 190; POSITION: absolute; TOP: 554px; LEFT: 296px"
					runat="server" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Height="18px" Width="166px" Font-Size="8pt"></asp:textbox><asp:textbox id="DPriceE3" style="Z-INDEX: 189; POSITION: absolute; TOP: 554px; LEFT: 224px"
					runat="server" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Height="18px" Width="64px" Font-Size="8pt"></asp:textbox><asp:dropdownlist id="DPriceE2" style="Z-INDEX: 188; POSITION: absolute; TOP: 554px; LEFT: 184px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="12px" Width="40px" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceE1" style="Z-INDEX: 187; POSITION: absolute; TOP: 554px; LEFT: 16px" runat="server"
					BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Height="18px" Width="166px" Font-Size="8pt"></asp:textbox><asp:textbox id="DPriceD3" style="Z-INDEX: 186; POSITION: absolute; TOP: 536px; LEFT: 504px"
					runat="server" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Height="18px" Width="64px" Font-Size="8pt"></asp:textbox><asp:dropdownlist id="DPriceD2" style="Z-INDEX: 185; POSITION: absolute; TOP: 536px; LEFT: 464px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="12px" Width="40px" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceD1" style="Z-INDEX: 184; POSITION: absolute; TOP: 536px; LEFT: 296px"
					runat="server" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Height="18px" Width="166px" Font-Size="8pt"></asp:textbox><asp:textbox id="DPriceC3" style="Z-INDEX: 183; POSITION: absolute; TOP: 536px; LEFT: 224px"
					runat="server" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Height="18px" Width="64px" Font-Size="8pt"></asp:textbox><asp:dropdownlist id="DPriceC2" style="Z-INDEX: 182; POSITION: absolute; TOP: 536px; LEFT: 184px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="12px" Width="40px" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceC1" style="Z-INDEX: 181; POSITION: absolute; TOP: 536px; LEFT: 16px" runat="server"
					BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Height="18px" Width="166px" Font-Size="8pt"></asp:textbox><asp:textbox id="DPriceB3" style="Z-INDEX: 180; POSITION: absolute; TOP: 518px; LEFT: 504px"
					runat="server" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Height="18px" Width="64px" Font-Size="8pt"></asp:textbox><asp:dropdownlist id="DPriceB2" style="Z-INDEX: 179; POSITION: absolute; TOP: 518px; LEFT: 464px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="12px" Width="40px" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceB1" style="Z-INDEX: 178; POSITION: absolute; TOP: 518px; LEFT: 296px"
					runat="server" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Height="18px" Width="166px" Font-Size="8pt"></asp:textbox><asp:textbox id="DPriceA1" style="Z-INDEX: 169; POSITION: absolute; TOP: 518px; LEFT: 16px" runat="server"
					BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Height="18px" Width="166px" Font-Size="8pt"></asp:textbox><asp:button id="BSliderCode" style="Z-INDEX: 167; POSITION: absolute; TOP: 120px; LEFT: 568px"
					runat="server" Height="20px" Width="20px" Text="....."></asp:button>
				<asp:hyperlink id="LRefFile" style="Z-INDEX: 164; POSITION: absolute; TOP: 624px; LEFT: 120px"
					runat="server" Height="8px" Width="64px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">其他文件</asp:hyperlink><asp:button id="BDate" style="Z-INDEX: 148; POSITION: absolute; TOP: 192px; LEFT: 736px" runat="server"
					Height="20px" Width="20px" Text="....."></asp:button><asp:textbox id="DDate" style="Z-INDEX: 147; POSITION: absolute; TOP: 192px; LEFT: 584px" runat="server"
					BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="152px" ReadOnly="True">DDate</asp:textbox><asp:textbox id="DSliderCode" style="Z-INDEX: 146; POSITION: absolute; TOP: 120px; LEFT: 224px"
					runat="server" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Height="56px" Width="342px" ReadOnly="True" Font-Size="8pt" TextMode="MultiLine">SliderCode</asp:textbox><asp:dropdownlist id="DPerson" style="Z-INDEX: 142; POSITION: absolute; TOP: 226px; LEFT: 584px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="20px" Width="176px">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DDivision" style="Z-INDEX: 141; POSITION: absolute; TOP: 226px; LEFT: 312px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="176px">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DSliderPrice" style="Z-INDEX: 129; POSITION: absolute; TOP: 476px; LEFT: 648px"
					runat="server" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="112px">SliderPrice</asp:textbox><asp:textbox id="DDevReason" style="Z-INDEX: 126; POSITION: absolute; TOP: 384px; LEFT: 16px"
					runat="server" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Height="80px" Width="744px" TextMode="MultiLine">DevReason</asp:textbox><asp:textbox id="DSellVendor" style="Z-INDEX: 121; POSITION: absolute; TOP: 296px; LEFT: 584px"
					runat="server" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="174px">DSellVendor</asp:textbox><asp:dropdownlist id="DManufPlace" style="Z-INDEX: 119; POSITION: absolute; TOP: 328px; LEFT: 584px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="176px">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DSliderType" style="Z-INDEX: 117; POSITION: absolute; TOP: 328px; LEFT: 312px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="176px">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><INPUT id="DRefFile" style="Z-INDEX: 114; POSITION: absolute; BACKGROUND-COLOR: #ffff00; WIDTH: 640px; HEIGHT: 20px; TOP: 622px; LEFT: 120px"
					type="file" size="87" name="File1" runat="server"><asp:button id="BSpec" style="Z-INDEX: 111; POSITION: absolute; TOP: 259px; LEFT: 736px" runat="server"
					Height="20px" Width="20px" Text="....."></asp:button><asp:textbox id="DSpec" style="Z-INDEX: 110; POSITION: absolute; TOP: 259px; LEFT: 312px" runat="server"
					BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="424px" ReadOnly="True" Font-Size="8pt">DSpec</asp:textbox><asp:textbox id="DSliderGRCode" style="Z-INDEX: 109; POSITION: absolute; TOP: 158px; LEFT: 604px"
					runat="server" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="154px" ReadOnly="True">DSliderGRCode</asp:textbox><asp:textbox id="DNo" style="Z-INDEX: 108; POSITION: absolute; TOP: 192px; LEFT: 312px" runat="server"
					BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="176px">DNo</asp:textbox><asp:textbox id="DFormSno" style="Z-INDEX: 107; POSITION: absolute; TOP: 648px; LEFT: 16px" runat="server"
					BorderStyle="None" ForeColor="Blue" BackColor="White" Height="20px" Width="97px">單號：123</asp:textbox><asp:image id="LMapFile" style="Z-INDEX: 245; POSITION: absolute; TOP: 120px; LEFT: 16px" runat="server"
					BorderStyle="Groove" Height="230px" Width="200px" ImageUrl="http://10.245.0.178//WorkFlow/Document/000003/1-Map2-2006124111750-2.jpg"></asp:image></FONT>
			<asp:imagebutton id="BPrint" style="Z-INDEX: 262; POSITION: absolute; TOP: 8px; LEFT: 8px" runat="server"
				Height="16px" Width="16px" ImageUrl="Images\Print.gif"></asp:imagebutton></FONT><INPUT id="BOK" style="Z-INDEX: 295; POSITION: absolute; WIDTH: 80px; HEIGHT: 28px; TOP: 656px; LEFT: 600px"
				onclick="Button('OK');" type="button" value="閱讀完成" name="BOK" runat="server"><INPUT id="BSAVE" style="Z-INDEX: 296; POSITION: absolute; WIDTH: 80px; HEIGHT: 28px; TOP: 656px; LEFT: 688px"
				onclick="Button('SAVE');" type="button" value="儲存" name="Button1" runat="server">
		</form>
	</body>
</HTML>
