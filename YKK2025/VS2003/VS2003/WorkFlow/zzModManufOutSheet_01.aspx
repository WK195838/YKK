<%@ Page Language="vb" AutoEventWireup="false" Codebehind="zzModManufOutSheet_01.aspx.vb" Inherits="SPD.ModManufOutSheet_01" aspCompat="True"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>ModManufOutSheet_01</title>
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
			//--Sample------------------------------------
			function SamplePicker(strField) {
				wPop=window.open('SamplePicker.aspx?field=' + strField,'SamplePopup','width=370,height=250,resizable=yes');
				if (document.Form1.DSample.value != "") {
					setTimeout("SendToSample()",200);
				}
			}
		    function SendToSample() {
				wPop.document.SampleForm.DContent.value=document.Form1.DSample.value;
			}
			//--Price------------------------------------
			function PricePicker(strField) {
				wPop=window.open('PricePicker.aspx?field=' + strField,'PricePopup','width=310,height=250,resizable=yes');
				if (document.Form1.DPrice.value != "") {
					setTimeout("SendToPrice()",200);
				}
			}
		    function SendToPrice() {
				wPop.document.PriceForm.DContent.value=document.Form1.DPrice.value;
			}
			//--Quality-1---------------------------------
			function QA1Picker(strField) {
				wPop=window.open('QA1Picker.aspx?field=' + strField,'QA1Popup','width=690,height=320,resizable=yes');
				if (document.Form1.DQuality1.value != "") {
					setTimeout("SendToQuality1()",200);
				}
			}
		    function SendToQuality1() {
				wPop.document.QAForm1.DContent.value=document.Form1.DQuality1.value;
			}
			//--Quality-2---------------------------------
			function QA2Picker(strField) {
				wPop=window.open('QA2Picker.aspx?field=' + strField,'QA2Popup','width=605,height=250,resizable=yes');
				if (document.Form1.DQuality2.value != "") {
					setTimeout("SendToQuality2()",200);
				}
			}
		    function SendToQuality2() {
				wPop.document.QAForm2.DContent.value=document.Form1.DQuality2.value;
			}
		</script>
		
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<FORM id="Form1" method="post" encType="multipart/form-data" runat="server">
			<FONT face="新細明體">
				<asp:image id="DManuaInSheet1" style="Z-INDEX: 100; LEFT: 8px; POSITION: absolute; TOP: 8px"
					runat="server" Height="670px" Width="593px" ImageUrl="Images\ManufOutSheet_001.jpg"></asp:image>
				<asp:textbox id="wQAAttachFile" style="Z-INDEX: 164; LEFT: 64px; POSITION: absolute; TOP: 56px"
					runat="server" Height="20px" Width="160px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow"
					ReadOnly="True" Visible="False">wQAAttachFile</asp:textbox>
				<asp:textbox id="wSampleFile" style="Z-INDEX: 163; LEFT: 56px; POSITION: absolute; TOP: 48px"
					runat="server" Height="20px" Width="160px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow"
					ReadOnly="True" Visible="False">wSampleFile</asp:textbox>
				<asp:textbox id="wAuthorizeFile" style="Z-INDEX: 162; LEFT: 48px; POSITION: absolute; TOP: 40px"
					runat="server" Height="20px" Width="160px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow"
					ReadOnly="True" Visible="False">wAuthorizeFile</asp:textbox>
				<asp:textbox id="wConfirmFile" style="Z-INDEX: 161; LEFT: 40px; POSITION: absolute; TOP: 32px"
					runat="server" Height="20px" Width="160px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow"
					ReadOnly="True" Visible="False">wConfirmFile</asp:textbox>
				<asp:textbox id="wMapFile2" style="Z-INDEX: 160; LEFT: 32px; POSITION: absolute; TOP: 24px" runat="server"
					Height="20px" Width="160px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" ReadOnly="True"
					Visible="False">wMapFile2</asp:textbox>
				<asp:textbox id="wMapFile1" style="Z-INDEX: 159; LEFT: 24px; POSITION: absolute; TOP: 16px" runat="server"
					Height="20px" Width="160px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" ReadOnly="True"
					Visible="False">wMapFile1</asp:textbox>
				<asp:button id="BDate" style="Z-INDEX: 158; LEFT: 568px; POSITION: absolute; TOP: 88px" runat="server"
					Height="20px" Width="20px" Text="....."></asp:button>
				<asp:textbox id="DDate" style="Z-INDEX: 157; LEFT: 416px; POSITION: absolute; TOP: 88px" runat="server"
					Height="20px" Width="152px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" ReadOnly="True">DDate</asp:textbox>
				<asp:hyperlink id="LQuality2" style="Z-INDEX: 156; LEFT: 552px; POSITION: absolute; TOP: 1048px"
					runat="server" Height="8px" Width="40px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">詳細</asp:hyperlink>
				<asp:hyperlink id="LQuality1" style="Z-INDEX: 155; LEFT: 552px; POSITION: absolute; TOP: 976px"
					runat="server" Height="8px" Width="40px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">詳細</asp:hyperlink>
				<asp:hyperlink id="LPrice" style="Z-INDEX: 154; LEFT: 552px; POSITION: absolute; TOP: 752px" runat="server"
					Height="8px" Width="40px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">詳細</asp:hyperlink>
				<asp:hyperlink id="LSample" style="Z-INDEX: 153; LEFT: 552px; POSITION: absolute; TOP: 504px" runat="server"
					Height="8px" Width="32px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">詳細</asp:hyperlink>
				<asp:hyperlink id="LSliderCode" style="Z-INDEX: 152; LEFT: 552px; POSITION: absolute; TOP: 192px"
					runat="server" Height="8px" Width="32px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">詳細</asp:hyperlink>
				<asp:textbox id="DQuality2" style="Z-INDEX: 151; LEFT: 120px; POSITION: absolute; TOP: 1016px"
					runat="server" Height="58px" Width="424px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow"
					ReadOnly="True" TextMode="MultiLine">Quality2</asp:textbox>
				<asp:textbox id="DQuality1" style="Z-INDEX: 150; LEFT: 120px; POSITION: absolute; TOP: 948px"
					runat="server" Height="58px" Width="424px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow"
					ReadOnly="True" TextMode="MultiLine">Quality1</asp:textbox>
				<asp:textbox id="DPrice" style="Z-INDEX: 149; LEFT: 120px; POSITION: absolute; TOP: 722px" runat="server"
					Height="58px" Width="424px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" ReadOnly="True"
					TextMode="MultiLine">Price</asp:textbox>
				<asp:textbox id="DSample" style="Z-INDEX: 148; LEFT: 120px; POSITION: absolute; TOP: 472px" runat="server"
					Height="58px" Width="424px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" ReadOnly="True"
					TextMode="MultiLine">Sample</asp:textbox>
				<asp:textbox id="DSliderCode" style="Z-INDEX: 147; LEFT: 120px; POSITION: absolute; TOP: 176px"
					runat="server" Height="38px" Width="424px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow"
					ReadOnly="True" TextMode="MultiLine">SliderCode</asp:textbox>
				<asp:button id="BMMapNo" style="Z-INDEX: 146; LEFT: 570px; POSITION: absolute; TOP: 234px" runat="server"
					Height="20px" Width="20px" Text="修"></asp:button>
				<asp:hyperlink id="LSampleFile" style="Z-INDEX: 145; LEFT: 536px; POSITION: absolute; TOP: 648px"
					runat="server" Height="8px" Width="50px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">樣品圖</asp:hyperlink><INPUT id="DSampleFile" style="Z-INDEX: 144; LEFT: 120px; WIDTH: 408px; POSITION: absolute; TOP: 648px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					type="file" size="48" name="File1" runat="server">
				<asp:dropdownlist id="DPerson" style="Z-INDEX: 143; LEFT: 416px; POSITION: absolute; TOP: 122px" runat="server"
					Height="20px" Width="176px" ForeColor="Blue" BackColor="Yellow" AutoPostBack="True">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DDivision" style="Z-INDEX: 142; LEFT: 120px; POSITION: absolute; TOP: 122px"
					runat="server" Height="20px" Width="176px" ForeColor="Blue" BackColor="Yellow" AutoPostBack="True">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist>
				<asp:hyperlink id="LQAAttachFile" style="Z-INDEX: 141; LEFT: 552px; POSITION: absolute; TOP: 1088px"
					runat="server" Height="8px" Width="40px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">附件</asp:hyperlink><INPUT id="DQAAttachFile" style="Z-INDEX: 140; LEFT: 120px; WIDTH: 424px; POSITION: absolute; TOP: 1090px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					type="file" size="51" name="File1" runat="server">
				<asp:button id="BQuality2" style="Z-INDEX: 139; LEFT: 544px; POSITION: absolute; TOP: 1016px"
					runat="server" Height="20px" Width="20px" Text="....."></asp:button>
				<asp:button id="BQuality1" style="Z-INDEX: 138; LEFT: 544px; POSITION: absolute; TOP: 948px"
					runat="server" Height="20px" Width="20px" Text="....."></asp:button>
				<asp:textbox id="DMoldPoint" style="Z-INDEX: 137; LEFT: 512px; POSITION: absolute; TOP: 872px"
					runat="server" Height="20px" Width="48px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow">MoldPoint</asp:textbox>
				<asp:textbox id="DMoldQty" style="Z-INDEX: 136; LEFT: 416px; POSITION: absolute; TOP: 872px"
					runat="server" Height="20px" Width="56px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow">DMoldQty</asp:textbox>
				<asp:textbox id="DMoldName" style="Z-INDEX: 135; LEFT: 120px; POSITION: absolute; TOP: 872px"
					runat="server" Height="20px" Width="176px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow">MoldName</asp:textbox>
				<asp:textbox id="DPullerPrice" style="Z-INDEX: 134; LEFT: 120px; POSITION: absolute; TOP: 828px"
					runat="server" Height="20px" Width="472px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow">PullerPrice</asp:textbox>
				<asp:textbox id="DPurMold" style="Z-INDEX: 133; LEFT: 416px; POSITION: absolute; TOP: 794px"
					runat="server" Height="20px" Width="176px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow">PurMold</asp:textbox>
				<asp:textbox id="DArMoldFee" style="Z-INDEX: 132; LEFT: 120px; POSITION: absolute; TOP: 794px"
					runat="server" Height="20px" Width="176px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow">ArMoldFee</asp:textbox>
				<asp:button id="BPrice" style="Z-INDEX: 131; LEFT: 544px; POSITION: absolute; TOP: 724px" runat="server"
					Height="20px" Width="20px" Text="....."></asp:button>
				<asp:button id="BSample" style="Z-INDEX: 129; LEFT: 544px; POSITION: absolute; TOP: 472px" runat="server"
					Height="20px" Width="20px" Text="....."></asp:button>
				<asp:button id="BSliderCode" style="Z-INDEX: 130; LEFT: 570px; POSITION: absolute; TOP: 154px"
					runat="server" Height="20px" Width="20px" Text="....."></asp:button>
				<asp:textbox id="DDevReason" style="Z-INDEX: 128; LEFT: 120px; POSITION: absolute; TOP: 404px"
					runat="server" Height="56px" Width="472px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow"
					TextMode="MultiLine">DevReason</asp:textbox>
				<asp:hyperlink id="LAuthorizeFile" style="Z-INDEX: 127; LEFT: 536px; POSITION: absolute; TOP: 612px"
					runat="server" Height="8px" Width="50px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">授權書</asp:hyperlink><INPUT id="DAuthorizeFile" style="Z-INDEX: 126; LEFT: 120px; WIDTH: 408px; POSITION: absolute; TOP: 612px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					type="file" size="48" name="File1" runat="server">
				<asp:hyperlink id="LConfirmFile" style="Z-INDEX: 125; LEFT: 536px; POSITION: absolute; TOP: 578px"
					runat="server" Height="8px" Width="50px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">確認書</asp:hyperlink><INPUT id="DConfirmFile" style="Z-INDEX: 124; LEFT: 120px; WIDTH: 408px; POSITION: absolute; TOP: 578px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					type="file" size="48" name="File1" runat="server">
				<asp:textbox id="DBuyer" style="Z-INDEX: 123; LEFT: 416px; POSITION: absolute; TOP: 371px" runat="server"
					Height="20px" Width="176px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow">DBuyer</asp:textbox>
				<asp:textbox id="DSellVendor" style="Z-INDEX: 122; LEFT: 120px; POSITION: absolute; TOP: 371px"
					runat="server" Height="20px" Width="176px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow">DSellVendor</asp:textbox>
				<asp:textbox id="DMaterialOther" style="Z-INDEX: 121; LEFT: 440px; POSITION: absolute; TOP: 336px"
					runat="server" Height="20px" Width="152px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow">DMaterialOther</asp:textbox>
				<asp:dropdownlist id="DMaterial" style="Z-INDEX: 120; LEFT: 120px; POSITION: absolute; TOP: 336px"
					runat="server" Height="20px" Width="120px" ForeColor="Blue" BackColor="Yellow" AutoPostBack="True">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DMaterialDetail" style="Z-INDEX: 110; LEFT: 240px; POSITION: absolute; TOP: 336px"
					runat="server" Height="20px" Width="200px" ForeColor="Blue" BackColor="Yellow" AutoPostBack="True">
					<asp:ListItem Value="3">CF</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DManufPlace" style="Z-INDEX: 119; LEFT: 416px; POSITION: absolute; TOP: 302px"
					runat="server" Height="20px" Width="176px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DSliderType2" style="Z-INDEX: 118; LEFT: 208px; POSITION: absolute; TOP: 302px"
					runat="server" Height="20px" Width="88px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DSliderType1" style="Z-INDEX: 117; LEFT: 120px; POSITION: absolute; TOP: 302px"
					runat="server" Height="20px" Width="86px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist>
				<asp:textbox id="DAssembler" style="Z-INDEX: 116; LEFT: 416px; POSITION: absolute; TOP: 268px"
					runat="server" Height="20px" Width="176px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow">DAssembler</asp:textbox>
				<asp:dropdownlist id="DLevel" style="Z-INDEX: 115; LEFT: 120px; POSITION: absolute; TOP: 268px" runat="server"
					Height="20px" Width="176px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist>
				<asp:hyperlink id="LMapFile2" style="Z-INDEX: 114; LEFT: 552px; POSITION: absolute; TOP: 544px"
					runat="server" Height="8px" Width="40px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">附圖2</asp:hyperlink><INPUT id="DMapFile2" style="Z-INDEX: 113; LEFT: 360px; WIDTH: 184px; POSITION: absolute; TOP: 544px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					type="file" size="11" name="File1" runat="server"><INPUT id="DMapFile1" style="Z-INDEX: 112; LEFT: 120px; WIDTH: 184px; POSITION: absolute; TOP: 544px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					type="file" size="11" name="File1" runat="server">
				<asp:hyperlink id="LMapFile1" style="Z-INDEX: 106; LEFT: 312px; POSITION: absolute; TOP: 544px"
					runat="server" Height="8px" Width="40px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">附圖1</asp:hyperlink>
				<asp:button id="BOMapNo" style="Z-INDEX: 111; LEFT: 548px; POSITION: absolute; TOP: 234px" runat="server"
					Height="20px" Width="20px" Text="原"></asp:button>
				<asp:textbox id="DMapNo" style="Z-INDEX: 109; LEFT: 416px; POSITION: absolute; TOP: 234px" runat="server"
					Height="20px" Width="132px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" AutoPostBack="True">DMapNo</asp:textbox>
				<asp:button id="BSpec" style="Z-INDEX: 108; LEFT: 280px; POSITION: absolute; TOP: 234px" runat="server"
					Height="20px" Width="20px" Text="....."></asp:button>
				<asp:textbox id="DSpec" style="Z-INDEX: 107; LEFT: 120px; POSITION: absolute; TOP: 234px" runat="server"
					Height="20px" Width="160px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" ReadOnly="True">DSpec</asp:textbox>
				<asp:textbox id="DSliderGRCode" style="Z-INDEX: 105; LEFT: 120px; POSITION: absolute; TOP: 154px"
					runat="server" Height="20px" Width="450px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow"
					ReadOnly="True">DSliderGRCode</asp:textbox>
				<asp:textbox id="DNo" style="Z-INDEX: 104; LEFT: 120px; POSITION: absolute; TOP: 88px" runat="server"
					Height="20px" Width="176px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow">DNo</asp:textbox>
				<asp:button id="BSave" style="Z-INDEX: 103; LEFT: 496px; POSITION: absolute; TOP: 1120px" runat="server"
					Height="32px" Width="105px" Text="儲存"></asp:button>
				<asp:textbox id="DFormSno" style="Z-INDEX: 102; LEFT: 16px; POSITION: absolute; TOP: 1120px"
					runat="server" Height="20px" Width="97px" BorderStyle="None" ForeColor="Blue" BackColor="White">單號：123</asp:textbox>
				<asp:image id="DManuaInSheet2" style="Z-INDEX: 101; LEFT: 8px; POSITION: absolute; TOP: 680px"
					runat="server" Height="440px" Width="593px" ImageUrl="Images\ManufInSheet_002.jpg"></asp:image></FONT></FORM>
	</body>
</HTML>
