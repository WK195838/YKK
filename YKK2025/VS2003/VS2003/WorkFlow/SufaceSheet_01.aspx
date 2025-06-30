<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SufaceSheet_01.aspx.vb" Inherits="SPD.SufaceSheet_01" aspCompat="True"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>表面處理委託書</title>
		<meta content="False" name="vs_snapToGrid">
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
			//--Spec------------------------------------
			function SpecPicker(strField, Fun) {
			    val=0;
				wPop=window.open('SpecPicker.aspx?field=' + strField + '&fun=' + Fun, 'SpecPopup','width=330,height=250,resizable=yes');
				if (document.Form1.DSpec.value != "") {
					setTimeout("SendToSpec()",200);
				}
			}
			function DevNoPicker(strField,FormNo) {
				window.open('DevNoPicker.aspx?field=' + strField + '&pFormNo=' + FormNo,'DevNoPopup','width=250,height=360,resizable=yes');
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
				<asp:textbox id="DOrderTime" style="Z-INDEX: 162; LEFT: 584px; POSITION: absolute; TOP: 288px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="150px" BorderStyle="Groove"
					ReadOnly="True">DOrderTime</asp:textbox><asp:hyperlink id="LPFASFile" style="Z-INDEX: 213; LEFT: 144px; POSITION: absolute; TOP: 936px"
					runat="server" Height="20px" Width="98px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">PFAS報告</asp:hyperlink><INPUT id="DPFASFile" style="Z-INDEX: 212; LEFT: 144px; WIDTH: 144px; POSITION: absolute; TOP: 936px; HEIGHT: 24px; BACKGROUND-COLOR: #ffff00"
					type="file" size="4" name="File1" runat="server">
				<asp:dropdownlist id="DQCCHECK16" style="Z-INDEX: 211; LEFT: 24px; POSITION: absolute; TOP: 936px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="568px" Width="112px">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DFReason" style="Z-INDEX: 210; LEFT: 509px; POSITION: absolute; TOP: 493px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="24px" Width="261px" BorderStyle="Groove">DFReason</asp:textbox><asp:dropdownlist id="DYearSeason" style="Z-INDEX: 209; LEFT: 586px; POSITION: absolute; TOP: 224px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="62px" Width="85px">
					<asp:ListItem></asp:ListItem>
					<asp:ListItem Value="年">年</asp:ListItem>
					<asp:ListItem Value="季">季</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQCCHECK15" style="Z-INDEX: 208; LEFT: 655px; POSITION: absolute; TOP: 734px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="107px">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DLOSS" style="Z-INDEX: 207; LEFT: 274px; POSITION: absolute; TOP: 867px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="20px" Width="110px" BorderStyle="Groove" MaxLength="10"></asp:textbox><asp:dropdownlist id="DQCCheck14" style="Z-INDEX: 206; LEFT: 146px; POSITION: absolute; TOP: 866px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="110px">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQCCheck13" style="Z-INDEX: 205; LEFT: 23px; POSITION: absolute; TOP: 867px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="110px">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:hyperlink id="LEACheckFile1" style="Z-INDEX: 204; LEFT: 624px; POSITION: absolute; TOP: 936px"
					runat="server" Height="12px" Width="72px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">A01-報告</asp:hyperlink><INPUT id="DEACheckFile1" style="Z-INDEX: 203; LEFT: 624px; WIDTH: 144px; POSITION: absolute; TOP: 936px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					type="file" size="4" name="File1" runat="server">
				<asp:textbox id="DSchedule" style="Z-INDEX: 202; LEFT: 673px; POSITION: absolute; TOP: 465px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="93px" BorderStyle="Groove">DSchedule</asp:textbox><asp:textbox id="DCap" style="Z-INDEX: 201; LEFT: 508px; POSITION: absolute; TOP: 464px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="20px" Width="75px" BorderStyle="Groove">DCap</asp:textbox><asp:textbox id="DOFormSno" style="Z-INDEX: 100; LEFT: 712px; POSITION: absolute; TOP: 160px"
					runat="server" Height="24px" Width="56px" ReadOnly="True"></asp:textbox><asp:button id="BImport" style="Z-INDEX: 200; LEFT: 475px; POSITION: absolute; TOP: 90px" runat="server"
					Height="20px" Width="20px" Text="進"></asp:button><asp:button id="BOut" style="Z-INDEX: 121; LEFT: 454px; POSITION: absolute; TOP: 90px" runat="server"
					Height="20px" Width="20px" Text="外"></asp:button><asp:button id="BIn" style="Z-INDEX: 117; LEFT: 433px; POSITION: absolute; TOP: 90px" runat="server"
					Height="20px" Width="20px" Text="內"></asp:button><asp:textbox id="DNo" style="Z-INDEX: 110; LEFT: 312px; POSITION: absolute; TOP: 90px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="20px" Width="121px" BorderStyle="Groove">DNo</asp:textbox><asp:textbox id="DPrice" style="Z-INDEX: 199; LEFT: 312px; POSITION: absolute; TOP: 320px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="20px" Width="180px" BorderStyle="Groove" MaxLength="20" AutoPostBack="True">DPrice</asp:textbox><asp:textbox id="DQCLT" style="Z-INDEX: 198; LEFT: 706px; POSITION: absolute; TOP: 664px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="18px" Width="53px" BorderStyle="Groove" Font-Size="8pt"></asp:textbox><asp:button id="BQCDate1" style="Z-INDEX: 197; LEFT: 96px; POSITION: absolute; TOP: 1024px"
					runat="server" Height="20px" Width="20px" Text="....."></asp:button><asp:button id="BQCDate2" style="Z-INDEX: 196; LEFT: 96px; POSITION: absolute; TOP: 1056px"
					runat="server" Height="20px" Width="20px" Text="....."></asp:button><asp:button id="BQCDate3" style="Z-INDEX: 195; LEFT: 96px; POSITION: absolute; TOP: 1088px"
					runat="server" Height="20px" Width="20px" Text="....."></asp:button><asp:dropdownlist id="DQCResult3" style="Z-INDEX: 194; LEFT: 120px; POSITION: absolute; TOP: 1088px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="78px">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQCResult2" style="Z-INDEX: 193; LEFT: 120px; POSITION: absolute; TOP: 1056px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="78px">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DQCRemark3" style="Z-INDEX: 192; LEFT: 208px; POSITION: absolute; TOP: 1088px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="554px" BorderStyle="Groove">DQCRemark3</asp:textbox><asp:textbox id="DQCRemark2" style="Z-INDEX: 191; LEFT: 208px; POSITION: absolute; TOP: 1056px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="554px" BorderStyle="Groove">DQCRemark2</asp:textbox><asp:textbox id="DQCDate3" style="Z-INDEX: 190; LEFT: 16px; POSITION: absolute; TOP: 1088px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="80px" BorderStyle="Groove">DQCDate3</asp:textbox><asp:textbox id="DQCDate2" style="Z-INDEX: 189; LEFT: 16px; POSITION: absolute; TOP: 1056px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="80px" BorderStyle="Groove">DQCDate2</asp:textbox><asp:textbox id="DQCRemark1" style="Z-INDEX: 188; LEFT: 208px; POSITION: absolute; TOP: 1024px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="554px" BorderStyle="Groove">DQCRemark1</asp:textbox><asp:dropdownlist id="DQCResult1" style="Z-INDEX: 187; LEFT: 120px; POSITION: absolute; TOP: 1024px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="78px">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DQCDate1" style="Z-INDEX: 186; LEFT: 16px; POSITION: absolute; TOP: 1024px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="80px" BorderStyle="Groove">DQCDate1</asp:textbox><asp:textbox id="DEADesc1" style="Z-INDEX: 185; LEFT: 520px; POSITION: absolute; TOP: 868px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="243px" BorderStyle="Groove">DEADesc1</asp:textbox><asp:dropdownlist id="DEACheck1" style="Z-INDEX: 184; LEFT: 399px; POSITION: absolute; TOP: 868px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="110px">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQCCheck12" style="Z-INDEX: 183; LEFT: 656px; POSITION: absolute; TOP: 800px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="110px">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQCCheck11" style="Z-INDEX: 182; LEFT: 528px; POSITION: absolute; TOP: 800px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="110px">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQCCheck10" style="Z-INDEX: 181; LEFT: 400px; POSITION: absolute; TOP: 800px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="110px">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQCCheck9" style="Z-INDEX: 180; LEFT: 272px; POSITION: absolute; TOP: 800px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="110px">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQCCheck8" style="Z-INDEX: 179; LEFT: 144px; POSITION: absolute; TOP: 800px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="110px">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQCCheck7" style="Z-INDEX: 178; LEFT: 24px; POSITION: absolute; TOP: 800px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="110px">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQCCheck5" style="Z-INDEX: 177; LEFT: 528px; POSITION: absolute; TOP: 736px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="110px">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQCCheck4" style="Z-INDEX: 176; LEFT: 400px; POSITION: absolute; TOP: 736px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="110px">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQCCheck3" style="Z-INDEX: 175; LEFT: 272px; POSITION: absolute; TOP: 736px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="110px">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQCCheck2" style="Z-INDEX: 174; LEFT: 144px; POSITION: absolute; TOP: 736px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="110px">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQCCheck1" style="Z-INDEX: 173; LEFT: 24px; POSITION: absolute; TOP: 736px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="110px">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><INPUT id="DFinalSampleFile" style="Z-INDEX: 172; LEFT: 16px; WIDTH: 200px; POSITION: absolute; TOP: 664px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					type="file" size="14" name="File1" runat="server">
				<asp:image id="LFinalSampleFile" style="Z-INDEX: 171; LEFT: 16px; POSITION: absolute; TOP: 432px"
					runat="server" Height="230px" Width="200px" BorderStyle="Groove" ImageUrl="http://10.245.0.178//WorkFlow/Document/000003/1-Map2-2006124111750-2.jpg"></asp:image><asp:textbox id="DEnglishName" style="Z-INDEX: 170; LEFT: 440px; POSITION: absolute; TOP: 632px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="320px" BorderStyle="Groove">DEnglishName</asp:textbox><asp:textbox id="DCode" style="Z-INDEX: 169; LEFT: 440px; POSITION: absolute; TOP: 592px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="20px" Width="320px" BorderStyle="Groove">DCode</asp:textbox><asp:button id="BBFinalDate" style="Z-INDEX: 168; LEFT: 736px; POSITION: absolute; TOP: 560px"
					runat="server" Height="20px" Width="20px" Text="....."></asp:button><asp:textbox id="DBFinalDate" style="Z-INDEX: 167; LEFT: 440px; POSITION: absolute; TOP: 560px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="294px" BorderStyle="Groove" ReadOnly="True">DBFinalDate</asp:textbox><asp:dropdownlist id="DAllowSample" style="Z-INDEX: 166; LEFT: 440px; POSITION: absolute; TOP: 528px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="320px">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DQty" style="Z-INDEX: 165; LEFT: 331px; POSITION: absolute; TOP: 496px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="20px" Width="102px" BorderStyle="Groove">DQty</asp:textbox><asp:textbox id="DColor" style="Z-INDEX: 164; LEFT: 330px; POSITION: absolute; TOP: 464px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="20px" Width="103px" BorderStyle="Groove">DColor</asp:textbox><asp:dropdownlist id="DManufType" style="Z-INDEX: 163; LEFT: 331px; POSITION: absolute; TOP: 400px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="432px">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:button id="BOrderTime" style="Z-INDEX: 161; LEFT: 736px; POSITION: absolute; TOP: 288px"
					runat="server" Height="20px" Width="20px" Text="....."></asp:button><asp:button id="BReqDelDate" style="Z-INDEX: 160; LEFT: 464px; POSITION: absolute; TOP: 224px"
					runat="server" Height="20px" Width="20px" Text="....."></asp:button><asp:textbox id="DSliderSample" style="Z-INDEX: 159; LEFT: 312px; POSITION: absolute; TOP: 256px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="180px" BorderStyle="Groove">DSliderSample</asp:textbox><asp:textbox id="DReqQty" style="Z-INDEX: 158; LEFT: 678px; POSITION: absolute; TOP: 224px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="20px" Width="87px" BorderStyle="Groove">DReqQty</asp:textbox><asp:textbox id="DReqDelDate" style="Z-INDEX: 157; LEFT: 312px; POSITION: absolute; TOP: 224px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="150px" BorderStyle="Groove" ReadOnly="True">DReqDelDate</asp:textbox><asp:textbox id="DDevReason" style="Z-INDEX: 120; LEFT: 312px; POSITION: absolute; TOP: 352px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="24px" Width="450px" BorderStyle="Groove" MaxLength="240" TextMode="MultiLine">DevReason</asp:textbox><asp:textbox id="DOPReadyDesc" style="Z-INDEX: 156; LEFT: 248px; POSITION: absolute; TOP: 1352px"
					runat="server" ForeColor="Red" BackColor="White" Height="20px" Width="104px" BorderStyle="None">需閱讀工程履歷</asp:textbox><asp:textbox id="DOPReady" style="Z-INDEX: 137; LEFT: 352px; POSITION: absolute; TOP: 1352px"
					runat="server" ForeColor="Red" BackColor="GreenYellow" Height="20px" Width="48px" BorderStyle="Groove" ReadOnly="True"></asp:textbox><asp:image id="DSufaceSheet1" style="Z-INDEX: 102; LEFT: 8px; POSITION: absolute; TOP: 8px"
					runat="server" Height="684px" Width="766px" ImageUrl="Images\SufaceSheet_001_C.jpg"></asp:image><asp:dropdownlist id="DSuppiler" style="Z-INDEX: 155; LEFT: 331px; POSITION: absolute; TOP: 432px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="432px">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DBuyer" style="Z-INDEX: 154; LEFT: 312px; POSITION: absolute; TOP: 188px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="20px" Width="180px">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:hyperlink id="LQCReqFile" style="Z-INDEX: 153; LEFT: 440px; POSITION: absolute; TOP: 664px"
					runat="server" Height="8px" Width="96px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">品質依賴書</asp:hyperlink><INPUT id="DQCReqFile" style="Z-INDEX: 152; LEFT: 440px; WIDTH: 262px; POSITION: absolute; TOP: 664px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					type="file" size="24" name="File1" runat="server"><asp:imagebutton id="BFlow" style="Z-INDEX: 151; LEFT: 8px; POSITION: absolute; TOP: 32px" runat="server"
					Height="16px" Width="16px" ImageUrl="Images\Flow-01.gif"></asp:imagebutton><INPUT id="DCustSampleFile" style="Z-INDEX: 145; LEFT: 16px; WIDTH: 200px; POSITION: absolute; TOP: 352px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					type="file" size="14" name="File1" runat="server"><asp:hyperlink id="LOPManualFile" style="Z-INDEX: 143; LEFT: 128px; POSITION: absolute; TOP: 1168px"
					runat="server" Height="8px" Width="96px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">作業標準書</asp:hyperlink><INPUT id="DOPManualFile" style="Z-INDEX: 142; LEFT: 128px; WIDTH: 240px; POSITION: absolute; TOP: 1168px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					type="file" name="File1" runat="server">
				<asp:hyperlink id="LEACheckFile" style="Z-INDEX: 141; LEFT: 452px; POSITION: absolute; TOP: 936px"
					runat="server" Height="8px" Width="96px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">Oeko-tex-報告</asp:hyperlink><asp:hyperlink id="LQCFinalFile" style="Z-INDEX: 139; LEFT: 300px; POSITION: absolute; TOP: 936px"
					runat="server" Height="12px" Width="96px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">測試報告書</asp:hyperlink><INPUT id="DQCFinalFile" style="Z-INDEX: 138; LEFT: 297px; WIDTH: 144px; POSITION: absolute; TOP: 936px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					type="file" size="4" name="File1" runat="server"><asp:button id="BDate" style="Z-INDEX: 136; LEFT: 736px; POSITION: absolute; TOP: 88px" runat="server"
					Height="20px" Width="20px" Text="....."></asp:button><asp:textbox id="DDate" style="Z-INDEX: 135; LEFT: 584px; POSITION: absolute; TOP: 88px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="20px" Width="150px" BorderStyle="Groove" ReadOnly="True">DDate</asp:textbox><asp:hyperlink id="LContactFile" style="Z-INDEX: 134; LEFT: 504px; POSITION: absolute; TOP: 1168px"
					runat="server" Height="8px" Width="96px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">切結書</asp:hyperlink><INPUT id="DContactFile" style="Z-INDEX: 133; LEFT: 504px; WIDTH: 260px; POSITION: absolute; TOP: 1168px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					type="file" size="24" name="File1" runat="server"><asp:dropdownlist id="DPerson" style="Z-INDEX: 132; LEFT: 584px; POSITION: absolute; TOP: 122px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="20px" Width="180px">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DDivision" style="Z-INDEX: 131; LEFT: 312px; POSITION: absolute; TOP: 122px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="180px">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DReasonDesc" style="Z-INDEX: 130; LEFT: 168px; POSITION: absolute; TOP: 1424px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="59px" Width="424px" BorderStyle="Groove" TextMode="MultiLine"
					Visible="False">DReasonDesc</asp:textbox><asp:textbox id="DReason" style="Z-INDEX: 129; LEFT: 240px; POSITION: absolute; TOP: 1392px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="352px" BorderStyle="Groove" ReadOnly="True" Visible="False">DReason</asp:textbox><asp:dropdownlist id="DReasonCode" style="Z-INDEX: 128; LEFT: 168px; POSITION: absolute; TOP: 1392px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="64px" AutoPostBack="True" Visible="False">
					<asp:ListItem Value="01">01</asp:ListItem>
					<asp:ListItem Value="02">02</asp:ListItem>
				</asp:dropdownlist><asp:hyperlink id="LBefOP" style="Z-INDEX: 127; LEFT: 168px; POSITION: absolute; TOP: 1352px" runat="server"
					Height="8px" Width="144px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">工程履歷</asp:hyperlink><asp:textbox id="DAEndTime" style="Z-INDEX: 126; LEFT: 440px; POSITION: absolute; TOP: 1320px"
					runat="server" ForeColor="Blue" BackColor="Gold" Height="20px" Width="152px" BorderStyle="Groove" ReadOnly="True" Font-Size="9pt">DAEndTime</asp:textbox><asp:textbox id="DAStartTime" style="Z-INDEX: 125; LEFT: 168px; POSITION: absolute; TOP: 1320px"
					runat="server" ForeColor="Blue" BackColor="Gold" Height="20px" Width="152px" BorderStyle="Groove" ReadOnly="True" Font-Size="9pt">DAStartTime</asp:textbox><asp:textbox id="DBEndTime" style="Z-INDEX: 124; LEFT: 440px; POSITION: absolute; TOP: 1288px"
					runat="server" ForeColor="Blue" BackColor="Gold" Height="20px" Width="152px" BorderStyle="Groove" ReadOnly="True" Font-Size="9pt">DBEndTime</asp:textbox><asp:textbox id="DBStartTime" style="Z-INDEX: 123; LEFT: 168px; POSITION: absolute; TOP: 1288px"
					runat="server" ForeColor="Blue" BackColor="Gold" Height="20px" Width="152px" BorderStyle="Groove" ReadOnly="True" Font-Size="9pt">DBStartTime</asp:textbox><asp:textbox id="DDecideDesc" style="Z-INDEX: 122; LEFT: 56px; POSITION: absolute; TOP: 1208px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="56px" Width="536px" BorderStyle="Groove" TextMode="MultiLine">DDecideDesc</asp:textbox><asp:hyperlink id="LForcastFile" style="Z-INDEX: 119; LEFT: 504px; POSITION: absolute; TOP: 1136px"
					runat="server" Height="8px" Width="96px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">報價單</asp:hyperlink><INPUT id="DForcastFile" style="Z-INDEX: 118; LEFT: 504px; WIDTH: 260px; POSITION: absolute; TOP: 1136px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					type="file" size="24" name="File1" runat="server">
				<asp:hyperlink id="LManufFlowFile" style="Z-INDEX: 116; LEFT: 128px; POSITION: absolute; TOP: 1136px"
					runat="server" Height="8px" Width="96px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">製造流程表</asp:hyperlink><INPUT id="DManufFlowFile" style="Z-INDEX: 115; LEFT: 128px; WIDTH: 240px; POSITION: absolute; TOP: 1136px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					type="file" name="File1" runat="server"><asp:textbox id="DSellVendor" style="Z-INDEX: 114; LEFT: 584px; POSITION: absolute; TOP: 188px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="180px" BorderStyle="Groove">DSellVendor</asp:textbox><asp:dropdownlist id="DAttachSample" style="Z-INDEX: 113; LEFT: 584px; POSITION: absolute; TOP: 256px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="180px">
					<asp:ListItem Value="Y">DAttachSample</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><INPUT id="DEACheckFile" style="Z-INDEX: 112; LEFT: 456px; WIDTH: 160px; POSITION: absolute; TOP: 936px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					type="file" size="7" name="File1" runat="server"><asp:textbox id="DORNO" style="Z-INDEX: 111; LEFT: 312px; POSITION: absolute; TOP: 288px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="20px" Width="180px" BorderStyle="Groove">DORNO</asp:textbox><asp:button id="BSpec" style="Z-INDEX: 109; LEFT: 736px; POSITION: absolute; TOP: 154px" runat="server"
					Height="20px" Width="20px" Text="....."></asp:button><asp:textbox id="DSpec" style="Z-INDEX: 108; LEFT: 312px; POSITION: absolute; TOP: 154px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="20px" Width="420px" BorderStyle="Groove" ReadOnly="True" Font-Size="8pt">DSpec</asp:textbox><asp:textbox id="DFormSno" style="Z-INDEX: 107; LEFT: 16px; POSITION: absolute; TOP: 1488px"
					runat="server" ForeColor="Blue" BackColor="White" Height="20px" Width="97px" BorderStyle="None">單號：123</asp:textbox><asp:image id="DDelivery" style="Z-INDEX: 106; LEFT: 8px; POSITION: absolute; TOP: 1272px"
					runat="server" Height="110px" Width="593px" ImageUrl="Images\Sheet_Delivery.jpg"></asp:image><asp:image id="DDelay" style="Z-INDEX: 105; LEFT: 8px; POSITION: absolute; TOP: 1384px" runat="server"
					Height="107px" Width="593px" ImageUrl="Images\Sheet_Delay.jpg" Visible="False"></asp:image><asp:image id="DDescSheet" style="Z-INDEX: 104; LEFT: 8px; POSITION: absolute; TOP: 1200px"
					runat="server" Height="75px" Width="593px" ImageUrl="Images\MapSheet_004.jpg"></asp:image><asp:image id="DSufaceSheet2" style="Z-INDEX: 103; LEFT: 14px; POSITION: absolute; TOP: 692px"
					runat="server" Height="508px" Width="766px" ImageUrl="Images\SufaceSheet_002_E.jpg"></asp:image><asp:image id="LCustSampleFile" style="Z-INDEX: 144; LEFT: 16px; POSITION: absolute; TOP: 120px"
					runat="server" Height="230px" Width="200px" BorderStyle="Groove" ImageUrl="http://10.245.0.178//WorkFlow/Document/000003/1-Map2-2006124111750-2.jpg"></asp:image></FONT><INPUT id="BSAVE" style="Z-INDEX: 146; LEFT: 256px; WIDTH: 80px; POSITION: absolute; TOP: 1496px; HEIGHT: 28px"
				onclick="Button('SAVE');" type="button" value="儲存" name="Button1" runat="server"><INPUT id="BNG2" style="Z-INDEX: 147; LEFT: 344px; WIDTH: 80px; POSITION: absolute; TOP: 1496px; HEIGHT: 28px"
				onclick="Button('NG2');" type="button" value="NG2" name="Button1" runat="server"><INPUT id="BNG1" style="Z-INDEX: 148; LEFT: 432px; WIDTH: 80px; POSITION: absolute; TOP: 1496px; HEIGHT: 28px"
				onclick="Button('NG1');" type="button" value="NG1" name="Button1" runat="server"><INPUT id="BOK" style="Z-INDEX: 149; LEFT: 520px; WIDTH: 80px; POSITION: absolute; TOP: 1496px; HEIGHT: 28px"
				onclick="Button('OK');" type="button" value="OK" name="Button2" runat="server">
			<asp:imagebutton id="BPrint" style="Z-INDEX: 150; LEFT: 8px; POSITION: absolute; TOP: 8px" runat="server"
				Height="16px" Width="16px" ImageUrl="Images\Print.gif"></asp:imagebutton><asp:textbox id="DOFormNo" style="Z-INDEX: 101; LEFT: 712px; POSITION: absolute; TOP: 136px"
				runat="server" Height="24px" Width="56px" ReadOnly="True"></asp:textbox></form>
		</FONT></FORM>
	</body>
</HTML>
