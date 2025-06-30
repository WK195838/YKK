<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SufaceSuppilerSheet_01.aspx.vb" Inherits="SPD.SufaceSuppilerSheet_01"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>表面處理-外注廠商追加委託書</title>
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
			//--Spec------------------------------------
			function SpecPicker(strField, Fun) {
			    val=0;
				wPop=window.open('SpecPicker.aspx?field=' + strField + '&fun=' + Fun, 'SpecPopup','width=330,height=250,resizable=yes');
				if (document.Form1.DSpec.value != "") {
					setTimeout("SendToSpec()",200);
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
				<asp:textbox id="DOrderTime" style="Z-INDEX: 158; LEFT: 584px; POSITION: absolute; TOP: 288px"
					runat="server" ReadOnly="True" BorderStyle="Groove" Width="150px" Height="20px" BackColor="Yellow"
					ForeColor="Blue">DOrderTime</asp:textbox>
				<asp:textbox id="DQCLT" style="Z-INDEX: 314; LEFT: 706px; POSITION: absolute; TOP: 664px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="18px" Width="53px" BorderStyle="Groove" Font-Size="8pt"></asp:textbox><asp:hyperlink id="LQCCheck6File" style="Z-INDEX: 195; LEFT: 648px; POSITION: absolute; TOP: 736px"
					runat="server" Width="96px" Height="8px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">電鍍膜厚</asp:hyperlink><INPUT id="DQCCheck6File" style="Z-INDEX: 194; LEFT: 648px; WIDTH: 120px; POSITION: absolute; TOP: 736px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					type="file" size="31" name="File1" runat="server">
				<asp:button id="BQCDate1" style="Z-INDEX: 193; LEFT: 96px; POSITION: absolute; TOP: 1024px"
					runat="server" Width="20px" Height="20px" Text="....."></asp:button><asp:button id="BQCDate2" style="Z-INDEX: 192; LEFT: 96px; POSITION: absolute; TOP: 1056px"
					runat="server" Width="20px" Height="20px" Text="....."></asp:button><asp:button id="BQCDate3" style="Z-INDEX: 191; LEFT: 96px; POSITION: absolute; TOP: 1088px"
					runat="server" Width="20px" Height="20px" Text="....."></asp:button><asp:dropdownlist id="DQCResult3" style="Z-INDEX: 190; LEFT: 120px; POSITION: absolute; TOP: 1088px"
					runat="server" Width="78px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQCResult2" style="Z-INDEX: 189; LEFT: 120px; POSITION: absolute; TOP: 1056px"
					runat="server" Width="78px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DQCRemark3" style="Z-INDEX: 188; LEFT: 208px; POSITION: absolute; TOP: 1088px"
					runat="server" BorderStyle="Groove" Width="554px" Height="20px" BackColor="Yellow" ForeColor="Blue">DQCRemark3</asp:textbox><asp:textbox id="DQCRemark2" style="Z-INDEX: 187; LEFT: 208px; POSITION: absolute; TOP: 1056px"
					runat="server" BorderStyle="Groove" Width="554px" Height="20px" BackColor="Yellow" ForeColor="Blue">DQCRemark2</asp:textbox><asp:textbox id="DQCDate3" style="Z-INDEX: 186; LEFT: 16px; POSITION: absolute; TOP: 1088px"
					runat="server" BorderStyle="Groove" Width="80px" Height="20px" BackColor="Yellow" ForeColor="Blue">DQCDate3</asp:textbox><asp:textbox id="DQCDate2" style="Z-INDEX: 185; LEFT: 16px; POSITION: absolute; TOP: 1056px"
					runat="server" BorderStyle="Groove" Width="80px" Height="20px" BackColor="Yellow" ForeColor="Blue">DQCDate2</asp:textbox><asp:textbox id="DQCRemark1" style="Z-INDEX: 184; LEFT: 208px; POSITION: absolute; TOP: 1024px"
					runat="server" BorderStyle="Groove" Width="554px" Height="20px" BackColor="Yellow" ForeColor="Blue">DQCRemark1</asp:textbox><asp:dropdownlist id="DQCResult1" style="Z-INDEX: 183; LEFT: 120px; POSITION: absolute; TOP: 1024px"
					runat="server" Width="78px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DQCDate1" style="Z-INDEX: 182; LEFT: 16px; POSITION: absolute; TOP: 1024px"
					runat="server" BorderStyle="Groove" Width="80px" Height="20px" BackColor="Yellow" ForeColor="Blue">DQCDate1</asp:textbox><asp:textbox id="DEADesc1" style="Z-INDEX: 181; LEFT: 144px; POSITION: absolute; TOP: 864px"
					runat="server" BorderStyle="Groove" Width="618px" Height="20px" BackColor="Yellow" ForeColor="Blue">DEADesc1</asp:textbox><asp:dropdownlist id="DEACheck1" style="Z-INDEX: 180; LEFT: 24px; POSITION: absolute; TOP: 864px"
					runat="server" Width="110px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQCCheck12" style="Z-INDEX: 179; LEFT: 656px; POSITION: absolute; TOP: 800px"
					runat="server" Width="110px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQCCheck11" style="Z-INDEX: 178; LEFT: 528px; POSITION: absolute; TOP: 800px"
					runat="server" Width="110px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQCCheck10" style="Z-INDEX: 177; LEFT: 400px; POSITION: absolute; TOP: 800px"
					runat="server" Width="110px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQCCheck9" style="Z-INDEX: 176; LEFT: 272px; POSITION: absolute; TOP: 800px"
					runat="server" Width="110px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQCCheck8" style="Z-INDEX: 175; LEFT: 144px; POSITION: absolute; TOP: 800px"
					runat="server" Width="110px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQCCheck7" style="Z-INDEX: 174; LEFT: 24px; POSITION: absolute; TOP: 800px"
					runat="server" Width="110px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQCCheck5" style="Z-INDEX: 173; LEFT: 528px; POSITION: absolute; TOP: 736px"
					runat="server" Width="110px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQCCheck4" style="Z-INDEX: 172; LEFT: 400px; POSITION: absolute; TOP: 736px"
					runat="server" Width="110px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQCCheck3" style="Z-INDEX: 171; LEFT: 272px; POSITION: absolute; TOP: 736px"
					runat="server" Width="110px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQCCheck2" style="Z-INDEX: 170; LEFT: 144px; POSITION: absolute; TOP: 736px"
					runat="server" Width="110px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQCCheck1" style="Z-INDEX: 169; LEFT: 24px; POSITION: absolute; TOP: 736px"
					runat="server" Width="110px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><INPUT id="DFinalSampleFile" style="Z-INDEX: 168; LEFT: 16px; WIDTH: 200px; POSITION: absolute; TOP: 664px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					type="file" size="14" name="File1" runat="server">
				<asp:image id="LFinalSampleFile" style="Z-INDEX: 167; LEFT: 16px; POSITION: absolute; TOP: 432px"
					runat="server" BorderStyle="Groove" Width="200px" Height="230px" ImageUrl="http://10.245.0.178//WorkFlow/Document/000003/1-Map2-2006124111750-2.jpg"></asp:image><asp:textbox id="DEnglishName" style="Z-INDEX: 166; LEFT: 440px; POSITION: absolute; TOP: 632px"
					runat="server" BorderStyle="Groove" Width="320px" Height="20px" BackColor="Yellow" ForeColor="Blue">DEnglishName</asp:textbox><asp:textbox id="DCode" style="Z-INDEX: 165; LEFT: 440px; POSITION: absolute; TOP: 592px" runat="server"
					BorderStyle="Groove" Width="320px" Height="20px" BackColor="Yellow" ForeColor="Blue">DCode</asp:textbox><asp:button id="BBFinalDate" style="Z-INDEX: 164; LEFT: 736px; POSITION: absolute; TOP: 560px"
					runat="server" Width="20px" Height="20px" Text="....."></asp:button><asp:textbox id="DBFinalDate" style="Z-INDEX: 163; LEFT: 440px; POSITION: absolute; TOP: 560px"
					runat="server" ReadOnly="True" BorderStyle="Groove" Width="294px" Height="20px" BackColor="Yellow" ForeColor="Blue">DBFinalDate</asp:textbox><asp:dropdownlist id="DAllowSample" style="Z-INDEX: 162; LEFT: 440px; POSITION: absolute; TOP: 528px"
					runat="server" Width="320px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DQty" style="Z-INDEX: 161; LEFT: 440px; POSITION: absolute; TOP: 496px" runat="server"
					BorderStyle="Groove" Width="320px" Height="20px" BackColor="Yellow" ForeColor="Blue">DQty</asp:textbox><asp:textbox id="DColor" style="Z-INDEX: 160; LEFT: 440px; POSITION: absolute; TOP: 464px" runat="server"
					BorderStyle="Groove" Width="320px" Height="20px" BackColor="Yellow" ForeColor="Blue">DColor</asp:textbox><asp:dropdownlist id="DManufType" style="Z-INDEX: 159; LEFT: 440px; POSITION: absolute; TOP: 400px"
					runat="server" Width="320px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:button id="BOrderTime" style="Z-INDEX: 157; LEFT: 736px; POSITION: absolute; TOP: 288px"
					runat="server" Width="20px" Height="20px" Text="....."></asp:button><asp:button id="BReqDelDate" style="Z-INDEX: 156; LEFT: 464px; POSITION: absolute; TOP: 224px"
					runat="server" Width="20px" Height="20px" Text="....."></asp:button><asp:textbox id="DSliderSample" style="Z-INDEX: 155; LEFT: 312px; POSITION: absolute; TOP: 256px"
					runat="server" BorderStyle="Groove" Width="180px" Height="20px" BackColor="Yellow" ForeColor="Blue">DSliderSample</asp:textbox><asp:textbox id="DReqQty" style="Z-INDEX: 154; LEFT: 584px; POSITION: absolute; TOP: 224px" runat="server"
					BorderStyle="Groove" Width="180px" Height="20px" BackColor="Yellow" ForeColor="Blue">DReqQty</asp:textbox><asp:textbox id="DReqDelDate" style="Z-INDEX: 153; LEFT: 312px; POSITION: absolute; TOP: 224px"
					runat="server" ReadOnly="True" BorderStyle="Groove" Width="150px" Height="20px" BackColor="Yellow" ForeColor="Blue">DReqDelDate</asp:textbox><asp:textbox id="DDevReason" style="Z-INDEX: 117; LEFT: 312px; POSITION: absolute; TOP: 320px"
					runat="server" BorderStyle="Groove" Width="450px" Height="56px" BackColor="Yellow" ForeColor="Blue" MaxLength="240" TextMode="MultiLine">DevReason</asp:textbox><asp:textbox id="DOPReadyDesc" style="Z-INDEX: 152; LEFT: 248px; POSITION: absolute; TOP: 1352px"
					runat="server" BorderStyle="None" Width="104px" Height="20px" BackColor="White" ForeColor="Red">需閱讀工程履歷</asp:textbox><asp:textbox id="DOPReady" style="Z-INDEX: 133; LEFT: 352px; POSITION: absolute; TOP: 1352px"
					runat="server" ReadOnly="True" BorderStyle="Groove" Width="48px" Height="20px" BackColor="GreenYellow" ForeColor="Red"></asp:textbox><asp:image id="DSufaceSheet1" style="Z-INDEX: 100; LEFT: 8px; POSITION: absolute; TOP: 8px"
					runat="server" Width="766px" Height="684px" ImageUrl="Images\SufaceSuppilerSheet_001_A.jpg"></asp:image><asp:dropdownlist id="DSuppiler" style="Z-INDEX: 151; LEFT: 440px; POSITION: absolute; TOP: 432px"
					runat="server" Width="320px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DBuyer" style="Z-INDEX: 150; LEFT: 312px; POSITION: absolute; TOP: 188px" runat="server"
					Width="180px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:hyperlink id="LQCReqFile" style="Z-INDEX: 149; LEFT: 440px; POSITION: absolute; TOP: 664px"
					runat="server" Width="96px" Height="8px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">品質依賴書</asp:hyperlink><INPUT id="DQCReqFile" style="Z-INDEX: 147; LEFT: 440px; WIDTH: 262px; POSITION: absolute; TOP: 664px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					type="file" size="24" name="File1" runat="server"><asp:imagebutton id="BFlow" style="Z-INDEX: 146; LEFT: 8px; POSITION: absolute; TOP: 32px" runat="server"
					Width="16px" Height="16px" ImageUrl="Images\Flow-01.gif"></asp:imagebutton><INPUT id="DCustSampleFile" style="Z-INDEX: 140; LEFT: 16px; WIDTH: 200px; POSITION: absolute; TOP: 352px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					type="file" size="14" name="File1" runat="server"><asp:hyperlink id="LOPManualFile" style="Z-INDEX: 138; LEFT: 128px; POSITION: absolute; TOP: 1168px"
					runat="server" Width="96px" Height="8px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">作業標準書</asp:hyperlink><INPUT id="DOPManualFile" style="Z-INDEX: 137; LEFT: 128px; WIDTH: 240px; POSITION: absolute; TOP: 1168px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					type="file" size="20" name="File1" runat="server">
				<asp:hyperlink id="LEACheckFile" style="Z-INDEX: 136; LEFT: 24px; POSITION: absolute; TOP: 932px"
					runat="server" Width="112px" Height="8px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">有害物質報告</asp:hyperlink><asp:hyperlink id="LQCFinalFile" style="Z-INDEX: 135; LEFT: 400px; POSITION: absolute; TOP: 932px"
					runat="server" Width="96px" Height="8px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">測試報告書</asp:hyperlink><INPUT id="DQCFinalFile" style="Z-INDEX: 134; LEFT: 400px; WIDTH: 362px; POSITION: absolute; TOP: 932px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					type="file" size="41" name="File1" runat="server"><asp:button id="BDate" style="Z-INDEX: 132; LEFT: 736px; POSITION: absolute; TOP: 88px" runat="server"
					Width="20px" Height="20px" Text="....."></asp:button><asp:textbox id="DDate" style="Z-INDEX: 131; LEFT: 584px; POSITION: absolute; TOP: 88px" runat="server"
					ReadOnly="True" BorderStyle="Groove" Width="150px" Height="20px" BackColor="Yellow" ForeColor="Blue">DDate</asp:textbox><asp:hyperlink id="LContactFile" style="Z-INDEX: 130; LEFT: 504px; POSITION: absolute; TOP: 1168px"
					runat="server" Width="96px" Height="8px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">切結書</asp:hyperlink><INPUT id="DContactFile" style="Z-INDEX: 129; LEFT: 504px; WIDTH: 260px; POSITION: absolute; TOP: 1168px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					type="file" size="24" name="File1" runat="server"><asp:dropdownlist id="DPerson" style="Z-INDEX: 128; LEFT: 584px; POSITION: absolute; TOP: 122px" runat="server"
					Width="180px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DDivision" style="Z-INDEX: 127; LEFT: 312px; POSITION: absolute; TOP: 122px"
					runat="server" Width="180px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DReasonDesc" style="Z-INDEX: 126; LEFT: 168px; POSITION: absolute; TOP: 1424px"
					runat="server" BorderStyle="Groove" Width="424px" Height="59px" BackColor="Yellow" ForeColor="Blue" TextMode="MultiLine"
					Visible="False">DReasonDesc</asp:textbox><asp:textbox id="DReason" style="Z-INDEX: 125; LEFT: 240px; POSITION: absolute; TOP: 1392px"
					runat="server" ReadOnly="True" BorderStyle="Groove" Width="352px" Height="20px" BackColor="Yellow" ForeColor="Blue" Visible="False">DReason</asp:textbox><asp:dropdownlist id="DReasonCode" style="Z-INDEX: 124; LEFT: 168px; POSITION: absolute; TOP: 1392px"
					runat="server" Width="64px" Height="20px" BackColor="Yellow" ForeColor="Blue" Visible="False" AutoPostBack="True">
					<asp:ListItem Value="01">01</asp:ListItem>
					<asp:ListItem Value="02">02</asp:ListItem>
				</asp:dropdownlist><asp:hyperlink id="LBefOP" style="Z-INDEX: 123; LEFT: 168px; POSITION: absolute; TOP: 1352px" runat="server"
					Width="144px" Height="8px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">工程履歷</asp:hyperlink><asp:textbox id="DAEndTime" style="Z-INDEX: 122; LEFT: 440px; POSITION: absolute; TOP: 1320px"
					runat="server" ReadOnly="True" BorderStyle="Groove" Width="152px" Height="20px" BackColor="Gold" ForeColor="Blue" Font-Size="9pt">DAEndTime</asp:textbox><asp:textbox id="DAStartTime" style="Z-INDEX: 121; LEFT: 168px; POSITION: absolute; TOP: 1320px"
					runat="server" ReadOnly="True" BorderStyle="Groove" Width="152px" Height="20px" BackColor="Gold" ForeColor="Blue" Font-Size="9pt">DAStartTime</asp:textbox><asp:textbox id="DBEndTime" style="Z-INDEX: 120; LEFT: 440px; POSITION: absolute; TOP: 1288px"
					runat="server" ReadOnly="True" BorderStyle="Groove" Width="152px" Height="20px" BackColor="Gold" ForeColor="Blue" Font-Size="9pt">DBEndTime</asp:textbox><asp:textbox id="DBStartTime" style="Z-INDEX: 119; LEFT: 168px; POSITION: absolute; TOP: 1288px"
					runat="server" ReadOnly="True" BorderStyle="Groove" Width="152px" Height="20px" BackColor="Gold" ForeColor="Blue" Font-Size="9pt">DBStartTime</asp:textbox><asp:textbox id="DDecideDesc" style="Z-INDEX: 118; LEFT: 56px; POSITION: absolute; TOP: 1208px"
					runat="server" BorderStyle="Groove" Width="536px" Height="56px" BackColor="Yellow" ForeColor="Blue" TextMode="MultiLine">DDecideDesc</asp:textbox><asp:hyperlink id="LForcastFile" style="Z-INDEX: 116; LEFT: 504px; POSITION: absolute; TOP: 1136px"
					runat="server" Width="96px" Height="8px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">報價單</asp:hyperlink><INPUT id="DForcastFile" style="Z-INDEX: 115; LEFT: 504px; WIDTH: 260px; POSITION: absolute; TOP: 1136px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					type="file" size="24" name="File1" runat="server">
				<asp:hyperlink id="LManufFlowFile" style="Z-INDEX: 114; LEFT: 128px; POSITION: absolute; TOP: 1136px"
					runat="server" Width="96px" Height="8px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">製造流程表</asp:hyperlink><INPUT id="DManufFlowFile" style="Z-INDEX: 113; LEFT: 128px; WIDTH: 240px; POSITION: absolute; TOP: 1136px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					type="file" size="20" name="File1" runat="server"><asp:textbox id="DSellVendor" style="Z-INDEX: 112; LEFT: 584px; POSITION: absolute; TOP: 188px"
					runat="server" BorderStyle="Groove" Width="180px" Height="20px" BackColor="Yellow" ForeColor="Blue">DSellVendor</asp:textbox><asp:dropdownlist id="DAttachSample" style="Z-INDEX: 111; LEFT: 584px; POSITION: absolute; TOP: 256px"
					runat="server" Width="180px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">DAttachSample</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><INPUT id="DEACheckFile" style="Z-INDEX: 110; LEFT: 24px; WIDTH: 360px; POSITION: absolute; TOP: 932px; HEIGHT: 20px; BACKGROUND-COLOR: #ffff00"
					type="file" size="40" name="File1" runat="server"><asp:textbox id="DORNO" style="Z-INDEX: 109; LEFT: 312px; POSITION: absolute; TOP: 288px" runat="server"
					BorderStyle="Groove" Width="180px" Height="20px" BackColor="Yellow" ForeColor="Blue">DORNO</asp:textbox><asp:button id="BSpec" style="Z-INDEX: 108; LEFT: 736px; POSITION: absolute; TOP: 154px" runat="server"
					Width="20px" Height="20px" Text="....."></asp:button><asp:textbox id="DSpec" style="Z-INDEX: 107; LEFT: 312px; POSITION: absolute; TOP: 154px" runat="server"
					ReadOnly="True" BorderStyle="Groove" Width="420px" Height="20px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">DSpec</asp:textbox><asp:textbox id="DNo" style="Z-INDEX: 106; LEFT: 312px; POSITION: absolute; TOP: 88px" runat="server"
					BorderStyle="Groove" Width="180px" Height="20px" BackColor="Yellow" ForeColor="Blue">DNo</asp:textbox><asp:textbox id="DFormSno" style="Z-INDEX: 105; LEFT: 16px; POSITION: absolute; TOP: 1488px"
					runat="server" BorderStyle="None" Width="97px" Height="20px" BackColor="White" ForeColor="Blue">單號：123</asp:textbox><asp:image id="DDelivery" style="Z-INDEX: 104; LEFT: 8px; POSITION: absolute; TOP: 1272px"
					runat="server" Width="593px" Height="110px" ImageUrl="Images\Sheet_Delivery.jpg"></asp:image><asp:image id="DDelay" style="Z-INDEX: 103; LEFT: 8px; POSITION: absolute; TOP: 1384px" runat="server"
					Width="593px" Height="107px" ImageUrl="Images\Sheet_Delay.jpg" Visible="False"></asp:image><asp:image id="DDescSheet" style="Z-INDEX: 102; LEFT: 8px; POSITION: absolute; TOP: 1200px"
					runat="server" Width="593px" Height="75px" ImageUrl="Images\MapSheet_004.jpg"></asp:image><asp:image id="DSufaceSheet2" style="Z-INDEX: 101; LEFT: 8px; POSITION: absolute; TOP: 688px"
					runat="server" Width="766px" Height="508px" ImageUrl="Images\SufaceSheet_002_A.jpg"></asp:image><asp:image id="LCustSampleFile" style="Z-INDEX: 139; LEFT: 16px; POSITION: absolute; TOP: 120px"
					runat="server" BorderStyle="Groove" Width="200px" Height="230px" ImageUrl="http://10.245.0.178//WorkFlow/Document/000003/1-Map2-2006124111750-2.jpg"></asp:image></FONT><INPUT id="BSAVE" style="Z-INDEX: 141; LEFT: 256px; WIDTH: 80px; POSITION: absolute; TOP: 1496px; HEIGHT: 28px"
				onclick="Button('SAVE');" type="button" value="儲存" name="Button1" runat="server"><INPUT id="BNG2" style="Z-INDEX: 142; LEFT: 344px; WIDTH: 80px; POSITION: absolute; TOP: 1496px; HEIGHT: 28px"
				onclick="Button('NG2');" type="button" value="NG2" name="Button1" runat="server"><INPUT id="BNG1" style="Z-INDEX: 143; LEFT: 432px; WIDTH: 80px; POSITION: absolute; TOP: 1496px; HEIGHT: 28px"
				onclick="Button('NG1');" type="button" value="NG1" name="Button1" runat="server"><INPUT id="BOK" style="Z-INDEX: 144; LEFT: 520px; WIDTH: 80px; POSITION: absolute; TOP: 1496px; HEIGHT: 28px"
				onclick="Button('OK');" type="button" value="OK" name="Button2" runat="server">
			<asp:imagebutton id="BPrint" style="Z-INDEX: 145; LEFT: 8px; POSITION: absolute; TOP: 8px" runat="server"
				Width="16px" Height="16px" ImageUrl="Images\Print.gif"></asp:imagebutton></form>
		</FONT></FORM>
	</body>
</HTML>
