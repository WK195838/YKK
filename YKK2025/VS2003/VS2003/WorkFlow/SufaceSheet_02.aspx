<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SufaceSheet_02.aspx.vb" Inherits="SPD.SufaceSheet_02"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>表面處理委託書</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form2" method="post" encType="multipart/form-data" runat="server">
			<FONT face="新細明體">
				<asp:textbox id="DOrderTime" style="Z-INDEX: 125; LEFT: 584px; POSITION: absolute; TOP: 288px"
					runat="server" ReadOnly="True" BorderStyle="Groove" Width="180px" Height="20px" BackColor="Yellow"
					ForeColor="Blue">DOrderTime</asp:textbox>
				<asp:dropdownlist id="DQCCheck16" style="Z-INDEX: 171; LEFT: 24px; POSITION: absolute; TOP: 948px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="2px" Width="112px">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist>
				<asp:hyperlink id="LPFASFile" style="Z-INDEX: 170; LEFT: 160px; POSITION: absolute; TOP: 952px"
					runat="server" Height="10px" Width="113px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">PFAS報告</asp:hyperlink>
				<asp:textbox style="Z-INDEX: 169; LEFT: 512px; POSITION: absolute; TOP: 496px" id="DFReason"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="24px" Width="248px" BorderStyle="Groove">DFReason</asp:textbox>
				<asp:textbox style="Z-INDEX: 168; LEFT: 584px; POSITION: absolute; TOP: 224px" id="DYearSeason"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="81px" BorderStyle="Groove">DReqQty</asp:textbox>
				<asp:dropdownlist style="Z-INDEX: 167; LEFT: 660px; POSITION: absolute; TOP: 736px" id="DQCCheck15"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="99px">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist>
				<asp:textbox id="DLOSS" style="Z-INDEX: 166; LEFT: 272px; POSITION: absolute; TOP: 872px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="24px" Width="110px" BorderStyle="Groove" MaxLength="10"></asp:textbox>
				<asp:dropdownlist id="DQCCheck13" style="Z-INDEX: 164; LEFT: 24px; POSITION: absolute; TOP: 876px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="54px" Width="110px">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DQCCheck14" style="Z-INDEX: 163; LEFT: 144px; POSITION: absolute; TOP: 872px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="14px" Width="110px">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist>
				<asp:hyperlink id="LEACheckFile1" style="Z-INDEX: 162; LEFT: 624px; POSITION: absolute; TOP: 952px"
					runat="server" Height="3px" Width="131px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">A01-報告</asp:hyperlink>
				<asp:textbox id="DSchedule" style="Z-INDEX: 160; LEFT: 671px; POSITION: absolute; TOP: 464px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="95px" BorderStyle="Groove"
					ReadOnly="True">DSchedule</asp:textbox>
				<asp:textbox id="DPrice" style="Z-INDEX: 159; LEFT: 312px; POSITION: absolute; TOP: 320px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="20px" Width="180px" BorderStyle="Groove" ReadOnly="True"
					AutoPostBack="True" MaxLength="20">DPrice</asp:textbox><asp:hyperlink id="LSuface" style="Z-INDEX: 158; LEFT: 656px; POSITION: absolute; TOP: 56px" runat="server"
					Width="112px" Height="8px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">有追加表面處理</asp:hyperlink><asp:textbox id="DStatus" style="Z-INDEX: 157; LEFT: 560px; POSITION: absolute; TOP: 16px" runat="server"
					ReadOnly="True" BorderStyle="Groove" Width="212px" Height="32px" BackColor="Red" ForeColor="White" Font-Size="10pt">修改工程進行中(xxxxxxxxxxxxxxxxxx)</asp:textbox><asp:panel id="Panel1" style="Z-INDEX: 156; LEFT: 560px; POSITION: absolute; TOP: 8px" runat="server"
					Width="210px" Height="64px" BackColor="White"></asp:panel><asp:dropdownlist id="DQCResult3" style="Z-INDEX: 155; LEFT: 120px; POSITION: absolute; TOP: 1108px"
					runat="server" Width="78px" Height="2px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQCResult2" style="Z-INDEX: 154; LEFT: 120px; POSITION: absolute; TOP: 1078px"
					runat="server" Width="78px" Height="6px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DQCRemark3" style="Z-INDEX: 153; LEFT: 208px; POSITION: absolute; TOP: 1108px"
					runat="server" ReadOnly="True" BorderStyle="Groove" Width="554px" Height="22px" BackColor="Yellow" ForeColor="Blue">DQCRemark3</asp:textbox><asp:textbox id="DQCRemark2" style="Z-INDEX: 152; LEFT: 208px; POSITION: absolute; TOP: 1076px"
					runat="server" ReadOnly="True" BorderStyle="Groove" Width="554px" Height="22px" BackColor="Yellow" ForeColor="Blue">DQCRemark2</asp:textbox><asp:textbox id="DQCDate3" style="Z-INDEX: 151; LEFT: 16px; POSITION: absolute; TOP: 1110px"
					runat="server" ReadOnly="True" BorderStyle="Groove" Width="96px" Height="23px" BackColor="Yellow" ForeColor="Blue">DQCDate3</asp:textbox><asp:textbox id="DQCDate2" style="Z-INDEX: 150; LEFT: 16px; POSITION: absolute; TOP: 1076px"
					runat="server" ReadOnly="True" BorderStyle="Groove" Width="96px" Height="22px" BackColor="Yellow" ForeColor="Blue">DQCDate2</asp:textbox><asp:textbox id="DQCRemark1" style="Z-INDEX: 149; LEFT: 208px; POSITION: absolute; TOP: 1042px"
					runat="server" ReadOnly="True" BorderStyle="Groove" Width="554px" Height="22px" BackColor="Yellow" ForeColor="Blue">DQCRemark1</asp:textbox><asp:dropdownlist id="DQCResult1" style="Z-INDEX: 148; LEFT: 120px; POSITION: absolute; TOP: 1040px"
					runat="server" Width="78px" Height="6px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DQCDate1" style="Z-INDEX: 147; LEFT: 16px; POSITION: absolute; TOP: 1042px"
					runat="server" ReadOnly="True" BorderStyle="Groove" Width="96px" Height="22px" BackColor="Yellow" ForeColor="Blue">DQCDate1</asp:textbox><asp:textbox id="DEADesc1" style="Z-INDEX: 146; LEFT: 520px; POSITION: absolute; TOP: 872px"
					runat="server" ReadOnly="True" BorderStyle="Groove" Width="243px" Height="26px" BackColor="Yellow" ForeColor="Blue">DEADesc1</asp:textbox><asp:dropdownlist id="DEACheck1" style="Z-INDEX: 145; LEFT: 400px; POSITION: absolute; TOP: 872px"
					runat="server" Width="110px" Height="14px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQCCheck12" style="Z-INDEX: 144; LEFT: 656px; POSITION: absolute; TOP: 800px"
					runat="server" Width="110px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQCCheck11" style="Z-INDEX: 143; LEFT: 528px; POSITION: absolute; TOP: 800px"
					runat="server" Width="110px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQCCheck10" style="Z-INDEX: 142; LEFT: 400px; POSITION: absolute; TOP: 800px"
					runat="server" Width="110px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQCCheck9" style="Z-INDEX: 141; LEFT: 272px; POSITION: absolute; TOP: 800px"
					runat="server" Width="110px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQCCheck8" style="Z-INDEX: 140; LEFT: 144px; POSITION: absolute; TOP: 800px"
					runat="server" Width="110px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQCCheck7" style="Z-INDEX: 139; LEFT: 24px; POSITION: absolute; TOP: 800px"
					runat="server" Width="110px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQCCheck5" style="Z-INDEX: 138; LEFT: 528px; POSITION: absolute; TOP: 736px"
					runat="server" Width="110px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQCCheck4" style="Z-INDEX: 137; LEFT: 400px; POSITION: absolute; TOP: 736px"
					runat="server" Width="110px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQCCheck3" style="Z-INDEX: 136; LEFT: 272px; POSITION: absolute; TOP: 736px"
					runat="server" Width="110px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQCCheck2" style="Z-INDEX: 135; LEFT: 144px; POSITION: absolute; TOP: 736px"
					runat="server" Width="110px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DQCCheck1" style="Z-INDEX: 134; LEFT: 24px; POSITION: absolute; TOP: 736px"
					runat="server" Width="110px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:image id="LFinalSampleFile" style="Z-INDEX: 133; LEFT: 16px; POSITION: absolute; TOP: 432px"
					runat="server" BorderStyle="Groove" Width="200px" Height="230px" ImageUrl="http://10.245.0.178//WorkFlow/Document/000003/1-Map2-2006124111750-2.jpg"></asp:image><asp:textbox id="DEnglishName" style="Z-INDEX: 132; LEFT: 440px; POSITION: absolute; TOP: 632px"
					runat="server" ReadOnly="True" BorderStyle="Groove" Width="320px" Height="20px" BackColor="Yellow" ForeColor="Blue">DEnglishName</asp:textbox><asp:textbox id="DCode" style="Z-INDEX: 131; LEFT: 440px; POSITION: absolute; TOP: 592px" runat="server"
					ReadOnly="True" BorderStyle="Groove" Width="320px" Height="20px" BackColor="Yellow" ForeColor="Blue">DCode</asp:textbox><asp:textbox id="DBFinalDate" style="Z-INDEX: 130; LEFT: 440px; POSITION: absolute; TOP: 560px"
					runat="server" ReadOnly="True" BorderStyle="Groove" Width="320px" Height="20px" BackColor="Yellow" ForeColor="Blue">DBFinalDate</asp:textbox><asp:dropdownlist id="DAllowSample" style="Z-INDEX: 129; LEFT: 440px; POSITION: absolute; TOP: 528px"
					runat="server" Width="320px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DQty" style="Z-INDEX: 128; LEFT: 330px; POSITION: absolute; TOP: 496px" runat="server"
					ReadOnly="True" BorderStyle="Groove" Width="88px" Height="20px" BackColor="Yellow" ForeColor="Blue" AutoPostBack="True">DQty</asp:textbox><asp:textbox id="DColor" style="Z-INDEX: 127; LEFT: 330px; POSITION: absolute; TOP: 464px" runat="server"
					ReadOnly="True" BorderStyle="Groove" Width="88px" Height="20px" BackColor="Yellow" ForeColor="Blue" AutoPostBack="True">DColor</asp:textbox><asp:dropdownlist id="DManufType" style="Z-INDEX: 126; LEFT: 330px; POSITION: absolute; TOP: 400px"
					runat="server" Width="424px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DSliderSample" style="Z-INDEX: 124; LEFT: 314px; POSITION: absolute; TOP: 256px"
					runat="server" BorderStyle="Groove" Width="173px" Height="20px" BackColor="Yellow" ForeColor="Blue">DSliderSample</asp:textbox><asp:textbox id="DReqQty" style="Z-INDEX: 123; LEFT: 672px; POSITION: absolute; TOP: 224px" runat="server"
					BorderStyle="Groove" Width="93px" Height="20px" BackColor="Yellow" ForeColor="Blue">DReqQty</asp:textbox><asp:textbox id="DReqDelDate" style="Z-INDEX: 122; LEFT: 312px; POSITION: absolute; TOP: 224px"
					runat="server" ReadOnly="True" BorderStyle="Groove" Width="180px" Height="20px" BackColor="Yellow" ForeColor="Blue">DReqDelDate</asp:textbox><asp:textbox id="DDevReason" style="Z-INDEX: 110; LEFT: 312px; POSITION: absolute; TOP: 352px"
					runat="server" BorderStyle="Groove" Width="450px" Height="24px" BackColor="Yellow" ForeColor="Blue" MaxLength="240" TextMode="MultiLine">DevReason</asp:textbox><asp:image id="DSufaceSheet1" style="Z-INDEX: 100; LEFT: 8px; POSITION: absolute; TOP: 8px"
					runat="server" Width="766px" Height="684px" ImageUrl="Images\SufaceSheet_001_C.jpg"></asp:image><asp:dropdownlist id="DSuppiler" style="Z-INDEX: 121; LEFT: 330px; POSITION: absolute; TOP: 432px"
					runat="server" Width="425px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DBuyer" style="Z-INDEX: 120; LEFT: 312px; POSITION: absolute; TOP: 188px" runat="server"
					Width="180px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:hyperlink id="LQCReqFile" style="Z-INDEX: 119; LEFT: 440px; POSITION: absolute; TOP: 664px"
					runat="server" Width="96px" Height="8px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">品質依賴書</asp:hyperlink><asp:hyperlink id="LOPManualFile" style="Z-INDEX: 117; LEFT: 128px; POSITION: absolute; TOP: 1190px"
					runat="server" Width="96px" Height="3px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">作業標準書</asp:hyperlink><asp:hyperlink id="LEACheckFile" style="Z-INDEX: 116; LEFT: 464px; POSITION: absolute; TOP: 952px"
					runat="server" Width="137px" Height="11px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">Oeko-tex-報告</asp:hyperlink><asp:hyperlink id="LQCFinalFile" style="Z-INDEX: 115; LEFT: 304px; POSITION: absolute; TOP: 950px"
					runat="server" Width="120px" Height="13px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">測試報告書</asp:hyperlink><asp:textbox id="DDate" style="Z-INDEX: 114; LEFT: 584px; POSITION: absolute; TOP: 88px" runat="server"
					ReadOnly="True" BorderStyle="Groove" Width="180px" Height="20px" BackColor="Yellow" ForeColor="Blue">DDate</asp:textbox><asp:hyperlink id="LContactFile" style="Z-INDEX: 113; LEFT: 504px; POSITION: absolute; TOP: 1200px"
					runat="server" Width="96px" Height="14px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">切結書</asp:hyperlink><asp:dropdownlist id="DPerson" style="Z-INDEX: 112; LEFT: 584px; POSITION: absolute; TOP: 122px" runat="server"
					Width="180px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DDivision" style="Z-INDEX: 111; LEFT: 312px; POSITION: absolute; TOP: 122px"
					runat="server" Width="180px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:hyperlink id="LForcastFile" style="Z-INDEX: 109; LEFT: 504px; POSITION: absolute; TOP: 1156px"
					runat="server" Width="96px" Height="1px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">報價單</asp:hyperlink><asp:hyperlink id="LManufFlowFile" style="Z-INDEX: 108; LEFT: 128px; POSITION: absolute; TOP: 1158px"
					runat="server" Width="96px" Height="3px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">製造流程表</asp:hyperlink><asp:textbox id="DSellVendor" style="Z-INDEX: 107; LEFT: 584px; POSITION: absolute; TOP: 188px"
					runat="server" BorderStyle="Groove" Width="180px" Height="20px" BackColor="Yellow" ForeColor="Blue">DSellVendor</asp:textbox><asp:dropdownlist id="DAttachSample" style="Z-INDEX: 106; LEFT: 584px; POSITION: absolute; TOP: 256px"
					runat="server" Width="180px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">DAttachSample</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DORNO" style="Z-INDEX: 105; LEFT: 312px; POSITION: absolute; TOP: 288px" runat="server"
					ReadOnly="True" BorderStyle="Groove" Width="180px" Height="20px" BackColor="Yellow" ForeColor="Blue" AutoPostBack="True">DORNO</asp:textbox><asp:textbox id="DSpec" style="Z-INDEX: 104; LEFT: 312px; POSITION: absolute; TOP: 154px" runat="server"
					ReadOnly="True" BorderStyle="Groove" Width="450px" Height="20px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">DSpec</asp:textbox><asp:textbox id="DNo" style="Z-INDEX: 103; LEFT: 312px; POSITION: absolute; TOP: 88px" runat="server"
					BorderStyle="Groove" Width="180px" Height="20px" BackColor="Yellow" ForeColor="Blue">DNo</asp:textbox><asp:textbox id="DFormSno" style="Z-INDEX: 102; LEFT: 16px; POSITION: absolute; TOP: 1224px"
					runat="server" BorderStyle="None" Width="97px" Height="14px" BackColor="White" ForeColor="Blue">單號：123</asp:textbox><asp:image id="DSufaceSheet2" style="Z-INDEX: 101; LEFT: 8px; POSITION: absolute; TOP: 688px"
					runat="server" Width="766px" Height="538px" ImageUrl="Images\SufaceSheet_002_E.jpg"></asp:image><asp:image id="LCustSampleFile" style="Z-INDEX: 118; LEFT: 16px; POSITION: absolute; TOP: 120px"
					runat="server" BorderStyle="Groove" Width="200px" Height="230px" ImageUrl="http://10.245.0.178//WorkFlow/Document/000003/1-Map2-2006124111750-2.jpg"></asp:image>
				<asp:textbox id="DCap" style="Z-INDEX: 161; LEFT: 512px; POSITION: absolute; TOP: 464px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="20px" Width="64px" BorderStyle="Groove" ReadOnly="True">DCap</asp:textbox></FONT></form>
		</FONT></FORM>
	</body>
</HTML>
