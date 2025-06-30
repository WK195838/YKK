<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SCD_SampleSheet_02.aspx.vb" Inherits="SPD.SCD_SampleSheet_02"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>開發見本委託書</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" encType="multipart/form-data" runat="server">
			<FONT face="新細明體">
				<asp:HyperLink id="LQCFILE5" style="Z-INDEX: 155; LEFT: 656px; POSITION: absolute; TOP: 680px"
					runat="server" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt" Height="8px">其它</asp:HyperLink>&lt;
				<asp:hyperlink id="LQCFILE1" style="Z-INDEX: 102; LEFT: 144px; POSITION: absolute; TOP: 680px"
					runat="server" Height="8px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">吋法檔案</asp:hyperlink>
				<asp:HyperLink ID="LQCFILE4" runat="server" Font-Size="12pt" Height="8px" NavigateUrl="BoardEdit.aspx"
					Style="Z-INDEX: 103; LEFT: 500px; POSITION: absolute; TOP: 680px" Target="_blank">式樣書與組織圖檔案</asp:HyperLink>
				<asp:HyperLink ID="LQCFILE3" runat="server" Font-Size="12pt" Height="8px" NavigateUrl="BoardEdit.aspx"
					Style="Z-INDEX: 105; LEFT: 376px; POSITION: absolute; TOP: 680px" Target="_blank">往覆測試檔案</asp:HyperLink>
				<asp:HyperLink ID="LQCFILE2" runat="server" Font-Size="12pt" Height="8px" NavigateUrl="BoardEdit.aspx"
					Style="Z-INDEX: 106; LEFT: 264px; POSITION: absolute; TOP: 680px" Target="_blank">強度檔案</asp:HyperLink>
				<asp:textbox id="DWF7" style="Z-INDEX: 107; LEFT: 640px; POSITION: absolute; TOP: 1040px" runat="server"
					Width="74px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">WF7</asp:textbox>
				<asp:textbox id="DWF6" style="Z-INDEX: 108; LEFT: 540px; POSITION: absolute; TOP: 1040px" runat="server"
					Width="74px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">WF6</asp:textbox>
				<asp:textbox id="DWF5" style="Z-INDEX: 109; LEFT: 440px; POSITION: absolute; TOP: 1040px" runat="server"
					Width="74px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">WF5</asp:textbox>
				<asp:textbox id="DWF4" style="Z-INDEX: 110; LEFT: 340px; POSITION: absolute; TOP: 1040px" runat="server"
					Width="74px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">WF4</asp:textbox>
				<asp:textbox id="DWF3" style="Z-INDEX: 111; LEFT: 240px; POSITION: absolute; TOP: 1040px" runat="server"
					Width="74px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">WF3</asp:textbox>
				<asp:textbox id="DWF2" style="Z-INDEX: 112; LEFT: 140px; POSITION: absolute; TOP: 1040px" runat="server"
					Width="74px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">WF2</asp:textbox>
				<asp:textbox id="DWF1" style="Z-INDEX: 113; LEFT: 40px; POSITION: absolute; TOP: 1040px" runat="server"
					Width="74px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">WF1</asp:textbox>
				<asp:textbox id="DWF7Name" style="Z-INDEX: 114; LEFT: 640px; POSITION: absolute; TOP: 1012px"
					runat="server" Width="74px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">WF7Name</asp:textbox>
				<asp:textbox id="DWF6Name" style="Z-INDEX: 115; LEFT: 540px; POSITION: absolute; TOP: 1012px"
					runat="server" Width="74px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">WF6Name</asp:textbox>
				<asp:textbox id="DWF5Name" style="Z-INDEX: 116; LEFT: 440px; POSITION: absolute; TOP: 1012px"
					runat="server" Width="74px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">WF5Name</asp:textbox>
				<asp:textbox id="DWF4Name" style="Z-INDEX: 117; LEFT: 340px; POSITION: absolute; TOP: 1012px"
					runat="server" Width="74px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">WF4Name</asp:textbox>
				<asp:textbox id="DWF3Name" style="Z-INDEX: 118; LEFT: 240px; POSITION: absolute; TOP: 1012px"
					runat="server" Width="74px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">WF3Name</asp:textbox><asp:textbox id="DOP6" style="Z-INDEX: 119; LEFT: 660px; POSITION: absolute; TOP: 948px" runat="server"
					ForeColor="Blue" Height="20px" Width="74px" BackColor="Yellow" BorderStyle="Groove">OP6</asp:textbox><asp:textbox id="DOP5" style="Z-INDEX: 120; LEFT: 560px; POSITION: absolute; TOP: 948px" runat="server"
					ForeColor="Blue" Height="20px" Width="74px" BackColor="Yellow" BorderStyle="Groove">OP5</asp:textbox><asp:textbox id="DOP4" style="Z-INDEX: 121; LEFT: 460px; POSITION: absolute; TOP: 948px" runat="server"
					ForeColor="Blue" Height="20px" Width="74px" BackColor="Yellow" BorderStyle="Groove">OP4</asp:textbox><asp:textbox id="DOP3" style="Z-INDEX: 122; LEFT: 360px; POSITION: absolute; TOP: 948px" runat="server"
					ForeColor="Blue" Height="20px" Width="74px" BackColor="Yellow" BorderStyle="Groove">OP3</asp:textbox><asp:textbox id="DOP2" style="Z-INDEX: 123; LEFT: 260px; POSITION: absolute; TOP: 948px" runat="server"
					ForeColor="Blue" Height="20px" Width="74px" BackColor="Yellow" BorderStyle="Groove">OP2</asp:textbox><asp:textbox id="DOP1" style="Z-INDEX: 124; LEFT: 160px; POSITION: absolute; TOP: 948px" runat="server"
					ForeColor="Blue" Height="20px" Width="74px" BackColor="Yellow" BorderStyle="Groove">OP1</asp:textbox><asp:textbox id="DCITEM" style="Z-INDEX: 125; LEFT: 280px; POSITION: absolute; TOP: 904px" runat="server"
					ForeColor="Blue" Height="19px" Width="312px" BackColor="Yellow" BorderStyle="Groove">CITEM</asp:textbox>
				<asp:TextBox ID="DO1Item" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
					Height="19px" Style="Z-INDEX: 126; LEFT: 280px; POSITION: absolute; TOP: 856px" Width="312px"></asp:TextBox>
				<asp:TextBox ID="DO2Item" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
					Height="19px" Style="Z-INDEX: 127; LEFT: 280px; POSITION: absolute; TOP: 880px" Width="312px"></asp:TextBox>
				<asp:textbox id="DCDITEM" style="Z-INDEX: 128; LEFT: 280px; POSITION: absolute; TOP: 836px" runat="server"
					ForeColor="Blue" Height="19px" Width="312px" BackColor="Yellow" BorderStyle="Groove">CDITEM</asp:textbox><asp:textbox id="DCSITEM" style="Z-INDEX: 129; LEFT: 280px; POSITION: absolute; TOP: 812px" runat="server"
					ForeColor="Blue" Height="19px" Width="312px" BackColor="Yellow" BorderStyle="Groove">CSITEM</asp:textbox><asp:textbox id="DTDRITEM" style="Z-INDEX: 130; LEFT: 500px; POSITION: absolute; TOP: 764px"
					runat="server" ForeColor="Blue" Height="19px" Width="92px" BackColor="Yellow" BorderStyle="Groove">TDRITEM</asp:textbox><asp:textbox id="DTSRITEM" style="Z-INDEX: 131; LEFT: 500px; POSITION: absolute; TOP: 740px"
					runat="server" ForeColor="Blue" Height="19px" Width="92px" BackColor="Yellow" BorderStyle="Groove">TSRITEM</asp:textbox><asp:textbox id="DTNRITEM" style="Z-INDEX: 132; LEFT: 500px; POSITION: absolute; TOP: 720px"
					runat="server" ForeColor="Blue" Height="19px" Width="92px" BackColor="Yellow" BorderStyle="Groove">TNRITEM</asp:textbox><asp:textbox id="DCNITEM" style="Z-INDEX: 133; LEFT: 280px; POSITION: absolute; TOP: 788px" runat="server"
					ForeColor="Blue" Height="19px" Width="312px" BackColor="Yellow" BorderStyle="Groove">CNITEM</asp:textbox><asp:textbox id="DTDLITEM" style="Z-INDEX: 134; LEFT: 280px; POSITION: absolute; TOP: 764px"
					runat="server" ForeColor="Blue" Height="19px" Width="92px" BackColor="Yellow" BorderStyle="Groove">TDLITEM</asp:textbox><asp:textbox id="DTSLITEM" style="Z-INDEX: 135; LEFT: 280px; POSITION: absolute; TOP: 740px"
					runat="server" ForeColor="Blue" Height="19px" Width="92px" BackColor="Yellow" BorderStyle="Groove">TSLITEM</asp:textbox><asp:textbox id="DTNLITEM" style="Z-INDEX: 136; LEFT: 280px; POSITION: absolute; TOP: 720px"
					runat="server" ForeColor="Blue" Height="19px" Width="92px" BackColor="Yellow" BorderStyle="Groove">TNLITEM</asp:textbox>
				<asp:textbox id="DOTHER" style="Z-INDEX: 137; LEFT: 140px; POSITION: absolute; TOP: 622px" runat="server"
					ForeColor="Blue" Height="48px" Width="608px" BackColor="Yellow" BorderStyle="Groove" TextMode="MultiLine"
					MaxLength="240" Font-Size="7pt">OTHER</asp:textbox><asp:textbox id="DTHCOL" style="Z-INDEX: 138; LEFT: 140px; POSITION: absolute; TOP: 566px" runat="server"
					ForeColor="Blue" Height="48px" Width="608px" BackColor="Yellow" BorderStyle="Groove" TextMode="MultiLine" MaxLength="240">THCOL</asp:textbox><asp:textbox id="DCCOL" style="Z-INDEX: 139; LEFT: 140px; POSITION: absolute; TOP: 538px" runat="server"
					ForeColor="Blue" Height="20px" Width="608px" BackColor="Yellow" BorderStyle="Groove">CCOL</asp:textbox><asp:textbox id="DECOL" style="Z-INDEX: 140; LEFT: 140px; POSITION: absolute; TOP: 510px" runat="server"
					ForeColor="Blue" Height="20px" Width="608px" BackColor="Yellow" BorderStyle="Groove">ECOL</asp:textbox><asp:textbox id="DTACOL" style="Z-INDEX: 141; LEFT: 140px; POSITION: absolute; TOP: 398px" runat="server"
					ForeColor="Blue" Height="48px" Width="608px" BackColor="Yellow" BorderStyle="Groove" TextMode="MultiLine" MaxLength="240">TACOL</asp:textbox><asp:textbox id="DDEVPRD" style="Z-INDEX: 142; LEFT: 580px; POSITION: absolute; TOP: 370px" runat="server"
					ForeColor="Blue" Height="20px" Width="168px" BackColor="Yellow" BorderStyle="Groove">DEVPRD</asp:textbox><asp:textbox id="DDEVNO" style="Z-INDEX: 143; LEFT: 360px; POSITION: absolute; TOP: 370px" runat="server"
					ForeColor="Blue" Height="20px" Width="112px" BackColor="Yellow" BorderStyle="Groove">DEVNO</asp:textbox><asp:textbox id="DTAWIDTH" style="Z-INDEX: 144; LEFT: 140px; POSITION: absolute; TOP: 370px"
					runat="server" ForeColor="Blue" Height="20px" Width="80px" BackColor="Yellow" BorderStyle="Groove">TAWIDTH</asp:textbox>
				<asp:image id="LSAMPLEFILE" style="Z-INDEX: 145; LEFT: 20px; POSITION: absolute; TOP: 202px"
					runat="server" Height="160px" Width="728px" BorderStyle="Groove" ImageUrl="F:\DMF04006-DS2W.jpg"></asp:image><asp:textbox id="DCODENO" style="Z-INDEX: 146; LEFT: 540px; POSITION: absolute; TOP: 150px" runat="server"
					ForeColor="Blue" Height="20px" Width="208px" BackColor="Yellow" BorderStyle="Groove">CODENO</asp:textbox><asp:textbox id="DITEM" style="Z-INDEX: 147; LEFT: 340px; POSITION: absolute; TOP: 150px" runat="server"
					ForeColor="Blue" Height="20px" Width="192px" BackColor="Yellow" BorderStyle="Groove">ITEM</asp:textbox><asp:textbox id="DSIZENO" style="Z-INDEX: 148; LEFT: 140px; POSITION: absolute; TOP: 150px" runat="server"
					ForeColor="Blue" Height="20px" Width="192px" BackColor="Yellow" BorderStyle="Groove">SIZENO</asp:textbox><asp:textbox id="DTALINE" style="Z-INDEX: 149; LEFT: 140px; POSITION: absolute; TOP: 454px" runat="server"
					ForeColor="Blue" Height="48px" Width="608px" BackColor="Yellow" BorderStyle="Groove" TextMode="MultiLine" MaxLength="240">TALINE</asp:textbox><asp:image id="DSampleSheet1" style="Z-INDEX: 100; LEFT: 10px; POSITION: absolute; TOP: 8px"
					runat="server" ImageUrl="Images\SCD_SampleSheet_011.jpg"></asp:image><asp:image id="DSampleSheet2" style="Z-INDEX: 101; LEFT: 8px; POSITION: absolute; TOP: 704px"
					runat="server" ImageUrl="Images\SCD_SampleSheet_02.jpg"></asp:image><asp:textbox id="DDATE" style="Z-INDEX: 150; LEFT: 640px; POSITION: absolute; TOP: 100px" runat="server"
					ForeColor="Blue" Height="20px" Width="109px" BackColor="Yellow" BorderStyle="Groove">DATE</asp:textbox><asp:textbox id="DAPPBUYER" style="Z-INDEX: 151; LEFT: 140px; POSITION: absolute; TOP: 100px"
					runat="server" ForeColor="Blue" Height="20px" Width="392px" BackColor="Yellow" BorderStyle="Groove">APPBUYER</asp:textbox><asp:textbox id="DNo" style="Z-INDEX: 152; LEFT: 8px; POSITION: absolute; TOP: 9px" runat="server"
					ForeColor="Black" Height="16px" Width="216px" Font-Names="Times New Roman" BackColor="White" BorderStyle="None" ReadOnly="True"> 123</asp:textbox>
				<asp:Label ID="DOther1" runat="server" Style="Z-INDEX: 153; LEFT: 160px; POSITION: absolute; TOP: 860px"
					Text="Label" Font-Size="12px">Label</asp:Label>
				<asp:Label ID="DOther2" runat="server" Style="Z-INDEX: 154; LEFT: 160px; POSITION: absolute; TOP: 884px"
					Text="Label" Font-Size="12px">Label</asp:Label></FONT></form>
		</FONT></FORM>
	</body>
</HTML>
