<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MapModSheet_02.aspx.vb" Inherits="SPD.MapSheetMod_02"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>圖面修改委託書</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體">
				<asp:image id="DMapSheet1" style="Z-INDEX: 100; LEFT: 8px; POSITION: absolute; TOP: 8px" runat="server"
					Height="554px" Width="593px" ImageUrl="Images\MapSheetMod_002_B.jpg"></asp:image>
				<asp:hyperlink id="LPdfFile" style="Z-INDEX: 157; LEFT: 448px; POSITION: absolute; TOP: 608px"
					runat="server" Width="32px" Height="16px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">圖檔</asp:hyperlink>
				<asp:textbox id="DSuppiler" style="Z-INDEX: 133; LEFT: 304px; POSITION: absolute; TOP: 528px"
					runat="server" Width="128px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow"
					ReadOnly="True">DSuppiler</asp:textbox>
				<asp:textbox id="DCpsc" style="Z-INDEX: 128; LEFT: 544px; POSITION: absolute; TOP: 528px" runat="server"
					Width="49px" Height="20px" ReadOnly="True" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DCpsc</asp:textbox>
				<asp:textbox id="DManufType" style="Z-INDEX: 127; LEFT: 168px; POSITION: absolute; TOP: 528px"
					runat="server" Width="128px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow"
					ReadOnly="True">DManufType</asp:textbox>
				<asp:hyperlink id="LBefMap" style="Z-INDEX: 125; LEFT: 456px; POSITION: absolute; TOP: 170px" runat="server"
					Width="50px" Height="8px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">前圖號</asp:hyperlink>
				<asp:hyperlink id="LOriMap" style="Z-INDEX: 124; LEFT: 168px; POSITION: absolute; TOP: 170px" runat="server"
					Width="50px" Height="8px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">原圖號</asp:hyperlink>
				<asp:textbox id="DSample" style="Z-INDEX: 123; LEFT: 540px; POSITION: absolute; TOP: 358px" runat="server"
					Width="52px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True">DSample</asp:textbox>
				<asp:textbox id="DLevel" style="Z-INDEX: 122; LEFT: 543px; POSITION: absolute; TOP: 570px" runat="server"
					Width="50px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True">DLevel</asp:textbox>
				<asp:textbox id="DMakeMap" style="Z-INDEX: 121; LEFT: 394px; POSITION: absolute; TOP: 570px"
					runat="server" Width="72px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow"
					ReadOnly="True">DMakeMap</asp:textbox>
				<asp:textbox id="DModReasonCode" style="Z-INDEX: 120; LEFT: 168px; POSITION: absolute; TOP: 290px"
					runat="server" Width="424px" Height="20px" ReadOnly="True" BorderStyle="Groove" ForeColor="Blue"
					BackColor="Yellow">DModReasonCode</asp:textbox>
				<asp:textbox id="DPerson" style="Z-INDEX: 119; LEFT: 456px; POSITION: absolute; TOP: 246px" runat="server"
					Width="136px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" ReadOnly="True">DPerson</asp:textbox>
				<asp:textbox id="DDivision" style="Z-INDEX: 118; LEFT: 168px; POSITION: absolute; TOP: 246px"
					runat="server" Width="128px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow"
					ReadOnly="True">DDivision</asp:textbox><asp:textbox id="DModContent" style="Z-INDEX: 117; LEFT: 64px; POSITION: absolute; TOP: 392px"
					runat="server" Height="88px" Width="528px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" TextMode="MultiLine" ReadOnly="True">DModContent</asp:textbox><asp:textbox id="DModReasonDesc" style="Z-INDEX: 105; LEFT: 168px; POSITION: absolute; TOP: 324px"
					runat="server" Height="20px" Width="424px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True">DModReasonDesc</asp:textbox><asp:textbox id="DBefFormSno" style="Z-INDEX: 116; LEFT: 520px; POSITION: absolute; TOP: 168px"
					runat="server" Height="20px" Width="48px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True" Visible="False">DBefFormSno</asp:textbox><asp:textbox id="DBefFormNo" style="Z-INDEX: 115; LEFT: 456px; POSITION: absolute; TOP: 168px"
					runat="server" Height="20px" Width="64px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True" Visible="False">DBefFormNo</asp:textbox><asp:textbox id="DOriFormSno" style="Z-INDEX: 114; LEFT: 224px; POSITION: absolute; TOP: 168px"
					runat="server" Height="20px" Width="32px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True" Visible="False">DOriFormSno</asp:textbox><asp:textbox id="DOriFormNo" style="Z-INDEX: 113; LEFT: 168px; POSITION: absolute; TOP: 168px"
					runat="server" Height="20px" Width="56px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True" Visible="False">DOriFormNo</asp:textbox><asp:textbox id="DBefMapNo" style="Z-INDEX: 112; LEFT: 456px; POSITION: absolute; TOP: 134px"
					runat="server" Height="20px" Width="137px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True">DBefMapNo</asp:textbox><asp:textbox id="DOriMapNo" style="Z-INDEX: 111; LEFT: 168px; POSITION: absolute; TOP: 134px"
					runat="server" Height="20px" Width="128px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" AutoPostBack="True" ReadOnly="True">DOriMapNo</asp:textbox>
				<asp:hyperlink id="LRefAttach" style="Z-INDEX: 110; LEFT: 168px; POSITION: absolute; TOP: 496px"
					runat="server" Height="8px" Width="40px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">附件</asp:hyperlink><asp:textbox id="DBuyer" style="Z-INDEX: 109; LEFT: 168px; POSITION: absolute; TOP: 212px" runat="server"
					Height="20px" Width="128px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True">DBuyer</asp:textbox><asp:textbox id="DSellVendor" style="Z-INDEX: 104; LEFT: 456px; POSITION: absolute; TOP: 212px"
					runat="server" Height="20px" Width="136px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True">DSellVendor</asp:textbox><asp:textbox id="DNo" style="Z-INDEX: 108; LEFT: 168px; POSITION: absolute; TOP: 90px" runat="server"
					Height="20px" Width="128px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True">DNo</asp:textbox><asp:textbox id="DDate" style="Z-INDEX: 102; LEFT: 456px; POSITION: absolute; TOP: 90px" runat="server"
					Height="20px" Width="137px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True">DDate</asp:textbox>
				<asp:hyperlink id="LMapFile" style="Z-INDEX: 106; LEFT: 168px; POSITION: absolute; TOP: 609px"
					runat="server" Height="8px" Width="40px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">圖檔</asp:hyperlink><asp:textbox id="DMapNo" style="Z-INDEX: 107; LEFT: 168px; POSITION: absolute; TOP: 570px" runat="server"
					Height="20px" Width="152px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True">DMapNo</asp:textbox><asp:textbox id="DFormSno" style="Z-INDEX: 103; LEFT: 16px; POSITION: absolute; TOP: 640px" runat="server"
					Height="20px" Width="97px" BackColor="White" ForeColor="Blue" BorderStyle="None">單號：123</asp:textbox><asp:image id="DMapSheet3" style="Z-INDEX: 101; LEFT: 8px; POSITION: absolute; TOP: 560px"
					runat="server" Height="107px" Width="593px" ImageUrl="Images\MapSheetMod_002_C1.jpg"></asp:image></FONT></form>
	</body>
</HTML>
