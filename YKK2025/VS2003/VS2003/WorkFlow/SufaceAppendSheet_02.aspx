<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SufaceAppendSheet_02.aspx.vb" Inherits="SPD.SufaceAppendSheet_02"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>表面處理委託-追加型別</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體">
				<asp:image id="DAppendSpecSheet" style="Z-INDEX: 100; LEFT: 8px; POSITION: absolute; TOP: 8px"
					runat="server" ImageUrl="Images\SufaceAppendSheet_001_A.jpg" Width="766px" Height="510px"></asp:image>
				<asp:textbox id="DSpec" style="Z-INDEX: 141; LEFT: 120px; POSITION: absolute; TOP: 228px" runat="server"
					Height="56px" Width="640px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" Font-Size="8pt"
					ReadOnly="True" TextMode="MultiLine">DSpec</asp:textbox>
				<asp:dropdownlist id="DQCNeed" style="Z-INDEX: 140; LEFT: 120px; POSITION: absolute; TOP: 376px" runat="server"
					Height="20px" Width="200px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DBuyer" style="Z-INDEX: 139; LEFT: 120px; POSITION: absolute; TOP: 196px" runat="server"
					Width="246px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow">DBuyer</asp:textbox><asp:hyperlink id="LAttachFile1" style="Z-INDEX: 138; LEFT: 128px; POSITION: absolute; TOP: 488px"
					runat="server" Width="64px" Height="8px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">參考附件</asp:hyperlink><asp:textbox id="DSellVendor" style="Z-INDEX: 136; LEFT: 504px; POSITION: absolute; TOP: 196px"
					runat="server" Width="260px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow">DSellVendor</asp:textbox><asp:hyperlink id="LONo" style="Z-INDEX: 127; LEFT: 720px; POSITION: absolute; TOP: 162px" runat="server"
					Width="41px" Height="8px" NavigateUrl="BoardEdit.aspx" Target="_blank" Font-Size="12pt">連結</asp:hyperlink><asp:textbox id="DOFormNo" style="Z-INDEX: 126; LEFT: 408px; POSITION: absolute; TOP: 376px"
					runat="server" Width="72px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" ReadOnly="True" AutoPostBack="True" Visible="False">DOFormNo</asp:textbox><asp:textbox id="DFormSno" style="Z-INDEX: 113; LEFT: 8px; POSITION: absolute; TOP: 520px" runat="server"
					Width="97px" Height="20px" BorderStyle="None" ForeColor="Blue" BackColor="White">單號：123</asp:textbox>
				<asp:textbox id="DQCRemark" style="Z-INDEX: 112; LEFT: 120px; POSITION: absolute; TOP: 408px"
					runat="server" Width="640px" Height="56px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow"
					TextMode="MultiLine">DQCRemark</asp:textbox><asp:textbox id="DAppendReason" style="Z-INDEX: 111; LEFT: 120px; POSITION: absolute; TOP: 296px"
					runat="server" Width="640px" Height="56px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" TextMode="MultiLine">DAppendReason</asp:textbox><asp:textbox id="DOFormSno" style="Z-INDEX: 110; LEFT: 520px; POSITION: absolute; TOP: 376px"
					runat="server" Width="86px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" ReadOnly="True" AutoPostBack="True" Visible="False">DOFormSno</asp:textbox><asp:textbox id="DCode" style="Z-INDEX: 108; LEFT: 120px; POSITION: absolute; TOP: 162px" runat="server"
					Width="246px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow">DCode</asp:textbox><asp:textbox id="DONo" style="Z-INDEX: 107; LEFT: 504px; POSITION: absolute; TOP: 162px" runat="server"
					Width="210px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow">DONo</asp:textbox><asp:textbox id="DNo" style="Z-INDEX: 106; LEFT: 120px; POSITION: absolute; TOP: 96px" runat="server"
					Width="246px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow">DNo</asp:textbox><asp:textbox id="DDate" style="Z-INDEX: 103; LEFT: 504px; POSITION: absolute; TOP: 96px" runat="server"
					Width="260px" Height="20px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" ReadOnly="True">DDate</asp:textbox><asp:dropdownlist id="DPerson" style="Z-INDEX: 102; LEFT: 504px; POSITION: absolute; TOP: 128px" runat="server"
					Width="260px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DDivision" style="Z-INDEX: 101; LEFT: 120px; POSITION: absolute; TOP: 128px"
					runat="server" Width="246px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist></FONT></form>
	</body>
</HTML>
