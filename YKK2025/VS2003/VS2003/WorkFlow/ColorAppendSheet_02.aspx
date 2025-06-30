<%@ Page Language="vb" AutoEventWireup="false" Codebehind="ColorAppendSheet_02.aspx.vb" Inherits="SPD.ColorAppendSheet_02"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>ColorAppendSheet_02</title>
		<meta content="False" name="vs_snapToGrid">
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<FORM id="Form1" method="post" runat="server">
			<FONT face="新細明體">
				<asp:image id="DColorAppendSheet" style="Z-INDEX: 100; LEFT: 8px; POSITION: absolute; TOP: -12px"
					runat="server" ImageUrl="Images\ColorAppendSheet_003.jpg" Height="440px" Width="632px"></asp:image>
				<asp:dropdownlist id="DDivision" style="Z-INDEX: 117; LEFT: 120px; POSITION: absolute; TOP: 124px"
					runat="server" Width="192px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DFormSno" style="Z-INDEX: 116; LEFT: 9px; POSITION: absolute; TOP: 431px" runat="server"
					Height="20px" Width="97px" ForeColor="Blue" BackColor="White" BorderStyle="None">單號：123</asp:textbox><asp:textbox id="DLevel" style="Z-INDEX: 115; LEFT: 434px; POSITION: absolute; TOP: 158px" runat="server"
					Height="20px" Width="198px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">DLevel</asp:textbox><asp:hyperlink id="LForCastFile" style="Z-INDEX: 114; LEFT: 120px; POSITION: absolute; TOP: 397px"
					runat="server" Height="8px" Width="64px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">報價單</asp:hyperlink><asp:textbox id="DPrice" style="Z-INDEX: 113; LEFT: 434px; POSITION: absolute; TOP: 364px" runat="server"
					Height="20px" Width="145px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">DPrice</asp:textbox><asp:textbox id="DPullerPrice" style="Z-INDEX: 112; LEFT: 120px; POSITION: absolute; TOP: 364px"
					runat="server" Height="20px" Width="145px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">DPullerPrice</asp:textbox><asp:hyperlink id="LOFormNo" style="Z-INDEX: 111; LEFT: 256px; POSITION: absolute; TOP: 226px"
					runat="server" Height="8px" Width="64px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">原委託</asp:hyperlink><asp:textbox id="DOFormNo" style="Z-INDEX: 110; LEFT: 120px; POSITION: absolute; TOP: 226px"
					runat="server" Height="20px" Width="72px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" AutoPostBack="True" ReadOnly="True">DOFormNo</asp:textbox><asp:textbox id="DManufFlow" style="Z-INDEX: 109; LEFT: 120px; POSITION: absolute; TOP: 293px"
					runat="server" Height="56px" Width="512px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" TextMode="MultiLine">DManufFlow</asp:textbox><asp:textbox id="DColorItem" style="Z-INDEX: 108; LEFT: 120px; POSITION: absolute; TOP: 260px"
					runat="server" Height="20px" Width="512px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">DColorItem</asp:textbox><asp:textbox id="DOFormSno" style="Z-INDEX: 107; LEFT: 200px; POSITION: absolute; TOP: 226px"
					runat="server" Height="20px" Width="48px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" AutoPostBack="True" ReadOnly="True">DOFormSno</asp:textbox><asp:textbox id="DMapNo" style="Z-INDEX: 106; LEFT: 120px; POSITION: absolute; TOP: 192px" runat="server"
					Height="20px" Width="200px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">DMapNo</asp:textbox><asp:textbox id="DSliderCode" style="Z-INDEX: 105; LEFT: 120px; POSITION: absolute; TOP: 158px"
					runat="server" Height="20px" Width="200px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">DSliderCode</asp:textbox><asp:textbox id="DNo" style="Z-INDEX: 104; LEFT: 120px; POSITION: absolute; TOP: 90px" runat="server"
					Height="20px" Width="200px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">DNo</asp:textbox><asp:textbox id="DDate" style="Z-INDEX: 103; LEFT: 434px; POSITION: absolute; TOP: 90px" runat="server"
					Height="20px" Width="197px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DDate</asp:textbox><asp:dropdownlist id="DPerson" style="Z-INDEX: 101; LEFT: 434px; POSITION: absolute; TOP: 124px" runat="server"
					Height="20px" Width="192px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist></FONT></FORM>
	</body>
</HTML>
