<%@ Page Language="vb" AutoEventWireup="false" Codebehind="ManufOutSDSheet_02.aspx.vb" Inherits="SPD.SliderDetailSheet_02"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>外注拉頭細目</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<FORM id="Form1" method="post" runat="server">
			<FONT face="新細明體">
				<asp:image id="DSliderDetailSheet" style="Z-INDEX: 100; LEFT: 8px; POSITION: absolute; TOP: 8px"
					runat="server" ImageUrl="Images\SliderDetailSheet_002.jpg" Width="593px" Height="314px"></asp:image>
				<asp:textbox id="DContent" style="Z-INDEX: 135; LEFT: 120px; POSITION: absolute; TOP: 188px"
					runat="server" Height="56px" Width="472px" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow"
					TextMode="MultiLine" ReadOnly="True">DContent</asp:textbox>
				<asp:textbox id="DPerson" style="Z-INDEX: 129; LEFT: 416px; POSITION: absolute; TOP: 122px" runat="server"
					Width="176px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True">DPerson</asp:textbox>
				<asp:textbox id="DDivision" style="Z-INDEX: 128; LEFT: 120px; POSITION: absolute; TOP: 122px"
					runat="server" Width="176px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove"
					ReadOnly="True">DDivision</asp:textbox>
				<asp:textbox id="DFormSno" style="Z-INDEX: 115; LEFT: 8px; POSITION: absolute; TOP: 320px" runat="server"
					Width="97px" Height="20px" BackColor="White" ForeColor="Blue" BorderStyle="None">單號：123</asp:textbox>
				<asp:textbox id="DOFormNo" style="Z-INDEX: 113; LEFT: 416px; POSITION: absolute; TOP: 156px"
					runat="server" Width="56px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove"
					AutoPostBack="True" ReadOnly="True">DOFormNo</asp:textbox>
				<asp:textbox id="DOFormSno" style="Z-INDEX: 111; LEFT: 472px; POSITION: absolute; TOP: 156px"
					runat="server" Width="56px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove"
					AutoPostBack="True" ReadOnly="True">DOFormSno</asp:textbox>
				<asp:hyperlink id="LOFormNo" style="Z-INDEX: 112; LEFT: 536px; POSITION: absolute; TOP: 158px"
					runat="server" Width="48px" Height="8px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">原委託</asp:hyperlink>
				<asp:hyperlink id="LAttachFile" style="Z-INDEX: 110; LEFT: 120px; POSITION: absolute; TOP: 294px"
					runat="server" Width="32px" Height="8px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">附件</asp:hyperlink>
				<asp:hyperlink id="LSliderFile" style="Z-INDEX: 107; LEFT: 120px; POSITION: absolute; TOP: 262px"
					runat="server" Width="80px" Height="8px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">拉頭細目表</asp:hyperlink>
				<asp:textbox id="DNo" style="Z-INDEX: 105; LEFT: 120px; POSITION: absolute; TOP: 88px" runat="server"
					Width="176px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True">DNo</asp:textbox>
				<asp:textbox id="DDate" style="Z-INDEX: 102; LEFT: 416px; POSITION: absolute; TOP: 88px" runat="server"
					Width="176px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True">DDate</asp:textbox>
				<asp:textbox id="DSliderGRCode" style="Z-INDEX: 103; LEFT: 120px; POSITION: absolute; TOP: 156px"
					runat="server" Width="176px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove"
					ReadOnly="True">DSliderGRCode</asp:textbox></FONT></FORM>
	</body>
</HTML>
