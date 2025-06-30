<%@ Page Language="vb" AutoEventWireup="false" Codebehind="ManufOutCTSheet_02.aspx.vb" Inherits="SPD.ManufOutCTSheet_02" aspCompat="True"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>外注業務連絡單</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<script language="javascript" type="text/javascript">
		</script>
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<FONT face="新細明體">
			<FORM id="Form1" method="post" runat="server">
				<FONT face="新細明體">
					<asp:textbox id="DLevel" style="Z-INDEX: 119; LEFT: 416px; POSITION: absolute; TOP: 158px" runat="server"
						Width="176px" Height="20px" ReadOnly="True" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DLevel</asp:textbox>
					<asp:hyperlink id="LNFormNo" style="Z-INDEX: 120; LEFT: 496px; POSITION: absolute; TOP: 228px"
						runat="server" Height="8px" Width="48px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">新委託</asp:hyperlink>
					<asp:image id="DOPContactSheet" style="Z-INDEX: 100; LEFT: 8px; POSITION: absolute; TOP: 8px"
						runat="server" ImageUrl="Images\ContactOutSheet_001.jpg" Width="593px" Height="454px"></asp:image>
					<asp:hyperlink id="LOFormNo" style="Z-INDEX: 116; LEFT: 272px; POSITION: absolute; TOP: 228px"
						runat="server" Width="48px" Height="8px" Target="_blank" NavigateUrl="BoardEdit.aspx" Font-Size="12pt">原委託</asp:hyperlink>
					<asp:hyperlink id="LMapNo" style="Z-INDEX: 115; LEFT: 304px; POSITION: absolute; TOP: 194px" runat="server"
						Width="40px" Height="8px" Target="_blank" NavigateUrl="BoardEdit.aspx" Font-Size="12pt">連結</asp:hyperlink>
					<asp:textbox id="DNFormSno" style="Z-INDEX: 114; LEFT: 416px; POSITION: absolute; TOP: 226px"
						runat="server" Width="72px" Height="20px" AutoPostBack="True" ReadOnly="True" BackColor="Yellow"
						ForeColor="Blue" BorderStyle="Groove">DNFormSno</asp:textbox>
					<asp:textbox id="DNFormNo" style="Z-INDEX: 113; LEFT: 344px; POSITION: absolute; TOP: 226px"
						runat="server" Width="72px" Height="20px" AutoPostBack="True" ReadOnly="True" BackColor="Yellow"
						ForeColor="Blue" BorderStyle="Groove">DNFormNo</asp:textbox>
					<asp:hyperlink id="LAttachFile" style="Z-INDEX: 112; LEFT: 120px; POSITION: absolute; TOP: 436px"
						runat="server" Width="32px" Height="8px" Target="_blank" NavigateUrl="BoardEdit.aspx" Font-Size="12pt">附件</asp:hyperlink>
					<asp:textbox id="DOFormNo" style="Z-INDEX: 111; LEFT: 120px; POSITION: absolute; TOP: 226px"
						runat="server" Width="72px" Height="20px" AutoPostBack="True" ReadOnly="True" BackColor="Yellow"
						ForeColor="Blue" BorderStyle="Groove">DOFormNo</asp:textbox>
					<asp:textbox id="DFormSno" style="Z-INDEX: 110; LEFT: 8px; POSITION: absolute; TOP: 464px" runat="server"
						Width="97px" Height="20px" BackColor="White" ForeColor="Blue" BorderStyle="None">單號：123</asp:textbox>
					<asp:textbox id="DDReason" style="Z-INDEX: 109; LEFT: 120px; POSITION: absolute; TOP: 364px"
						runat="server" Width="472px" Height="56px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove"
						TextMode="MultiLine" ReadOnly="True">DReason</asp:textbox>
					<asp:textbox id="DContent" style="Z-INDEX: 108; LEFT: 120px; POSITION: absolute; TOP: 294px"
						runat="server" Width="472px" Height="56px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove"
						TextMode="MultiLine" ReadOnly="True">DContent</asp:textbox>
					<asp:textbox id="DTarget" style="Z-INDEX: 107; LEFT: 120px; POSITION: absolute; TOP: 260px" runat="server"
						Width="472px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True">DTarget</asp:textbox>
					<asp:textbox id="DOFormSno" style="Z-INDEX: 106; LEFT: 192px; POSITION: absolute; TOP: 226px"
						runat="server" Width="72px" Height="20px" AutoPostBack="True" ReadOnly="True" BackColor="Yellow"
						ForeColor="Blue" BorderStyle="Groove">DOFormSno</asp:textbox>
					<asp:textbox id="DMapNo" style="Z-INDEX: 104; LEFT: 120px; POSITION: absolute; TOP: 192px" runat="server"
						Width="176px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True">DMapNo</asp:textbox>
					<asp:textbox id="DSliderCode" style="Z-INDEX: 103; LEFT: 120px; POSITION: absolute; TOP: 158px"
						runat="server" Width="176px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove"
						ReadOnly="True">DSliderCode</asp:textbox>
					<asp:textbox id="DNo" style="Z-INDEX: 102; LEFT: 120px; POSITION: absolute; TOP: 90px" runat="server"
						Width="176px" Height="20px" ReadOnly="True" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DNo</asp:textbox>
					<asp:textbox id="DDate" style="Z-INDEX: 101; LEFT: 416px; POSITION: absolute; TOP: 90px" runat="server"
						Width="176px" Height="20px" ReadOnly="True" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DDate</asp:textbox>
					<asp:textbox id="DDivision" style="Z-INDEX: 117; LEFT: 120px; POSITION: absolute; TOP: 124px"
						runat="server" Width="176px" Height="20px" ReadOnly="True" BackColor="Yellow" ForeColor="Blue"
						BorderStyle="Groove">DDivision</asp:textbox>
					<asp:textbox id="DPerson" style="Z-INDEX: 118; LEFT: 416px; POSITION: absolute; TOP: 124px" runat="server"
						Width="176px" Height="20px" ReadOnly="True" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DPerson</asp:textbox>
				</FONT>
			</FORM>
		</FONT>
	</body>
</HTML>
