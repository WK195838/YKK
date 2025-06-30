<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MapSheet_02.aspx.vb" Inherits="SPD.MapSheet_02"%>
<HTML>
	<HEAD>
		<title>製圖委託書</title>
		<meta content="True" name="vs_showGrid">
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<style type="text/css">BODY { FONT-SIZE: 9pt; FONT-FAMILY: Verdana, Geneva, Sans-Serif }
	TABLE { FONT-SIZE: 9pt; FONT-FAMILY: Verdana, Geneva, Sans-Serif }
	TR { FONT-SIZE: 9pt; FONT-FAMILY: Verdana, Geneva, Sans-Serif }
	TD { FONT-SIZE: 9pt; FONT-FAMILY: Verdana, Geneva, Sans-Serif }
	UL { FONT-SIZE: 9pt; FONT-FAMILY: Verdana, Geneva, Sans-Serif }
	LI { FONT-SIZE: 9pt; FONT-FAMILY: Verdana, Geneva, Sans-Serif }
	.normal { FONT-SIZE: 9pt; FONT-FAMILY: Verdana, Geneva, Sans-Serif }
	H1 { FONT-WEIGHT: 900; FONT-SIZE: 10.5pt; COLOR: #666666; FONT-FAMILY: Verdana, Geneva, Sans-Serif }
	.small { FONT-SIZE: 7.5pt; FONT-FAMILY: Verdana, Geneva, Sans-Serif }
	.error { COLOR: #ff0033 }
	.required { FONT-WEIGHT: 900; COLOR: #ff0033; FONT-FAMILY: Verdana, Geneva, Sans-Serif }
	.success { FONT-WEIGHT: 900; FONT-SIZE: 11pt; MARGIN: 10px 0px; COLOR: #009933 }
		</style>
		<script language="javascript" type="text/javascript">
			function calendarPicker(strField)
			{
				window.open('DatePicker.aspx?field=' + strField,'calendarPopup','width=250,height=190,resizable=yes');
			}
			function MapPicker(strField)
			{
				window.open('MapPicker.aspx?field=' + strField,'MapPopup','width=168,height=328,resizable=yes');
			}
		</script>
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" encType="multipart/form-data" runat="server">
			<FONT face="新細明體">
				<asp:image id="DMapSheet1" style="Z-INDEX: 101; LEFT: 8px; POSITION: absolute; TOP: 8px" runat="server"
					Height="597px" Width="593px" ImageUrl="Images\MapSheet_001_d.jpg"></asp:image>
				<asp:hyperlink id="LPdfFile" style="Z-INDEX: 157; LEFT: 448px; POSITION: absolute; TOP: 656px"
					runat="server" Width="40px" Height="2px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">圖檔</asp:hyperlink>
				<asp:textbox id="DSuppiler" style="Z-INDEX: 133; LEFT: 304px; POSITION: absolute; TOP: 438px"
					runat="server" Width="128px" Height="20px" ReadOnly="True" ForeColor="Blue" BackColor="Yellow"
					BorderStyle="Groove">DSuppiler</asp:textbox>
				<asp:textbox id="DManufType" style="Z-INDEX: 128; LEFT: 168px; POSITION: absolute; TOP: 438px"
					runat="server" Width="128px" Height="20px" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue"
					ReadOnly="True">DManufType</asp:textbox>
				<asp:textbox id="DFormSno" style="Z-INDEX: 127; LEFT: 8px; POSITION: absolute; TOP: 688px" runat="server"
					Width="97px" Height="20px" ForeColor="Blue" BackColor="White" BorderStyle="None">單號：123</asp:textbox>
				<asp:image id="Image1" style="Z-INDEX: 126; LEFT: 0px; POSITION: absolute; TOP: 680px" runat="server"
					ImageUrl="Images\MapSheet_006.jpg" Width="593px" Height="34px"></asp:image>
				<asp:textbox id="DSample" style="Z-INDEX: 125; LEFT: 544px; POSITION: absolute; TOP: 336px" runat="server"
					Width="48px" Height="20px" ReadOnly="True" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue">DSample</asp:textbox>
				<asp:textbox id="DMakeMap" style="Z-INDEX: 124; LEFT: 394px; POSITION: absolute; TOP: 618px"
					runat="server" Width="72px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove"
					ReadOnly="True">DMakeMap</asp:textbox>
				<asp:textbox id="DHalfFinish" style="Z-INDEX: 123; LEFT: 456px; POSITION: absolute; TOP: 404px"
					runat="server" Width="137px" Height="20px" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue"
					ReadOnly="True">DHalfFinish</asp:textbox>
				<asp:textbox id="DCPSC" style="Z-INDEX: 122; LEFT: 168px; POSITION: absolute; TOP: 234px" runat="server"
					Width="128px" Height="20px" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue" ReadOnly="True">DCPSC</asp:textbox>
				<asp:textbox id="DPerson" style="Z-INDEX: 121; LEFT: 456px; POSITION: absolute; TOP: 166px" runat="server"
					Width="136px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DPerson</asp:textbox>
				<asp:textbox id="DDivision" style="Z-INDEX: 120; LEFT: 168px; POSITION: absolute; TOP: 166px"
					runat="server" Width="128px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove"
					ReadOnly="True">DDivision</asp:textbox>
				<asp:textbox id="DLight" style="Z-INDEX: 119; LEFT: 168px; POSITION: absolute; TOP: 268px" runat="server"
					Width="128px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DLight</asp:textbox>
				<asp:textbox id="DSpec" style="Z-INDEX: 118; LEFT: 56px; POSITION: absolute; TOP: 336px" runat="server"
					Width="472px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DBody</asp:textbox>
				<asp:textbox id="DFrontBack" style="Z-INDEX: 117; LEFT: 168px; POSITION: absolute; TOP: 404px"
					runat="server" Width="128px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove"
					ReadOnly="True">DFrontBack</asp:textbox>
				<asp:textbox id="DMaterialDetail" style="Z-INDEX: 116; LEFT: 168px; POSITION: absolute; TOP: 504px"
					runat="server" Width="424px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove"
					ReadOnly="True">DMaterialDetail</asp:textbox>
				<asp:textbox id="DMaterial" style="Z-INDEX: 115; LEFT: 168px; POSITION: absolute; TOP: 472px"
					runat="server" Width="128px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove"
					ReadOnly="True">DMaterial</asp:textbox><asp:textbox id="DLevel" style="Z-INDEX: 114; LEFT: 542px; POSITION: absolute; TOP: 618px" runat="server"
					Height="20px" Width="50px" ReadOnly="True" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">DLevel</asp:textbox><asp:hyperlink id="LMapFile" style="Z-INDEX: 113; LEFT: 168px; POSITION: absolute; TOP: 656px"
					runat="server" Height="8px" Width="64px" Target="_blank" NavigateUrl="BoardEdit.aspx" Font-Size="12pt">圖檔</asp:hyperlink><asp:hyperlink id="LRefMapFile" style="Z-INDEX: 112; LEFT: 168px; POSITION: absolute; TOP: 544px"
					runat="server" Height="8px" Width="64px" Target="_blank" NavigateUrl="BoardEdit.aspx" Font-Size="12pt">草圖</asp:hyperlink>
				<asp:textbox id="DMapNo" style="Z-INDEX: 111; LEFT: 168px; POSITION: absolute; TOP: 618px" runat="server"
					Height="20px" Width="152px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DMapNo</asp:textbox><asp:textbox id="DSurface" style="Z-INDEX: 108; LEFT: 456px; POSITION: absolute; TOP: 370px"
					runat="server" Height="20px" Width="137px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DSurface</asp:textbox><asp:textbox id="DNo" style="Z-INDEX: 102; LEFT: 168px; POSITION: absolute; TOP: 88px" runat="server"
					Height="20px" Width="128px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DNo</asp:textbox></FONT><asp:textbox id="DDate" style="Z-INDEX: 103; LEFT: 456px; POSITION: absolute; TOP: 88px" runat="server"
				Height="20px" Width="136px" ReadOnly="True" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">DDate</asp:textbox><asp:textbox id="DBuyer" style="Z-INDEX: 104; LEFT: 168px; POSITION: absolute; TOP: 132px" runat="server"
				Height="20px" Width="128px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DBuyer</asp:textbox><asp:textbox id="DSellVendor" style="Z-INDEX: 105; LEFT: 456px; POSITION: absolute; TOP: 132px"
				runat="server" Height="20px" Width="136px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DSellVendor</asp:textbox><asp:textbox id="DBackground" style="Z-INDEX: 106; LEFT: 168px; POSITION: absolute; TOP: 200px"
				runat="server" Height="20px" Width="424px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DBackground</asp:textbox><asp:textbox id="DCramper" style="Z-INDEX: 107; LEFT: 168px; POSITION: absolute; TOP: 370px"
				runat="server" Height="20px" Width="128px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DCramper</asp:textbox><asp:textbox id="DDescription" style="Z-INDEX: 109; LEFT: 168px; POSITION: absolute; TOP: 575px"
				runat="server" Height="20px" Width="424px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DDescription</asp:textbox><asp:textbox id="DMapReqDelDate" style="Z-INDEX: 110; LEFT: 456px; POSITION: absolute; TOP: 234px"
				runat="server" Height="20px" Width="136px" ReadOnly="True" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">DMapReqDelDate</asp:textbox><asp:image id="DMapSheet3" style="Z-INDEX: 100; LEFT: 8px; POSITION: absolute; TOP: 608px"
				runat="server" Height="77px" Width="593px" ImageUrl="Images\MapSheet_003_A1.JPG"></asp:image>
			<asp:Panel id="Panel1" style="Z-INDEX: 130; LEFT: 392px; POSITION: absolute; TOP: 8px" runat="server"
				BackColor="White" Width="210px" Height="64px"></asp:Panel></FONT>
			<asp:textbox id="DStatus" style="Z-INDEX: 132; LEFT: 392px; POSITION: absolute; TOP: 16px" runat="server"
				Font-Size="10pt" ForeColor="White" BackColor="Red" Height="32px" Width="212px" BorderStyle="Groove"
				ReadOnly="True">修改工程進行中(xxxxxxxxxxxxxxxxxx)</asp:textbox>
			<asp:hyperlink id="LMMap" style="Z-INDEX: 131; LEFT: 520px; POSITION: absolute; TOP: 56px" runat="server"
				Font-Size="12pt" Height="8px" Width="80px" Target="_blank" NavigateUrl="BoardEdit.aspx">有修改圖面</asp:hyperlink>
		</form>
	</body>
</HTML>
