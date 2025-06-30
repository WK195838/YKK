<%@ Page Language="vb" AutoEventWireup="false" Codebehind="ImportModSheet_02.aspx.vb" Inherits="SPD.ImportModSheet_02"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>ImportModSheet_02</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" encType="multipart/form-data" runat="server">
			<FONT face="新細明體">
				<asp:image id="DImportSheet" style="Z-INDEX: 100; LEFT: 8px; POSITION: absolute; TOP: 8px"
					runat="server" Width="761px" Height="644px" ImageUrl="Images\ImportSheet_001.jpg"></asp:image><asp:label id="Label22" style="Z-INDEX: 291; LEFT: 504px; POSITION: absolute; TOP: 506px" runat="server"
					Width="24px" Height="10px" Font-Size="8pt">單價</asp:label><asp:label id="Label21" style="Z-INDEX: 290; LEFT: 464px; POSITION: absolute; TOP: 506px" runat="server"
					Width="24px" Height="10px" Font-Size="8pt">幣別</asp:label><asp:label id="Label20" style="Z-INDEX: 289; LEFT: 296px; POSITION: absolute; TOP: 506px" runat="server"
					Width="48px" Height="10px" Font-Size="8pt">拉頭種類</asp:label><asp:label id="Label19" style="Z-INDEX: 288; LEFT: 224px; POSITION: absolute; TOP: 506px" runat="server"
					Width="24px" Height="10px" Font-Size="8pt">單價</asp:label><asp:label id="Label18" style="Z-INDEX: 287; LEFT: 184px; POSITION: absolute; TOP: 506px" runat="server"
					Width="24px" Height="10px" Font-Size="8pt">幣別</asp:label><asp:label id="Label17" style="Z-INDEX: 286; LEFT: 56px; POSITION: absolute; TOP: 506px" runat="server"
					Width="48px" Height="10px" Font-Size="8pt">拉頭種類</asp:label><asp:dropdownlist id="DBuyer" style="Z-INDEX: 270; LEFT: 312px; POSITION: absolute; TOP: 296px" runat="server"
					Width="176px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceJ1" style="Z-INDEX: 202; LEFT: 296px; POSITION: absolute; TOP: 590px"
					runat="server" Width="166px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DPriceI3" style="Z-INDEX: 201; LEFT: 224px; POSITION: absolute; TOP: 590px"
					runat="server" Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:dropdownlist id="DPriceI2" style="Z-INDEX: 200; LEFT: 184px; POSITION: absolute; TOP: 590px"
					runat="server" Width="40px" Height="12px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceI1" style="Z-INDEX: 199; LEFT: 16px; POSITION: absolute; TOP: 590px" runat="server"
					Width="166px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DPriceH3" style="Z-INDEX: 198; LEFT: 504px; POSITION: absolute; TOP: 572px"
					runat="server" Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:dropdownlist id="DPriceH2" style="Z-INDEX: 197; LEFT: 464px; POSITION: absolute; TOP: 572px"
					runat="server" Width="40px" Height="12px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceJ3" style="Z-INDEX: 204; LEFT: 504px; POSITION: absolute; TOP: 590px"
					runat="server" Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:dropdownlist id="DPriceJ2" style="Z-INDEX: 203; LEFT: 464px; POSITION: absolute; TOP: 590px"
					runat="server" Width="40px" Height="12px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceA3" style="Z-INDEX: 150; LEFT: 224px; POSITION: absolute; TOP: 518px"
					runat="server" Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:dropdownlist id="DPriceA2" style="Z-INDEX: 149; LEFT: 184px; POSITION: absolute; TOP: 518px"
					runat="server" Width="40px" Height="12px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceH1" style="Z-INDEX: 196; LEFT: 296px; POSITION: absolute; TOP: 572px"
					runat="server" Width="166px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DPriceG3" style="Z-INDEX: 195; LEFT: 224px; POSITION: absolute; TOP: 572px"
					runat="server" Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:dropdownlist id="DPriceG2" style="Z-INDEX: 194; LEFT: 184px; POSITION: absolute; TOP: 572px"
					runat="server" Width="40px" Height="12px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceG1" style="Z-INDEX: 193; LEFT: 16px; POSITION: absolute; TOP: 572px" runat="server"
					Width="166px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DPriceF3" style="Z-INDEX: 192; LEFT: 504px; POSITION: absolute; TOP: 554px"
					runat="server" Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:dropdownlist id="DPriceF2" style="Z-INDEX: 191; LEFT: 464px; POSITION: absolute; TOP: 554px"
					runat="server" Width="40px" Height="12px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceF1" style="Z-INDEX: 190; LEFT: 296px; POSITION: absolute; TOP: 554px"
					runat="server" Width="166px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DPriceE3" style="Z-INDEX: 189; LEFT: 224px; POSITION: absolute; TOP: 554px"
					runat="server" Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:dropdownlist id="DPriceE2" style="Z-INDEX: 188; LEFT: 184px; POSITION: absolute; TOP: 554px"
					runat="server" Width="40px" Height="12px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceE1" style="Z-INDEX: 187; LEFT: 16px; POSITION: absolute; TOP: 554px" runat="server"
					Width="166px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DPriceD3" style="Z-INDEX: 186; LEFT: 504px; POSITION: absolute; TOP: 536px"
					runat="server" Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:dropdownlist id="DPriceD2" style="Z-INDEX: 185; LEFT: 464px; POSITION: absolute; TOP: 536px"
					runat="server" Width="40px" Height="12px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceD1" style="Z-INDEX: 184; LEFT: 296px; POSITION: absolute; TOP: 536px"
					runat="server" Width="166px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DPriceC3" style="Z-INDEX: 183; LEFT: 224px; POSITION: absolute; TOP: 536px"
					runat="server" Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:dropdownlist id="DPriceC2" style="Z-INDEX: 182; LEFT: 184px; POSITION: absolute; TOP: 536px"
					runat="server" Width="40px" Height="12px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceC1" style="Z-INDEX: 181; LEFT: 16px; POSITION: absolute; TOP: 536px" runat="server"
					Width="166px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DPriceB3" style="Z-INDEX: 180; LEFT: 504px; POSITION: absolute; TOP: 518px"
					runat="server" Width="64px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:dropdownlist id="DPriceB2" style="Z-INDEX: 179; LEFT: 464px; POSITION: absolute; TOP: 518px"
					runat="server" Width="40px" Height="12px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceB1" style="Z-INDEX: 178; LEFT: 296px; POSITION: absolute; TOP: 518px"
					runat="server" Width="166px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox><asp:textbox id="DPriceA1" style="Z-INDEX: 169; LEFT: 16px; POSITION: absolute; TOP: 518px" runat="server"
					Width="166px" Height="18px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove"></asp:textbox>
				<asp:hyperlink id="LRefFile" style="Z-INDEX: 164; LEFT: 120px; POSITION: absolute; TOP: 624px"
					runat="server" Width="64px" Height="8px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">其他文件</asp:hyperlink><asp:textbox id="DDate" style="Z-INDEX: 147; LEFT: 584px; POSITION: absolute; TOP: 192px" runat="server"
					Width="168px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True">DDate</asp:textbox><asp:textbox id="DSliderCode" style="Z-INDEX: 146; LEFT: 224px; POSITION: absolute; TOP: 120px"
					runat="server" Width="368px" Height="56px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove" ReadOnly="True" TextMode="MultiLine">SliderCode</asp:textbox><asp:dropdownlist id="DPerson" style="Z-INDEX: 142; LEFT: 584px; POSITION: absolute; TOP: 226px" runat="server"
					Width="176px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DDivision" style="Z-INDEX: 141; LEFT: 312px; POSITION: absolute; TOP: 226px"
					runat="server" Width="176px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DSliderPrice" style="Z-INDEX: 129; LEFT: 648px; POSITION: absolute; TOP: 476px"
					runat="server" Width="112px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">SliderPrice</asp:textbox><asp:textbox id="DDevReason" style="Z-INDEX: 126; LEFT: 16px; POSITION: absolute; TOP: 384px"
					runat="server" Width="744px" Height="80px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" TextMode="MultiLine">DevReason</asp:textbox><asp:textbox id="DSellVendor" style="Z-INDEX: 121; LEFT: 584px; POSITION: absolute; TOP: 296px"
					runat="server" Width="174px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DSellVendor</asp:textbox><asp:dropdownlist id="DManufPlace" style="Z-INDEX: 119; LEFT: 584px; POSITION: absolute; TOP: 328px"
					runat="server" Width="176px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DSliderType" style="Z-INDEX: 117; LEFT: 312px; POSITION: absolute; TOP: 328px"
					runat="server" Width="176px" Height="20px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DSpec" style="Z-INDEX: 110; LEFT: 312px; POSITION: absolute; TOP: 259px" runat="server"
					Width="440px" Height="20px" BackColor="Yellow" ForeColor="Blue" Font-Size="8pt" BorderStyle="Groove" ReadOnly="True">DSpec</asp:textbox><asp:textbox id="DSliderGRCode" style="Z-INDEX: 109; LEFT: 604px; POSITION: absolute; TOP: 158px"
					runat="server" Width="154px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True">DSliderGRCode</asp:textbox><asp:textbox id="DNo" style="Z-INDEX: 108; LEFT: 312px; POSITION: absolute; TOP: 192px" runat="server"
					Width="176px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DNo</asp:textbox><asp:textbox id="DFormSno" style="Z-INDEX: 107; LEFT: 16px; POSITION: absolute; TOP: 656px" runat="server"
					Width="97px" Height="20px" BackColor="White" ForeColor="Blue" BorderStyle="None">單號：123</asp:textbox><asp:image id="LMapFile" style="Z-INDEX: 245; LEFT: 16px; POSITION: absolute; TOP: 120px" runat="server"
					Width="200px" Height="230px" BorderStyle="Groove" ImageUrl="http://10.245.0.178//WorkFlow/Document/000003/1-Map2-2006124111750-2.jpg"></asp:image></FONT></form>
		</FONT></FORM>
	</body>
</HTML>
