<%@ Page Language="vb" AutoEventWireup="false" Codebehind="ImportSheet_02.aspx.vb" Inherits="SPD.ImportSheet_02"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>ImportSheet_02</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" encType="multipart/form-data" runat="server">
			<FONT face="新細明體">
				<asp:image id="DImportSheet" style="Z-INDEX: 101; LEFT: 8px; POSITION: absolute; TOP: 8px"
					runat="server" ImageUrl="Images\ImportSheet_001.jpg" Height="644px" Width="761px"></asp:image>
				<asp:hyperlink id="LSliderDetail" style="Z-INDEX: 159; LEFT: 648px; POSITION: absolute; TOP: 64px"
					runat="server" Height="16px" Width="112px" Target="_blank" NavigateUrl="BoardEdit.aspx" Font-Size="12pt">有追加拉頭細目</asp:hyperlink>
				<asp:Panel id="Panel1" style="Z-INDEX: 158; LEFT: 560px; POSITION: absolute; TOP: 8px" runat="server"
					Height="72px" Width="220px" BackColor="White">
					<asp:textbox id="DStatus" runat="server" Width="226px" Height="48px" Font-Size="10pt" BackColor="Red"
						ReadOnly="True" BorderStyle="Groove" ForeColor="White">修改工程進行中(xxxxxxxxxxxxxxxxxx)</asp:textbox>
				</asp:Panel><asp:hyperlink id="LSurface" style="Z-INDEX: 157; LEFT: 96px; POSITION: absolute; TOP: 64px" runat="server"
					Height="8px" Width="80px" Target="_blank" NavigateUrl="BoardEdit.aspx" Font-Size="12pt">有表面處理</asp:hyperlink><asp:hyperlink id="LColorAppend" style="Z-INDEX: 156; LEFT: 184px; POSITION: absolute; TOP: 64px"
					runat="server" Height="8px" Width="112px" Target="_blank" NavigateUrl="BoardEdit.aspx" Font-Size="12pt">有外注色番追加</asp:hyperlink><asp:hyperlink id="LOPContact" style="Z-INDEX: 155; LEFT: 312px; POSITION: absolute; TOP: 64px"
					runat="server" Height="8px" Width="112px" Target="_blank" NavigateUrl="BoardEdit.aspx" Font-Size="12pt">有型別胴體追加</asp:hyperlink><asp:hyperlink id="LContact" style="Z-INDEX: 154; LEFT: 440px; POSITION: absolute; TOP: 64px" runat="server"
					Height="8px" Width="112px" Target="_blank" NavigateUrl="BoardEdit.aspx" Font-Size="12pt">有追加業務連絡</asp:hyperlink><asp:label id="Label22" style="Z-INDEX: 153; LEFT: 504px; POSITION: absolute; TOP: 506px" runat="server"
					Height="10px" Width="24px" Font-Size="8pt">單價</asp:label><asp:label id="Label21" style="Z-INDEX: 152; LEFT: 464px; POSITION: absolute; TOP: 506px" runat="server"
					Height="10px" Width="24px" Font-Size="8pt">幣別</asp:label><asp:label id="Label20" style="Z-INDEX: 151; LEFT: 296px; POSITION: absolute; TOP: 506px" runat="server"
					Height="10px" Width="48px" Font-Size="8pt">拉頭種類</asp:label><asp:label id="Label19" style="Z-INDEX: 150; LEFT: 224px; POSITION: absolute; TOP: 506px" runat="server"
					Height="10px" Width="24px" Font-Size="8pt">單價</asp:label><asp:label id="Label18" style="Z-INDEX: 149; LEFT: 184px; POSITION: absolute; TOP: 506px" runat="server"
					Height="10px" Width="24px" Font-Size="8pt">幣別</asp:label><asp:label id="Label17" style="Z-INDEX: 148; LEFT: 56px; POSITION: absolute; TOP: 506px" runat="server"
					Height="10px" Width="48px" Font-Size="8pt">拉頭種類</asp:label><asp:dropdownlist id="DBuyer" style="Z-INDEX: 147; LEFT: 312px; POSITION: absolute; TOP: 296px" runat="server"
					Height="20px" Width="176px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceJ1" style="Z-INDEX: 143; LEFT: 296px; POSITION: absolute; TOP: 590px"
					runat="server" Height="18px" Width="166px" Font-Size="8pt" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue"></asp:textbox><asp:textbox id="DPriceI3" style="Z-INDEX: 142; LEFT: 224px; POSITION: absolute; TOP: 590px"
					runat="server" Height="18px" Width="64px" Font-Size="8pt" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue"></asp:textbox><asp:dropdownlist id="DPriceI2" style="Z-INDEX: 141; LEFT: 184px; POSITION: absolute; TOP: 590px"
					runat="server" Height="12px" Width="40px" Font-Size="8pt" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceI1" style="Z-INDEX: 140; LEFT: 16px; POSITION: absolute; TOP: 590px" runat="server"
					Height="18px" Width="166px" Font-Size="8pt" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue"></asp:textbox><asp:textbox id="DPriceH3" style="Z-INDEX: 139; LEFT: 504px; POSITION: absolute; TOP: 572px"
					runat="server" Height="18px" Width="64px" Font-Size="8pt" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue"></asp:textbox><asp:dropdownlist id="DPriceH2" style="Z-INDEX: 138; LEFT: 464px; POSITION: absolute; TOP: 572px"
					runat="server" Height="12px" Width="40px" Font-Size="8pt" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceJ3" style="Z-INDEX: 145; LEFT: 504px; POSITION: absolute; TOP: 590px"
					runat="server" Height="18px" Width="64px" Font-Size="8pt" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue"></asp:textbox><asp:dropdownlist id="DPriceJ2" style="Z-INDEX: 144; LEFT: 464px; POSITION: absolute; TOP: 590px"
					runat="server" Height="12px" Width="40px" Font-Size="8pt" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceA3" style="Z-INDEX: 116; LEFT: 224px; POSITION: absolute; TOP: 518px"
					runat="server" Height="18px" Width="64px" Font-Size="8pt" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue"></asp:textbox><asp:dropdownlist id="DPriceA2" style="Z-INDEX: 115; LEFT: 184px; POSITION: absolute; TOP: 518px"
					runat="server" Height="12px" Width="40px" Font-Size="8pt" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceH1" style="Z-INDEX: 137; LEFT: 296px; POSITION: absolute; TOP: 572px"
					runat="server" Height="18px" Width="166px" Font-Size="8pt" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue"></asp:textbox><asp:textbox id="DPriceG3" style="Z-INDEX: 136; LEFT: 224px; POSITION: absolute; TOP: 572px"
					runat="server" Height="18px" Width="64px" Font-Size="8pt" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue"></asp:textbox><asp:dropdownlist id="DPriceG2" style="Z-INDEX: 135; LEFT: 184px; POSITION: absolute; TOP: 572px"
					runat="server" Height="12px" Width="40px" Font-Size="8pt" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceG1" style="Z-INDEX: 134; LEFT: 16px; POSITION: absolute; TOP: 572px" runat="server"
					Height="18px" Width="166px" Font-Size="8pt" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue"></asp:textbox><asp:textbox id="DPriceF3" style="Z-INDEX: 133; LEFT: 504px; POSITION: absolute; TOP: 554px"
					runat="server" Height="18px" Width="64px" Font-Size="8pt" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue"></asp:textbox><asp:dropdownlist id="DPriceF2" style="Z-INDEX: 132; LEFT: 464px; POSITION: absolute; TOP: 554px"
					runat="server" Height="12px" Width="40px" Font-Size="8pt" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceF1" style="Z-INDEX: 131; LEFT: 296px; POSITION: absolute; TOP: 554px"
					runat="server" Height="18px" Width="166px" Font-Size="8pt" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue"></asp:textbox><asp:textbox id="DPriceE3" style="Z-INDEX: 130; LEFT: 224px; POSITION: absolute; TOP: 554px"
					runat="server" Height="18px" Width="64px" Font-Size="8pt" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue"></asp:textbox><asp:dropdownlist id="DPriceE2" style="Z-INDEX: 129; LEFT: 184px; POSITION: absolute; TOP: 554px"
					runat="server" Height="12px" Width="40px" Font-Size="8pt" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceE1" style="Z-INDEX: 128; LEFT: 16px; POSITION: absolute; TOP: 554px" runat="server"
					Height="18px" Width="166px" Font-Size="8pt" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue"></asp:textbox><asp:textbox id="DPriceD3" style="Z-INDEX: 127; LEFT: 504px; POSITION: absolute; TOP: 536px"
					runat="server" Height="18px" Width="64px" Font-Size="8pt" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue"></asp:textbox><asp:dropdownlist id="DPriceD2" style="Z-INDEX: 126; LEFT: 464px; POSITION: absolute; TOP: 536px"
					runat="server" Height="12px" Width="40px" Font-Size="8pt" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceD1" style="Z-INDEX: 125; LEFT: 296px; POSITION: absolute; TOP: 536px"
					runat="server" Height="18px" Width="166px" Font-Size="8pt" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue"></asp:textbox><asp:textbox id="DPriceC3" style="Z-INDEX: 124; LEFT: 224px; POSITION: absolute; TOP: 536px"
					runat="server" Height="18px" Width="64px" Font-Size="8pt" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue"></asp:textbox><asp:dropdownlist id="DPriceC2" style="Z-INDEX: 123; LEFT: 184px; POSITION: absolute; TOP: 536px"
					runat="server" Height="12px" Width="40px" Font-Size="8pt" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceC1" style="Z-INDEX: 122; LEFT: 16px; POSITION: absolute; TOP: 536px" runat="server"
					Height="18px" Width="166px" Font-Size="8pt" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue"></asp:textbox><asp:textbox id="DPriceB3" style="Z-INDEX: 121; LEFT: 504px; POSITION: absolute; TOP: 518px"
					runat="server" Height="18px" Width="64px" Font-Size="8pt" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue"></asp:textbox><asp:dropdownlist id="DPriceB2" style="Z-INDEX: 120; LEFT: 464px; POSITION: absolute; TOP: 518px"
					runat="server" Height="12px" Width="40px" Font-Size="8pt" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="NT">NT</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DPriceB1" style="Z-INDEX: 119; LEFT: 296px; POSITION: absolute; TOP: 518px"
					runat="server" Height="18px" Width="166px" Font-Size="8pt" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue"></asp:textbox><asp:textbox id="DPriceA1" style="Z-INDEX: 118; LEFT: 16px; POSITION: absolute; TOP: 518px" runat="server"
					Height="18px" Width="166px" Font-Size="8pt" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue"></asp:textbox><asp:hyperlink id="LRefFile" style="Z-INDEX: 117; LEFT: 120px; POSITION: absolute; TOP: 624px"
					runat="server" Height="8px" Width="64px" Target="_blank" NavigateUrl="BoardEdit.aspx" Font-Size="12pt">其他文件</asp:hyperlink><asp:textbox id="DDate" style="Z-INDEX: 114; LEFT: 584px; POSITION: absolute; TOP: 192px" runat="server"
					Height="20px" Width="168px" ReadOnly="True" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue">DDate</asp:textbox><asp:textbox id="DSliderCode" style="Z-INDEX: 113; LEFT: 224px; POSITION: absolute; TOP: 120px"
					runat="server" Height="56px" Width="368px" Font-Size="8pt" ReadOnly="True" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue" TextMode="MultiLine">SliderCode</asp:textbox><asp:dropdownlist id="DPerson" style="Z-INDEX: 112; LEFT: 584px; POSITION: absolute; TOP: 226px" runat="server"
					Height="20px" Width="176px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DDivision" style="Z-INDEX: 111; LEFT: 312px; POSITION: absolute; TOP: 226px"
					runat="server" Height="20px" Width="176px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="3">3</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DSliderPrice" style="Z-INDEX: 110; LEFT: 648px; POSITION: absolute; TOP: 476px"
					runat="server" Height="20px" Width="112px" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue">SliderPrice</asp:textbox><asp:textbox id="DDevReason" style="Z-INDEX: 109; LEFT: 16px; POSITION: absolute; TOP: 384px"
					runat="server" Height="80px" Width="744px" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue" TextMode="MultiLine">DevReason</asp:textbox><asp:textbox id="DSellVendor" style="Z-INDEX: 108; LEFT: 584px; POSITION: absolute; TOP: 296px"
					runat="server" Height="20px" Width="174px" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue">DSellVendor</asp:textbox><asp:dropdownlist id="DManufPlace" style="Z-INDEX: 107; LEFT: 584px; POSITION: absolute; TOP: 328px"
					runat="server" Height="20px" Width="176px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DSliderType" style="Z-INDEX: 106; LEFT: 312px; POSITION: absolute; TOP: 328px"
					runat="server" Height="20px" Width="176px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="Y">Yes</asp:ListItem>
					<asp:ListItem Value="N">No</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DSpec" style="Z-INDEX: 105; LEFT: 312px; POSITION: absolute; TOP: 259px" runat="server"
					Height="20px" Width="440px" Font-Size="8pt" ReadOnly="True" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue">DSpec</asp:textbox><asp:textbox id="DSliderGRCode" style="Z-INDEX: 104; LEFT: 604px; POSITION: absolute; TOP: 158px"
					runat="server" Height="20px" Width="154px" ReadOnly="True" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue">DSliderGRCode</asp:textbox><asp:textbox id="DNo" style="Z-INDEX: 103; LEFT: 312px; POSITION: absolute; TOP: 192px" runat="server"
					Height="20px" Width="176px" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue">DNo</asp:textbox><asp:textbox id="DFormSno" style="Z-INDEX: 102; LEFT: 16px; POSITION: absolute; TOP: 656px" runat="server"
					Height="20px" Width="97px" BorderStyle="None" BackColor="White" ForeColor="Blue">單號：123</asp:textbox><asp:image id="LMapFile" style="Z-INDEX: 146; LEFT: 16px; POSITION: absolute; TOP: 120px" runat="server"
					ImageUrl="http://10.245.0.178//WorkFlow/Document/000003/1-Map2-2006124111750-2.jpg" Height="230px" Width="200px" BorderStyle="Groove"></asp:image></FONT></form>
		</FONT></FORM>
	</body>
</HTML>
