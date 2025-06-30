<%@ Page Language="vb" AspCompat="true" AutoEventWireup="false" Codebehind="Simulation.aspx.vb" Inherits="SPD.Simulation"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>工程模擬</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體">
				<asp:textbox id="DOP10D" style="Z-INDEX: 130; POSITION: absolute; TOP: 872px; LEFT: 408px" runat="server"
					BackColor="Yellow" Font-Size="8pt" Height="80px" BorderStyle="Groove" ReadOnly="True" Width="280px"
					TextMode="MultiLine"></asp:textbox>
				<asp:image id="DOP25I" style="Z-INDEX: 201; POSITION: absolute; TOP: 2344px; LEFT: 248px" runat="server"
					Width="104px" Height="56px" ImageUrl="Images\sirusi.gif"></asp:image>
				<asp:hyperlink id="LOP25" style="Z-INDEX: 200; POSITION: absolute; TOP: 2312px; LEFT: 696px" runat="server"
					Width="24px" Height="8px" Font-Size="10pt" Target="_blank">顯示待處理</asp:hyperlink>
				<asp:textbox id="DOP25D" style="Z-INDEX: 199; POSITION: absolute; TOP: 2312px; LEFT: 408px" runat="server"
					TextMode="MultiLine" Width="280px" ReadOnly="True" BorderStyle="Groove" Height="80px" Font-Size="8pt"
					BackColor="Yellow"></asp:textbox>
				<asp:textbox id="DOP25" style="Z-INDEX: 198; POSITION: absolute; TOP: 2312px; LEFT: 208px" runat="server"
					Width="192px" ReadOnly="True" BorderStyle="Groove" Height="28px" Font-Size="11pt" BackColor="Yellow"
					ForeColor="#C00000">aaaaaaaa</asp:textbox>
				<asp:hyperlink id="LOP24" style="Z-INDEX: 197; POSITION: absolute; TOP: 2216px; LEFT: 696px" runat="server"
					Width="24px" Height="8px" Font-Size="10pt" Target="_blank">顯示待處理</asp:hyperlink>
				<asp:textbox id="DOP24D" style="Z-INDEX: 196; POSITION: absolute; TOP: 2216px; LEFT: 408px" runat="server"
					TextMode="MultiLine" Width="280px" ReadOnly="True" BorderStyle="Groove" Height="80px" Font-Size="8pt"
					BackColor="Yellow"></asp:textbox>
				<asp:textbox id="DOP24" style="Z-INDEX: 195; POSITION: absolute; TOP: 2216px; LEFT: 208px" runat="server"
					Width="192px" ReadOnly="True" BorderStyle="Groove" Height="28px" Font-Size="11pt" BackColor="Yellow"
					ForeColor="#C00000">aaaaaaaa</asp:textbox>
				<asp:image id="DOP23I" style="Z-INDEX: 194; POSITION: absolute; TOP: 2152px; LEFT: 248px" runat="server"
					Width="104px" Height="56px" ImageUrl="Images\sirusi.gif"></asp:image>
				<asp:hyperlink id="LOP23" style="Z-INDEX: 193; POSITION: absolute; TOP: 2120px; LEFT: 696px" runat="server"
					Width="24px" Height="8px" Font-Size="10pt" Target="_blank">顯示待處理</asp:hyperlink>
				<asp:textbox id="DOP23D" style="Z-INDEX: 192; POSITION: absolute; TOP: 2120px; LEFT: 408px" runat="server"
					TextMode="MultiLine" Width="280px" ReadOnly="True" BorderStyle="Groove" Height="80px" Font-Size="8pt"
					BackColor="Yellow"></asp:textbox>
				<asp:textbox id="DOP23" style="Z-INDEX: 191; POSITION: absolute; TOP: 2120px; LEFT: 208px" runat="server"
					Width="192px" ReadOnly="True" BorderStyle="Groove" Height="28px" Font-Size="11pt" BackColor="Yellow"
					ForeColor="#C00000">aaaaaaaa</asp:textbox>
				<asp:image id="DOP22I" style="Z-INDEX: 190; POSITION: absolute; TOP: 2056px; LEFT: 248px" runat="server"
					Width="104px" Height="56px" ImageUrl="Images\sirusi.gif"></asp:image>
				<asp:hyperlink id="LOP22" style="Z-INDEX: 189; POSITION: absolute; TOP: 2024px; LEFT: 696px" runat="server"
					Width="24px" Height="8px" Font-Size="10pt" Target="_blank">顯示待處理</asp:hyperlink>
				<asp:textbox id="DOP22D" style="Z-INDEX: 188; POSITION: absolute; TOP: 2024px; LEFT: 408px" runat="server"
					TextMode="MultiLine" Width="280px" ReadOnly="True" BorderStyle="Groove" Height="80px" Font-Size="8pt"
					BackColor="Yellow"></asp:textbox>
				<asp:textbox id="DOP22" style="Z-INDEX: 187; POSITION: absolute; TOP: 2024px; LEFT: 208px" runat="server"
					Width="192px" ReadOnly="True" BorderStyle="Groove" Height="28px" Font-Size="11pt" BackColor="Yellow"
					ForeColor="#C00000">aaaaaaaa</asp:textbox>
				<asp:image id="DOP21I" style="Z-INDEX: 186; POSITION: absolute; TOP: 1960px; LEFT: 248px" runat="server"
					Width="104px" Height="56px" ImageUrl="Images\sirusi.gif"></asp:image>
				<asp:hyperlink id="LOP21" style="Z-INDEX: 185; POSITION: absolute; TOP: 1928px; LEFT: 696px" runat="server"
					Width="24px" Height="8px" Font-Size="10pt" Target="_blank">顯示待處理</asp:hyperlink>
				<asp:textbox id="DOP21D" style="Z-INDEX: 184; POSITION: absolute; TOP: 1928px; LEFT: 408px" runat="server"
					TextMode="MultiLine" Width="280px" ReadOnly="True" BorderStyle="Groove" Height="80px" Font-Size="8pt"
					BackColor="Yellow"></asp:textbox>
				<asp:textbox id="DOP21" style="Z-INDEX: 183; POSITION: absolute; TOP: 1928px; LEFT: 208px" runat="server"
					Width="192px" ReadOnly="True" BorderStyle="Groove" Height="28px" Font-Size="11pt" BackColor="Yellow"
					ForeColor="#C00000">aaaaaaaa</asp:textbox>
				<asp:image id="DOP20I" style="Z-INDEX: 182; POSITION: absolute; TOP: 1864px; LEFT: 248px" runat="server"
					Width="104px" Height="56px" ImageUrl="Images\sirusi.gif"></asp:image>
				<asp:image id="DOP24I" style="Z-INDEX: 180; POSITION: absolute; TOP: 2248px; LEFT: 248px" runat="server"
					Width="104px" Height="56px" ImageUrl="Images\sirusi.gif"></asp:image>
				<asp:hyperlink id="LOP20" style="Z-INDEX: 179; POSITION: absolute; TOP: 1832px; LEFT: 696px" runat="server"
					Width="24px" Height="8px" Font-Size="10pt" Target="_blank">顯示待處理</asp:hyperlink>
				<asp:hyperlink id="LOP19" style="Z-INDEX: 178; POSITION: absolute; TOP: 1736px; LEFT: 696px" runat="server"
					Width="24px" Height="8px" Font-Size="10pt" Target="_blank">顯示待處理</asp:hyperlink>
				<asp:hyperlink id="LOP18" style="Z-INDEX: 177; POSITION: absolute; TOP: 1640px; LEFT: 696px" runat="server"
					Width="24px" Height="8px" Font-Size="10pt" Target="_blank">顯示待處理</asp:hyperlink>
				<asp:hyperlink id="LOP17" style="Z-INDEX: 176; POSITION: absolute; TOP: 1544px; LEFT: 696px" runat="server"
					Width="24px" Height="8px" Font-Size="10pt" Target="_blank">顯示待處理</asp:hyperlink>
				<asp:hyperlink id="LOP16" style="Z-INDEX: 175; POSITION: absolute; TOP: 1448px; LEFT: 696px" runat="server"
					Width="24px" Height="8px" Font-Size="10pt" Target="_blank">顯示待處理</asp:hyperlink>
				<asp:hyperlink id="LOP15" style="Z-INDEX: 174; POSITION: absolute; TOP: 1352px; LEFT: 696px" runat="server"
					Width="24px" Height="8px" Font-Size="10pt" Target="_blank">顯示待處理</asp:hyperlink>
				<asp:hyperlink id="LOP14" style="Z-INDEX: 173; POSITION: absolute; TOP: 1256px; LEFT: 696px" runat="server"
					Width="24px" Height="8px" Font-Size="10pt" Target="_blank">顯示待處理</asp:hyperlink>
				<asp:hyperlink id="LOP13" style="Z-INDEX: 172; POSITION: absolute; TOP: 1160px; LEFT: 696px" runat="server"
					Width="24px" Height="8px" Font-Size="10pt" Target="_blank">顯示待處理</asp:hyperlink>
				<asp:hyperlink id="LOP12" style="Z-INDEX: 171; POSITION: absolute; TOP: 1064px; LEFT: 696px" runat="server"
					Width="24px" Height="8px" Font-Size="10pt" Target="_blank">顯示待處理</asp:hyperlink>
				<asp:hyperlink id="LOP11" style="Z-INDEX: 170; POSITION: absolute; TOP: 968px; LEFT: 696px" runat="server"
					Width="24px" Height="8px" Font-Size="10pt" Target="_blank">顯示待處理</asp:hyperlink>
				<asp:hyperlink id="LOP10" style="Z-INDEX: 169; POSITION: absolute; TOP: 872px; LEFT: 696px" runat="server"
					Width="24px" Height="8px" Font-Size="10pt" Target="_blank">顯示待處理</asp:hyperlink>
				<asp:hyperlink id="LOP9" style="Z-INDEX: 168; POSITION: absolute; TOP: 776px; LEFT: 696px" runat="server"
					Width="24px" Height="8px" Font-Size="10pt" Target="_blank">顯示待處理</asp:hyperlink>
				<asp:hyperlink id="LOP8" style="Z-INDEX: 167; POSITION: absolute; TOP: 680px; LEFT: 696px" runat="server"
					Width="24px" Height="8px" Font-Size="10pt" Target="_blank">顯示待處理</asp:hyperlink>
				<asp:hyperlink id="LOP7" style="Z-INDEX: 166; POSITION: absolute; TOP: 584px; LEFT: 696px" runat="server"
					Width="24px" Height="8px" Font-Size="10pt" Target="_blank">顯示待處理</asp:hyperlink>
				<asp:hyperlink id="LOP4" style="Z-INDEX: 165; POSITION: absolute; TOP: 296px; LEFT: 696px" runat="server"
					Width="24px" Height="8px" Font-Size="10pt" Target="_blank">顯示待處理</asp:hyperlink>
				<asp:hyperlink id="LOP3" style="Z-INDEX: 164; POSITION: absolute; TOP: 200px; LEFT: 696px" runat="server"
					Width="24px" Height="8px" Font-Size="10pt" Target="_blank">顯示待處理</asp:hyperlink>
				<asp:hyperlink id="LOP2" style="Z-INDEX: 163; POSITION: absolute; TOP: 104px; LEFT: 696px" runat="server"
					Width="24px" Height="8px" Font-Size="10pt" Target="_blank">顯示待處理</asp:hyperlink>
				<asp:hyperlink id="LOP5" style="Z-INDEX: 162; POSITION: absolute; TOP: 392px; LEFT: 696px" runat="server"
					Width="24px" Height="8px" Font-Size="10pt" Target="_blank">顯示待處理</asp:hyperlink>
				<asp:hyperlink id="LOP6" style="Z-INDEX: 161; POSITION: absolute; TOP: 488px; LEFT: 696px" runat="server"
					Width="24px" Height="8px" Font-Size="10pt" Target="_blank">顯示待處理</asp:hyperlink>
				<asp:hyperlink id="LOP1" style="Z-INDEX: 160; POSITION: absolute; TOP: 8px; LEFT: 696px" runat="server"
					Width="24px" Height="8px" Font-Size="10pt" Target="_blank">顯示待處理</asp:hyperlink>
				<asp:image id="DOP19I" style="Z-INDEX: 159; POSITION: absolute; TOP: 1768px; LEFT: 248px" runat="server"
					Width="104px" Height="56px" ImageUrl="Images\sirusi.gif"></asp:image>
				<asp:textbox id="DOP19D" style="Z-INDEX: 158; POSITION: absolute; TOP: 1736px; LEFT: 408px" runat="server"
					TextMode="MultiLine" Width="280px" ReadOnly="True" BorderStyle="Groove" Height="80px" Font-Size="8pt"
					BackColor="Yellow"></asp:textbox>
				<asp:textbox id="DOP19" style="Z-INDEX: 157; POSITION: absolute; TOP: 1736px; LEFT: 208px" runat="server"
					Width="192px" ReadOnly="True" BorderStyle="Groove" Height="28px" Font-Size="11pt" BackColor="Yellow"
					ForeColor="#C00000">aaaaaaaa</asp:textbox>
				<asp:image id="DOP18I" style="Z-INDEX: 156; POSITION: absolute; TOP: 1672px; LEFT: 248px" runat="server"
					Width="104px" Height="56px" ImageUrl="Images\sirusi.gif"></asp:image>
				<asp:textbox id="DOP18D" style="Z-INDEX: 155; POSITION: absolute; TOP: 1640px; LEFT: 408px" runat="server"
					TextMode="MultiLine" Width="280px" ReadOnly="True" BorderStyle="Groove" Height="80px" Font-Size="8pt"
					BackColor="Yellow"></asp:textbox>
				<asp:textbox id="DOP18" style="Z-INDEX: 154; POSITION: absolute; TOP: 1640px; LEFT: 208px" runat="server"
					Width="192px" ReadOnly="True" BorderStyle="Groove" Height="28px" Font-Size="11pt" BackColor="Yellow"
					ForeColor="#C00000">aaaaaaaa</asp:textbox>
				<asp:image id="DOP17I" style="Z-INDEX: 153; POSITION: absolute; TOP: 1576px; LEFT: 248px" runat="server"
					Width="104px" Height="56px" ImageUrl="Images\sirusi.gif"></asp:image>
				<asp:textbox id="DOP17D" style="Z-INDEX: 152; POSITION: absolute; TOP: 1544px; LEFT: 408px" runat="server"
					TextMode="MultiLine" Width="280px" ReadOnly="True" BorderStyle="Groove" Height="80px" Font-Size="8pt"
					BackColor="Yellow"></asp:textbox>
				<asp:textbox id="DOP17" style="Z-INDEX: 151; POSITION: absolute; TOP: 1544px; LEFT: 208px" runat="server"
					Width="192px" ReadOnly="True" BorderStyle="Groove" Height="28px" Font-Size="11pt" BackColor="Yellow"
					ForeColor="#C00000">aaaaaaaa</asp:textbox>
				<asp:image id="DOP16I" style="Z-INDEX: 150; POSITION: absolute; TOP: 1480px; LEFT: 248px" runat="server"
					Width="104px" Height="56px" ImageUrl="Images\sirusi.gif"></asp:image>
				<asp:textbox id="DOP16D" style="Z-INDEX: 149; POSITION: absolute; TOP: 1448px; LEFT: 408px" runat="server"
					TextMode="MultiLine" Width="280px" ReadOnly="True" BorderStyle="Groove" Height="80px" Font-Size="8pt"
					BackColor="Yellow"></asp:textbox>
				<asp:textbox id="DOP16" style="Z-INDEX: 148; POSITION: absolute; TOP: 1448px; LEFT: 208px" runat="server"
					Width="192px" ReadOnly="True" BorderStyle="Groove" Height="28px" Font-Size="11pt" BackColor="Yellow"
					ForeColor="#C00000">aaaaaaaa</asp:textbox>
				<asp:textbox id="DOP20D" style="Z-INDEX: 147; POSITION: absolute; TOP: 1832px; LEFT: 408px" runat="server"
					TextMode="MultiLine" Width="280px" ReadOnly="True" BorderStyle="Groove" Height="80px" Font-Size="8pt"
					BackColor="Yellow"></asp:textbox>
				<asp:textbox id="DOP20" style="Z-INDEX: 146; POSITION: absolute; TOP: 1832px; LEFT: 208px" runat="server"
					Width="192px" ReadOnly="True" BorderStyle="Groove" Height="28px" Font-Size="11pt" BackColor="Yellow"
					ForeColor="#C00000">aaaaaaaa</asp:textbox>
				<asp:image id="DOP15I" style="Z-INDEX: 145; POSITION: absolute; TOP: 1384px; LEFT: 248px" runat="server"
					Width="104px" Height="56px" ImageUrl="Images\sirusi.gif"></asp:image><asp:textbox id="DOP15D" style="Z-INDEX: 143; POSITION: absolute; TOP: 1352px; LEFT: 408px" runat="server"
					BackColor="Yellow" Font-Size="8pt" Height="80px" BorderStyle="Groove" ReadOnly="True" Width="280px" TextMode="MultiLine"></asp:textbox><asp:image id="DOP11I" style="Z-INDEX: 142; POSITION: absolute; TOP: 1000px; LEFT: 248px" runat="server"
					Height="56px" Width="104px" ImageUrl="Images\sirusi.gif"></asp:image><asp:image id="DOP12I" style="Z-INDEX: 141; POSITION: absolute; TOP: 1096px; LEFT: 248px" runat="server"
					Height="56px" Width="104px" ImageUrl="Images\sirusi.gif"></asp:image><asp:image id="DOP13I" style="Z-INDEX: 140; POSITION: absolute; TOP: 1192px; LEFT: 248px" runat="server"
					Height="56px" Width="104px" ImageUrl="Images\sirusi.gif"></asp:image><asp:textbox id="DOP11D" style="Z-INDEX: 139; POSITION: absolute; TOP: 968px; LEFT: 408px" runat="server"
					BackColor="Yellow" Font-Size="8pt" Height="80px" BorderStyle="Groove" ReadOnly="True" Width="280px" TextMode="MultiLine"></asp:textbox><asp:textbox id="DOP12D" style="Z-INDEX: 138; POSITION: absolute; TOP: 1064px; LEFT: 408px" runat="server"
					BackColor="Yellow" Font-Size="8pt" Height="80px" BorderStyle="Groove" ReadOnly="True" Width="280px" TextMode="MultiLine"></asp:textbox><asp:textbox id="DOP13D" style="Z-INDEX: 137; POSITION: absolute; TOP: 1160px; LEFT: 408px" runat="server"
					BackColor="Yellow" Font-Size="8pt" Height="80px" BorderStyle="Groove" ReadOnly="True" Width="280px" TextMode="MultiLine"></asp:textbox><asp:textbox id="DOP11" style="Z-INDEX: 136; POSITION: absolute; TOP: 968px; LEFT: 208px" runat="server"
					ForeColor="#C00000" BackColor="Yellow" Font-Size="11pt" Height="28px" BorderStyle="Groove" ReadOnly="True" Width="192px">aaaaaaaa</asp:textbox><asp:textbox id="DOP15" style="Z-INDEX: 135; POSITION: absolute; TOP: 1352px; LEFT: 208px" runat="server"
					ForeColor="#C00000" BackColor="Yellow" Font-Size="11pt" Height="28px" BorderStyle="Groove" ReadOnly="True" Width="192px">aaaaaaaa</asp:textbox><asp:textbox id="DOP12" style="Z-INDEX: 134; POSITION: absolute; TOP: 1064px; LEFT: 208px" runat="server"
					ForeColor="#C00000" BackColor="Yellow" Font-Size="11pt" Height="28px" BorderStyle="Groove" ReadOnly="True" Width="192px">aaaaaaaa</asp:textbox><asp:textbox id="DOP13" style="Z-INDEX: 133; POSITION: absolute; TOP: 1160px; LEFT: 208px" runat="server"
					ForeColor="#C00000" BackColor="Yellow" Font-Size="11pt" Height="28px" BorderStyle="Groove" ReadOnly="True" Width="192px">aaaaaaaa</asp:textbox><asp:textbox id="DOP14D" style="Z-INDEX: 132; POSITION: absolute; TOP: 1256px; LEFT: 408px" runat="server"
					BackColor="Yellow" Font-Size="8pt" Height="80px" BorderStyle="Groove" ReadOnly="True" Width="280px" TextMode="MultiLine"></asp:textbox><asp:textbox id="DOP14" style="Z-INDEX: 131; POSITION: absolute; TOP: 1256px; LEFT: 208px" runat="server"
					ForeColor="#C00000" BackColor="Yellow" Font-Size="11pt" Height="28px" BorderStyle="Groove" ReadOnly="True" Width="192px">aaaaaaaa</asp:textbox><FONT face="新細明體"></FONT><asp:image id="DOP14I" style="Z-INDEX: 127; POSITION: absolute; TOP: 1288px; LEFT: 248px" runat="server"
					Height="56px" Width="104px" ImageUrl="Images\sirusi.gif"></asp:image><asp:textbox id="DOP10" style="Z-INDEX: 129; POSITION: absolute; TOP: 872px; LEFT: 208px" runat="server"
					ForeColor="#C00000" BackColor="Yellow" Font-Size="11pt" Height="28px" BorderStyle="Groove" ReadOnly="True" Width="192px">aaaaaaaa</asp:textbox><asp:image id="DOP9I" style="Z-INDEX: 128; POSITION: absolute; TOP: 808px; LEFT: 248px" runat="server"
					Height="56px" Width="104px" ImageUrl="Images\sirusi.gif"></asp:image><asp:image id="DOP10I" style="Z-INDEX: 126; POSITION: absolute; TOP: 904px; LEFT: 248px" runat="server"
					Height="56px" Width="104px" ImageUrl="Images\sirusi.gif"></asp:image><asp:image id="DOP8I" style="Z-INDEX: 125; POSITION: absolute; TOP: 712px; LEFT: 248px" runat="server"
					Height="56px" Width="104px" ImageUrl="Images\sirusi.gif"></asp:image><asp:image id="DOP4I" style="Z-INDEX: 124; POSITION: absolute; TOP: 328px; LEFT: 248px" runat="server"
					Height="56px" Width="104px" ImageUrl="Images\sirusi.gif"></asp:image><asp:image id="DOP5I" style="Z-INDEX: 123; POSITION: absolute; TOP: 424px; LEFT: 248px" runat="server"
					Height="56px" Width="104px" ImageUrl="Images\sirusi.gif"></asp:image><asp:image id="DOP6I" style="Z-INDEX: 122; POSITION: absolute; TOP: 520px; LEFT: 248px" runat="server"
					Height="56px" Width="104px" ImageUrl="Images\sirusi.gif"></asp:image><asp:image id="DOP7I" style="Z-INDEX: 121; POSITION: absolute; TOP: 616px; LEFT: 248px" runat="server"
					Height="56px" Width="104px" ImageUrl="Images\sirusi.gif"></asp:image><asp:image id="DOP3I" style="Z-INDEX: 120; POSITION: absolute; TOP: 232px; LEFT: 248px" runat="server"
					Height="56px" Width="104px" ImageUrl="Images\sirusi.gif"></asp:image><asp:image id="DOP2I" style="Z-INDEX: 119; POSITION: absolute; TOP: 136px; LEFT: 248px" runat="server"
					Height="56px" Width="104px" ImageUrl="Images\sirusi.gif"></asp:image><asp:image id="DOP1I" style="Z-INDEX: 118; POSITION: absolute; TOP: 40px; LEFT: 248px" runat="server"
					Height="56px" Width="104px" ImageUrl="Images\sirusi.gif"></asp:image><asp:textbox id="DOP2D" style="Z-INDEX: 117; POSITION: absolute; TOP: 104px; LEFT: 408px" runat="server"
					BackColor="Yellow" Font-Size="8pt" Height="80px" BorderStyle="Groove" ReadOnly="True" Width="280px" TextMode="MultiLine"></asp:textbox><asp:textbox id="DOP3D" style="Z-INDEX: 116; POSITION: absolute; TOP: 200px; LEFT: 408px" runat="server"
					BackColor="Yellow" Font-Size="8pt" Height="80px" BorderStyle="Groove" ReadOnly="True" Width="280px" TextMode="MultiLine"></asp:textbox><asp:textbox id="DOP4D" style="Z-INDEX: 115; POSITION: absolute; TOP: 296px; LEFT: 408px" runat="server"
					BackColor="Yellow" Font-Size="8pt" Height="80px" BorderStyle="Groove" ReadOnly="True" Width="280px" TextMode="MultiLine"></asp:textbox><asp:textbox id="DOP5D" style="Z-INDEX: 114; POSITION: absolute; TOP: 392px; LEFT: 408px" runat="server"
					BackColor="Yellow" Font-Size="8pt" Height="80px" BorderStyle="Groove" ReadOnly="True" Width="280px" TextMode="MultiLine"></asp:textbox><asp:textbox id="DOP6D" style="Z-INDEX: 113; POSITION: absolute; TOP: 488px; LEFT: 408px" runat="server"
					BackColor="Yellow" Font-Size="8pt" Height="80px" BorderStyle="Groove" ReadOnly="True" Width="280px" TextMode="MultiLine"></asp:textbox><asp:textbox id="DOP7D" style="Z-INDEX: 112; POSITION: absolute; TOP: 584px; LEFT: 408px" runat="server"
					BackColor="Yellow" Font-Size="8pt" Height="80px" BorderStyle="Groove" ReadOnly="True" Width="280px" TextMode="MultiLine"></asp:textbox><asp:textbox id="DOP8D" style="Z-INDEX: 111; POSITION: absolute; TOP: 680px; LEFT: 408px" runat="server"
					BackColor="Yellow" Font-Size="8pt" Height="80px" BorderStyle="Groove" ReadOnly="True" Width="280px" TextMode="MultiLine"></asp:textbox>
				<asp:TextBox id="DOP9D" style="Z-INDEX: 110; POSITION: absolute; TOP: 776px; LEFT: 408px" runat="server"
					BackColor="Yellow" Font-Size="8pt" Height="80px" BorderStyle="Groove" ReadOnly="True" Width="280px"
					TextMode="MultiLine"></asp:TextBox>
				<asp:TextBox id="DOP2" style="Z-INDEX: 109; POSITION: absolute; TOP: 104px; LEFT: 208px" runat="server"
					ForeColor="#C00000" BackColor="Yellow" Font-Size="11pt" Height="28px" BorderStyle="Groove"
					ReadOnly="True" Width="192px">aaaaaaaa</asp:TextBox>
				<asp:TextBox id="DOP3" style="Z-INDEX: 108; POSITION: absolute; TOP: 200px; LEFT: 208px" runat="server"
					ForeColor="#C00000" BackColor="Yellow" Font-Size="11pt" Height="28px" BorderStyle="Groove"
					ReadOnly="True" Width="192px">aaaaaaaa</asp:TextBox>
				<asp:TextBox id="DOP4" style="Z-INDEX: 107; POSITION: absolute; TOP: 296px; LEFT: 208px" runat="server"
					ForeColor="#C00000" BackColor="Yellow" Font-Size="11pt" Height="28px" BorderStyle="Groove"
					ReadOnly="True" Width="192px">aaaaaaaa</asp:TextBox>
				<asp:TextBox id="DOP5" style="Z-INDEX: 106; POSITION: absolute; TOP: 392px; LEFT: 208px" runat="server"
					ForeColor="#C00000" BackColor="Yellow" Font-Size="11pt" Height="28px" BorderStyle="Groove"
					ReadOnly="True" Width="192px">aaaaaaaa</asp:TextBox>
				<asp:TextBox id="DOP6" style="Z-INDEX: 105; POSITION: absolute; TOP: 488px; LEFT: 208px" runat="server"
					ForeColor="#C00000" BackColor="Yellow" Font-Size="11pt" Height="28px" BorderStyle="Groove"
					ReadOnly="True" Width="192px">aaaaaaaa</asp:TextBox>
				<asp:TextBox id="DOP7" style="Z-INDEX: 104; POSITION: absolute; TOP: 584px; LEFT: 208px" runat="server"
					ForeColor="#C00000" BackColor="Yellow" Font-Size="11pt" Height="28px" BorderStyle="Groove"
					ReadOnly="True" Width="192px">aaaaaaaa</asp:TextBox>
				<asp:TextBox id="DOP8" style="Z-INDEX: 103; POSITION: absolute; TOP: 680px; LEFT: 208px" runat="server"
					ForeColor="#C00000" BackColor="Yellow" Font-Size="11pt" Height="28px" BorderStyle="Groove"
					ReadOnly="True" Width="192px">aaaaaaaa</asp:TextBox>
				<asp:TextBox id="DOP9" style="Z-INDEX: 102; POSITION: absolute; TOP: 776px; LEFT: 208px" runat="server"
					ForeColor="#C00000" BackColor="Yellow" Font-Size="11pt" Height="28px" BorderStyle="Groove"
					ReadOnly="True" Width="192px">aaaaaaaa</asp:TextBox>
				<asp:TextBox id="DOP1D" style="Z-INDEX: 101; POSITION: absolute; TOP: 8px; LEFT: 408px" runat="server"
					BackColor="Yellow" Font-Size="8pt" Height="80px" BorderStyle="Groove" ReadOnly="True" Width="280px"
					TextMode="MultiLine"></asp:TextBox>
				<asp:TextBox id="DOP1" style="Z-INDEX: 100; POSITION: absolute; TOP: 8px; LEFT: 208px" runat="server"
					ForeColor="#C00000" BackColor="Yellow" Font-Size="11pt" Height="28px" BorderStyle="Groove"
					ReadOnly="True" Width="192px">aaaaaaaa</asp:TextBox></FONT>
			<asp:TextBox id="DFormName" style="Z-INDEX: 144; POSITION: absolute; TOP: 8px; LEFT: 8px" runat="server"
				ForeColor="#C00000" BackColor="YellowGreen" Font-Size="11pt" Height="28px" BorderStyle="Groove"
				ReadOnly="True" Width="192px">aaaaaaaa</asp:TextBox></form>
	</body>
</HTML>
