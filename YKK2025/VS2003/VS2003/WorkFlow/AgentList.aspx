<%@ Page Language="vb" AutoEventWireup="false" Codebehind="AgentList.aspx.vb" Inherits="SPD.AgentList"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>代理人</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體">
				<DIV title="篩選項目：" style="DISPLAY: inline; Z-INDEX: 105; LEFT: 8px; WIDTH: 88px; COLOR: #0000ff; POSITION: absolute; TOP: 8px; HEIGHT: 24px"
					ms_positioning="FlowLayout">篩選項目：</DIV>
				<asp:imagebutton id="BExcel" style="Z-INDEX: 150; LEFT: 632px; POSITION: absolute; TOP: 8px" runat="server"
					Width="21px" Height="21px" ImageUrl="Images\msexcel.gif"></asp:imagebutton>
				<asp:dropdownlist id="DAgentName" style="Z-INDEX: 107; LEFT: 480px; POSITION: absolute; TOP: 8px"
					runat="server" Height="40px" Width="100px" ForeColor="Blue" BackColor="Yellow"></asp:dropdownlist>
				<asp:dropdownlist id="DAllForm" style="Z-INDEX: 106; LEFT: 96px; POSITION: absolute; TOP: 8px" runat="server"
					Height="40px" Width="110px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="ALL" Selected="True">ALL</asp:ListItem>
					<asp:ListItem Value="0">全委託單</asp:ListItem>
					<asp:ListItem Value="1">單一委託單</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DFormName" style="Z-INDEX: 101; LEFT: 208px; POSITION: absolute; TOP: 8px" runat="server"
					Height="40px" Width="165px" ForeColor="Blue" BackColor="Yellow"></asp:dropdownlist>
				<asp:dropdownlist id="DUserName" style="Z-INDEX: 103; LEFT: 376px; POSITION: absolute; TOP: 8px" runat="server"
					Height="40px" Width="100px" ForeColor="Blue" BackColor="Yellow"></asp:dropdownlist>
				<asp:button id="Go" style="Z-INDEX: 100; LEFT: 584px; POSITION: absolute; TOP: 8px" runat="server"
					Height="24px" Width="40px" ForeColor="Blue" BackColor="White" Text="Go"></asp:button>
				<asp:datagrid id="DataGrid1" style="Z-INDEX: 104; LEFT: 8px; POSITION: absolute; TOP: 40px" runat="server"
					Height="300px" Width="722px" BackColor="White" BorderStyle="None" AutoGenerateColumns="False"
					CellPadding="4" BorderWidth="1px" BorderColor="#CC9966" Font-Size="9pt">
					<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
					<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
					<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
					<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
					<Columns>
						<asp:BoundColumn DataField="AllFormDesc" ReadOnly="True" HeaderText="代理類型">
							<HeaderStyle Width="72px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="FormName" ReadOnly="True" HeaderText="委託單">
							<HeaderStyle Width="100px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="RangeDesc" ReadOnly="True" HeaderText="代理期間">
							<HeaderStyle Width="230px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="UserName" ReadOnly="True" HeaderText="原擔當">
							<HeaderStyle Width="60px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="AgentName" ReadOnly="True" HeaderText="代理人">
							<HeaderStyle Width="60px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="Description" ReadOnly="True" HeaderText="說明">
							<HeaderStyle Width="200px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
					</Columns>
					<PagerStyle NextPageText="下一頁" PrevPageText="上一頁" HorizontalAlign="Center" ForeColor="#993300"
						BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
				</asp:datagrid></FONT>
		</form>
	</body>
</HTML>
