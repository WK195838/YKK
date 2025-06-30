<%@ Page Language="vb" AutoEventWireup="false" Codebehind="InqManufOutList.aspx.vb" Inherits="SPD.InqManufOutList" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>外注委託統計查詢</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<DIV title="篩選項目：" style="DISPLAY: inline; Z-INDEX: 100; LEFT: 8px; WIDTH: 96px; COLOR: #0000ff; POSITION: absolute; TOP: 8px; HEIGHT: 24px"
				ms_positioning="FlowLayout">篩選項目：</DIV>
			<asp:imagebutton id="BExcel" style="Z-INDEX: 122; LEFT: 8px; POSITION: absolute; TOP: 40px" runat="server"
				ImageUrl="Images\msexcel.gif" Height="21px" Width="21px"></asp:imagebutton><asp:datagrid id="DataGrid1" style="Z-INDEX: 113; LEFT: 8px; POSITION: absolute; TOP: 64px" runat="server"
				Height="300px" Width="650px" BackColor="White" BorderStyle="None" Font-Size="9pt" AutoGenerateColumns="False" CellPadding="4" BorderWidth="1px" BorderColor="#CC9966" PageSize="15">
				<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
				<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
				<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
				<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
				<Columns>
					<asp:BoundColumn DataField="SOrB" ReadOnly="True" HeaderText="外注廠/Buyer">
						<HeaderStyle Wrap="False" Width="150px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="MonthlyCount" ReadOnly="True" HeaderText="當月件數">
						<HeaderStyle Wrap="False" Width="100px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Right"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="MonthlyRate" ReadOnly="True" HeaderText="當月件數/ 總件數%">
						<HeaderStyle Wrap="False" Width="150px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Right"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="CumulativeCount" ReadOnly="True" HeaderText="累計件數">
						<HeaderStyle Wrap="False" Width="100px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Right"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="CumulativeRate" ReadOnly="True" HeaderText="累計件數/ 總件數%">
						<HeaderStyle Wrap="False" Width="150px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Right"></ItemStyle>
					</asp:BoundColumn>
				</Columns>
				<PagerStyle NextPageText="下一頁" PrevPageText="上一頁" HorizontalAlign="Center" ForeColor="#993300"
					BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
			</asp:datagrid><asp:button id="Go" style="Z-INDEX: 109; LEFT: 512px; POSITION: absolute; TOP: 8px" runat="server"
				Height="24px" Width="40px" BackColor="White" ForeColor="Blue" Text="Go"></asp:button><asp:dropdownlist id="DSOrB" style="Z-INDEX: 107; LEFT: 280px; POSITION: absolute; TOP: 8px" runat="server"
				Height="40px" Width="80px" BackColor="Yellow" ForeColor="Blue">
				<asp:ListItem Value="Suppiler">外注廠</asp:ListItem>
				<asp:ListItem Value="Buyer">Buyer</asp:ListItem>
			</asp:dropdownlist><asp:dropdownlist id="DSts" style="Z-INDEX: 106; LEFT: 376px; POSITION: absolute; TOP: 8px" runat="server"
				Height="40px" Width="120px" BackColor="Yellow" ForeColor="Blue">
				<asp:ListItem Value="0">當月委託</asp:ListItem>
				<asp:ListItem Value="1">當月開發完成</asp:ListItem>
			</asp:dropdownlist>
			<DIV title="" style="DISPLAY: inline; Z-INDEX: 103; LEFT: 176px; WIDTH: 1px; COLOR: #0000ff; POSITION: absolute; TOP: 8px; HEIGHT: 24px"
				ms_positioning="FlowLayout">年</DIV>
			<asp:dropdownlist id="DYear" style="Z-INDEX: 102; LEFT: 104px; POSITION: absolute; TOP: 8px" runat="server"
				Height="40px" Width="72px" BackColor="Yellow" ForeColor="Blue"></asp:dropdownlist><asp:dropdownlist id="DMonth" style="Z-INDEX: 104; LEFT: 200px; POSITION: absolute; TOP: 8px" runat="server"
				Height="40px" Width="48px" BackColor="Yellow" ForeColor="Blue">
				<asp:ListItem Value="1">01</asp:ListItem>
				<asp:ListItem Value="2">02</asp:ListItem>
				<asp:ListItem Value="3">03</asp:ListItem>
				<asp:ListItem Value="4">04</asp:ListItem>
				<asp:ListItem Value="5">05</asp:ListItem>
				<asp:ListItem Value="6">06</asp:ListItem>
				<asp:ListItem Value="7">07</asp:ListItem>
				<asp:ListItem Value="8">08</asp:ListItem>
				<asp:ListItem Value="9">09</asp:ListItem>
				<asp:ListItem Value="10">10</asp:ListItem>
				<asp:ListItem Value="11">11</asp:ListItem>
				<asp:ListItem Value="12">12</asp:ListItem>
			</asp:dropdownlist>
			<DIV title="" style="DISPLAY: inline; Z-INDEX: 105; LEFT: 248px; WIDTH: 1px; COLOR: #0000ff; POSITION: absolute; TOP: 8px; HEIGHT: 24px"
				ms_positioning="FlowLayout">
				<P>月</P>
			</DIV>
		</form>
	</body>
</HTML>
