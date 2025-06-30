<%@ Page Language="vb" AutoEventWireup="false" Codebehind="FlowList.aspx.vb" Inherits="SPD.FlowList"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>流程</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體">
				<DIV title="篩選項目：" style="DISPLAY: inline; Z-INDEX: 127; LEFT: 8px; WIDTH: 88px; COLOR: #0000ff; POSITION: absolute; TOP: 8px; HEIGHT: 24px"
					ms_positioning="FlowLayout">篩選項目：</DIV>
				<asp:imagebutton id="BExcel" style="Z-INDEX: 282; LEFT: 320px; POSITION: absolute; TOP: 8px" runat="server"
					Width="21px" Height="21px" ImageUrl="Images\msexcel.gif"></asp:imagebutton>
				<asp:hyperlink id="LFlowChart" style="Z-INDEX: 281; LEFT: 720px; POSITION: absolute; TOP: 16px"
					runat="server" Height="8px" Width="80px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">工作流程圖</asp:hyperlink><asp:dropdownlist id="DFormName" style="Z-INDEX: 118; LEFT: 96px; POSITION: absolute; TOP: 8px" runat="server"
					Height="40px" Width="165px" ForeColor="Blue" BackColor="Yellow"></asp:dropdownlist><asp:button id="Go" style="Z-INDEX: 117; LEFT: 272px; POSITION: absolute; TOP: 8px" runat="server"
					Height="24px" Width="40px" ForeColor="Blue" BackColor="White" Text="Go"></asp:button><asp:datagrid id="DataGrid1" style="Z-INDEX: 126; LEFT: 8px; POSITION: absolute; TOP: 40px" runat="server"
					Height="300px" Width="900px" BackColor="White" BorderStyle="None" AutoGenerateColumns="False" CellPadding="4" BorderWidth="1px" BorderColor="#CC9966" Font-Size="9pt">
					<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
					<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
					<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
					<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
					<Columns>
						<asp:BoundColumn DataField="FormName" ReadOnly="True" HeaderText="委託單">
							<HeaderStyle Width="80px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="Step" ReadOnly="True" HeaderText="No">
							<HeaderStyle Width="20px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="StepNameDesc" ReadOnly="True" HeaderText="工程名">
							<HeaderStyle Width="170px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="RelatedTypeDesc" ReadOnly="True" HeaderText="核定類別">
							<HeaderStyle Width="50px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="RelatedName" ReadOnly="True" HeaderText="核定者">
							<HeaderStyle Width="50px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="SignTypeDesc" ReadOnly="True" HeaderText="核定模式">
							<HeaderStyle Width="50px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="FlowTypeDesc" ReadOnly="True" HeaderText="處理">
							<HeaderStyle Width="40px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="LoadingDesc" ReadOnly="True" HeaderText="負荷">
							<HeaderStyle Width="60px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="LeadTimeDesc" ReadOnly="True" HeaderText="L / T">
							<HeaderStyle Width="70px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="OKFunDesc" ReadOnly="True" HeaderText="按鈕(1)">
							<HeaderStyle Width="60px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="NGFunDesc1" ReadOnly="True" HeaderText="按鈕(2)">
							<HeaderStyle Width="60px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="NGFunDesc2" ReadOnly="True" HeaderText="按鈕(3)">
							<HeaderStyle Width="60px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
					</Columns>
					<PagerStyle NextPageText="下一頁" PrevPageText="上一頁" HorizontalAlign="Center" ForeColor="#993300"
						BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
				</asp:datagrid></FONT></form>
	</body>
</HTML>
