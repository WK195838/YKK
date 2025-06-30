<%@ Page Language="vb" AutoEventWireup="false" Codebehind="BefOPList.aspx.vb" Inherits="SPD.BefOPList"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>工程履歷</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體">
				<asp:datagrid id="DataGrid1" style="Z-INDEX: 101; POSITION: absolute; TOP: 40px; LEFT: 8px" runat="server"
					Font-Size="9pt" BorderStyle="None" BorderColor="#CC9966" BorderWidth="1px" CellPadding="4"
					AutoGenerateColumns="False" Width="800px" Height="300px" BackColor="White">
					<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
					<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
					<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
					<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
					<Columns>
						<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="POPURL" DataNavigateUrlFormatString="{0}"
							DataTextField="POP">
							<HeaderStyle Width="20px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:HyperLinkColumn>
						<asp:BoundColumn DataField="StepNameDesc" ReadOnly="True" HeaderText="工程">
							<HeaderStyle Width="150px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="DecideName" ReadOnly="True" HeaderText="擔當">
							<HeaderStyle Width="50px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="AgentName" ReadOnly="True" HeaderText="代理/兼職">
							<HeaderStyle Width="70px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="FlowTypeDesc" ReadOnly="True" HeaderText="類別">
							<HeaderStyle Width="50px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="DelaySts" ReadOnly="True" HeaderText="處理進度">
							<HeaderStyle Width="50px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="StsDesc" ReadOnly="True" HeaderText="處理結果">
							<HeaderStyle Width="50px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="DecideDescA" ReadOnly="True" HeaderText="說明">
							<HeaderStyle Width="170px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="Description" ReadOnly="True" HeaderText="參考資料">
							<HeaderStyle Width="200px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
					</Columns>
					<PagerStyle NextPageText="下一頁" PrevPageText="上一頁" HorizontalAlign="Center" ForeColor="#993300"
						BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
				</asp:datagrid>
				<asp:imagebutton id="BExcel" style="Z-INDEX: 296; POSITION: absolute; TOP: 8px; LEFT: 96px" runat="server"
					Height="21px" Width="21px" ImageUrl="Images\msexcel.gif"></asp:imagebutton><INPUT id="BRead" style="Z-INDEX: 295; POSITION: absolute; WIDTH: 80px; HEIGHT: 28px; TOP: 8px; LEFT: 8px"
					type="button" value="閱讀完成" name="BOK" runat="server"></FONT></form>
	</body>
</HTML>
