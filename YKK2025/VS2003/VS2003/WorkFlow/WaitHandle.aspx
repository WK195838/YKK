<%@ Page Language="vb" AutoEventWireup="false" Codebehind="WaitHandle.aspx.vb" Inherits="SPD.WaitHandle"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>待處理</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體">
				<asp:dropdownlist id="DSts" style="Z-INDEX: 100; LEFT: 264px; POSITION: absolute; TOP: 8px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="40px" Width="100px">
					<asp:ListItem Value="ALL" Selected="True">ALL</asp:ListItem>
					<asp:ListItem Value="0">未處理</asp:ListItem>
					<asp:ListItem Value="3">已閱讀</asp:ListItem>
				</asp:dropdownlist>
				<asp:textbox style="Z-INDEX: 113; LEFT: 368px; POSITION: absolute; TOP: 40px" id="DKey1" runat="server"
					Width="96px" Height="22px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DKey1</asp:textbox>
				<DIV title="篩選項目：" style="DISPLAY: inline; Z-INDEX: 109; LEFT: 8px; WIDTH: 88px; COLOR: #0000ff; POSITION: absolute; TOP: 8px; HEIGHT: 24px"
					ms_positioning="FlowLayout">篩選項目：</DIV>
				<asp:dropdownlist id="DFormName" style="Z-INDEX: 108; LEFT: 96px; POSITION: absolute; TOP: 8px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="40px" Width="165px" AutoPostBack="True"></asp:dropdownlist><asp:button id="Go" style="Z-INDEX: 106; LEFT: 472px; POSITION: absolute; TOP: 40px" runat="server"
					ForeColor="Blue" BackColor="White" Height="24px" Width="40px" Text="Go"></asp:button><asp:dropdownlist id="DSort" style="Z-INDEX: 105; LEFT: 264px; POSITION: absolute; TOP: 40px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="40px" Width="100px">
					<asp:ListItem Value="ASC">昇順</asp:ListItem>
					<asp:ListItem Value="DESC" Selected="True">降順</asp:ListItem>
				</asp:dropdownlist><FONT face="新細明體">
					<DIV title="排序：" style="DISPLAY: inline; Z-INDEX: 104; LEFT: 8px; WIDTH: 88px; COLOR: #0000ff; POSITION: absolute; TOP: 40px; HEIGHT: 24px"
						ms_positioning="FlowLayout">排序：</DIV>
					<asp:dropdownlist id="DSortKey" style="Z-INDEX: 103; LEFT: 96px; POSITION: absolute; TOP: 40px" runat="server"
						ForeColor="Blue" BackColor="Yellow" Height="40px" Width="165px">
						<asp:ListItem Value="No" Selected="True">委託No</asp:ListItem>
						<asp:ListItem Value="ApplyTime">委託時間</asp:ListItem>
						<asp:ListItem Value="ReceiptTime">收件時間</asp:ListItem>
						<asp:ListItem Value="ReadTimeLimit">閱讀期限</asp:ListItem>
						<asp:ListItem Value="FirstReadTime">第一次閱讀時間</asp:ListItem>
						<asp:ListItem Value="LastReadTime">最後一次閱讀時間</asp:ListItem>
						<asp:ListItem Value="BStartTime">預定開始時間</asp:ListItem>
						<asp:ListItem Value="BEndTime">預定完成時間</asp:ListItem>
						<asp:ListItem Value="AStartTime">實際開始時間</asp:ListItem>
					</asp:dropdownlist><asp:dropdownlist id="DApplyName" style="Z-INDEX: 102; LEFT: 472px; POSITION: absolute; TOP: 8px"
						runat="server" ForeColor="Blue" BackColor="Yellow" Height="40px" Width="100px"></asp:dropdownlist><asp:dropdownlist id="DFlowType" style="Z-INDEX: 101; LEFT: 368px; POSITION: absolute; TOP: 8px" runat="server"
						ForeColor="Blue" BackColor="Yellow" Height="40px" Width="100px">
						<asp:ListItem Value="ALL" Selected="True">ALL</asp:ListItem>
						<asp:ListItem Value="0">通知</asp:ListItem>
						<asp:ListItem Value="1">核定</asp:ListItem>
					</asp:dropdownlist></FONT><asp:datagrid id="DataGrid1" style="Z-INDEX: 110; LEFT: 8px; POSITION: absolute; TOP: 72px" runat="server"
					BackColor="White" Height="100px" Width="1080px" AutoGenerateColumns="False" CellPadding="4" BorderWidth="1px" BorderColor="#CC9966"
					AllowPaging="True" BorderStyle="None" Font-Size="9pt" PageSize="15">
					<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
					<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
					<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
					<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
					<Columns>
						<asp:BoundColumn DataField="StsDesc" ReadOnly="True" HeaderText="狀態">
							<HeaderStyle Width="48px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:HyperLinkColumn Target="_self" DataNavigateUrlField="URL" DataNavigateUrlFormatString="{0}" DataTextField="FormName"
							HeaderText="委託單">
							<HeaderStyle Width="84px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:HyperLinkColumn>
						<asp:BoundColumn DataField="No" ReadOnly="True" HeaderText="委託No">
							<HeaderStyle Width="80px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="MapNo" ReadOnly="True" HeaderText="委託部門/圖號/客訴內容">
							<HeaderStyle Width="130px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="ApplyName" ReadOnly="True" HeaderText="委託人">
							<HeaderStyle Width="60px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="StepNameDesc" ReadOnly="True" HeaderText="工程/請假起迄">
							<HeaderStyle Width="180px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="FlowTypeDesc" ReadOnly="True" HeaderText="類別">
							<HeaderStyle Width="60px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="AgentName" ReadOnly="True" HeaderText="代理/兼職">
							<HeaderStyle Width="60px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="Description" ReadOnly="True" HeaderText="參考資料">
							<HeaderStyle Width="464px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
					</Columns>
					<PagerStyle NextPageText="下一頁" PrevPageText="上一頁" HorizontalAlign="Center" ForeColor="#993300"
						BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
				</asp:datagrid></FONT></form>
	</body>
</HTML>
