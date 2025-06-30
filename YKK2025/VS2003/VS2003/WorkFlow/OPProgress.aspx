<%@ Page Language="vb" AutoEventWireup="false" Codebehind="OPProgress.aspx.vb" Inherits="SPD.OPProgress"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>進度查詢</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體"></FONT>
			<DIV title="篩選項目：" style="DISPLAY: inline; Z-INDEX: 117; LEFT: 8px; WIDTH: 88px; COLOR: #0000ff; POSITION: absolute; TOP: 8px; HEIGHT: 24px"
				ms_positioning="FlowLayout">篩選項目：</DIV>
			<asp:dropdownlist id="DFlowType" style="Z-INDEX: 133; LEFT: 472px; POSITION: absolute; TOP: 8px" runat="server"
				ForeColor="Blue" BackColor="Yellow" Width="104px" Height="40px">
				<asp:ListItem Value="1" Selected="True">核定</asp:ListItem>
			</asp:dropdownlist><asp:dropdownlist id="DActive" style="Z-INDEX: 118; LEFT: 264px; POSITION: absolute; TOP: 8px" runat="server"
				ForeColor="Blue" BackColor="Yellow" Width="100px" Height="40px">
				<asp:ListItem Value="1">開發中</asp:ListItem>
			</asp:dropdownlist><asp:datagrid id="DataGrid1" style="Z-INDEX: 119; LEFT: 8px; POSITION: absolute; TOP: 40px" runat="server"
				BackColor="White" Width="790px" Height="300px" AllowPaging="True" PageSize="15" BorderStyle="None" Font-Size="9pt"
				BorderColor="#CC9966" BorderWidth="1px" CellPadding="4" AutoGenerateColumns="False">
				<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
				<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
				<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
				<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
				<Columns>
					<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="ViewURL" DataNavigateUrlFormatString="{0}"
						DataTextField="No" HeaderText="委託No">
						<HeaderStyle Width="70px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:HyperLinkColumn>
					<asp:BoundColumn DataField="Person" ReadOnly="True" HeaderText="依賴者">
						<HeaderStyle Width="80px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="ApplyTime" ReadOnly="True" HeaderText="依賴日">
						<HeaderStyle Width="100px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="StepName" ReadOnly="True" HeaderText="工程">
						<HeaderStyle Width="90px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="SCURL" DataNavigateUrlFormatString="{0}" DataTextField="StepPerson"
						HeaderText="工程擔當">
						<HeaderStyle Width="70px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:HyperLinkColumn>
					<asp:BoundColumn DataField="BStartTime" ReadOnly="True" HeaderText="開始預定">
						<HeaderStyle Width="90px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="BEndTime" ReadOnly="True" HeaderText="完成預定">
						<HeaderStyle Width="90px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="ReadTime" ReadOnly="True" HeaderText="閱讀期限">
						<HeaderStyle Width="150px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="OPURL" DataNavigateUrlFormatString="{0}" DataTextField="WorkFlow"
						HeaderText="履歷">
						<HeaderStyle Width="30px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Center"></ItemStyle>
					</asp:HyperLinkColumn>
				</Columns>
				<PagerStyle NextPageText="下一頁" PrevPageText="上一頁" HorizontalAlign="Center" ForeColor="#993300"
					BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
			</asp:datagrid><asp:dropdownlist id="DFormName" style="Z-INDEX: 121; LEFT: 96px; POSITION: absolute; TOP: 8px" runat="server"
				ForeColor="Blue" BackColor="Yellow" Width="165px" Height="40px" AutoPostBack="True"></asp:dropdownlist><asp:dropdownlist id="DSts" style="Z-INDEX: 122; LEFT: 368px; POSITION: absolute; TOP: 8px" runat="server"
				ForeColor="Blue" BackColor="Yellow" Width="100px" Height="40px">
				<asp:ListItem Value="Normal">正常</asp:ListItem>
				<asp:ListItem Value="Delay">延遲</asp:ListItem>
				<asp:ListItem Value="ReadDelay">未閱讀</asp:ListItem>
			</asp:dropdownlist><asp:button id="Go" style="Z-INDEX: 132; LEFT: 584px; POSITION: absolute; TOP: 8px" runat="server"
				ForeColor="Blue" BackColor="White" Width="40px" Height="24px" Text="Go"></asp:button></form>
	</body>
</HTML>
