<%@ Page Language="VB" AutoEventWireup="false" CodeFile="WaitHandle_IRW.aspx.vb" Inherits="WaitHandle_IRW" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

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
			<FONT face="新細明體">&nbsp;
				<DIV title="篩選項目：" style="DISPLAY: inline; Z-INDEX: 109; LEFT: 8px; WIDTH: 88px; COLOR: #0000ff; POSITION: absolute; TOP: 8px; HEIGHT: 24px"
					ms_positioning="FlowLayout">篩選項目：</DIV>
				<asp:dropdownlist id="DFormName" style="Z-INDEX: 108; LEFT: 96px; POSITION: absolute; TOP: 8px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="40px" Width="165px"></asp:dropdownlist><asp:button id="Go" style="Z-INDEX: 106; LEFT: 616px; POSITION: absolute; TOP: 7px" runat="server"
					ForeColor="Blue" BackColor="White" Height="24px" Width="40px" Text="Go"></asp:button><FONT face="新細明體">
					<asp:dropdownlist id="DFlowType" style="Z-INDEX: 101; LEFT: 264px; POSITION: absolute; TOP: 8px" runat="server"
						ForeColor="Blue" BackColor="Yellow" Height="40px" Width="100px">
						<asp:ListItem Value="1" Selected="True">核定</asp:ListItem>
						<asp:ListItem Value="0">通知</asp:ListItem>
					</asp:dropdownlist></FONT><asp:datagrid id="DataGrid1" style="Z-INDEX: 110; LEFT: 8px; POSITION: absolute; TOP: 41px" runat="server"
					BackColor="White" Height="100px" Width="1080px" AutoGenerateColumns="False" CellPadding="4" BorderWidth="1px" BorderColor="#CC9966" BorderStyle="None" Font-Size="9pt" PageSize="30">
					<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
					<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
					<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
					<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
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
						<asp:BoundColumn DataField="MapNo" ReadOnly="True" HeaderText="委託部門/圖號">
							<HeaderStyle Width="130px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="ApplyName" ReadOnly="True" HeaderText="委託人">
							<HeaderStyle Width="60px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="StepNameDesc" ReadOnly="True" HeaderText="工程">
							<HeaderStyle Width="115px"></HeaderStyle>
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
				</asp:datagrid>
                <asp:TextBox ID="DApplyName" runat="server" BackColor="Yellow" BorderStyle="Groove"
                    ForeColor="Blue" Height="18px" MaxLength="7" Style="z-index: 126; left: 367px;
                    position: absolute; top: 9px; text-align: left" Width="116px">申請者</asp:TextBox>
                <asp:TextBox ID="DNo" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
                    Height="18px" MaxLength="7" Style="z-index: 126; left: 491px; position: absolute;
                    top: 9px; text-align: left" Width="116px">No</asp:TextBox>
            </FONT></form>
	</body>
</HTML>
