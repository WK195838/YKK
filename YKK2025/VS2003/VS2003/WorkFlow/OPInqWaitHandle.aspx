<%@ Page Language="vb" AutoEventWireup="false" Codebehind="OPInqWaitHandle.aspx.vb" Inherits="SPD.OPInqWaitHandle"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>OPInqWaitHandle</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體">
				<DIV title="篩選項目：" style="DISPLAY: inline; Z-INDEX: 101; LEFT: 8px; WIDTH: 88px; COLOR: #0000ff; POSITION: absolute; TOP: 8px; HEIGHT: 24px"
					ms_positioning="FlowLayout">篩選項目：</DIV>
				<asp:textbox id="DBuyer" style="Z-INDEX: 110; LEFT: 528px; POSITION: absolute; TOP: 32px" runat="server"
					ReadOnly="True" BorderStyle="Groove" BackColor="Yellow" Height="20px" Width="120px" ForeColor="Blue">DBuyer</asp:textbox>
				<asp:textbox id="DDecide" style="Z-INDEX: 109; LEFT: 408px; POSITION: absolute; TOP: 32px" runat="server"
					ForeColor="Blue" Width="120px" Height="20px" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DDecide</asp:textbox>
				<asp:datagrid id="DataGrid1" style="Z-INDEX: 108; LEFT: 8px; POSITION: absolute; TOP: 64px" runat="server"
					Width="923px" Height="300px" BackColor="White" BorderStyle="None" PageSize="15" BorderColor="#CC9966"
					BorderWidth="1px" CellPadding="4" AutoGenerateColumns="False" Font-Size="9pt">
					<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
					<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
					<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
					<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
					<Columns>
						<asp:BoundColumn DataField="T_StsDesc" ReadOnly="True" HeaderText="處理結果">
							<HeaderStyle Width="48px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="T_FlowTypeDesc" ReadOnly="True" HeaderText="類別">
							<HeaderStyle Width="48px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:HyperLinkColumn Target="blank" DataNavigateUrlField="T_ViewURL" DataNavigateUrlFormatString="{0}"
							DataTextField="T_FormName" HeaderText="委託單">
							<HeaderStyle Width="84px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:HyperLinkColumn>
						<asp:BoundColumn DataField="No" ReadOnly="True" HeaderText="委託No">
							<HeaderStyle Width="80px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="Buyer" ReadOnly="True" HeaderText="Buyer">
							<HeaderStyle Width="80px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="T_StepNameDesc" ReadOnly="True" HeaderText="工程">
							<HeaderStyle Width="115px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="T_DecideName" ReadOnly="True" HeaderText="擔當">
							<HeaderStyle Width="60px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="Division" ReadOnly="True" HeaderText="部門">
							<HeaderStyle Width="80px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="T_ApplyName" ReadOnly="True" HeaderText="委託人">
							<HeaderStyle Width="48px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="Description" ReadOnly="True" HeaderText="參考資料">
							<HeaderStyle Width="280px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
					</Columns>
					<PagerStyle NextPageText="下一頁" PrevPageText="上一頁" HorizontalAlign="Center" ForeColor="#993300"
						BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
				</asp:datagrid><asp:dropdownlist id="DDivision" style="Z-INDEX: 107; LEFT: 96px; POSITION: absolute; TOP: 32px" runat="server"
					ForeColor="Blue" Width="184px" Height="40px" BackColor="Yellow"></asp:dropdownlist><asp:textbox id="DPerson" style="Z-INDEX: 106; LEFT: 288px; POSITION: absolute; TOP: 32px" runat="server"
					ForeColor="Blue" Width="120px" Height="20px" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DPerson</asp:textbox><asp:dropdownlist id="DStep" style="Z-INDEX: 103; LEFT: 288px; POSITION: absolute; TOP: 8px" runat="server"
					ForeColor="Blue" Width="344px" Height="40px" BackColor="Yellow"></asp:dropdownlist><asp:dropdownlist id="DFormName" style="Z-INDEX: 100; LEFT: 96px; POSITION: absolute; TOP: 8px" runat="server"
					ForeColor="Blue" Width="184px" Height="40px" BackColor="Yellow" AutoPostBack="True"></asp:dropdownlist></FONT><asp:button id="Go" style="Z-INDEX: 104; LEFT: 656px; POSITION: absolute; TOP: 32px" runat="server"
				ForeColor="Blue" Width="40px" Height="24px" BackColor="White" Text="Go"></asp:button></form>
	</body>
</HTML>
