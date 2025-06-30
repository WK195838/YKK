<%@ Page Language="vb" AutoEventWireup="false" Codebehind="GetCommissionNo.aspx.vb" Inherits="SPD.GetCommissionNo"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>GetCommissionNo</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體">
				<DIV title="篩選項目：" style="DISPLAY: inline; Z-INDEX: 100; LEFT: 8px; WIDTH: 88px; COLOR: #0000ff; POSITION: absolute; TOP: 8px; HEIGHT: 24px"
					ms_positioning="FlowLayout">篩選項目：</DIV>
				<asp:imagebutton id="BExcel" style="Z-INDEX: 150; LEFT: 432px; POSITION: absolute; TOP: 8px" runat="server"
					Width="21px" Height="21px" ImageUrl="Images\msexcel.gif"></asp:imagebutton><asp:textbox id="DMsg" style="Z-INDEX: 106; LEFT: 8px; POSITION: absolute; TOP: 40px" runat="server"
					ForeColor="Red" BackColor="White" Width="208px" Height="20px" BorderStyle="None" ReadOnly="True">顯示最後發行20個委託單號碼</asp:textbox><asp:datagrid id="DataGrid1" style="Z-INDEX: 105; LEFT: 8px; POSITION: absolute; TOP: 64px" runat="server"
					BackColor="White" Width="580px" Height="300px" BorderStyle="None" PageSize="15" Font-Size="9pt" BorderColor="#CC9966" BorderWidth="1px" CellPadding="4" AutoGenerateColumns="False">
					<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
					<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
					<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
					<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
					<Columns>
						<asp:BoundColumn DataField="No" ReadOnly="True" HeaderText="委託單號碼">
							<HeaderStyle Width="100px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="Sts" ReadOnly="True" HeaderText="狀態">
							<HeaderStyle Width="100px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="CreateTime" ReadOnly="True" HeaderText="發行日">
							<HeaderStyle Width="180px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="CreateUser" ReadOnly="True" HeaderText="發行人">
							<HeaderStyle Width="100px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="FormSno" ReadOnly="True" HeaderText="系統單號">
							<HeaderStyle Width="100px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
					</Columns>
					<PagerStyle NextPageText="下一頁" PrevPageText="上一頁" HorizontalAlign="Center" ForeColor="#993300"
						BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
				</asp:datagrid><asp:button id="Go" style="Z-INDEX: 104; LEFT: 384px; POSITION: absolute; TOP: 8px" runat="server"
					ForeColor="Blue" BackColor="White" Width="40px" Height="24px" Text="Go"></asp:button><asp:dropdownlist id="DDivision" style="Z-INDEX: 102; LEFT: 256px; POSITION: absolute; TOP: 8px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Width="120px" Height="40px"></asp:dropdownlist><asp:dropdownlist id="DFormName" style="Z-INDEX: 101; LEFT: 96px; POSITION: absolute; TOP: 8px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Width="160px" Height="40px"></asp:dropdownlist></FONT></form>
	</body>
</HTML>
