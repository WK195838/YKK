<%@ Page Language="vb" AutoEventWireup="false" Codebehind="DevelopLeadTimeList.aspx.vb" Inherits="SPD.DevelopLeadTimeList"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>開發時間</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體">
				<asp:datagrid id="DataGrid1" style="Z-INDEX: 113; LEFT: 8px; POSITION: absolute; TOP: 32px" runat="server"
					BackColor="White" Height="300px" Width="450px" BorderStyle="None" Font-Size="9pt" AutoGenerateColumns="False"
					CellPadding="4" BorderWidth="1px" BorderColor="#CC9966" PageSize="15">
					<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
					<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
					<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
					<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
					<Columns>
						<asp:BoundColumn DataField="Step" ReadOnly="True" HeaderText="工程No.">
							<HeaderStyle Width="60px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="BTimeDesc" ReadOnly="True" HeaderText="收件時間 ~ 實際完工">
							<HeaderStyle Width="250px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="BTime" ReadOnly="True" HeaderText="時間">
							<HeaderStyle Width="80px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
					</Columns>
					<PagerStyle NextPageText="下一頁" PrevPageText="上一頁" HorizontalAlign="Center" ForeColor="#993300"
						BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
				</asp:datagrid>
				<DIV title="" style="DISPLAY: inline; Z-INDEX: 121; LEFT: 376px; WIDTH: 79px; COLOR: #0000ff; POSITION: absolute; TOP: 8px; HEIGHT: 24px"
					ms_positioning="FlowLayout">單位：分</DIV>
				<asp:imagebutton id="BExcel" style="Z-INDEX: 120; LEFT: 8px; POSITION: absolute; TOP: 8px" runat="server"
					Height="21px" Width="21px" ImageUrl="Images\msexcel.gif"></asp:imagebutton></FONT></form>
	</body>
</HTML>
