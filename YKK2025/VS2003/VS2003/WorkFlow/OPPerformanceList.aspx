<%@ Page Language="vb" AutoEventWireup="false" Codebehind="OPPerformanceList.aspx.vb" Inherits="SPD.OPPerformanceList"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>OPPerformanceList</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="�s�ө���">
				<asp:datagrid id="DataGrid1" style="Z-INDEX: 100; LEFT: 8px; POSITION: absolute; TOP: 8px" runat="server"
					BorderStyle="None" BackColor="White" Height="300px" Width="1000px" PageSize="15" BorderColor="#CC9966"
					BorderWidth="1px" CellPadding="4" AutoGenerateColumns="False" Font-Size="9pt">
					<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
					<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
					<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
					<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
					<Columns>
						<asp:BoundColumn DataField="Buyer" ReadOnly="True" HeaderText="Buyer">
							<HeaderStyle Width="80px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="URL" DataNavigateUrlFormatString="{0}" DataTextField="No"
							HeaderText="No">
							<HeaderStyle Width="60px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:HyperLinkColumn>
						<asp:BoundColumn DataField="Sts" ReadOnly="True" HeaderText="���A">
							<HeaderStyle Width="40px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="Division" ReadOnly="True" HeaderText="�̿ೡ��">
							<HeaderStyle Width="80px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="Finished" ReadOnly="True" HeaderText="����">
							<HeaderStyle Width="60px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="Suppiler" ReadOnly="True" HeaderText="�~�`">
							<HeaderStyle Width="60px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="Level" ReadOnly="True" HeaderText="������">
							<HeaderStyle Width="40px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="Step" ReadOnly="True" HeaderText="�u�{No.">
							<HeaderStyle Width="60px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="TimeRange" ReadOnly="True" HeaderText="�_ ~ ��">
							<HeaderStyle Width="120px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="BTime" ReadOnly="True" HeaderText="�w�w(A)">
							<HeaderStyle Width="50px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Right"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="ATime" ReadOnly="True" HeaderText="���Z(B)">
							<HeaderStyle Width="50px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Right"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="Balance" ReadOnly="True" HeaderText="A-B">
							<HeaderStyle Width="50px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Right"></ItemStyle>
						</asp:BoundColumn>
					</Columns>
					<PagerStyle NextPageText="�U�@��" PrevPageText="�W�@��" HorizontalAlign="Center" ForeColor="#993300"
						BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
				</asp:datagrid></FONT></form>
	</body>
</HTML>
