<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MapList.aspx.vb" Inherits="SPD.MapList"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>MapList</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="�s�ө���">
				<asp:datagrid id="DataGrid1" style="Z-INDEX: 113; LEFT: 8px; POSITION: absolute; TOP: 8px" runat="server"
					Font-Size="9pt" BorderColor="#CC9966" BorderWidth="1px" CellPadding="4" AutoGenerateColumns="False"
					BorderStyle="None" BackColor="White" Width="720px" Height="300px">
					<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
					<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
					<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
					<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
					<Columns>
						<asp:BoundColumn DataField="StsDesc" ReadOnly="True" HeaderText="���A">
							<HeaderStyle Width="80px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="URL" DataNavigateUrlFormatString="{0}" DataTextField="No"
							HeaderText="No">
							<HeaderStyle Width="80px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:HyperLinkColumn>
						<asp:BoundColumn DataField="DateDesc" ReadOnly="True" HeaderText="���">
							<HeaderStyle Width="100px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="MapNo" ReadOnly="True" HeaderText="�ϸ�">
							<HeaderStyle Width="100px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="OriMapNo" ReadOnly="True" HeaderText="��ϸ�">
							<HeaderStyle Width="100px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="BefMapNo" ReadOnly="True" HeaderText="�e�ϸ�">
							<HeaderStyle Width="100px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="ModReason" ReadOnly="True" HeaderText="�ק��]">
							<HeaderStyle Width="160px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
					</Columns>
					<PagerStyle NextPageText="�U�@��" PrevPageText="�W�@��" HorizontalAlign="Center" ForeColor="#993300"
						BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
				</asp:datagrid></FONT>
		</form>
	</body>
</HTML>
