<%@ Page Language="vb" AutoEventWireup="false" Codebehind="ManufOutList.aspx.vb" Inherits="SPD.ManufOutOPList"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>�~�`�U���s����</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="�s�ө���">
				<asp:datagrid style="Z-INDEX: 113; LEFT: 8px; POSITION: absolute; TOP: 8px" id="DataGrid1" runat="server"
					Height="300px" Width="800px" BackColor="White" BorderStyle="None" AutoGenerateColumns="False"
					CellPadding="4" BorderWidth="1px" BorderColor="#CC9966" Font-Size="9pt">
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
						<asp:BoundColumn DataField="SliderCode" ReadOnly="True" HeaderText="���Y">
							<HeaderStyle Width="100px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="MapNo" ReadOnly="True" HeaderText="�ϸ�">
							<HeaderStyle Width="100px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="OFormDesc" ReadOnly="True" HeaderText="��e�U">
							<HeaderStyle Width="100px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="NFormDesc" ReadOnly="True" HeaderText="�s�e�U/Item">
							<HeaderStyle Width="100px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
					</Columns>
					<PagerStyle NextPageText="�U�@��" PrevPageText="�W�@��" HorizontalAlign="Center" ForeColor="#993300"
						BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
				</asp:datagrid></FONT></form>
	</body>
</HTML>
