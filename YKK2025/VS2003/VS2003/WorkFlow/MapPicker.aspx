<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MapPicker.aspx.vb" Inherits="SPD.OMapPicker"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>�ϭ����</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="�s�ө���">
				<asp:textbox id="DKey" style="Z-INDEX: 104; LEFT: 0px; POSITION: absolute; TOP: 8px" runat="server"
					ForeColor="Blue" Width="146px" Height="20px" BackColor="Yellow" BorderStyle="Groove">DKey</asp:textbox>
				<asp:button id="BKey" style="Z-INDEX: 103; LEFT: 148px; POSITION: absolute; TOP: 8px" runat="server"
					Width="20px" Height="20px" Text="....."></asp:button><FONT face="�s�ө���">
					<asp:datagrid id="DataGrid1" style="Z-INDEX: 101; LEFT: 0px; POSITION: absolute; TOP: 32px" runat="server"
						Width="168px" Height="328px" BackColor="White" BorderStyle="None" AllowPaging="True" BorderColor="#CC9966"
						BorderWidth="1px" CellPadding="4" AutoGenerateColumns="False" Font-Size="Smaller" DataKeyField="MapNo">
						<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
						<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
						<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
						<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
						<Columns>
							<asp:ButtonColumn Text="���" HeaderText="�I��" CommandName="Select"></asp:ButtonColumn>
							<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="URL" DataNavigateUrlFormatString="{0}" DataTextField="MapNo"
								HeaderText="�ϸ�"></asp:HyperLinkColumn>
						</Columns>
						<PagerStyle NextPageText="�U�@��" PrevPageText="�W�@��" HorizontalAlign="Center" ForeColor="#993300"
							BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
					</asp:datagrid></FONT></FONT></form>
	</body>
</HTML>
