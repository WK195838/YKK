<%@ Page Language="vb" AutoEventWireup="false" Codebehind="CodeNoPicker.aspx.vb" Inherits="SPD.CodeNoPicker"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>Code No���</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="�s�ө���">
				<asp:textbox id="DKey" style="Z-INDEX: 104; LEFT: 0px; POSITION: absolute; TOP: 8px" runat="server"
					BorderStyle="Groove" BackColor="Yellow" Height="20px" Width="224px" ForeColor="Blue">DKey</asp:textbox>
				<asp:datagrid id="DataGrid1" style="Z-INDEX: 105; LEFT: 0px; POSITION: absolute; TOP: 32px" runat="server"
					BorderStyle="None" BackColor="White" Height="328px" Width="250px" DataKeyField="Code" Font-Size="Smaller"
					AutoGenerateColumns="False" CellPadding="4" BorderWidth="1px" BorderColor="#CC9966" AllowPaging="True">
					<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
					<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
					<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
					<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
					<Columns>
						<asp:ButtonColumn Text="���" HeaderText="�I��" CommandName="Select"></asp:ButtonColumn>
						<asp:BoundColumn DataField="Code" ReadOnly="True" HeaderText="Code">
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="URL" DataNavigateUrlFormatString="{0}" DataTextField="No"
							HeaderText="No"></asp:HyperLinkColumn>
					</Columns>
					<PagerStyle NextPageText="�U�@��" PrevPageText="�W�@��" HorizontalAlign="Center" ForeColor="#993300"
						BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
				</asp:datagrid>
				<asp:button id="BKey" style="Z-INDEX: 103; LEFT: 224px; POSITION: absolute; TOP: 8px" runat="server"
					Height="20px" Width="20px" Text="....."></asp:button></FONT>
		</form>
	</body>
</HTML>
