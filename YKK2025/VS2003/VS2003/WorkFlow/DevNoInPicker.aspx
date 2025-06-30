<%@ Page Language="vb" AutoEventWireup="false" Codebehind="DevNoInPicker.aspx.vb" Inherits="SPD.DevNoPicker"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>DevNoPicker</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體">
				<asp:textbox id="DKey" style="Z-INDEX: 104; LEFT: 0px; POSITION: absolute; TOP: 8px" runat="server"
					BorderStyle="Groove" BackColor="Yellow" Height="20px" Width="224px" ForeColor="Blue">DKey</asp:textbox>
				<asp:datagrid id="DataGrid1" style="Z-INDEX: 105; LEFT: 0px; POSITION: absolute; TOP: 32px" runat="server"
					BorderStyle="None" BackColor="White" Height="328px" Width="250px" DataKeyField="MapNo" Font-Size="Smaller"
					AutoGenerateColumns="False" CellPadding="4" BorderWidth="1px" BorderColor="#CC9966" AllowPaging="True">
					<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
					<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
					<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
					<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
					<Columns>
						<asp:ButtonColumn Text="選取" HeaderText="點選" CommandName="Select"></asp:ButtonColumn>
						<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="URL" DataNavigateUrlFormatString="{0}" DataTextField="No"
							HeaderText="開發No"></asp:HyperLinkColumn>
						<asp:BoundColumn DataField="FormNoDesc" ReadOnly="True" HeaderText="表單號碼-單號">
							<HeaderStyle Width="84px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
					</Columns>
					<PagerStyle NextPageText="下一頁" PrevPageText="上一頁" HorizontalAlign="Center" ForeColor="#993300"
						BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
				</asp:datagrid>
				<asp:button id="BKey" style="Z-INDEX: 103; LEFT: 224px; POSITION: absolute; TOP: 8px" runat="server"
					Height="20px" Width="20px" Text="....."></asp:button></FONT>
		</form>
	</body>
</HTML>
