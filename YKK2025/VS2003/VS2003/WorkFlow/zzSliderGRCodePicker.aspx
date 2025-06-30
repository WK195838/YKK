<%@ Page Language="vb" AutoEventWireup="false" Codebehind="zzSliderGRCodePicker.aspx.vb" Inherits="SPD.SliderGRCodePicker"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>SliderGRCodePicker</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體">
				<asp:datagrid id="DataGrid1" style="Z-INDEX: 105; LEFT: 8px; POSITION: absolute; TOP: 32px" runat="server"
					DataKeyField="No" Font-Size="Smaller" AutoGenerateColumns="False" CellPadding="4" BorderWidth="1px"
					BorderColor="#CC9966" AllowPaging="True" BorderStyle="None" BackColor="White" Height="328px"
					Width="170px">
					<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
					<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
					<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
					<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
					<Columns>
						<asp:ButtonColumn Text="選取" HeaderText="點選" CommandName="Select"></asp:ButtonColumn>
						<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="URL" DataNavigateUrlFormatString="{0}" DataTextField="FormNoDesc"
							HeaderText="單號"></asp:HyperLinkColumn>
						<asp:BoundColumn DataField="No" ReadOnly="True" HeaderText="拉頭群組">
							<HeaderStyle Width="60px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
					</Columns>
					<PagerStyle NextPageText="下一頁" PrevPageText="上一頁" HorizontalAlign="Center" ForeColor="#993300"
						BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
				</asp:datagrid>
				<asp:textbox id="DKey" style="Z-INDEX: 106; LEFT: 8px; POSITION: absolute; TOP: 8px" runat="server"
					BorderStyle="Groove" BackColor="Yellow" Height="20px" Width="152px" ForeColor="Blue">DKey</asp:textbox>
				<asp:button id="BKey" style="Z-INDEX: 103; LEFT: 160px; POSITION: absolute; TOP: 8px" runat="server"
					Height="20px" Width="20px" Text="....."></asp:button></FONT>
		</form>
	</body>
</HTML>
