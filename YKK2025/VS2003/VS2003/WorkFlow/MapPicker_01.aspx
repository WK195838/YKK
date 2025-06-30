<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MapPicker_01.aspx.vb" Inherits="SPD.MapPicker_01"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>MapPicker_01</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體">
				<asp:textbox id="DKey" style="Z-INDEX: 104; LEFT: 0px; POSITION: absolute; TOP: 8px" runat="server"
					BorderStyle="Groove" BackColor="Yellow" Height="20px" Width="146px" ForeColor="Blue">DKey</asp:textbox><asp:button id="BKey" style="Z-INDEX: 103; LEFT: 148px; POSITION: absolute; TOP: 8px" runat="server"
					Height="20px" Width="20px" Text="....."></asp:button><FONT face="新細明體"><asp:datagrid id="DataGrid1" style="Z-INDEX: 101; LEFT: 0px; POSITION: absolute; TOP: 32px" runat="server"
						BorderStyle="None" BackColor="White" Height="328px" Width="168px" DataKeyField="MapNo" Font-Size="Smaller" AutoGenerateColumns="False" CellPadding="4" BorderWidth="1px"
						BorderColor="#CC9966" AllowPaging="True">
						<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
						<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
						<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
						<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
						<Columns>
							<asp:ButtonColumn Text="選取" HeaderText="點選" CommandName="Select"></asp:ButtonColumn>
							<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="URL" DataNavigateUrlFormatString="{0}" DataTextField="MapNo"
								HeaderText="圖號"></asp:HyperLinkColumn>
						</Columns>
						<PagerStyle NextPageText="下一頁" PrevPageText="上一頁" HorizontalAlign="Center" ForeColor="#993300"
							BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
					</asp:datagrid></FONT></FONT></form>
	</body>
</HTML>
