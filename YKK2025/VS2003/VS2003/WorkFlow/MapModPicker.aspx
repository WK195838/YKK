<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MapModPicker.aspx.vb" Inherits="SPD.MapPicker" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>圖面修改選擇</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<style type="text/css">BODY { FONT-SIZE: 9pt; FONT-FAMILY: Verdana, Geneva, Sans-Serif }
	TABLE { FONT-SIZE: 9pt; FONT-FAMILY: Verdana, Geneva, Sans-Serif }
	TR { FONT-SIZE: 9pt; FONT-FAMILY: Verdana, Geneva, Sans-Serif }
	TD { FONT-SIZE: 9pt; FONT-FAMILY: Verdana, Geneva, Sans-Serif }
	UL { FONT-SIZE: 9pt; FONT-FAMILY: Verdana, Geneva, Sans-Serif }
	LI { FONT-SIZE: 9pt; FONT-FAMILY: Verdana, Geneva, Sans-Serif }
	.normal { FONT-SIZE: 9pt; FONT-FAMILY: Verdana, Geneva, Sans-Serif }
	H1 { FONT-WEIGHT: 900; FONT-SIZE: 10.5pt; COLOR: #666666; FONT-FAMILY: Verdana, Geneva, Sans-Serif }
	.small { FONT-SIZE: 7.5pt; FONT-FAMILY: Verdana, Geneva, Sans-Serif }
	.error { COLOR: #ff0033 }
	.required { FONT-WEIGHT: 900; COLOR: #ff0033; FONT-FAMILY: Verdana, Geneva, Sans-Serif }
	.success { FONT-WEIGHT: 900; FONT-SIZE: 11pt; MARGIN: 10px 0px; COLOR: #009933 }
		</style>
	</HEAD>
	<body bottomMargin="0" leftMargin="0" topMargin="0" rightMargin="0" MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體">
				<asp:datagrid id="DataGrid1" style="Z-INDEX: 101; LEFT: 0px; POSITION: absolute; TOP: 32px" runat="server"
					Width="168px" Height="328px" BackColor="White" BorderStyle="None" AllowPaging="True" BorderColor="#CC9966"
					BorderWidth="1px" CellPadding="4" AutoGenerateColumns="False" Font-Size="Smaller" DataKeyField="MapNo">
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
				</asp:datagrid><asp:button id="BKey" style="Z-INDEX: 103; LEFT: 148px; POSITION: absolute; TOP: 8px" runat="server"
					Width="20px" Height="20px" Text="....."></asp:button><asp:textbox id="DKey" style="Z-INDEX: 102; LEFT: 0px; POSITION: absolute; TOP: 8px" runat="server"
					Width="146px" Height="20px" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue">DKey</asp:textbox></FONT></form>
	</body>
</HTML>
