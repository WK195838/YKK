<%@ Page Language="vb" AutoEventWireup="false" Codebehind="DevelopDataPicker.aspx.vb" Inherits="SPD.DevelopDataPicker"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>�}�o��ƿ��</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="�s�ө���">
				<asp:textbox id="DKey" style="Z-INDEX: 102; LEFT: 0px; POSITION: absolute; TOP: 0px" runat="server"
					ForeColor="Blue" Width="280px" Height="20px" BackColor="Yellow" BorderStyle="Groove">DKey</asp:textbox><asp:datagrid id="DataGrid1" style="Z-INDEX: 103; LEFT: 0px; POSITION: absolute; TOP: 24px" runat="server"
					Width="300px" Height="328px" BackColor="White" BorderStyle="None" AllowPaging="True" BorderColor="#CC9966" BorderWidth="1px" CellPadding="4" AutoGenerateColumns="False" Font-Size="Smaller" DataKeyField="Rno">
					<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
					<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
					<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
					<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
					<Columns>
						<asp:ButtonColumn Text="���" HeaderText="�I��" CommandName="Select"></asp:ButtonColumn>
						<asp:BoundColumn DataField="RNO" ReadOnly="True" HeaderText="�e�U-No">
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="CodeNo" ReadOnly="True" HeaderText="Code-No">
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="DevNo" ReadOnly="True" HeaderText="�}�o-No">
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
					</Columns>
					<PagerStyle NextPageText="�U�@��" PrevPageText="�W�@��" HorizontalAlign="Center" ForeColor="#993300"
						BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
				</asp:datagrid><asp:button id="BKey" style="Z-INDEX: 101; LEFT: 280px; POSITION: absolute; TOP: 0px" runat="server"
					Width="20px" Height="20px" Text="....."></asp:button><INPUT style="Z-INDEX: 104; LEFT: 312px; POSITION: absolute; TOP: 16px" type="button" value="Button"
					onclick="test();"></FONT></form>
	</body>
</HTML>
