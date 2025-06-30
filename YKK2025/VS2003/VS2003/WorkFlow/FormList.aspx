<%@ Page Language="vb" AutoEventWireup="false" Codebehind="FormList.aspx.vb" Inherits="SPD.FormList"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>委託單</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體">
				<asp:datagrid id="DataGrid1" style="Z-INDEX: 126; LEFT: 8px; POSITION: absolute; TOP: 32px" runat="server"
					Font-Size="9pt" BorderColor="#CC9966" BorderWidth="1px" CellPadding="4" AutoGenerateColumns="False"
					BorderStyle="None" BackColor="White" Width="790px" Height="300px">
					<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
					<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
					<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
					<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
					<Columns>
						<asp:BoundColumn DataField="FormName" ReadOnly="True" HeaderText="委託單">
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="FormSno" ReadOnly="True" HeaderText="可使用單號">
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="FormUse" ReadOnly="True" HeaderText="用途">
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="IniAuthorityDesc" ReadOnly="True" HeaderText="起單權限">
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="IniUserDesc" ReadOnly="True" HeaderText="可起單者">
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
<%------------------------------------------------------
						<asp:BoundColumn DataField="InqAuthorityDesc" ReadOnly="True" HeaderText="查詢權限">
							<HeaderStyle Width="50px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="InqUserDesc" ReadOnly="True" HeaderText="可查詢者">
							<HeaderStyle Width="230px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
-----------%>						
					</Columns>
					<PagerStyle NextPageText="下一頁" PrevPageText="上一頁" HorizontalAlign="Center" ForeColor="#993300"
						BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
				</asp:datagrid>
				<asp:imagebutton id="BExcel" style="Z-INDEX: 150; LEFT: 8px; POSITION: absolute; TOP: 8px" runat="server"
					Height="21px" Width="21px" ImageUrl="Images\msexcel.gif"></asp:imagebutton></FONT>
		</form>
	</body>
</HTML>
