<%@ Page Language="vb" AutoEventWireup="false" Codebehind="BefOPList.aspx.vb" Inherits="SPD.BefOPList"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>�u�{�i��</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="�s�ө���">
				<asp:datagrid id="DataGrid1" style="Z-INDEX: 101; POSITION: absolute; TOP: 40px; LEFT: 8px" runat="server"
					Font-Size="9pt" BorderStyle="None" BorderColor="#CC9966" BorderWidth="1px" CellPadding="4"
					AutoGenerateColumns="False" Width="800px" Height="300px" BackColor="White">
					<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
					<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
					<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
					<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
					<Columns>
						<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="POPURL" DataNavigateUrlFormatString="{0}"
							DataTextField="POP">
							<HeaderStyle Width="20px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:HyperLinkColumn>
						<asp:BoundColumn DataField="StepNameDesc" ReadOnly="True" HeaderText="�u�{">
							<HeaderStyle Width="150px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="DecideName" ReadOnly="True" HeaderText="���">
							<HeaderStyle Width="50px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="AgentName" ReadOnly="True" HeaderText="�N�z/��¾">
							<HeaderStyle Width="70px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="FlowTypeDesc" ReadOnly="True" HeaderText="���O">
							<HeaderStyle Width="50px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="DelaySts" ReadOnly="True" HeaderText="�B�z�i��">
							<HeaderStyle Width="50px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="StsDesc" ReadOnly="True" HeaderText="�B�z���G">
							<HeaderStyle Width="50px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="DecideDescA" ReadOnly="True" HeaderText="����">
							<HeaderStyle Width="170px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="Description" ReadOnly="True" HeaderText="�ѦҸ��">
							<HeaderStyle Width="200px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
					</Columns>
					<PagerStyle NextPageText="�U�@��" PrevPageText="�W�@��" HorizontalAlign="Center" ForeColor="#993300"
						BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
				</asp:datagrid>
				<asp:imagebutton id="BExcel" style="Z-INDEX: 296; POSITION: absolute; TOP: 8px; LEFT: 96px" runat="server"
					Height="21px" Width="21px" ImageUrl="Images\msexcel.gif"></asp:imagebutton><INPUT id="BRead" style="Z-INDEX: 295; POSITION: absolute; WIDTH: 80px; HEIGHT: 28px; TOP: 8px; LEFT: 8px"
					type="button" value="�\Ū����" name="BOK" runat="server"></FONT></form>
	</body>
</HTML>
