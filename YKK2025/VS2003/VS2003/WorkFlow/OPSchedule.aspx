<%@ Page Language="vb" AutoEventWireup="false" Codebehind="OPSchedule.aspx.vb" Inherits="SPD.OPScheduleList"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>��{�d��</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="�s�ө���"></FONT>
			<DIV title="�z�ﶵ�ءG" style="DISPLAY: inline; Z-INDEX: 100; LEFT: 8px; WIDTH: 88px; COLOR: #0000ff; POSITION: absolute; TOP: 8px; HEIGHT: 24px"
				ms_positioning="FlowLayout">�z�ﶵ�ءG</DIV>
			<asp:dropdownlist id="DFlowType" style="Z-INDEX: 106; LEFT: 96px; POSITION: absolute; TOP: 8px" runat="server"
				ForeColor="Blue" BackColor="Yellow" Width="104px" Height="40px">
				<asp:ListItem Value="ALL">ALL</asp:ListItem>
				<asp:ListItem Value="0">�q��</asp:ListItem>
				<asp:ListItem Value="1">�֩w</asp:ListItem>
			</asp:dropdownlist>
			<asp:datagrid id="DataGrid1" style="Z-INDEX: 103; LEFT: 8px; POSITION: absolute; TOP: 40px" runat="server"
				BackColor="White" Width="770px" Height="300px" AutoGenerateColumns="False" CellPadding="4"
				BorderWidth="1px" BorderColor="#CC9966" Font-Size="9pt" BorderStyle="None" PageSize="20">
				<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
				<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
				<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
				<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
				<Columns>
					<asp:BoundColumn DataField="StepPerson" ReadOnly="True" HeaderText="�u�{���">
						<HeaderStyle Width="70px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="StepName" ReadOnly="True" HeaderText="�u�{">
						<HeaderStyle Width="140px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="BOPTime" ReadOnly="True" HeaderText="�w�w��{">
						<HeaderStyle Width="240px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="No" ReadOnly="True" HeaderText="�e�UNo">
						<HeaderStyle Width="120px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="Person" ReadOnly="True" HeaderText="�̿��">
						<HeaderStyle Width="80px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="ApplyTime" ReadOnly="True" HeaderText="�̿��">
						<HeaderStyle Width="120px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
				</Columns>
				<PagerStyle NextPageText="�U�@��" PrevPageText="�W�@��" HorizontalAlign="Center" ForeColor="#993300"
					BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
			</asp:datagrid><asp:button id="Go" style="Z-INDEX: 101; LEFT: 200px; POSITION: absolute; TOP: 8px" runat="server"
				ForeColor="Blue" BackColor="White" Width="40px" Height="24px" Text="Go"></asp:button></form>
	</body>
</HTML>
