<%@ Page Language="vb" AutoEventWireup="false" Codebehind="OPSchedule_01.aspx.vb" Inherits="SPD.OPSchedule_01"%>
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
			<FONT face="�s�ө���">
				<DIV title="�z�ﶵ�ءG" style="DISPLAY: inline; Z-INDEX: 107; LEFT: 8px; WIDTH: 88px; COLOR: #0000ff; POSITION: absolute; TOP: 8px; HEIGHT: 24px"
					ms_positioning="FlowLayout">�z�ﶵ�ءG</DIV>
				<asp:datagrid id="DataGrid1" style="Z-INDEX: 103; LEFT: 8px; POSITION: absolute; TOP: 40px" runat="server"
					AutoGenerateColumns="False" CellPadding="4" BorderWidth="1px" BorderColor="#CC9966" Font-Size="9pt"
					BorderStyle="None" BackColor="White" Width="744px" Height="300px">
					<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
					<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
					<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
					<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
					<Columns>
						<asp:BoundColumn DataField="StepNameDesc" ReadOnly="True" HeaderText="�u�{">
							<HeaderStyle Width="148px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="Description" ReadOnly="True" HeaderText="�w�w��{">
							<HeaderStyle Width="232px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="OPURL" DataNavigateUrlFormatString="{0}" DataTextField="WorkFlow"
							HeaderText="�u�{��T">
							<HeaderStyle Width="60px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:HyperLinkColumn>
						<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="ViewURL" DataNavigateUrlFormatString="{0}"
							DataTextField="FormName" HeaderText="�e�U��">
							<HeaderStyle Width="100px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:HyperLinkColumn>
						<asp:BoundColumn DataField="FormSno" ReadOnly="True" HeaderText="�渹">
							<HeaderStyle Width="36px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="FlowTypeDesc" ReadOnly="True" HeaderText="���O">
							<HeaderStyle Width="48px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="ApplyName" ReadOnly="True" HeaderText="�e�U�H">
							<HeaderStyle Width="60px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="DecideName" ReadOnly="True" HeaderText="�֩w�H">
							<HeaderStyle Width="60px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
					</Columns>
					<PagerStyle NextPageText="�U�@��" PrevPageText="�W�@��" HorizontalAlign="Center" ForeColor="#993300"
						BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
				</asp:datagrid>
				<asp:dropdownlist id="DFlowType" style="Z-INDEX: 106; LEFT: 96px; POSITION: absolute; TOP: 8px" runat="server"
					BackColor="Yellow" Width="104px" Height="40px" ForeColor="Blue">
					<asp:ListItem Value="ALL">ALL</asp:ListItem>
					<asp:ListItem Value="0">�q��</asp:ListItem>
					<asp:ListItem Value="1" Selected="True">�֩w</asp:ListItem>
				</asp:dropdownlist>
				<asp:button id="Go" style="Z-INDEX: 101; LEFT: 200px; POSITION: absolute; TOP: 8px" runat="server"
					BackColor="White" Width="40px" Height="24px" ForeColor="Blue" Text="Go"></asp:button></FONT></form>
	</body>
</HTML>
