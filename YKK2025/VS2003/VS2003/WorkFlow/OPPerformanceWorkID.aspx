<%@ Page Language="vb" AutoEventWireup="false" Codebehind="OPPerformanceWorkID.aspx.vb" Inherits="SPD.OPPerformanceWorkID"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>OPPerformanceWorkID</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="�s�ө���">
				<asp:datagrid id="DataGrid1" style="Z-INDEX: 100; LEFT: 8px; POSITION: absolute; TOP: 40px" runat="server"
					Font-Size="9pt" AutoGenerateColumns="False" CellPadding="4" BorderWidth="1px" BorderColor="#CC9966"
					PageSize="15" Width="950px" Height="300px" BackColor="White" BorderStyle="None">
					<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
					<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
					<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
					<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
					<Columns>
						<asp:BoundColumn DataField="No" ReadOnly="True" HeaderText="No">
							<HeaderStyle Width="80px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="FormName" ReadOnly="True" HeaderText="�e�U��">
							<HeaderStyle Width="100px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="StepName" ReadOnly="True" HeaderText="�u�{�W��">
							<HeaderStyle Width="250px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="DecideName" ReadOnly="True" HeaderText="����">
							<HeaderStyle Width="80px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="DivName" ReadOnly="True" HeaderText="�����">
							<HeaderStyle Width="80px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="TimeRange" ReadOnly="True" HeaderText="�u�{�_���ɶ�">
							<HeaderStyle Width="250px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="AWorkTime" ReadOnly="True" HeaderText="�u�{�ɶ�">
							<HeaderStyle Width="60px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
					</Columns>
					<PagerStyle NextPageText="�U�@��" PrevPageText="�W�@��" HorizontalAlign="Center" ForeColor="#993300"
						BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
				</asp:datagrid>
				<DIV title="�z�ﶵ�ءG" style="DISPLAY: inline; Z-INDEX: 106; LEFT: 464px; WIDTH: 192px; COLOR: #0000ff; POSITION: absolute; TOP: 8px; HEIGHT: 24px"
					ms_positioning="FlowLayout"><FONT color="midnightblue">�]�u��ܢ�������ơ^</FONT></DIV>
				<asp:textbox id="DAEndTime" style="Z-INDEX: 105; LEFT: 288px; POSITION: absolute; TOP: 8px" runat="server"
					Width="160px" Height="20px" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue" ReadOnly="True">DAEndTime</asp:textbox>
				<DIV title="" style="DISPLAY: inline; Z-INDEX: 104; LEFT: 272px; WIDTH: 16px; COLOR: #0000ff; POSITION: absolute; TOP: 8px; HEIGHT: 24px"
					ms_positioning="FlowLayout">��</DIV>
				<DIV title="�z�ﶵ�ءG" style="DISPLAY: inline; Z-INDEX: 103; LEFT: 8px; WIDTH: 88px; COLOR: #0000ff; POSITION: absolute; TOP: 8px; HEIGHT: 24px"
					ms_positioning="FlowLayout">�֩w�ɶ��G</DIV>
				<asp:textbox id="DAStartTime" style="Z-INDEX: 102; LEFT: 112px; POSITION: absolute; TOP: 8px"
					runat="server" Width="160px" Height="20px" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
					ReadOnly="True">DAStartTime</asp:textbox></FONT></form>
	</body>
</HTML>
