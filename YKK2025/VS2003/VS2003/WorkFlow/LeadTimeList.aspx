<%@ Page Language="vb" AutoEventWireup="false" Codebehind="LeadTimeList.aspx.vb" Inherits="SPD.LeadTimeList"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>標準Lead-Time</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體"></FONT>
			<DIV title="篩選項目：" style="DISPLAY: inline; Z-INDEX: 117; LEFT: 8px; WIDTH: 88px; COLOR: #0000ff; POSITION: absolute; TOP: 8px; HEIGHT: 24px"
				ms_positioning="FlowLayout">篩選項目：</DIV>
			<asp:hyperlink id="LQCLeadTime" style="Z-INDEX: 151; LEFT: 416px; POSITION: absolute; TOP: 40px"
				runat="server" Width="120px" Height="8px" Font-Size="12pt" NavigateUrl="Images\QC_LeadTime.xls"
				Target="_blank">品質測試Leadtime</asp:hyperlink>
			<asp:imagebutton id="BExcel" style="Z-INDEX: 150; LEFT: 448px; POSITION: absolute; TOP: 8px" runat="server"
				Height="21px" Width="21px" ImageUrl="Images\msexcel.gif"></asp:imagebutton>
			<asp:button id="Go" style="Z-INDEX: 105; LEFT: 400px; POSITION: absolute; TOP: 8px" runat="server"
				Text="Go" BackColor="White" ForeColor="Blue" Width="40px" Height="24px"></asp:button><asp:dropdownlist id="DFormName" style="Z-INDEX: 118; LEFT: 96px; POSITION: absolute; TOP: 8px" runat="server"
				BackColor="Yellow" ForeColor="Blue" Width="165px" Height="40px" AutoPostBack="True"></asp:dropdownlist><asp:dropdownlist id="DLevel" style="Z-INDEX: 120; LEFT: 264px; POSITION: absolute; TOP: 8px" runat="server"
				BackColor="Yellow" ForeColor="Blue" Width="130px" Height="40px">
				<asp:ListItem Value="ALL" Selected="True">ALL</asp:ListItem>
				<asp:ListItem Value="0">通知</asp:ListItem>
				<asp:ListItem Value="1">核定</asp:ListItem>
			</asp:dropdownlist><asp:datagrid id="DataGrid1" style="Z-INDEX: 126; LEFT: 8px; POSITION: absolute; TOP: 64px" runat="server"
				BackColor="White" Width="530px" Height="300px" Font-Size="9pt" BorderColor="#CC9966" BorderWidth="1px" CellPadding="4"
				AutoGenerateColumns="False" BorderStyle="None">
				<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
				<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
				<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
				<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
				<Columns>
					<asp:BoundColumn DataField="FormName" ReadOnly="True" HeaderText="委託單">
						<HeaderStyle Width="150px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="Step" ReadOnly="True" HeaderText="No">
						<HeaderStyle Width="60px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Center"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="StepName" ReadOnly="True" HeaderText="工程名">
						<HeaderStyle Width="200px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="Level" ReadOnly="True" HeaderText="難易度">
						<HeaderStyle Width="60px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Center"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="Hours" ReadOnly="True" HeaderText="分數">
						<HeaderStyle Width="60px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Center"></ItemStyle>
					</asp:BoundColumn>
				</Columns>
				<PagerStyle NextPageText="下一頁" PrevPageText="上一頁" HorizontalAlign="Center" ForeColor="#993300"
					BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
			</asp:datagrid></form>
	</body>
</HTML>
