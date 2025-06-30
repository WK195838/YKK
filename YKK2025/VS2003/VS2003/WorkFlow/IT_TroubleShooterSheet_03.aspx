<%@ Page Language="vb" AutoEventWireup="false" Codebehind="IT_TroubleShooterSheet_03.aspx.vb" Inherits="SPD.IT_TroubleShooterSheet_03"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>故障檢測申請書</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<script language="javascript" type="text/javascript">
		</script>
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" encType="multipart/form-data" runat="server">
			<FONT face="新細明體">
				<asp:dropdownlist id="DSystem" style="Z-INDEX: 120; LEFT: 456px; POSITION: absolute; TOP: 152px" runat="server"
					Height="20px" Width="224px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="要" Selected="True">要</asp:ListItem>
					<asp:ListItem Value="不要">不要</asp:ListItem>
				</asp:dropdownlist>
				<asp:dropdownlist id="DPriority" style="Z-INDEX: 122; LEFT: 120px; POSITION: absolute; TOP: 152px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Width="224px" Height="20px">
					<asp:ListItem Value="要" Selected="True">要</asp:ListItem>
					<asp:ListItem Value="不要">不要</asp:ListItem>
				</asp:dropdownlist>
				<asp:label id="DHistoryLabel" style="Z-INDEX: 121; LEFT: 8px; POSITION: absolute; TOP: 488px"
					runat="server" ForeColor="Blue" Width="64px" Height="16px" Font-Size="11pt" Font-Names="新細明體">核定履歷</asp:label>
				<asp:datagrid id="DataGrid9" style="Z-INDEX: 105; LEFT: 8px; POSITION: absolute; TOP: 504px" runat="server"
					BackColor="White" Width="780px" Height="100px" BorderStyle="None" Font-Size="9pt" AutoGenerateColumns="False"
					CellPadding="4" BorderWidth="1px" BorderColor="#CC9966">
					<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
					<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
					<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
					<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
					<Columns>
						<asp:BoundColumn DataField="StepNameDesc" ReadOnly="True" HeaderText="工程">
							<HeaderStyle Width="170px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="DecideName" ReadOnly="True" HeaderText="擔當">
							<HeaderStyle Width="50px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="AgentName" ReadOnly="True" HeaderText="代理">
							<HeaderStyle Width="50px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="FlowTypeDesc" ReadOnly="True" HeaderText="類別">
							<HeaderStyle Width="50px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="StsDesc" ReadOnly="True" HeaderText="處理結果">
							<HeaderStyle Width="50px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="DecideDescA" ReadOnly="True" HeaderText="說明">
							<HeaderStyle Width="220px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="Description" ReadOnly="True" HeaderText="核定時間">
							<HeaderStyle Width="150px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
					</Columns>
					<PagerStyle NextPageText="下一頁" PrevPageText="上一頁" HorizontalAlign="Center" ForeColor="#993300"
						BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
				</asp:datagrid>
				<asp:textbox id="DBDays" style="Z-INDEX: 119; LEFT: 608px; POSITION: absolute; TOP: 352px" runat="server"
					Height="20px" Width="70px" BorderStyle="Groove" BackColor="Yellow" ForeColor="Blue">DBDays</asp:textbox><asp:textbox id="DDepoName" style="Z-INDEX: 118; LEFT: 120px; POSITION: absolute; TOP: 120px"
					runat="server" Width="40px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">台北</asp:textbox><asp:textbox id="DDepoCode" style="Z-INDEX: 117; LEFT: 160px; POSITION: absolute; TOP: 120px"
					runat="server" Width="25px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">01</asp:textbox><asp:textbox id="DDivision" style="Z-INDEX: 106; LEFT: 192px; POSITION: absolute; TOP: 120px"
					runat="server" Width="88px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DDivision</asp:textbox><asp:textbox id="DDivisionCode" style="Z-INDEX: 107; LEFT: 280px; POSITION: absolute; TOP: 120px"
					runat="server" Width="63px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DDivisionCode</asp:textbox><asp:textbox id="DJobCode" style="Z-INDEX: 116; LEFT: 600px; POSITION: absolute; TOP: 120px"
					runat="server" Width="80px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DJobCode</asp:textbox><asp:textbox id="DEmpID" style="Z-INDEX: 115; LEFT: 600px; POSITION: absolute; TOP: 88px" runat="server"
					Width="80px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DEmpID</asp:textbox><asp:textbox id="DBEndDate" style="Z-INDEX: 114; LEFT: 336px; POSITION: absolute; TOP: 352px"
					runat="server" Width="176px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">DBEndDate</asp:textbox><asp:textbox id="DBStartDate" style="Z-INDEX: 113; LEFT: 120px; POSITION: absolute; TOP: 352px"
					runat="server" Width="184px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">DBStartDate</asp:textbox><asp:textbox id="DRemark" style="Z-INDEX: 112; LEFT: 120px; POSITION: absolute; TOP: 382px" runat="server"
					Width="560px" Height="90px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" TextMode="MultiLine" MaxLength="250">DRemark</asp:textbox><asp:hyperlink id="LRefFile" style="Z-INDEX: 111; LEFT: 120px; POSITION: absolute; TOP: 288px"
					runat="server" Width="72px" Height="8px" Font-Size="12pt" Target="_blank" NavigateUrl="BoardEdit.aspx">參考附件</asp:hyperlink><asp:textbox id="DTarget" style="Z-INDEX: 110; LEFT: 120px; POSITION: absolute; TOP: 184px" runat="server"
					Width="560px" Height="88px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" TextMode="MultiLine" MaxLength="250">DTarget</asp:textbox><asp:textbox id="DName" style="Z-INDEX: 109; LEFT: 456px; POSITION: absolute; TOP: 88px" runat="server"
					Width="144px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DName</asp:textbox><asp:dropdownlist id="DEngineer" style="Z-INDEX: 104; LEFT: 120px; POSITION: absolute; TOP: 318px"
					runat="server" Width="224px" Height="20px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="要" Selected="True">要</asp:ListItem>
					<asp:ListItem Value="不要">不要</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DJobTitle" style="Z-INDEX: 103; LEFT: 456px; POSITION: absolute; TOP: 120px"
					runat="server" Width="144px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DJobTitle</asp:textbox><asp:textbox id="DDate" style="Z-INDEX: 102; LEFT: 120px; POSITION: absolute; TOP: 88px" runat="server"
					Width="224px" Height="20px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DDate</asp:textbox><asp:image id="DTroubleShooterSheet1" style="Z-INDEX: 100; LEFT: 8px; POSITION: absolute; TOP: 8px"
					runat="server" Width="678px" Height="472px" ImageUrl="Images\IT_TroubleShooterSheet_001.jpg"></asp:image><asp:textbox id="DNo" style="Z-INDEX: 101; LEFT: 8px; POSITION: absolute; TOP: 9px" runat="server"
					Width="216px" Height="16px" ForeColor="Black" BackColor="White" BorderStyle="None" ReadOnly="True" Font-Names="Times New Roman"> 123</asp:textbox></FONT></form>
		</FONT></FORM>
	</body>
</HTML>
