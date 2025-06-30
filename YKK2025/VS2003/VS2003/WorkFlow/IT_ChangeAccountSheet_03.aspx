<%@ Page Language="vb" AutoEventWireup="false" Codebehind="IT_ChangeAccountSheet_03.aspx.vb" Inherits="SPD.IT_ChangeAccountSheet_03"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>帳號變更申請書</title>
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
				<asp:button id="BFinishDate" style="Z-INDEX: 128; LEFT: 320px; POSITION: absolute; TOP: 152px"
					runat="server" Width="20px" Height="20px" CausesValidation="False" Text="....."></asp:button>
				<asp:label id="DHistoryLabel" style="Z-INDEX: 129; LEFT: 8px; POSITION: absolute; TOP: 488px"
					runat="server" Height="16px" Width="64px" ForeColor="Blue" Font-Names="新細明體" Font-Size="11pt">核定履歷</asp:label>
				<asp:datagrid id="DataGrid9" style="Z-INDEX: 105; LEFT: 8px; POSITION: absolute; TOP: 504px" runat="server"
					Height="100px" Width="780px" BorderStyle="None" BackColor="White" Font-Size="9pt" AutoGenerateColumns="False"
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
				</asp:datagrid><asp:textbox id="DFinishDate" style="Z-INDEX: 127; LEFT: 120px; POSITION: absolute; TOP: 152px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Width="198px" Height="20px" BorderStyle="Groove" ReadOnly="True">DFinishDate</asp:textbox><asp:dropdownlist id="DSystem" style="Z-INDEX: 117; LEFT: 456px; POSITION: absolute; TOP: 152px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Width="224px" Height="20px">
					<asp:ListItem Value="要" Selected="True">要</asp:ListItem>
					<asp:ListItem Value="不要">不要</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DFun4" style="Z-INDEX: 126; LEFT: 16px; POSITION: absolute; TOP: 283px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Width="96px" Height="20px">
					<asp:ListItem Value="新增" Selected="True">新增</asp:ListItem>
					<asp:ListItem Value="修改">修改</asp:ListItem>
					<asp:ListItem Value="刪除">刪除</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DFun3" style="Z-INDEX: 125; LEFT: 16px; POSITION: absolute; TOP: 251px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Width="96px" Height="20px">
					<asp:ListItem Value="新增" Selected="True">新增</asp:ListItem>
					<asp:ListItem Value="修改">修改</asp:ListItem>
					<asp:ListItem Value="刪除">刪除</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DFun2" style="Z-INDEX: 124; LEFT: 16px; POSITION: absolute; TOP: 219px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Width="96px" Height="20px">
					<asp:ListItem Value="新增" Selected="True">新增</asp:ListItem>
					<asp:ListItem Value="修改">修改</asp:ListItem>
					<asp:ListItem Value="刪除">刪除</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DAccount4" style="Z-INDEX: 123; LEFT: 120px; POSITION: absolute; TOP: 283px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Width="560px" Height="20px" BorderStyle="Groove">DAccount4</asp:textbox><asp:textbox id="DAccount3" style="Z-INDEX: 122; LEFT: 120px; POSITION: absolute; TOP: 251px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Width="560px" Height="20px" BorderStyle="Groove">DAccount3</asp:textbox><asp:textbox id="DAccount2" style="Z-INDEX: 121; LEFT: 120px; POSITION: absolute; TOP: 219px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Width="560px" Height="20px" BorderStyle="Groove">DAccount2</asp:textbox><asp:textbox id="DAccount1" style="Z-INDEX: 120; LEFT: 120px; POSITION: absolute; TOP: 186px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Width="560px" Height="20px" BorderStyle="Groove">DAccount1</asp:textbox><asp:dropdownlist id="DFun1" style="Z-INDEX: 119; LEFT: 16px; POSITION: absolute; TOP: 186px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Width="96px" Height="20px">
					<asp:ListItem Value="新增" Selected="True">新增</asp:ListItem>
					<asp:ListItem Value="修改">修改</asp:ListItem>
					<asp:ListItem Value="刪除">刪除</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DBDays" style="Z-INDEX: 116; LEFT: 608px; POSITION: absolute; TOP: 352px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Width="70px" Height="20px" BorderStyle="Groove">DBDays</asp:textbox><asp:textbox id="DDepoName" style="Z-INDEX: 115; LEFT: 120px; POSITION: absolute; TOP: 120px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Width="40px" Height="20px" BorderStyle="Groove" ReadOnly="True">台北</asp:textbox><asp:textbox id="DDepoCode" style="Z-INDEX: 114; LEFT: 160px; POSITION: absolute; TOP: 120px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Width="25px" Height="20px" BorderStyle="Groove" ReadOnly="True">01</asp:textbox><asp:textbox id="DDivision" style="Z-INDEX: 106; LEFT: 192px; POSITION: absolute; TOP: 120px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Width="88px" Height="20px" BorderStyle="Groove" ReadOnly="True">DDivision</asp:textbox><asp:textbox id="DDivisionCode" style="Z-INDEX: 107; LEFT: 280px; POSITION: absolute; TOP: 120px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Width="63px" Height="20px" BorderStyle="Groove" ReadOnly="True">DDivisionCode</asp:textbox><asp:textbox id="DJobCode" style="Z-INDEX: 113; LEFT: 600px; POSITION: absolute; TOP: 120px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Width="80px" Height="20px" BorderStyle="Groove" ReadOnly="True">DJobCode</asp:textbox><asp:textbox id="DEmpID" style="Z-INDEX: 112; LEFT: 600px; POSITION: absolute; TOP: 88px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Width="80px" Height="20px" BorderStyle="Groove" ReadOnly="True">DEmpID</asp:textbox><asp:textbox id="DBEndDate" style="Z-INDEX: 111; LEFT: 336px; POSITION: absolute; TOP: 352px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Width="176px" Height="20px" BorderStyle="Groove">DBEndDate</asp:textbox><asp:textbox id="DBStartDate" style="Z-INDEX: 110; LEFT: 120px; POSITION: absolute; TOP: 352px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Width="184px" Height="20px" BorderStyle="Groove">DBStartDate</asp:textbox><asp:textbox id="DRemark" style="Z-INDEX: 109; LEFT: 120px; POSITION: absolute; TOP: 382px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Width="560px" Height="90px" BorderStyle="Groove" MaxLength="250" TextMode="MultiLine">DRemark</asp:textbox><asp:textbox id="DName" style="Z-INDEX: 108; LEFT: 456px; POSITION: absolute; TOP: 88px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Width="144px" Height="20px" BorderStyle="Groove" ReadOnly="True">DName</asp:textbox><asp:dropdownlist id="DEngineer" style="Z-INDEX: 104; LEFT: 120px; POSITION: absolute; TOP: 318px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Width="224px" Height="20px">
					<asp:ListItem Value="要" Selected="True">要</asp:ListItem>
					<asp:ListItem Value="不要">不要</asp:ListItem>
				</asp:dropdownlist><asp:textbox id="DJobTitle" style="Z-INDEX: 103; LEFT: 456px; POSITION: absolute; TOP: 120px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Width="144px" Height="20px" BorderStyle="Groove" ReadOnly="True">DJobTitle</asp:textbox><asp:textbox id="DDate" style="Z-INDEX: 102; LEFT: 120px; POSITION: absolute; TOP: 88px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Width="224px" Height="20px" BorderStyle="Groove" ReadOnly="True">DDate</asp:textbox><asp:image id="DChangeAccountSheet1" style="Z-INDEX: 100; LEFT: 8px; POSITION: absolute; TOP: 8px"
					runat="server" Width="678px" Height="472px" ImageUrl="Images\IT_ChangeAccountSheet_001.jpg"></asp:image><asp:textbox id="DNo" style="Z-INDEX: 101; LEFT: 8px; POSITION: absolute; TOP: 9px" runat="server"
					ForeColor="Black" BackColor="White" Width="216px" Height="16px" BorderStyle="None" ReadOnly="True" Font-Names="Times New Roman"> 123</asp:textbox></FONT></form>
		</FONT></FORM>
	</body>
</HTML>
