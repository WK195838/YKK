<%@ Page Language="vb" AutoEventWireup="false" Codebehind="zzOperationReport.aspx.vb" Inherits="SPD.OperationReport" aspCompat="True"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>OperationReport</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<script language="javascript" type="text/javascript">
			//--Calendar------------------------------------
			function calendarPicker(strField)
			{
				window.open('DatePicker.aspx?field=' + strField,'calendarPopup','width=250,height=190,resizable=yes');
			}
		</script>
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體">
				<DIV title="篩選項目：" style="DISPLAY: inline; Z-INDEX: 102; LEFT: 8px; WIDTH: 88px; COLOR: #0000ff; POSITION: absolute; TOP: 8px; HEIGHT: 24px"
					ms_positioning="FlowLayout">篩選項目：</DIV>
				<asp:datagrid id="DataGrid1" style="Z-INDEX: 122; LEFT: 8px; POSITION: absolute; TOP: 472px" runat="server"
					Width="810px" Height="300px" Font-Size="9pt" BackColor="White" BorderStyle="None" AutoGenerateColumns="False"
					CellPadding="4" BorderWidth="1px" BorderColor="#CC9966" PageSize="15">
					<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
					<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
					<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
					<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
					<Columns>
						<asp:BoundColumn DataField="Buyer" ReadOnly="True" HeaderText="Buyer">
							<HeaderStyle Width="70px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="No" ReadOnly="True" HeaderText="No">
							<HeaderStyle Width="70px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="Sts" ReadOnly="True" HeaderText="狀態">
							<HeaderStyle Width="40px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="BTime" ReadOnly="True" HeaderText="預定">
							<HeaderStyle Width="230px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="BWorkTime" ReadOnly="True" HeaderText="工時(A)">
							<HeaderStyle Width="50px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Right"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="ATime" ReadOnly="True" HeaderText="實績">
							<HeaderStyle Width="230px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="AWorkTime" ReadOnly="True" HeaderText="工時(B)">
							<HeaderStyle Width="50px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Right"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="Balance" ReadOnly="True" HeaderText="A-B">
							<HeaderStyle Width="50px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Right"></ItemStyle>
						</asp:BoundColumn>
					</Columns>
					<PagerStyle NextPageText="下一頁" PrevPageText="上一頁" HorizontalAlign="Center" ForeColor="#993300"
						BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
				</asp:datagrid>
				<asp:label id="ChartUnit4" style="Z-INDEX: 121; LEFT: 672px; POSITION: absolute; TOP: 288px"
					runat="server" Font-Size="12pt" ForeColor="Navy" Height="16px" Width="56px">單位=時</asp:label><asp:label id="ChartUnit3" style="Z-INDEX: 120; LEFT: 304px; POSITION: absolute; TOP: 288px"
					runat="server" Font-Size="12pt" ForeColor="Navy" Height="16px" Width="57px">單位=時</asp:label><asp:label id="ChartUnit2" style="Z-INDEX: 119; LEFT: 672px; POSITION: absolute; TOP: 88px"
					runat="server" Font-Size="12pt" ForeColor="Navy" Height="16px" Width="57px">單位=%</asp:label><asp:label id="ChartTitle4" style="Z-INDEX: 118; LEFT: 376px; POSITION: absolute; TOP: 272px"
					runat="server" Font-Size="14pt" ForeColor="Navy" Height="16px" Width="348px">　　　　　　　平均處理時間</asp:label><asp:label id="ChartTitle3" style="Z-INDEX: 117; LEFT: 8px; POSITION: absolute; TOP: 272px"
					runat="server" Font-Size="14pt" ForeColor="Navy" Height="16px" Width="348px">　　　　　　　　　時數</asp:label><asp:label id="ChartTitle2" style="Z-INDEX: 116; LEFT: 376px; POSITION: absolute; TOP: 72px"
					runat="server" Font-Size="14pt" ForeColor="Navy" Height="16px" Width="348px">　　　　　　　　　比率</asp:label><asp:label id="ChartUnit1" style="Z-INDEX: 115; LEFT: 312px; POSITION: absolute; TOP: 88px"
					runat="server" Font-Size="12pt" ForeColor="Navy" Height="16px" Width="48px">單位=1</asp:label><asp:label id="ChartTitle1" style="Z-INDEX: 114; LEFT: 8px; POSITION: absolute; TOP: 72px"
					runat="server" Font-Size="14pt" ForeColor="Navy" Height="16px" Width="348px">　　　　　　　　　件數</asp:label><asp:image id="BarImage4" style="Z-INDEX: 113; LEFT: 376px; POSITION: absolute; TOP: 304px"
					runat="server" Height="161px" Width="348px" ImageUrl="Chart\Operation-Time-Ratio.bmp"></asp:image><asp:image id="BarImage3" style="Z-INDEX: 112; LEFT: 8px; POSITION: absolute; TOP: 304px" runat="server"
					Height="161px" Width="348px" ImageUrl="Chart\Operation-Time.bmp"></asp:image><asp:image id="BarImage2" style="Z-INDEX: 111; LEFT: 376px; POSITION: absolute; TOP: 104px"
					runat="server" Height="161px" Width="348px" ImageUrl="Chart\Operation-Record-Ratio.bmp"></asp:image><asp:dropdownlist id="DStep" style="Z-INDEX: 103; LEFT: 224px; POSITION: absolute; TOP: 8px" runat="server"
					ForeColor="Blue" Height="40px" Width="344px" BackColor="Yellow"></asp:dropdownlist><asp:dropdownlist id="DFormName" style="Z-INDEX: 101; LEFT: 96px; POSITION: absolute; TOP: 8px" runat="server"
					ForeColor="Blue" Height="40px" Width="120px" BackColor="Yellow" AutoPostBack="True"></asp:dropdownlist></FONT><asp:textbox id="DSDate" style="Z-INDEX: 104; LEFT: 96px; POSITION: absolute; TOP: 40px" runat="server"
				ForeColor="Blue" Height="20px" Width="96px" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DSDate</asp:textbox><asp:button id="BSDate" style="Z-INDEX: 105; LEFT: 192px; POSITION: absolute; TOP: 40px" runat="server"
				Height="20px" Width="20px" Text="....."></asp:button>
			<DIV title="" style="DISPLAY: inline; Z-INDEX: 106; LEFT: 224px; WIDTH: 16px; COLOR: #0000ff; POSITION: absolute; TOP: 40px; HEIGHT: 24px"
				ms_positioning="FlowLayout">～</DIV>
			<asp:textbox id="DEDate" style="Z-INDEX: 107; LEFT: 248px; POSITION: absolute; TOP: 40px" runat="server"
				ForeColor="Blue" Height="20px" Width="96px" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DEDate</asp:textbox><asp:button id="BEDate" style="Z-INDEX: 108; LEFT: 344px; POSITION: absolute; TOP: 40px" runat="server"
				Height="20px" Width="20px" Text="....."></asp:button><asp:button id="Go" style="Z-INDEX: 109; LEFT: 416px; POSITION: absolute; TOP: 40px" runat="server"
				ForeColor="Blue" Height="24px" Width="40px" BackColor="White" Text="Go"></asp:button><asp:image id="BarImage1" style="Z-INDEX: 110; LEFT: 8px; POSITION: absolute; TOP: 104px" runat="server"
				Height="161px" Width="348px" ImageUrl="Chart\Operation-Record.bmp"></asp:image></form>
	</body>
</HTML>
