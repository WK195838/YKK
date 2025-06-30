<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MultiOPPerformance.aspx.vb" Inherits="SPD.MultiOPPerformance"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>MultiOPPerformance</title>
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
				<DIV title="篩選項目：" style="DISPLAY: inline; Z-INDEX: 101; LEFT: 8px; WIDTH: 88px; COLOR: #0000ff; POSITION: absolute; TOP: 8px; HEIGHT: 24px"
					ms_positioning="FlowLayout">篩選項目：</DIV>
				<asp:dropdownlist id="DSts" style="Z-INDEX: 115; LEFT: 96px; POSITION: absolute; TOP: 56px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="40px" Width="120px">
					<asp:ListItem Value="ALL" Selected="True">ALL</asp:ListItem>
					<asp:ListItem Value="0">開發中</asp:ListItem>
					<asp:ListItem Value="1">OK完成</asp:ListItem>
					<asp:ListItem Value="2">NG完成</asp:ListItem>
					<asp:ListItem Value="3">取消</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DDivision" style="Z-INDEX: 114; LEFT: 224px; POSITION: absolute; TOP: 56px"
					runat="server" ForeColor="Blue" BackColor="Yellow" Height="40px" Width="152px"></asp:dropdownlist><asp:textbox id="DPerson" style="Z-INDEX: 111; LEFT: 384px; POSITION: absolute; TOP: 56px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="20px" Width="120px" ReadOnly="True" BorderStyle="Groove">DPerson</asp:textbox><asp:textbox id="DBuyer" style="Z-INDEX: 102; LEFT: 512px; POSITION: absolute; TOP: 56px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="20px" Width="120px" ReadOnly="True" BorderStyle="Groove">DBuyer</asp:textbox><asp:datagrid id="DataGrid1" style="Z-INDEX: 113; LEFT: 8px; POSITION: absolute; TOP: 112px" runat="server"
					BackColor="White" Height="300px" Width="790px" BorderStyle="None" Font-Size="9pt" AutoGenerateColumns="False" CellPadding="4" BorderWidth="1px" BorderColor="#CC9966" PageSize="15">
					<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
					<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
					<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
					<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
					<Columns>
						<asp:BoundColumn DataField="Buyer" ReadOnly="True" HeaderText="Buyer">
							<HeaderStyle Width="80px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="URL" DataNavigateUrlFormatString="{0}" DataTextField="No"
							HeaderText="No">
							<HeaderStyle Width="60px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:HyperLinkColumn>
						<asp:BoundColumn DataField="Sts" ReadOnly="True" HeaderText="狀態">
							<HeaderStyle Width="40px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="Division" ReadOnly="True" HeaderText="依賴部門">
							<HeaderStyle Width="80px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="Finished" ReadOnly="True" HeaderText="類型">
							<HeaderStyle Width="60px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="Suppiler" ReadOnly="True" HeaderText="外注">
							<HeaderStyle Width="60px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="Level" ReadOnly="True" HeaderText="難易度">
							<HeaderStyle Width="40px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="Step" ReadOnly="True" HeaderText="工程No.">
							<HeaderStyle Width="60px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="BTime" ReadOnly="True" HeaderText="預定(A)">
							<HeaderStyle Width="50px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Right"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="ATime" ReadOnly="True" HeaderText="實績(B)">
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
				</asp:datagrid><asp:dropdownlist id="DEStep" style="Z-INDEX: 112; LEFT: 360px; POSITION: absolute; TOP: 32px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="40px" Width="224px"></asp:dropdownlist><asp:dropdownlist id="DSStep" style="Z-INDEX: 103; LEFT: 96px; POSITION: absolute; TOP: 32px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="40px" Width="224px"></asp:dropdownlist><asp:dropdownlist id="DFormName" style="Z-INDEX: 100; LEFT: 96px; POSITION: absolute; TOP: 8px" runat="server"
					ForeColor="Blue" BackColor="Yellow" Height="40px" Width="264px" AutoPostBack="True"></asp:dropdownlist></FONT><asp:textbox id="DSDate" style="Z-INDEX: 104; LEFT: 96px; POSITION: absolute; TOP: 80px" runat="server"
				ForeColor="Blue" BackColor="Yellow" Height="20px" Width="96px" ReadOnly="True" BorderStyle="Groove">DSDate</asp:textbox><asp:button id="BSDate" style="Z-INDEX: 105; LEFT: 192px; POSITION: absolute; TOP: 80px" runat="server"
				Height="20px" Width="20px" Text="....."></asp:button>
			<DIV title="" style="DISPLAY: inline; Z-INDEX: 106; LEFT: 224px; WIDTH: 16px; COLOR: #0000ff; POSITION: absolute; TOP: 80px; HEIGHT: 24px"
				ms_positioning="FlowLayout">～</DIV>
			<asp:textbox id="DEDate" style="Z-INDEX: 107; LEFT: 248px; POSITION: absolute; TOP: 80px" runat="server"
				ForeColor="Blue" BackColor="Yellow" Height="20px" Width="96px" ReadOnly="True" BorderStyle="Groove">DEDate</asp:textbox><asp:button id="BEDate" style="Z-INDEX: 108; LEFT: 344px; POSITION: absolute; TOP: 80px" runat="server"
				Height="20px" Width="20px" Text="....."></asp:button><asp:button id="Go" style="Z-INDEX: 109; LEFT: 376px; POSITION: absolute; TOP: 80px" runat="server"
				ForeColor="Blue" BackColor="White" Height="24px" Width="40px" Text="Go"></asp:button>
			<DIV title="" style="DISPLAY: inline; Z-INDEX: 110; LEFT: 328px; WIDTH: 16px; COLOR: #0000ff; POSITION: absolute; TOP: 32px; HEIGHT: 24px"
				ms_positioning="FlowLayout">～</DIV>
		</form>
	</body>
</HTML>
