<%@ Page Language="vb" AutoEventWireup="false" Codebehind="DevelopLeadTime.aspx.vb" Inherits="SPD.DevelopLeadTime" aspCompat="True"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>開發時間</title>
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
				<DIV title="篩選項目：" style="Z-INDEX: 101; POSITION: absolute; WIDTH: 88px; DISPLAY: inline; HEIGHT: 24px; COLOR: #0000ff; TOP: 8px; LEFT: 8px"
					ms_positioning="FlowLayout">篩選項目：</DIV>
				<asp:imagebutton id="BExcel" style="Z-INDEX: 117; POSITION: absolute; TOP: 82px; LEFT: 440px" runat="server"
					ImageUrl="Images\msexcel.gif" Width="21px" Height="21px"></asp:imagebutton><asp:dropdownlist id="DSts" style="Z-INDEX: 116; POSITION: absolute; TOP: 56px; LEFT: 96px" runat="server"
					Width="120px" Height="40px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="1" Selected="True">OK完成</asp:ListItem>
					<asp:ListItem Value="2">NG完成</asp:ListItem>
					<asp:ListItem Value="3">取消</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DDivision" style="Z-INDEX: 115; POSITION: absolute; TOP: 56px; LEFT: 224px"
					runat="server" Width="152px" Height="40px" ForeColor="Blue" BackColor="Yellow"></asp:dropdownlist><asp:textbox id="DPerson" style="Z-INDEX: 112; POSITION: absolute; TOP: 56px; LEFT: 384px" runat="server"
					Width="120px" Height="20px" ForeColor="Blue" BackColor="Yellow" ReadOnly="True" BorderStyle="Groove">DPerson</asp:textbox><asp:textbox id="DBuyer" style="Z-INDEX: 102; POSITION: absolute; TOP: 56px; LEFT: 512px" runat="server"
					Width="120px" Height="20px" ForeColor="Blue" BackColor="Yellow" ReadOnly="True" BorderStyle="Groove">DBuyer</asp:textbox><asp:datagrid id="DataGrid1" style="Z-INDEX: 114; POSITION: absolute; TOP: 112px; LEFT: 8px" runat="server"
					Width="550px" Height="300px" BackColor="White" BorderStyle="None" Font-Size="9pt" AutoGenerateColumns="False" CellPadding="4" BorderWidth="1px" BorderColor="#CC9966" PageSize="15">
					<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
					<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
					<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
					<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
					<Columns>
						<asp:BoundColumn DataField="Buyer" ReadOnly="True" HeaderText="Buyer">
							<HeaderStyle Width="120px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="URL" DataNavigateUrlFormatString="{0}" DataTextField="No"
							HeaderText="No">
							<HeaderStyle Width="80px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:HyperLinkColumn>
						<asp:BoundColumn DataField="Sts" ReadOnly="True" HeaderText="狀態">
							<HeaderStyle Width="60px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="Division" ReadOnly="True" HeaderText="依賴部門">
							<HeaderStyle Width="80px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="Step" ReadOnly="True" HeaderText="工程No.">
							<HeaderStyle Width="50px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="BTime" ReadOnly="True" HeaderText="時間">
							<HeaderStyle Width="80px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Right"></ItemStyle>
						</asp:BoundColumn>
					</Columns>
					<PagerStyle NextPageText="下一頁" PrevPageText="上一頁" HorizontalAlign="Center" ForeColor="#993300"
						BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
				</asp:datagrid><asp:dropdownlist id="DEStep" style="Z-INDEX: 113; POSITION: absolute; TOP: 32px; LEFT: 360px" runat="server"
					Width="224px" Height="40px" ForeColor="Blue" BackColor="Yellow"></asp:dropdownlist><asp:dropdownlist id="DSStep" style="Z-INDEX: 103; POSITION: absolute; TOP: 32px; LEFT: 96px" runat="server"
					Width="224px" Height="40px" ForeColor="Blue" BackColor="Yellow"></asp:dropdownlist><asp:dropdownlist id="DFormName" style="Z-INDEX: 100; POSITION: absolute; TOP: 8px; LEFT: 96px" runat="server"
					Width="264px" Height="40px" ForeColor="Blue" BackColor="Yellow" AutoPostBack="True"></asp:dropdownlist></FONT><asp:textbox id="DSDate" style="Z-INDEX: 104; POSITION: absolute; TOP: 80px; LEFT: 96px" runat="server"
				Width="96px" Height="20px" ForeColor="Blue" BackColor="Yellow" ReadOnly="True" BorderStyle="Groove">DSDate</asp:textbox><asp:button id="BSDate" style="Z-INDEX: 105; POSITION: absolute; TOP: 80px; LEFT: 192px" runat="server"
				Width="20px" Height="20px" Text="....."></asp:button>
			<DIV title="" style="Z-INDEX: 106; POSITION: absolute; WIDTH: 16px; DISPLAY: inline; HEIGHT: 24px; COLOR: #0000ff; TOP: 80px; LEFT: 224px"
				ms_positioning="FlowLayout">～</DIV>
			<asp:textbox id="DEDate" style="Z-INDEX: 107; POSITION: absolute; TOP: 80px; LEFT: 248px" runat="server"
				Width="96px" Height="20px" ForeColor="Blue" BackColor="Yellow" ReadOnly="True" BorderStyle="Groove">DEDate</asp:textbox><asp:button id="BEDate" style="Z-INDEX: 108; POSITION: absolute; TOP: 80px; LEFT: 344px" runat="server"
				Width="20px" Height="20px" Text="....."></asp:button><asp:button id="Go" style="Z-INDEX: 109; POSITION: absolute; TOP: 80px; LEFT: 376px" runat="server"
				Width="40px" Height="24px" ForeColor="Blue" BackColor="White" Text="Go"></asp:button>
			<DIV title="" style="Z-INDEX: 110; POSITION: absolute; WIDTH: 16px; DISPLAY: inline; HEIGHT: 24px; COLOR: #0000ff; TOP: 32px; LEFT: 328px"
				ms_positioning="FlowLayout">～</DIV>
		</form>
	</body>
</HTML>
