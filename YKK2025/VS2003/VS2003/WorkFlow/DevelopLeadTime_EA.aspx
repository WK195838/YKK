<%@ Page Language="vb" AutoEventWireup="false" Codebehind="DevelopLeadTime_EA.aspx.vb" Inherits="SPD.DevelopLeadTime_EA" aspCompat="True"%>
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
				<DIV title="" style="Z-INDEX: 119; POSITION: absolute; WIDTH: 81px; DISPLAY: inline; HEIGHT: 24px; COLOR: #0000ff; TOP: 112px; LEFT: 624px"
					ms_positioning="FlowLayout">單位：分</DIV>
				<asp:textbox id="TextBox1" style="Z-INDEX: 118; POSITION: absolute; TOP: 112px; LEFT: 8px" runat="server"
					Height="24px" Width="496px" BorderStyle="None" ReadOnly="True" TextMode="MultiLine">(a)時間：實際開始(系統計算)～實際完成　(b)時間：收件開始～實際完成</asp:textbox><asp:imagebutton id="BExcel" style="Z-INDEX: 117; POSITION: absolute; TOP: 82px; LEFT: 440px" runat="server"
					Height="21px" Width="21px" ImageUrl="Images\msexcel.gif"></asp:imagebutton><asp:dropdownlist id="DSts" style="Z-INDEX: 116; POSITION: absolute; TOP: 56px; LEFT: 96px" runat="server"
					Height="40px" Width="120px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="ALL" Selected="True">ALL</asp:ListItem>
					<asp:ListItem Value="0">開發中</asp:ListItem>
					<asp:ListItem Value="1">OK完成</asp:ListItem>
					<asp:ListItem Value="2">NG完成</asp:ListItem>
					<asp:ListItem Value="3">取消</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DDivision" style="Z-INDEX: 115; POSITION: absolute; TOP: 56px; LEFT: 224px"
					runat="server" Height="40px" Width="152px" BackColor="Yellow" ForeColor="Blue"></asp:dropdownlist><asp:textbox id="DPerson" style="Z-INDEX: 112; POSITION: absolute; TOP: 56px; LEFT: 384px" runat="server"
					Height="20px" Width="120px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True">DPerson</asp:textbox><asp:textbox id="DBuyer" style="Z-INDEX: 102; POSITION: absolute; TOP: 56px; LEFT: 512px" runat="server"
					Height="20px" Width="120px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True">DBuyer</asp:textbox><asp:datagrid id="DataGrid1" style="Z-INDEX: 114; POSITION: absolute; TOP: 144px; LEFT: 8px" runat="server"
					Height="300px" Width="700px" BackColor="White" BorderStyle="None" PageSize="15" BorderColor="#CC9966" BorderWidth="1px" CellPadding="4" AutoGenerateColumns="False" Font-Size="9pt">
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
							<HeaderStyle Width="80px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="ATime" ReadOnly="True" HeaderText="(a)時間">
							<HeaderStyle Width="60px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Right"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="BTime" ReadOnly="True" HeaderText="(b)時間">
							<HeaderStyle Width="60px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Right"></ItemStyle>
						</asp:BoundColumn>
					</Columns>
					<PagerStyle NextPageText="下一頁" PrevPageText="上一頁" HorizontalAlign="Center" ForeColor="#993300"
						BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
				</asp:datagrid><asp:dropdownlist id="DEStep" style="Z-INDEX: 113; POSITION: absolute; TOP: 32px; LEFT: 360px" runat="server"
					Height="40px" Width="224px" BackColor="Yellow" ForeColor="Blue"></asp:dropdownlist><asp:dropdownlist id="DSStep" style="Z-INDEX: 103; POSITION: absolute; TOP: 32px; LEFT: 96px" runat="server"
					Height="40px" Width="224px" BackColor="Yellow" ForeColor="Blue"></asp:dropdownlist><asp:dropdownlist id="DFormName" style="Z-INDEX: 100; POSITION: absolute; TOP: 8px; LEFT: 96px" runat="server"
					Height="40px" Width="264px" BackColor="Yellow" ForeColor="Blue" AutoPostBack="True"></asp:dropdownlist></FONT><asp:textbox id="DSDate" style="Z-INDEX: 104; POSITION: absolute; TOP: 80px; LEFT: 96px" runat="server"
				Height="20px" Width="96px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True">DSDate</asp:textbox><asp:button id="BSDate" style="Z-INDEX: 105; POSITION: absolute; TOP: 80px; LEFT: 192px" runat="server"
				Height="20px" Width="20px" Text="....."></asp:button>
			<DIV title="" style="Z-INDEX: 106; POSITION: absolute; WIDTH: 16px; DISPLAY: inline; HEIGHT: 24px; COLOR: #0000ff; TOP: 80px; LEFT: 224px"
				ms_positioning="FlowLayout">～</DIV>
			<asp:textbox id="DEDate" style="Z-INDEX: 107; POSITION: absolute; TOP: 80px; LEFT: 248px" runat="server"
				Height="20px" Width="96px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True">DEDate</asp:textbox><asp:button id="BEDate" style="Z-INDEX: 108; POSITION: absolute; TOP: 80px; LEFT: 344px" runat="server"
				Height="20px" Width="20px" Text="....."></asp:button><asp:button id="Go" style="Z-INDEX: 109; POSITION: absolute; TOP: 80px; LEFT: 376px" runat="server"
				Height="24px" Width="40px" BackColor="White" ForeColor="Blue" Text="Go"></asp:button>
			<DIV title="" style="Z-INDEX: 111; POSITION: absolute; WIDTH: 16px; DISPLAY: inline; HEIGHT: 24px; COLOR: #0000ff; TOP: 32px; LEFT: 328px"
				ms_positioning="FlowLayout">～</DIV>
		</form>
	</body>
</HTML>
