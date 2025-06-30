<%@ Page Language="vb" AutoEventWireup="false" Codebehind="DevelopLeadTime@.aspx.vb" Inherits="SPD.DevelopLeadTime_" aspCompat="True"%>
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
				<DIV title="篩選項目：" style="DISPLAY: inline; Z-INDEX: 101; LEFT: 8px; WIDTH: 88px; COLOR: #0000ff; POSITION: absolute; TOP: 8px; HEIGHT: 24px"
					ms_positioning="FlowLayout">篩選項目：</DIV>
				<asp:imagebutton id="BExcel" style="Z-INDEX: 117; LEFT: 440px; POSITION: absolute; TOP: 82px" runat="server"
					Height="21px" Width="21px" ImageUrl="Images\msexcel.gif"></asp:imagebutton>
				<asp:dropdownlist id="DSts" style="Z-INDEX: 116; LEFT: 96px; POSITION: absolute; TOP: 56px" runat="server"
					Width="120px" Height="40px" BackColor="Yellow" ForeColor="Blue">
					<asp:ListItem Value="1" Selected="True">OK完成</asp:ListItem>
					<asp:ListItem Value="2">NG完成</asp:ListItem>
					<asp:ListItem Value="3">取消</asp:ListItem>
				</asp:dropdownlist><asp:dropdownlist id="DDivision" style="Z-INDEX: 115; LEFT: 224px; POSITION: absolute; TOP: 56px"
					runat="server" Width="152px" Height="40px" BackColor="Yellow" ForeColor="Blue"></asp:dropdownlist><asp:textbox id="DPerson" style="Z-INDEX: 112; LEFT: 384px; POSITION: absolute; TOP: 56px" runat="server"
					Width="120px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True">DPerson</asp:textbox><asp:textbox id="DBuyer" style="Z-INDEX: 102; LEFT: 512px; POSITION: absolute; TOP: 56px" runat="server"
					Width="120px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True">DBuyer</asp:textbox><asp:datagrid id="DataGrid1" style="Z-INDEX: 114; LEFT: 8px; POSITION: absolute; TOP: 112px" runat="server"
					Width="1050px" Height="300px" BackColor="White" BorderStyle="None" PageSize="15" BorderColor="#CC9966" BorderWidth="1px" CellPadding="4" AutoGenerateColumns="False" Font-Size="9pt">
					<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
					<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
					<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
					<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
					<Columns>
						<asp:BoundColumn DataField="Buyer" ReadOnly="True" HeaderText="Buyer">
							<HeaderStyle Width="70px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="URL" DataNavigateUrlFormatString="{0}" DataTextField="No"
							HeaderText="No">
							<HeaderStyle Width="70px"></HeaderStyle>
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
						<asp:BoundColumn DataField="SliderCode" ReadOnly="True" HeaderText="SliderCode">
							<HeaderStyle Width="80px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="Spec" ReadOnly="True" HeaderText="Size-Chain-胴體">
							<HeaderStyle Width="80px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="Finished" ReadOnly="True" HeaderText="類型">
							<HeaderStyle Width="50px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="Suppiler" ReadOnly="True" HeaderText="外注">
							<HeaderStyle Width="60px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="Level" ReadOnly="True" HeaderText="難易度">
							<HeaderStyle Width="50px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="Step" ReadOnly="True" HeaderText="工程No.">
							<HeaderStyle Width="50px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="TotalTimeRange" ReadOnly="True" HeaderText="起 ~ 迄">
							<HeaderStyle Width="120px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="TotalLeadTime" ReadOnly="True" HeaderText="經過時間(分)">
							<HeaderStyle Width="70px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Right"></ItemStyle>
						</asp:BoundColumn>
					</Columns>
					<PagerStyle NextPageText="下一頁" PrevPageText="上一頁" HorizontalAlign="Center" ForeColor="#993300"
						BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
				</asp:datagrid><asp:dropdownlist id="DEStep" style="Z-INDEX: 113; LEFT: 360px; POSITION: absolute; TOP: 32px" runat="server"
					Width="224px" Height="40px" BackColor="Yellow" ForeColor="Blue"></asp:dropdownlist><asp:dropdownlist id="DSStep" style="Z-INDEX: 103; LEFT: 96px; POSITION: absolute; TOP: 32px" runat="server"
					Width="224px" Height="40px" BackColor="Yellow" ForeColor="Blue"></asp:dropdownlist><asp:dropdownlist id="DFormName" style="Z-INDEX: 100; LEFT: 96px; POSITION: absolute; TOP: 8px" runat="server"
					Width="264px" Height="40px" BackColor="Yellow" ForeColor="Blue" AutoPostBack="True"></asp:dropdownlist></FONT><asp:textbox id="DSDate" style="Z-INDEX: 104; LEFT: 96px; POSITION: absolute; TOP: 80px" runat="server"
				Width="96px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True">DSDate</asp:textbox><asp:button id="BSDate" style="Z-INDEX: 105; LEFT: 192px; POSITION: absolute; TOP: 80px" runat="server"
				Width="20px" Height="20px" Text="....."></asp:button>
			<DIV title="" style="DISPLAY: inline; Z-INDEX: 106; LEFT: 224px; WIDTH: 16px; COLOR: #0000ff; POSITION: absolute; TOP: 80px; HEIGHT: 24px"
				ms_positioning="FlowLayout">～</DIV>
			<asp:textbox id="DEDate" style="Z-INDEX: 107; LEFT: 248px; POSITION: absolute; TOP: 80px" runat="server"
				Width="96px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True">DEDate</asp:textbox><asp:button id="BEDate" style="Z-INDEX: 108; LEFT: 344px; POSITION: absolute; TOP: 80px" runat="server"
				Width="20px" Height="20px" Text="....."></asp:button><asp:button id="Go" style="Z-INDEX: 109; LEFT: 376px; POSITION: absolute; TOP: 80px" runat="server"
				Width="40px" Height="24px" BackColor="White" ForeColor="Blue" Text="Go"></asp:button>
			<DIV title="" style="DISPLAY: inline; Z-INDEX: 110; LEFT: 328px; WIDTH: 16px; COLOR: #0000ff; POSITION: absolute; TOP: 32px; HEIGHT: 24px"
				ms_positioning="FlowLayout">～</DIV>
		</form>
	</body>
</HTML>
