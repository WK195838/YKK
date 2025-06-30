<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SCD_DelayList.aspx.vb" Inherits="SPD.SCD_DelayList"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>延遲未核定</title>
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
			<FONT face="新細明體"></FONT>
			<DIV title="篩選項目：" style="Z-INDEX: 100; POSITION: absolute; WIDTH: 88px; DISPLAY: inline; HEIGHT: 24px; COLOR: #0000ff; TOP: 8px; LEFT: 8px"
				ms_positioning="FlowLayout">篩選項目：</DIV>
			<asp:textbox style="Z-INDEX: 110; POSITION: absolute; TOP: 32px; LEFT: 352px" id="DDecideName"
				runat="server" Width="120px" Height="20px" ReadOnly="True" BorderStyle="Groove" BackColor="Yellow"
				ForeColor="Blue">DDecideName</asp:textbox>
			<asp:dropdownlist style="Z-INDEX: 109; POSITION: absolute; TOP: 8px; LEFT: 352px" id="DDelay" runat="server"
				Width="120px" Height="40px" BackColor="Yellow" ForeColor="Blue">
				<asp:ListItem Value="1" Selected="True">延遲</asp:ListItem>
				<asp:ListItem Value="ALL">ALL</asp:ListItem>
			</asp:dropdownlist>
			<asp:imagebutton id="BExcel" style="Z-INDEX: 108; POSITION: absolute; TOP: 32px; LEFT: 536px" runat="server"
				ImageUrl="Images\msexcel.gif" Height="21px" Width="21px"></asp:imagebutton><asp:textbox id="DNo" style="Z-INDEX: 107; POSITION: absolute; TOP: 32px; LEFT: 96px" runat="server"
				Height="20px" Width="120px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DNo</asp:textbox><asp:datagrid id="DataGrid1" style="Z-INDEX: 106; POSITION: absolute; TOP: 64px; LEFT: 8px" runat="server"
				Height="100px" Width="960px" BackColor="White" BorderStyle="None" AutoGenerateColumns="False" CellPadding="4" BorderWidth="1px" BorderColor="#CC9966" Font-Size="9pt" AllowPaging="True" PageSize="20">
				<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
				<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
				<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
				<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
				<Columns>
					<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="ViewURL" DataNavigateUrlFormatString="{0}"
						DataTextField="Field1">
						<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:HyperLinkColumn>
					<asp:BoundColumn DataField="Field2" ReadOnly="True" HeaderText="狀態">
						<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
						<ItemStyle HorizontalAlign="Center"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="Field3" ReadOnly="True" HeaderText="依賴日">
						<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="Field4" ReadOnly="True" HeaderText="工程">
						<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="Field5" ReadOnly="True" HeaderText="工程擔當">
						<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="Field6" ReadOnly="True" HeaderText="開始預定">
						<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="Field7" ReadOnly="True" HeaderText="開始預定">
						<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="Field8" ReadOnly="True" HeaderText="開始預定">
						<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="Field9" ReadOnly="True" HeaderText="開始預定">
						<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="OPURL" DataNavigateUrlFormatString="{0}" DataTextField="WorkFlow"
						HeaderText="履歷">
						<HeaderStyle HorizontalAlign="Center" Width="30px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Center"></ItemStyle>
					</asp:HyperLinkColumn>
				</Columns>
				<PagerStyle NextPageText="下一頁" PrevPageText="上一頁" HorizontalAlign="Center" ForeColor="#993300"
					BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
			</asp:datagrid><asp:textbox id="DName" style="Z-INDEX: 104; POSITION: absolute; TOP: 32px; LEFT: 224px" runat="server"
				Height="20px" Width="120px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" ReadOnly="True">DName</asp:textbox><asp:dropdownlist id="DDivision" style="Z-INDEX: 101; POSITION: absolute; TOP: 32px; LEFT: 632px"
				runat="server" Height="40px" Width="120px" ForeColor="Blue" BackColor="Yellow" Visible="False"></asp:dropdownlist><asp:dropdownlist id="DFormName" style="Z-INDEX: 102; POSITION: absolute; TOP: 8px; LEFT: 96px" runat="server"
				Height="40px" Width="248px" ForeColor="Blue" BackColor="Yellow" AutoPostBack="True"></asp:dropdownlist><asp:button id="Go" style="Z-INDEX: 103; POSITION: absolute; TOP: 32px; LEFT: 488px" runat="server"
				Height="24px" Width="40px" ForeColor="Blue" BackColor="White" Text="Go"></asp:button></form>
	</body>
</HTML>
