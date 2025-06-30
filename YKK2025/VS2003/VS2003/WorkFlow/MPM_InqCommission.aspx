<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MPM_InqCommission.aspx.vb" Inherits="SPD.MPM_InqCommission"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>調閱資料</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
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
			<DIV style="Z-INDEX: 100; POSITION: absolute; WIDTH: 88px; DISPLAY: inline; HEIGHT: 24px; COLOR: #0000ff; TOP: 8px; LEFT: 8px"
				title="篩選項目：" ms_positioning="FlowLayout">篩選項目：</DIV> <!--EndFragment--><asp:dropdownlist style="Z-INDEX: 120; POSITION: absolute; TOP: 56px; LEFT: 224px" id="Dtype2" runat="server"
				Width="120px" Height="40px" BackColor="Yellow" ForeColor="Blue"></asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 119; POSITION: absolute; TOP: 56px; LEFT: 96px" id="Dtype1" runat="server"
				Width="120px" Height="40px" BackColor="Yellow" ForeColor="Blue"></asp:dropdownlist><asp:imagebutton style="Z-INDEX: 118; POSITION: absolute; TOP: 80px; LEFT: 400px" id="BExcel" runat="server"
				ImageUrl="Images\msexcel.gif" Width="21px" Height="21px"></asp:imagebutton><asp:dropdownlist style="Z-INDEX: 117; POSITION: absolute; TOP: 8px; LEFT: 352px" id="DProgress" runat="server"
				Width="120px" Height="40px" BackColor="Yellow" ForeColor="Blue">
				<asp:ListItem Value="ALL" Selected="True">ALL</asp:ListItem>
				<asp:ListItem Value="1">開發中</asp:ListItem>
				<asp:ListItem Value="2">開發完成</asp:ListItem>
			</asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 116; POSITION: absolute; TOP: 8px; LEFT: 480px" id="DSts" runat="server"
				Width="120px" Height="40px" BackColor="Yellow" ForeColor="Blue">
				<asp:ListItem Value="ALL" Selected="True">ALL</asp:ListItem>
				<asp:ListItem Value="1">OK</asp:ListItem>
				<asp:ListItem Value="2">NG</asp:ListItem>
				<asp:ListItem Value="3">取消</asp:ListItem>
			</asp:dropdownlist><asp:textbox style="Z-INDEX: 114; POSITION: absolute; TOP: 32px; LEFT: 96px" id="DNo" runat="server"
				Width="120px" Height="20px" BackColor="Yellow" ForeColor="Blue" ReadOnly="True" BorderStyle="Groove">DNo</asp:textbox><asp:datagrid style="Z-INDEX: 113; POSITION: absolute; TOP: 120px; LEFT: 8px" id="DataGrid1" runat="server"
				Width="790px" Height="300px" BackColor="White" BorderStyle="None" PageSize="15" AllowPaging="True" Font-Size="9pt" BorderColor="#CC9966" BorderWidth="1px" CellPadding="4" AutoGenerateColumns="False">
				<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
				<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
				<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
				<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
				<Columns>
					<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="ViewURL" DataNavigateUrlFormatString="{0}"
						DataTextField="Field1">
						<HeaderStyle Width="90px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:HyperLinkColumn>
					<asp:BoundColumn DataField="Field2" ReadOnly="True" HeaderText="開發狀態">
						<HeaderStyle Width="90px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="Field3" ReadOnly="True" HeaderText="依賴日">
						<HeaderStyle Width="90px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="Field4" ReadOnly="True" HeaderText="工程">
						<HeaderStyle Width="90px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="Field5" ReadOnly="True" HeaderText="工程擔當">
						<HeaderStyle Width="90px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="Field6" ReadOnly="True" HeaderText="開始預定">
						<HeaderStyle Width="90px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="Field7" ReadOnly="True" HeaderText="完成預定">
						<HeaderStyle Width="90px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="Field8" ReadOnly="True" HeaderText="閱讀期限">
						<HeaderStyle Width="90px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="OPURL" DataNavigateUrlFormatString="{0}" DataTextField="WorkFlow"
						HeaderText="履歷">
						<HeaderStyle Width="30px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Center"></ItemStyle>
					</asp:HyperLinkColumn>
				</Columns>
				<PagerStyle NextPageText="下一頁" PrevPageText="上一頁" HorizontalAlign="Center" ForeColor="#993300"
					BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
			</asp:datagrid><asp:textbox style="Z-INDEX: 111; POSITION: absolute; TOP: 32px; LEFT: 352px" id="DMapNo" runat="server"
				Width="120px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove"></asp:textbox><asp:dropdownlist style="Z-INDEX: 101; POSITION: absolute; TOP: 32px; LEFT: 224px" id="DDivision"
				runat="server" Width="120px" Height="40px" BackColor="Yellow" ForeColor="Blue"></asp:dropdownlist><asp:dropdownlist style="Z-INDEX: 103; POSITION: absolute; TOP: 8px; LEFT: 96px" id="DFormName" runat="server"
				Width="248px" Height="40px" BackColor="Yellow" ForeColor="Blue" AutoPostBack="True"></asp:dropdownlist><asp:button style="Z-INDEX: 104; POSITION: absolute; TOP: 80px; LEFT: 320px" id="BEDate" runat="server"
				Width="20px" Height="20px" Text="....."></asp:button><asp:textbox style="Z-INDEX: 105; POSITION: absolute; TOP: 80px; LEFT: 232px" id="DEDate" runat="server"
				Width="88px" Height="20px" BackColor="Yellow" ForeColor="Blue" ReadOnly="True" BorderStyle="Groove">DEDate</asp:textbox>
			<DIV style="Z-INDEX: 106; POSITION: absolute; WIDTH: 16px; DISPLAY: inline; HEIGHT: 24px; COLOR: #0000ff; TOP: 80px; LEFT: 208px"
				title="" ms_positioning="FlowLayout">～</DIV>
			<asp:button style="Z-INDEX: 107; POSITION: absolute; TOP: 80px; LEFT: 184px" id="BSDate" runat="server"
				Width="24px" Height="20px" Text="....."></asp:button><asp:textbox style="Z-INDEX: 108; POSITION: absolute; TOP: 80px; LEFT: 96px" id="DSDate" runat="server"
				Width="88px" Height="20px" BackColor="Yellow" ForeColor="Blue" ReadOnly="True" BorderStyle="Groove">DSDate</asp:textbox>
			<DIV style="Z-INDEX: 109; POSITION: absolute; WIDTH: 88px; DISPLAY: inline; HEIGHT: 24px; COLOR: #0000ff; TOP: 80px; LEFT: 8px"
				title="完成日期：" ms_positioning="FlowLayout">完成日期：</DIV>
			<asp:button style="Z-INDEX: 110; POSITION: absolute; TOP: 80px; LEFT: 352px" id="Go" runat="server"
				Width="40px" Height="24px" BackColor="White" ForeColor="Blue" Text="Go"></asp:button></form>
	</body>
</HTML>
