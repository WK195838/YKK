<%@ Page Language="vb" AutoEventWireup="false" Codebehind="OPInqCommission_List.aspx.vb" Inherits="SPD.OPInqCommission_List"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>OPInqCommission_List</title>
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
			<DIV title="篩選項目：" style="DISPLAY: inline; Z-INDEX: 100; LEFT: 8px; WIDTH: 88px; COLOR: #0000ff; POSITION: absolute; TOP: 8px; HEIGHT: 24px"
				ms_positioning="FlowLayout">篩選項目：</DIV>
			<asp:hyperlink id="LISIP" style="Z-INDEX: 128; LEFT: 816px; POSITION: absolute; TOP: 80px" runat="server"
				Height="16px" Width="97px" Font-Size="20pt" Target="_blank" NavigateUrl="BoardEdit.aspx">ISIP</asp:hyperlink>
			<asp:textbox id="DSUPPILER" style="Z-INDEX: 127; LEFT: 608px; POSITION: absolute; TOP: 80px"
				runat="server" ReadOnly="True" BorderStyle="Groove" ForeColor="Blue" BackColor="Yellow" Height="20px"
				Width="120px">DSUPPILER</asp:textbox>
			<asp:textbox style="Z-INDEX: 126; LEFT: 480px; POSITION: absolute; TOP: 80px" id="DSellVendor"
				runat="server" Width="120px" Height="20px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove"
				ReadOnly="True">DSellVendor</asp:textbox>
			<asp:dropdownlist style="Z-INDEX: 125; LEFT: 352px; POSITION: absolute; TOP: 80px" id="DKeepData"
				runat="server" ForeColor="Blue" BackColor="Yellow" Height="40px" Width="120px">
				<asp:ListItem Value="0">未封存</asp:ListItem>
				<asp:ListItem Value="1">已封存</asp:ListItem>
				<asp:ListItem Value="9" Selected="True">委託單</asp:ListItem>
			</asp:dropdownlist>
			<asp:textbox id="DCode" style="Z-INDEX: 123; LEFT: 608px; POSITION: absolute; TOP: 32px" runat="server"
				Width="120px" Height="20px" BackColor="Yellow" ForeColor="Blue" ReadOnly="True" BorderStyle="Groove">DCode</asp:textbox>
			<asp:dropdownlist id="DMPSelect" style="Z-INDEX: 122; LEFT: 608px; POSITION: absolute; TOP: 8px" runat="server"
				Enabled="False" AutoPostBack="True" ForeColor="Blue" BackColor="Yellow" Height="40px" Width="120px">
				<asp:ListItem Value="0">一般查詢</asp:ListItem>
				<asp:ListItem Value="1">模具費用</asp:ListItem>
			</asp:dropdownlist><asp:imagebutton id="BExcel" style="Z-INDEX: 121; LEFT: 784px; POSITION: absolute; TOP: 80px" runat="server"
				Height="21px" Width="21px" ImageUrl="Images\msexcel.gif"></asp:imagebutton><asp:dropdownlist id="DProgress" style="Z-INDEX: 120; LEFT: 352px; POSITION: absolute; TOP: 8px" runat="server"
				ForeColor="Blue" BackColor="Yellow" Height="40px" Width="120px">
				<asp:ListItem Value="ALL" Selected="True">ALL</asp:ListItem>
				<asp:ListItem Value="1">開發中</asp:ListItem>
				<asp:ListItem Value="2">開發完成</asp:ListItem>
			</asp:dropdownlist><asp:dropdownlist id="DSts" style="Z-INDEX: 119; LEFT: 480px; POSITION: absolute; TOP: 8px" runat="server"
				ForeColor="Blue" BackColor="Yellow" Height="40px" Width="120px">
				<asp:ListItem Value="ALL" Selected="True">ALL</asp:ListItem>
				<asp:ListItem Value="1">OK</asp:ListItem>
				<asp:ListItem Value="2">NG</asp:ListItem>
				<asp:ListItem Value="3">取消</asp:ListItem>
			</asp:dropdownlist><asp:textbox id="DCpsc" style="Z-INDEX: 118; LEFT: 608px; POSITION: absolute; TOP: 56px" runat="server"
				ForeColor="Blue" BackColor="Yellow" Height="20px" Width="120px" BorderStyle="Groove" ReadOnly="True">DCpsc</asp:textbox><asp:textbox id="DMakeMap" style="Z-INDEX: 117; LEFT: 224px; POSITION: absolute; TOP: 56px" runat="server"
				ForeColor="Blue" BackColor="Yellow" Height="20px" Width="120px" BorderStyle="Groove" ReadOnly="True">DMakeMap</asp:textbox><asp:textbox id="DNo" style="Z-INDEX: 116; LEFT: 96px; POSITION: absolute; TOP: 32px" runat="server"
				ForeColor="Blue" BackColor="Yellow" Height="20px" Width="120px" BorderStyle="Groove" ReadOnly="True">DNo</asp:textbox><asp:textbox id="DSpec" style="Z-INDEX: 115; LEFT: 480px; POSITION: absolute; TOP: 56px" runat="server"
				ForeColor="Blue" BackColor="Yellow" Height="20px" Width="120px" BorderStyle="Groove" ReadOnly="True">DSpec</asp:textbox><asp:datagrid id="DataGrid1" style="Z-INDEX: 114; LEFT: 8px; POSITION: absolute; TOP: 112px" runat="server"
				BackColor="White" Height="300px" Width="790px" BorderStyle="None" AutoGenerateColumns="False" CellPadding="4" BorderWidth="1px" BorderColor="#CC9966" Font-Size="9pt" AllowPaging="True" PageSize="15">
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
						HeaderText="履歷/封存">
						<HeaderStyle Width="60px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Center"></ItemStyle>
					</asp:HyperLinkColumn>
				</Columns>
				<PagerStyle NextPageText="下一頁" PrevPageText="上一頁" HorizontalAlign="Center" ForeColor="#993300"
					BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
			</asp:datagrid><asp:textbox id="DSliderCode" style="Z-INDEX: 113; LEFT: 352px; POSITION: absolute; TOP: 56px"
				runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="120px" BorderStyle="Groove" ReadOnly="True">DSliderCode</asp:textbox><asp:textbox id="DMapNo" style="Z-INDEX: 112; LEFT: 96px; POSITION: absolute; TOP: 56px" runat="server"
				ForeColor="Blue" BackColor="Yellow" Height="20px" Width="120px" BorderStyle="Groove" ReadOnly="True">DMapNo</asp:textbox><asp:textbox id="DPerson" style="Z-INDEX: 111; LEFT: 352px; POSITION: absolute; TOP: 32px" runat="server"
				ForeColor="Blue" BackColor="Yellow" Height="20px" Width="120px" BorderStyle="Groove" ReadOnly="True">DPerson</asp:textbox><asp:textbox id="DBuyer" style="Z-INDEX: 101; LEFT: 480px; POSITION: absolute; TOP: 32px" runat="server"
				ForeColor="Blue" BackColor="Yellow" Height="20px" Width="120px" BorderStyle="Groove" ReadOnly="True">DBuyer</asp:textbox><asp:dropdownlist id="DDivision" style="Z-INDEX: 102; LEFT: 224px; POSITION: absolute; TOP: 32px"
				runat="server" ForeColor="Blue" BackColor="Yellow" Height="40px" Width="120px"></asp:dropdownlist><asp:dropdownlist id="DFormName" style="Z-INDEX: 103; LEFT: 96px; POSITION: absolute; TOP: 8px" runat="server"
				AutoPostBack="True" ForeColor="Blue" BackColor="Yellow" Height="40px" Width="248px"></asp:dropdownlist><asp:button id="BEDate" style="Z-INDEX: 104; LEFT: 320px; POSITION: absolute; TOP: 80px" runat="server"
				Height="20px" Width="20px" Text="....."></asp:button><asp:textbox id="DEDate" style="Z-INDEX: 105; LEFT: 232px; POSITION: absolute; TOP: 80px" runat="server"
				ForeColor="Blue" BackColor="Yellow" Height="20px" Width="88px" BorderStyle="Groove">DEDate</asp:textbox>
			<DIV title="" style="DISPLAY: inline; Z-INDEX: 106; LEFT: 208px; WIDTH: 16px; COLOR: #0000ff; POSITION: absolute; TOP: 80px; HEIGHT: 24px"
				ms_positioning="FlowLayout">～</DIV>
			<asp:button id="BSDate" style="Z-INDEX: 107; LEFT: 184px; POSITION: absolute; TOP: 80px" runat="server"
				Height="20px" Width="24px" Text="....."></asp:button><asp:textbox id="DSDate" style="Z-INDEX: 108; LEFT: 96px; POSITION: absolute; TOP: 80px" runat="server"
				ForeColor="Blue" BackColor="Yellow" Height="20px" Width="88px" BorderStyle="Groove">DSDate</asp:textbox>
			<DIV title="完成日期：" style="DISPLAY: inline; Z-INDEX: 109; LEFT: 8px; WIDTH: 88px; COLOR: #0000ff; POSITION: absolute; TOP: 80px; HEIGHT: 24px"
				ms_positioning="FlowLayout">完成日期：</DIV>
			<asp:button id="Go" style="Z-INDEX: 110; LEFT: 736px; POSITION: absolute; TOP: 80px" runat="server"
				ForeColor="Blue" BackColor="White" Height="24px" Width="40px" Text="Go"></asp:button></form>
	</body>
</HTML>
