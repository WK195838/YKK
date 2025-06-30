<%@ Page Language="vb" AutoEventWireup="false" Codebehind="OPInqCommission_ing.aspx.vb" Inherits="SPD.OPInqCommission_ing"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>OPInqCommission</title>
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
			<asp:textbox id="DCpsc" style="Z-INDEX: 120; LEFT: 96px; POSITION: absolute; TOP: 56px" runat="server"
				ReadOnly="True" BorderStyle="Groove" Height="20px" Width="120px" BackColor="Yellow" ForeColor="Blue">DCpsc</asp:textbox>
			<asp:textbox id="DMakeMap" style="Z-INDEX: 119; LEFT: 352px; POSITION: absolute; TOP: 32px" runat="server"
				ForeColor="Blue" BackColor="Yellow" Width="120px" Height="20px" BorderStyle="Groove" ReadOnly="True">DMakeMap</asp:textbox>
			<asp:textbox id="DNo" style="Z-INDEX: 118; LEFT: 352px; POSITION: absolute; TOP: 8px" runat="server"
				ForeColor="Blue" BackColor="Yellow" Width="120px" Height="20px" BorderStyle="Groove" ReadOnly="True">DNo</asp:textbox>
			<asp:dropdownlist id="DSts" style="Z-INDEX: 117; LEFT: 224px; POSITION: absolute; TOP: 8px" runat="server"
				Height="40px" Width="120px" BackColor="Yellow" ForeColor="Blue">
				<asp:ListItem Value="0" Selected="True">開發中</asp:ListItem>
			</asp:dropdownlist><asp:textbox id="DSpec" style="Z-INDEX: 116; LEFT: 608px; POSITION: absolute; TOP: 32px" runat="server"
				Height="20px" Width="120px" BackColor="Yellow" ForeColor="Blue" ReadOnly="True" BorderStyle="Groove">DSpec</asp:textbox><asp:datagrid id="DataGrid1" style="Z-INDEX: 115; LEFT: 8px; POSITION: absolute; TOP: 112px" runat="server"
				Height="300px" Width="790px" BackColor="White" BorderStyle="None" PageSize="15" AllowPaging="True" Font-Size="9pt" BorderColor="#CC9966" BorderWidth="1px" CellPadding="4" AutoGenerateColumns="False">
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
					<asp:BoundColumn DataField="Field2" ReadOnly="True" HeaderText="依賴者">
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
			</asp:datagrid><asp:textbox id="DSliderCode" style="Z-INDEX: 114; LEFT: 480px; POSITION: absolute; TOP: 32px"
				runat="server" Height="20px" Width="120px" BackColor="Yellow" ForeColor="Blue" ReadOnly="True" BorderStyle="Groove">DSliderCode</asp:textbox><asp:textbox id="DMapNo" style="Z-INDEX: 113; LEFT: 224px; POSITION: absolute; TOP: 32px" runat="server"
				Height="20px" Width="120px" BackColor="Yellow" ForeColor="Blue" ReadOnly="True" BorderStyle="Groove">DMapNo</asp:textbox><asp:textbox id="DPerson" style="Z-INDEX: 112; LEFT: 608px; POSITION: absolute; TOP: 8px" runat="server"
				Height="20px" Width="120px" BackColor="Yellow" ForeColor="Blue" ReadOnly="True" BorderStyle="Groove">DPerson</asp:textbox><asp:textbox id="DBuyer" style="Z-INDEX: 101; LEFT: 96px; POSITION: absolute; TOP: 32px" runat="server"
				Height="20px" Width="120px" BackColor="Yellow" ForeColor="Blue" ReadOnly="True" BorderStyle="Groove">DBuyer</asp:textbox><asp:dropdownlist id="DDivision" style="Z-INDEX: 102; LEFT: 480px; POSITION: absolute; TOP: 8px" runat="server"
				Height="40px" Width="120px" BackColor="Yellow" ForeColor="Blue"></asp:dropdownlist><asp:dropdownlist id="DFormName" style="Z-INDEX: 103; LEFT: 96px; POSITION: absolute; TOP: 8px" runat="server"
				Height="40px" Width="120px" BackColor="Yellow" ForeColor="Blue" AutoPostBack="True"></asp:dropdownlist><asp:button id="BEDate" style="Z-INDEX: 104; LEFT: 320px; POSITION: absolute; TOP: 80px" runat="server"
				Height="20px" Width="20px" Text="....."></asp:button><asp:textbox id="DEDate" style="Z-INDEX: 105; LEFT: 232px; POSITION: absolute; TOP: 80px" runat="server"
				Height="20px" Width="88px" BackColor="Yellow" ForeColor="Blue" ReadOnly="True" BorderStyle="Groove">DEDate</asp:textbox>
			<DIV title="" style="DISPLAY: inline; Z-INDEX: 106; LEFT: 208px; WIDTH: 16px; COLOR: #0000ff; POSITION: absolute; TOP: 80px; HEIGHT: 24px"
				ms_positioning="FlowLayout">～</DIV>
			<asp:button id="BSDate" style="Z-INDEX: 107; LEFT: 184px; POSITION: absolute; TOP: 80px" runat="server"
				Height="20px" Width="24px" Text="....."></asp:button><asp:textbox id="DSDate" style="Z-INDEX: 108; LEFT: 96px; POSITION: absolute; TOP: 80px" runat="server"
				Height="20px" Width="88px" BackColor="Yellow" ForeColor="Blue" ReadOnly="True" BorderStyle="Groove">DSDate</asp:textbox>
			<DIV title="完成日期：" style="DISPLAY: inline; Z-INDEX: 109; LEFT: 8px; WIDTH: 88px; COLOR: #0000ff; POSITION: absolute; TOP: 80px; HEIGHT: 24px"
				ms_positioning="FlowLayout">完成日期：</DIV>
			<asp:button id="Go" style="Z-INDEX: 110; LEFT: 352px; POSITION: absolute; TOP: 80px" runat="server"
				Height="24px" Width="40px" BackColor="White" ForeColor="Blue" Text="Go"></asp:button></form>
	</body>
</HTML>
