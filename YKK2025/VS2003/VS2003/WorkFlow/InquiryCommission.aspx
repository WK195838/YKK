<%@ Page Language="vb" AutoEventWireup="false" Codebehind="InquiryCommission.aspx.vb" Inherits="SPD.InquiryCommission"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>調閱歷史記錄</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<script language="javascript" type="text/javascript">
			function calendarPicker(strField)
			{
				window.open('DatePicker.aspx?field=' + strField,'calendarPopup','width=250,height=190,resizable=yes');
			}
			function MapPicker(strField)
			{
				window.open('MapPicker.aspx?field=' + strField,'MapPopup','width=168,height=328,resizable=yes');
			}
		</script>
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體"></FONT>
			<DIV title="篩選項目：" style="DISPLAY: inline; Z-INDEX: 100; LEFT: 8px; WIDTH: 88px; COLOR: #0000ff; POSITION: absolute; TOP: 8px; HEIGHT: 24px"
				ms_positioning="FlowLayout">篩選項目：</DIV>
			<asp:dropdownlist id="DFormName" style="Z-INDEX: 101; LEFT: 96px; POSITION: absolute; TOP: 8px" runat="server"
				ForeColor="Blue" BackColor="Yellow" Width="165px" Height="40px"></asp:dropdownlist><asp:button id="BEDate" style="Z-INDEX: 102; LEFT: 320px; POSITION: absolute; TOP: 40px" runat="server"
				Width="20px" Height="20px" Text="....."></asp:button><asp:textbox id="DEDate" style="Z-INDEX: 103; LEFT: 232px; POSITION: absolute; TOP: 40px" runat="server"
				ForeColor="Blue" BackColor="Yellow" Width="88px" Height="20px" ReadOnly="True" BorderStyle="Groove">DEDate</asp:textbox>
			<DIV title="篩選項目：" style="DISPLAY: inline; Z-INDEX: 104; LEFT: 208px; WIDTH: 16px; COLOR: #0000ff; POSITION: absolute; TOP: 40px; HEIGHT: 24px"
				ms_positioning="FlowLayout">～</DIV>
			<asp:button id="BSDate" style="Z-INDEX: 105; LEFT: 184px; POSITION: absolute; TOP: 40px" runat="server"
				Width="20px" Height="20px" Text="....."></asp:button><asp:textbox id="DSDate" style="Z-INDEX: 106; LEFT: 96px; POSITION: absolute; TOP: 40px" runat="server"
				ForeColor="Blue" BackColor="Yellow" Width="88px" Height="20px" ReadOnly="True" BorderStyle="Groove">DSDate</asp:textbox>
			<DIV title="篩選項目：" style="DISPLAY: inline; Z-INDEX: 108; LEFT: 8px; WIDTH: 88px; COLOR: #0000ff; POSITION: absolute; TOP: 40px; HEIGHT: 24px"
				ms_positioning="FlowLayout">委託日期：</DIV>
			<DIV title="排序：" style="DISPLAY: inline; Z-INDEX: 109; LEFT: 8px; WIDTH: 88px; COLOR: #0000ff; POSITION: absolute; TOP: 72px; HEIGHT: 24px"
				ms_positioning="FlowLayout">排序：</DIV>
			<asp:dropdownlist id="DSortKey" style="Z-INDEX: 110; LEFT: 96px; POSITION: absolute; TOP: 72px" runat="server"
				ForeColor="Blue" BackColor="Yellow" Width="165px" Height="40px">
				<asp:ListItem Value="T_FormSno" Selected="True">表單單號</asp:ListItem>
				<asp:ListItem Value="T_ApplyTime">委託時間</asp:ListItem>
			</asp:dropdownlist><asp:dropdownlist id="DSort" style="Z-INDEX: 111; LEFT: 264px; POSITION: absolute; TOP: 72px" runat="server"
				ForeColor="Blue" BackColor="Yellow" Width="100px" Height="40px">
				<asp:ListItem Value="ASC">昇順</asp:ListItem>
				<asp:ListItem Value="DESC" Selected="True">降順</asp:ListItem>
			</asp:dropdownlist><asp:button id="Go" style="Z-INDEX: 112; LEFT: 368px; POSITION: absolute; TOP: 72px" runat="server"
				ForeColor="Blue" BackColor="White" Width="40px" Height="24px" Text="Go"></asp:button><asp:datagrid id="DataGrid1" style="Z-INDEX: 113; LEFT: 8px; POSITION: absolute; TOP: 104px" runat="server"
				BackColor="White" Width="494px" Height="300px" BorderStyle="None" Font-Size="9pt" AllowPaging="True" BorderColor="#CC9966" BorderWidth="1px" CellPadding="4" AutoGenerateColumns="False">
				<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
				<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
				<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
				<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
				<Columns>
					<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="T_OPURL" DataNavigateUrlFormatString="{0}"
						DataTextField="T_WorkFlow" HeaderText="工程">
						<HeaderStyle Width="84px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Center"></ItemStyle>
					</asp:HyperLinkColumn>
					<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="T_ViewURL" DataNavigateUrlFormatString="{0}"
						DataTextField="T_FormName" HeaderText="委託單">
						<HeaderStyle Width="116px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:HyperLinkColumn>
					<asp:BoundColumn DataField="T_FormSno" ReadOnly="True" HeaderText="單號">
						<HeaderStyle Width="60px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Center"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="T_Step" ReadOnly="True" HeaderText="No">
						<HeaderStyle Width="12px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Center"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="T_ApplyName" ReadOnly="True" HeaderText="委託人">
						<HeaderStyle Width="72px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Center"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="T_ApplyTime" ReadOnly="True" HeaderText="委託時間">
						<HeaderStyle Width="148px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Center"></ItemStyle>
					</asp:BoundColumn>
				</Columns>
				<PagerStyle NextPageText="下一頁" PrevPageText="上一頁" HorizontalAlign="Center" ForeColor="#993300"
					BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
			</asp:datagrid></form>
	</body>
</HTML>
