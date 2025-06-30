<%@ Page Language="vb" AutoEventWireup="false" Codebehind="InquiryCommissionSPD.aspx.vb" Inherits="SPD.InquiryCommissionSPD"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>調閱歷史記錄(SPD)</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體"></FONT>
			<DIV title="篩選項目：" style="DISPLAY: inline; Z-INDEX: 116; LEFT: 8px; WIDTH: 88px; COLOR: #0000ff; POSITION: absolute; TOP: 8px; HEIGHT: 24px"
				ms_positioning="FlowLayout">篩選項目：</DIV>
			<asp:textbox id="DBuyer" style="Z-INDEX: 117; LEFT: 376px; POSITION: absolute; TOP: 8px" runat="server"
				Height="20px" Width="120px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DBuyer</asp:textbox>
			<asp:dropdownlist id="DDivision" style="Z-INDEX: 118; LEFT: 264px; POSITION: absolute; TOP: 8px" runat="server"
				Height="40px" Width="112px" BackColor="Yellow" ForeColor="Blue"></asp:dropdownlist>
			<asp:dropdownlist id="DFormName" style="Z-INDEX: 119; LEFT: 96px; POSITION: absolute; TOP: 8px" runat="server"
				Height="40px" Width="165px" BackColor="Yellow" ForeColor="Blue"></asp:dropdownlist>
			<asp:button id="BEDate" style="Z-INDEX: 120; LEFT: 320px; POSITION: absolute; TOP: 40px" runat="server"
				Height="20px" Width="20px" Text="....."></asp:button>
			<asp:textbox id="DEDate" style="Z-INDEX: 121; LEFT: 232px; POSITION: absolute; TOP: 40px" runat="server"
				Height="20px" Width="88px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True">DEDate</asp:textbox>
			<DIV title="篩選項目：" style="DISPLAY: inline; Z-INDEX: 122; LEFT: 208px; WIDTH: 16px; COLOR: #0000ff; POSITION: absolute; TOP: 40px; HEIGHT: 24px"
				ms_positioning="FlowLayout">～</DIV>
			<asp:button id="BSDate" style="Z-INDEX: 123; LEFT: 184px; POSITION: absolute; TOP: 40px" runat="server"
				Height="20px" Width="20px" Text="....."></asp:button>
			<asp:textbox id="DSDate" style="Z-INDEX: 124; LEFT: 96px; POSITION: absolute; TOP: 40px" runat="server"
				Height="20px" Width="88px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True">DSDate</asp:textbox>
			<DIV title="篩選項目：" style="DISPLAY: inline; Z-INDEX: 125; LEFT: 8px; WIDTH: 88px; COLOR: #0000ff; POSITION: absolute; TOP: 40px; HEIGHT: 24px"
				ms_positioning="FlowLayout">委託日期：</DIV>
			<DIV title="排序：" style="DISPLAY: inline; Z-INDEX: 126; LEFT: 8px; WIDTH: 88px; COLOR: #0000ff; POSITION: absolute; TOP: 72px; HEIGHT: 24px"
				ms_positioning="FlowLayout">排序：</DIV>
			<asp:dropdownlist id="DSortKey" style="Z-INDEX: 127; LEFT: 96px; POSITION: absolute; TOP: 72px" runat="server"
				Height="40px" Width="165px" BackColor="Yellow" ForeColor="Blue">
				<asp:ListItem Value="T_FormSno" Selected="True">表單單號</asp:ListItem>
				<asp:ListItem Value="T_ApplyTime">委託時間</asp:ListItem>
			</asp:dropdownlist>
			<asp:dropdownlist id="DSort" style="Z-INDEX: 128; LEFT: 264px; POSITION: absolute; TOP: 72px" runat="server"
				Height="40px" Width="100px" BackColor="Yellow" ForeColor="Blue">
				<asp:ListItem Value="ASC">昇順</asp:ListItem>
				<asp:ListItem Value="DESC" Selected="True">降順</asp:ListItem>
			</asp:dropdownlist>
			<asp:button id="Go" style="Z-INDEX: 129; LEFT: 368px; POSITION: absolute; TOP: 72px" runat="server"
				Height="24px" Width="40px" BackColor="White" ForeColor="Blue" Text="Go"></asp:button>
			<asp:datagrid id="DataGrid1" style="Z-INDEX: 130; LEFT: 8px; POSITION: absolute; TOP: 104px" runat="server"
				Height="300px" Width="482px" BackColor="White" BorderStyle="None" Font-Size="9pt" AllowPaging="True"
				BorderColor="#CC9966" BorderWidth="1px" CellPadding="4" AutoGenerateColumns="False">
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
			</asp:datagrid>
		</form>
	</body>
</HTML>
