<%@ Page Language="vb" AutoEventWireup="false" Codebehind="InquiryDocument.aspx.vb" Inherits="SPD.InquiryDocument"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>各類文件</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體"></FONT>
			<DIV title="篩選項目：" style="DISPLAY: inline; Z-INDEX: 100; LEFT: 8px; WIDTH: 88px; COLOR: #0000ff; POSITION: absolute; TOP: 8px; HEIGHT: 24px"
				ms_positioning="FlowLayout">類 別：</DIV>
			<asp:button id="Go" style="Z-INDEX: 109; LEFT: 288px; POSITION: absolute; TOP: 72px" runat="server"
				ForeColor="Blue" BackColor="White" Height="24px" Width="40px" Text="Go"></asp:button><asp:datagrid id="DataGrid1" style="Z-INDEX: 108; LEFT: 8px; POSITION: absolute; TOP: 104px" runat="server"
				BackColor="White" Height="300px" Width="598px" AutoGenerateColumns="False" CellPadding="4" BorderWidth="1px" BorderColor="#CC9966" AllowPaging="True" BorderStyle="None" Font-Size="9pt">
				<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
				<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
				<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
				<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
				<Columns>
					<asp:BoundColumn DataField="Class" ReadOnly="True" HeaderText="類別">
						<HeaderStyle Width="84px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Center"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="TypeDesc" ReadOnly="True" HeaderText="性質">
						<HeaderStyle Width="48px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Center"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="Year" ReadOnly="True" HeaderText="年">
						<HeaderStyle Width="48px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Center"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="Month" ReadOnly="True" HeaderText="月">
						<HeaderStyle Width="48px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Center"></ItemStyle>
					</asp:BoundColumn>
					<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="URL" DataNavigateUrlFormatString="{0}" DataTextField="Description"
						HeaderText="文件">
						<HeaderStyle Width="220px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:HyperLinkColumn>
					<asp:BoundColumn DataField="Maker" ReadOnly="True" HeaderText="作成者">
						<HeaderStyle Width="60px"></HeaderStyle>
						<ItemStyle HorizontalAlign="center"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="MakerTimeDesc" ReadOnly="True" HeaderText="作成時間">
						<HeaderStyle Width="90px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Center"></ItemStyle>
					</asp:BoundColumn>
				</Columns>
				<PagerStyle NextPageText="下一頁" PrevPageText="上一頁" HorizontalAlign="Center" ForeColor="#993300"
					BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
			</asp:datagrid><asp:dropdownlist id="DMonth" style="Z-INDEX: 107; LEFT: 192px; POSITION: absolute; TOP: 72px" runat="server"
				ForeColor="Blue" BackColor="Yellow" Height="20px" Width="92px">
				<asp:ListItem Value="01" Selected="True">01</asp:ListItem>
				<asp:ListItem Value="02">02</asp:ListItem>
				<asp:ListItem Value="03">03</asp:ListItem>
				<asp:ListItem Value="04">04</asp:ListItem>
				<asp:ListItem Value="05">05</asp:ListItem>
				<asp:ListItem Value="06">06</asp:ListItem>
				<asp:ListItem Value="07">07</asp:ListItem>
				<asp:ListItem Value="08">08</asp:ListItem>
				<asp:ListItem Value="09">09</asp:ListItem>
				<asp:ListItem Value="10">10</asp:ListItem>
				<asp:ListItem Value="11">11</asp:ListItem>
				<asp:ListItem Value="12">12</asp:ListItem>
			</asp:dropdownlist><asp:dropdownlist id="DYear" style="Z-INDEX: 106; LEFT: 96px; POSITION: absolute; TOP: 72px" runat="server"
				ForeColor="Blue" BackColor="Yellow" Height="20px" Width="92px">
				<asp:ListItem Value="2005" Selected="True">2005</asp:ListItem>
				<asp:ListItem Value="2006">2006</asp:ListItem>
				<asp:ListItem Value="2007">2007</asp:ListItem>
				<asp:ListItem Value="2008">2008</asp:ListItem>
				<asp:ListItem Value="2009">2009</asp:ListItem>
				<asp:ListItem Value="2010">2010</asp:ListItem>
				<asp:ListItem Value="2011">2011</asp:ListItem>
				<asp:ListItem Value="2012">2012</asp:ListItem>
				<asp:ListItem Value="2013">2013</asp:ListItem>
				<asp:ListItem Value="2014">2014</asp:ListItem>
				<asp:ListItem Value="2015">2015</asp:ListItem>
			</asp:dropdownlist><asp:dropdownlist id="DType" style="Z-INDEX: 104; LEFT: 96px; POSITION: absolute; TOP: 40px" runat="server"
				ForeColor="Blue" BackColor="Yellow" Height="20px" Width="188px">
				<asp:ListItem Value="0" Selected="True">月報</asp:ListItem>
				<asp:ListItem Value="1">年報</asp:ListItem>
			</asp:dropdownlist><asp:dropdownlist id="DClass" style="Z-INDEX: 103; LEFT: 96px; POSITION: absolute; TOP: 8px" runat="server"
				ForeColor="Blue" BackColor="Yellow" Height="40px" Width="188px"></asp:dropdownlist>
			<DIV title="排序：" style="DISPLAY: inline; Z-INDEX: 101; LEFT: 8px; WIDTH: 88px; COLOR: #0000ff; POSITION: absolute; TOP: 40px; HEIGHT: 24px"
				ms_positioning="FlowLayout">性 質：</DIV>
			<DIV title="排序：" style="DISPLAY: inline; Z-INDEX: 102; LEFT: 8px; WIDTH: 88px; COLOR: #0000ff; POSITION: absolute; TOP: 72px; HEIGHT: 24px"
				ms_positioning="FlowLayout">年 月：</DIV>
		</form>
	</body>
</HTML>
