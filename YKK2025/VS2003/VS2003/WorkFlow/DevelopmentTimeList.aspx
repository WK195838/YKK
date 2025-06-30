<%@ Page Language="vb" AutoEventWireup="false" Codebehind="DevelopmentTimeList.aspx.vb" Inherits="SPD.DevelopmentTimeList"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>開發納期一覽</title>
		<meta content="False" name="vs_snapToGrid">
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體"></FONT>
			<DIV title="" style="Z-INDEX: 101; POSITION: absolute; WIDTH: 8px; DISPLAY: inline; HEIGHT: 24px; COLOR: #0000ff; TOP: 8px; LEFT: 248px"
				ms_positioning="FlowLayout">
				<P>月</P>
			</DIV>
			<asp:label id="Label3" style="Z-INDEX: 118; POSITION: absolute; TEXT-ALIGN: center; TOP: 88px; LEFT: 844px"
				runat="server" Visible="False" ForeColor="White" BackColor="Chocolate" Font-Size="X-Small"
				BorderColor="White" Height="24px" Width="325px">外注委託書</asp:label><asp:label id="Label2" style="Z-INDEX: 117; POSITION: absolute; TEXT-ALIGN: center; TOP: 88px; LEFT: 521px"
				runat="server" Visible="False" ForeColor="White" BackColor="Chocolate" Font-Size="X-Small" BorderColor="White" Height="24px" Width="322px">內製委託書</asp:label>
			<DIV id="DIV1" title="" style="Z-INDEX: 115; POSITION: absolute; WIDTH: 8px; DISPLAY: inline; HEIGHT: 24px; COLOR: #0000ff; TOP: 8px; LEFT: 280px"
				ms_positioning="FlowLayout">
				<P>～</P>
			</DIV>
			<asp:imagebutton id="BExcel" style="Z-INDEX: 110; POSITION: absolute; TOP: 40px; LEFT: 336px" runat="server"
				Height="21px" Width="21px" ImageUrl="Images\msexcel.gif"></asp:imagebutton><asp:datagrid id="DataGrid1" style="Z-INDEX: 109; POSITION: absolute; TOP: 112px; LEFT: 8px" runat="server"
				BackColor="White" Font-Size="9pt" BorderColor="#CC9966" Height="20px" Width="1161px" PageSize="15" BorderWidth="1px" CellPadding="4" AutoGenerateColumns="False" BorderStyle="None">
				<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
				<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
				<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
				<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
				<Columns>
					<asp:BoundColumn DataField="MNo" ReadOnly="True" HeaderText="委託No">
						<HeaderStyle Width="80px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="Buyer" ReadOnly="True" HeaderText="Buyer">
						<HeaderStyle Width="80px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="MLevel" ReadOnly="True" HeaderText="難易度">
						<HeaderStyle Width="40px"></HeaderStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="MapCreateTime" ReadOnly="True" HeaderText="委託起始日">
						<HeaderStyle Width="80px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Right"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="MapFinishTime" ReadOnly="True" HeaderText="最終完成日">
						<HeaderStyle Width="80px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Right"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="MapModTime" ReadOnly="True" HeaderText="經過時間">
						<HeaderStyle Width="40px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Right"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="MapModCount" ReadOnly="True" HeaderText="修改次數">
						<HeaderStyle Width="20px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Center"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="MapSts" ReadOnly="True" HeaderText="狀態">
						<HeaderStyle Width="30px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Center"></ItemStyle>
					</asp:BoundColumn>
					<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="MISUrl" DataNavigateUrlFormatString="{0}"
						DataTextField="ManufInCreateTime" HeaderText="委託起始日">
						<HeaderStyle Width="80px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Right"></ItemStyle>
					</asp:HyperLinkColumn>
					<asp:BoundColumn DataField="ManufInFinishTime" ReadOnly="True" HeaderText="最終完成日">
						<HeaderStyle Width="80px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Right"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="ManufInModTime" ReadOnly="True" HeaderText="經過時間">
						<HeaderStyle Width="40px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Right"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="ManufInModCount" ReadOnly="True" HeaderText="開發次數">
						<HeaderStyle Width="20px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Center"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="ManufInSts" ReadOnly="True" HeaderText="狀態">
						<HeaderStyle Width="30px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Center"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="MapToInET" ReadOnly="True" HeaderText="繪圖內製經過時間">
						<HeaderStyle Width="30px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Right" VerticalAlign="Middle"></ItemStyle>
					</asp:BoundColumn>
					<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="MOSUrl" DataNavigateUrlFormatString="{0}"
						DataTextField="ManufOutCreateTime" HeaderText="委託起始日">
						<HeaderStyle Width="80px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Right"></ItemStyle>
					</asp:HyperLinkColumn>
					<asp:BoundColumn DataField="ManufOutFinishTime" ReadOnly="True" HeaderText="最終完成日">
						<HeaderStyle Width="80px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Right"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="ManufOutModTime" ReadOnly="True" HeaderText="經過時間">
						<HeaderStyle Width="40px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Right"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="ManufOutModCount" ReadOnly="True" HeaderText="開發次數">
						<HeaderStyle Width="20px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Center"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="ManufOutSts" ReadOnly="True" HeaderText="狀態">
						<HeaderStyle Width="30px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Center"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="MapToOutET" ReadOnly="True" HeaderText="繪圖外注經過時間">
						<HeaderStyle Width="30px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Right" VerticalAlign="Middle"></ItemStyle>
					</asp:BoundColumn>
				</Columns>
				<PagerStyle NextPageText="下一頁" PrevPageText="上一頁" HorizontalAlign="Center" ForeColor="#993300"
					BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
			</asp:datagrid><asp:button id="Go" style="Z-INDEX: 108; POSITION: absolute; TOP: 40px; LEFT: 272px" runat="server"
				ForeColor="Blue" BackColor="White" Height="24px" Width="40px" Text="Go"></asp:button><asp:dropdownlist id="DSMonth" style="Z-INDEX: 102; POSITION: absolute; TOP: 8px; LEFT: 200px" runat="server"
				ForeColor="Blue" BackColor="Yellow" Height="40px" Width="48px" AutoPostBack="True">
				<asp:ListItem Value="1">01</asp:ListItem>
				<asp:ListItem Value="2">02</asp:ListItem>
				<asp:ListItem Value="3">03</asp:ListItem>
				<asp:ListItem Value="4">04</asp:ListItem>
				<asp:ListItem Value="5">05</asp:ListItem>
				<asp:ListItem Value="6">06</asp:ListItem>
				<asp:ListItem Value="7">07</asp:ListItem>
				<asp:ListItem Value="8">08</asp:ListItem>
				<asp:ListItem Value="9">09</asp:ListItem>
				<asp:ListItem Value="10">10</asp:ListItem>
				<asp:ListItem Value="11">11</asp:ListItem>
				<asp:ListItem Value="12">12</asp:ListItem>
			</asp:dropdownlist>
			<DIV title="" style="Z-INDEX: 103; POSITION: absolute; WIDTH: 1px; DISPLAY: inline; HEIGHT: 24px; COLOR: #0000ff; TOP: 8px; LEFT: 176px"
				ms_positioning="FlowLayout">年</DIV>
			<asp:dropdownlist id="DSYear" style="Z-INDEX: 104; POSITION: absolute; TOP: 8px; LEFT: 104px" runat="server"
				ForeColor="Blue" BackColor="Yellow" Height="40px" Width="72px" AutoPostBack="True"></asp:dropdownlist>
			<DIV title="委託年月：" style="Z-INDEX: 105; POSITION: absolute; WIDTH: 96px; DISPLAY: inline; HEIGHT: 24px; COLOR: #0000ff; TOP: 8px; LEFT: 8px"
				ms_positioning="FlowLayout">委託年月：</DIV>
			<asp:dropdownlist id="DBuyer" style="Z-INDEX: 106; POSITION: absolute; TOP: 40px; LEFT: 104px" runat="server"
				ForeColor="Blue" BackColor="Yellow" Height="40px" Width="144px"></asp:dropdownlist>
			<DIV title="BUYER：" style="Z-INDEX: 107; POSITION: absolute; WIDTH: 80px; DISPLAY: inline; HEIGHT: 24px; COLOR: #0000ff; TOP: 40px; LEFT: 8px"
				ms_positioning="FlowLayout">
				<P>BUYER：</P>
			</DIV>
			<asp:dropdownlist id="DEYear" style="Z-INDEX: 111; POSITION: absolute; TOP: 8px; LEFT: 312px" runat="server"
				ForeColor="Blue" BackColor="Yellow" Height="40px" Width="72px" AutoPostBack="True"></asp:dropdownlist>
			<DIV title="" style="Z-INDEX: 112; POSITION: absolute; WIDTH: 1px; DISPLAY: inline; HEIGHT: 24px; COLOR: #0000ff; TOP: 8px; LEFT: 384px"
				ms_positioning="FlowLayout">年</DIV>
			<asp:dropdownlist id="DEMonth" style="Z-INDEX: 113; POSITION: absolute; TOP: 8px; LEFT: 408px" runat="server"
				ForeColor="Blue" BackColor="Yellow" Height="40px" Width="48px" AutoPostBack="True">
				<asp:ListItem Value="1">01</asp:ListItem>
				<asp:ListItem Value="2">02</asp:ListItem>
				<asp:ListItem Value="3">03</asp:ListItem>
				<asp:ListItem Value="4">04</asp:ListItem>
				<asp:ListItem Value="5">05</asp:ListItem>
				<asp:ListItem Value="6">06</asp:ListItem>
				<asp:ListItem Value="7">07</asp:ListItem>
				<asp:ListItem Value="8">08</asp:ListItem>
				<asp:ListItem Value="9">09</asp:ListItem>
				<asp:ListItem Value="10">10</asp:ListItem>
				<asp:ListItem Value="11">11</asp:ListItem>
				<asp:ListItem Value="12">12</asp:ListItem>
			</asp:dropdownlist>
			<DIV title="" style="Z-INDEX: 114; POSITION: absolute; WIDTH: 16px; DISPLAY: inline; HEIGHT: 24px; COLOR: #0000ff; TOP: 8px; LEFT: 456px"
				ms_positioning="FlowLayout">
				<P>月</P>
			</DIV>
			<asp:label id="Label1" style="Z-INDEX: 116; POSITION: absolute; TEXT-ALIGN: center; TOP: 88px; LEFT: 186px"
				runat="server" Visible="False" ForeColor="White" BackColor="Chocolate" Font-Size="X-Small"
				BorderColor="White" Height="24px" Width="334px">繪圖階段</asp:label><asp:label id="Label4" style="Z-INDEX: 119; POSITION: absolute; TOP: 72px; LEFT: 1082px" runat="server"
				Visible="False" Font-Size="XX-Small" Height="9px" Width="81px">時間單位：小時</asp:label></form>
	</body>
</HTML>
