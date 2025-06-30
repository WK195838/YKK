<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MapManagement.aspx.vb" Inherits="SPD.MapManagement"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>圖面管理</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體">
				<DIV title="篩選項目：" style="DISPLAY: inline; Z-INDEX: 100; LEFT: 8px; WIDTH: 88px; COLOR: #0000ff; POSITION: absolute; TOP: 8px; HEIGHT: 24px"
					ms_positioning="FlowLayout">篩選項目：</DIV>
				<asp:imagebutton id="BExcel1" style="Z-INDEX: 108; LEFT: 776px; POSITION: absolute; TOP: 384px" runat="server"
					Width="21px" Height="21px" ImageUrl="Images\msexcel.gif" Visible="False"></asp:imagebutton>
				<asp:imagebutton id="BExcel" style="Z-INDEX: 107; LEFT: 736px; POSITION: absolute; TOP: 32px" runat="server"
					Width="21px" Height="21px" ImageUrl="Images\msexcel.gif"></asp:imagebutton>
				<asp:dropdownlist id="DDivision" style="Z-INDEX: 106; LEFT: 96px; POSITION: absolute; TOP: 8px" runat="server"
					BackColor="Yellow" Height="40px" Width="112px" ForeColor="Blue"></asp:dropdownlist><asp:textbox id="DBuyer" style="Z-INDEX: 105; LEFT: 208px; POSITION: absolute; TOP: 8px" runat="server"
					BorderStyle="Groove" BackColor="Yellow" Height="20px" Width="120px" ForeColor="Blue">DBuyer</asp:textbox><asp:datagrid id="Datagrid2" style="Z-INDEX: 104; LEFT: 8px; POSITION: absolute; TOP: 56px" runat="server"
					DataKeyField="MapNo" AutoGenerateColumns="False" CellPadding="4" BorderWidth="1px" BorderColor="#CC9966" AllowPaging="True" Font-Size="Smaller" BorderStyle="None" BackColor="White" Height="328px" Width="750px">
					<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
					<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
					<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
					<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
					<Columns>
						<asp:ButtonColumn Text="顯示" HeaderText="修改履歷" CommandName="ShowInformation">
							<HeaderStyle Width="80px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:ButtonColumn>
						<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="URL" DataNavigateUrlFormatString="{0}" DataTextField="MapNo"
							HeaderText="原始圖號">
							<HeaderStyle Width="110px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:HyperLinkColumn>
						<asp:BoundColumn DataField="StsDesc" HeaderText="狀態">
							<HeaderStyle Width="100px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="CompletedTime" HeaderText="完成日">
							<HeaderStyle Width="150px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="Buyer" HeaderText="Buyer">
							<HeaderStyle Width="200px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="No" HeaderText="No">
							<HeaderStyle Width="80px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="FormSno" HeaderText="單號">
							<HeaderStyle Width="60px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
					</Columns>
					<PagerStyle NextPageText="下一頁" PrevPageText="上一頁" HorizontalAlign="Center" ForeColor="#993300"
						BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
				</asp:datagrid><asp:datagrid id="DataGrid1" style="Z-INDEX: 103; LEFT: 8px; POSITION: absolute; TOP: 408px" runat="server"
					DataKeyField="MapNo" AutoGenerateColumns="False" CellPadding="4" BorderWidth="1px" BorderColor="#CC9966"
					AllowPaging="True" Font-Size="Smaller" BorderStyle="None" BackColor="White" Height="328px" Width="790px">
					<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
					<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
					<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
					<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
					<Columns>
						<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="URL" DataNavigateUrlFormatString="{0}" DataTextField="MapNo"
							HeaderText="修改圖號">
							<HeaderStyle Width="100px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:HyperLinkColumn>
						<asp:BoundColumn DataField="StsDesc" HeaderText="狀態">
							<HeaderStyle Width="100px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="CompletedTime" HeaderText="完成日">
							<HeaderStyle Width="140px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="No" HeaderText="No">
							<HeaderStyle Width="40px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="FormSno" HeaderText="單號">
							<HeaderStyle Width="40px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="ModReasonCode" HeaderText="修改原因">
							<HeaderStyle Width="150px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="ModContent" HeaderText="修改內容">
							<HeaderStyle Width="220px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
					</Columns>
					<PagerStyle NextPageText="下一頁" PrevPageText="上一頁" HorizontalAlign="Center" ForeColor="#993300"
						BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
				</asp:datagrid><asp:button id="Go" style="Z-INDEX: 102; LEFT: 480px; POSITION: absolute; TOP: 8px" runat="server"
					ForeColor="Blue" Width="40px" Height="24px" BackColor="White" Text="Go"></asp:button><asp:textbox id="DMapNo" style="Z-INDEX: 101; LEFT: 328px; POSITION: absolute; TOP: 8px" runat="server"
					ForeColor="Blue" Width="152px" Height="20px" BackColor="Yellow" BorderStyle="Groove">DMapNo</asp:textbox></FONT></form>
	</body>
</HTML>
