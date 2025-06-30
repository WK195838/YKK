<%@ Page Language="vb" AutoEventWireup="false" Codebehind="DevelopLeadTime_EAList.aspx.vb" Inherits="SPD.DevelopLeadTime_EAList"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>開發時間</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體">
				<asp:datagrid id="DataGrid1" style="Z-INDEX: 113; LEFT: 8px; POSITION: absolute; TOP: 40px" runat="server"
					PageSize="15" BorderColor="#CC9966" BorderWidth="1px" CellPadding="4" AutoGenerateColumns="False"
					Font-Size="9pt" BorderStyle="None" Width="1100px" Height="300px" BackColor="White">
					<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
					<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
					<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
					<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
					<Columns>
						<asp:BoundColumn DataField="No" ReadOnly="True" HeaderText="No">
							<HeaderStyle Width="80px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="Step" ReadOnly="True" HeaderText="工程No.">
							<HeaderStyle Width="50px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="ATimeDesc" ReadOnly="True" HeaderText="(a)起 ~ 迄">
							<HeaderStyle Width="250px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="ATime" ReadOnly="True" HeaderText="(a)時間">
							<HeaderStyle Width="50px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="BTimeDesc" ReadOnly="True" HeaderText="(b)起 ~ 迄">
							<HeaderStyle Width="250px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="BTime" ReadOnly="True" HeaderText="(b)時間">
							<HeaderStyle Width="50px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="CTimeDesc" ReadOnly="True" HeaderText="(c)起 ~ 迄">
							<HeaderStyle Width="250px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="CTime" ReadOnly="True" HeaderText="(c)時間">
							<HeaderStyle Width="50px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
					</Columns>
					<PagerStyle NextPageText="下一頁" PrevPageText="上一頁" HorizontalAlign="Center" ForeColor="#993300"
						BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
				</asp:datagrid>
				<DIV title="" style="DISPLAY: inline; Z-INDEX: 120; LEFT: 1024px; WIDTH: 79px; COLOR: #0000ff; POSITION: absolute; TOP: 8px; HEIGHT: 24px"
					ms_positioning="FlowLayout">單位：分</DIV>
				<asp:textbox id="TextBox1" style="Z-INDEX: 119; LEFT: 8px; POSITION: absolute; TOP: 8px" runat="server"
					BorderStyle="None" Width="712px" Height="24px" ReadOnly="True" TextMode="MultiLine">(a)時間：實際開始(系統計算)～實際完成　(b)時間：收件開始～實際完成　(c)時間：委託開始～實際完成</asp:textbox><asp:imagebutton id="BExcel" style="Z-INDEX: 117; LEFT: 992px; POSITION: absolute; TOP: 8px" runat="server"
					Width="21px" Height="21px" ImageUrl="Images\msexcel.gif"></asp:imagebutton></FONT></form>
	</body>
</HTML>
