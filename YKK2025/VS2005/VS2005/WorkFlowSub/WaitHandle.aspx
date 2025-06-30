<%@ Page Language="VB" AutoEventWireup="false" CodeFile="WaitHandle.aspx.vb" Inherits="WaitHandle" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>待處理</title>
</head>
<body style="font-size: 12pt">
    <form id="form1" runat="server">
    <div>

			<FONT face="新細明體">

				<asp:dropdownlist id="DSts" style="Z-INDEX: 100; POSITION: absolute; TOP: 8px; LEFT: 264px" runat="server"
					ForeColor="Blue" BackColor="Yellow"  Width="100px">
					<asp:ListItem Value="ALL" Selected="True">ALL</asp:ListItem>
					<asp:ListItem Value="0">未處理</asp:ListItem>
					<asp:ListItem Value="3">已閱讀</asp:ListItem>
				</asp:dropdownlist>
				
				<asp:textbox style="Z-INDEX: 113; POSITION: absolute; TOP: 40px; LEFT: 368px" id="DKey1" runat="server"
					Width="96px" Height="22px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove">DKey1</asp:textbox><DIV title="篩選項目：" style="Z-INDEX: 109; POSITION: absolute; WIDTH: 88px; DISPLAY: inline; HEIGHT: 24px; COLOR: #0000ff; TOP: 8px; LEFT: 8px"
					ms_positioning="FlowLayout">篩選項目：</DIV>
				<asp:dropdownlist id="DFormName" style="Z-INDEX: 108; POSITION: absolute; TOP: 8px; LEFT: 96px" runat="server"
					ForeColor="Blue" BackColor="Yellow"  Width="165px" AutoPostBack="True">
					</asp:dropdownlist>

				<asp:button id="Go" style="Z-INDEX: 106; POSITION: absolute; TOP: 40px; LEFT: 472px" runat="server"
					ForeColor="Blue" BackColor="White" Height="24px" Width="40px" Text="Go"></asp:button>
					
				<asp:button id="Refresh" style="Z-INDEX: 106; POSITION: absolute; TOP: 40px; LEFT: 516px" runat="server"
					ForeColor="Blue" BackColor="White" Height="24px" Width="65px" Text="重新整理"></asp:button>

        <asp:HyperLink ID="LPortalPage" runat="server" Font-Size="12pt" Height="8px"
            NavigateUrl="http://10.245.1.10/Portal" style="Z-INDEX: 100; POSITION: absolute; TOP: 8px; LEFT: 600px" Target="_self" Width="66px">回首頁</asp:HyperLink>

					
					
					<asp:dropdownlist id="DSort" style="Z-INDEX: 105; POSITION: absolute; TOP: 40px; LEFT: 264px" runat="server"
					ForeColor="Blue" BackColor="Yellow"  Width="100px">
					<asp:ListItem Value="ASC">昇順</asp:ListItem>
					<asp:ListItem Value="DESC" Selected="True">降順</asp:ListItem>
				</asp:dropdownlist><FONT face="新細明體">
					<DIV title="排序：" style="Z-INDEX: 104; POSITION: absolute; WIDTH: 88px; DISPLAY: inline; HEIGHT: 24px; COLOR: #0000ff; TOP: 40px; LEFT: 8px"
						ms_positioning="FlowLayout">排序：</DIV>
					<asp:dropdownlist id="DSortKey" style="Z-INDEX: 103; POSITION: absolute; TOP: 40px; LEFT: 96px" runat="server"
						ForeColor="Blue" BackColor="Yellow"  Width="165px">
						<asp:ListItem Value="No" Selected="True">委託No</asp:ListItem>
						<asp:ListItem Value="ApplyTime">委託時間</asp:ListItem>
						<asp:ListItem Value="ReceiptTime">收件時間</asp:ListItem>
						<asp:ListItem Value="ReadTimeLimit">閱讀期限</asp:ListItem>
						<asp:ListItem Value="FirstReadTime">第一次閱讀時間</asp:ListItem>
						<asp:ListItem Value="LastReadTime">最後一次閱讀時間</asp:ListItem>
						<asp:ListItem Value="BStartTime">預定開始時間</asp:ListItem>
						<asp:ListItem Value="BEndTime">預定完成時間</asp:ListItem>
						<asp:ListItem Value="AStartTime">實際開始時間</asp:ListItem>
					</asp:dropdownlist><asp:dropdownlist id="DApplyName" style="Z-INDEX: 102; POSITION: absolute; TOP: 8px; LEFT: 472px"
						runat="server" ForeColor="Blue" BackColor="Yellow"  Width="100px"></asp:dropdownlist><asp:dropdownlist id="DFlowType" style="Z-INDEX: 101; POSITION: absolute; TOP: 8px; LEFT: 368px" runat="server"
						ForeColor="Blue" BackColor="Yellow"  Width="100px">
						<asp:ListItem Value="ALL" Selected="True">ALL</asp:ListItem>
						<asp:ListItem Value="0">通知</asp:ListItem>
						<asp:ListItem Value="1">核定</asp:ListItem>
					</asp:dropdownlist></FONT><asp:datagrid id="DataGrid1" style="Z-INDEX: 110; POSITION: absolute; TOP: 72px; LEFT: 8px" runat="server"
					BackColor="White" Height="100px" Width="1080px" AutoGenerateColumns="False" CellPadding="4" BorderWidth="1px" BorderColor="#CC9966"
					AllowPaging="True" BorderStyle="None" Font-Size="9pt" PageSize="15">
					<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
					<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
					<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
					<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
					<Columns>
						<asp:BoundColumn DataField="StsDesc" ReadOnly="True" HeaderText="狀態">
							<HeaderStyle Width="48px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="URL" DataNavigateUrlFormatString="{0}" DataTextField="FormName"
							HeaderText="委託單">
							<HeaderStyle Width="84px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:HyperLinkColumn>
						<asp:BoundColumn DataField="No" ReadOnly="True" HeaderText="委託No">
							<HeaderStyle Width="80px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="MapNo" ReadOnly="True" HeaderText="委託部門/圖號/客訴內容">
							<HeaderStyle Width="130px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="ApplyName" ReadOnly="True" HeaderText="委託人">
							<HeaderStyle Width="60px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="StepNameDesc" ReadOnly="True" HeaderText="工程/請假起迄">
							<HeaderStyle Width="180px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="FlowTypeDesc" ReadOnly="True" HeaderText="類別">
							<HeaderStyle Width="60px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="AgentName" ReadOnly="True" HeaderText="代理/兼職">
							<HeaderStyle Width="60px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="Description" ReadOnly="True" HeaderText="參考資料">
							<HeaderStyle Width="464px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
					</Columns>
					<PagerStyle NextPageText="下一頁" PrevPageText="上一頁" HorizontalAlign="Center" ForeColor="#993300"
						BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
				</asp:datagrid></FONT>
        <asp:HyperLink ID="LWaitPage" runat="server" Font-Size="12pt" Height="8px" NavigateUrl="http://10.245.1.10/Portal"
            Style="z-index: 100; left: 600px; position: absolute; top: 40px" Target="_self"
            Width="66px">回待處理</asp:HyperLink>
    </div>
    </form>
</body>
</html>
