<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MPM_InqCommission.aspx.vb" Inherits="MPM_InqCommission" %>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<HTML>
	<HEAD>
		<title>調閱資料</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<script language="javascript" type="text/javascript">
			//--Calendar------------------------------------
			function calendarPicker(field1)
			{
				window.open('DatePicker.aspx?field1=' + field1,'calendarPopup','width=250,height=190,resizable=yes');
			}
		</script>
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體"></FONT>
			<DIV title="篩選項目：" style="DISPLAY: inline; Z-INDEX: 118; LEFT: 3px; WIDTH: 88px; COLOR: #0000ff; POSITION: absolute; TOP: 8px; HEIGHT: 24px"
				ms_positioning="FlowLayout">篩選項目：</DIV>
			<asp:imagebutton id="BExcel" style="Z-INDEX: 100; LEFT: 429px; POSITION: absolute; TOP: 105px" runat="server"
				ImageUrl="Images\msexcel.gif" Height="21px" Width="21px"></asp:imagebutton><asp:textbox id="DNo" style="Z-INDEX: 103; LEFT: 96px; POSITION: absolute; TOP: 40px" runat="server"
				Height="20px" Width="120px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">委託單No.</asp:textbox><asp:datagrid id="DataGrid1" style="Z-INDEX: 104; LEFT: 10px; POSITION: absolute; TOP: 147px" runat="server"
				Height="100px" Width="780px" BackColor="White" BorderStyle="None" AutoGenerateColumns="False" CellPadding="4" BorderWidth="1px" BorderColor="#CC9966" Font-Size="9pt" AllowPaging="True" PageSize="15">
				<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
				<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
				<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
				<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
				<Columns>
					<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="ViewURL" DataNavigateUrlFormatString="{0}"
						DataTextField="Field1">
						<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:HyperLinkColumn>
					<asp:BoundColumn DataField="Field2" ReadOnly="True" HeaderText="狀態">
						<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="Field3" ReadOnly="True" HeaderText="依賴日">
						<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="Field4" ReadOnly="True" HeaderText="工程">
						<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="Field5" ReadOnly="True" HeaderText="工程擔當">
						<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="Field6" ReadOnly="True" HeaderText="開始預定">
						<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="Field7" ReadOnly="True" HeaderText="完成預定">
						<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="Field8" ReadOnly="True" HeaderText="完成預定">
						<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
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
			</asp:datagrid><asp:textbox id="DMapno" style="Z-INDEX: 105; LEFT: 352px; POSITION: absolute; TOP: 40px" runat="server"
				Height="20px" Width="114px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">圖號</asp:textbox><asp:button id="BEDate" style="Z-INDEX: 108; LEFT: 346px; POSITION: absolute; TOP: 100px" runat="server"
				Height="25px" Width="28px" Text="....."></asp:button><asp:textbox id="DEDate" style="Z-INDEX: 100; LEFT: 249px; POSITION: absolute; TOP: 105px" runat="server"
				Height="20px" Width="88px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">DEDate</asp:textbox>
			<DIV title="" style="DISPLAY: inline; Z-INDEX: 119; LEFT: 228px; WIDTH: 16px; COLOR: #0000ff; POSITION: absolute; TOP: 103px; HEIGHT: 24px"
				ms_positioning="FlowLayout">～</DIV>
			<asp:button id="BSDate" style="Z-INDEX: 110; LEFT: 198px; POSITION: absolute; TOP: 104px" runat="server"
				Height="26px" Width="28px" Text="....."></asp:button><asp:textbox id="DSDate" style="Z-INDEX: 111; LEFT: 95px; POSITION: absolute; TOP: 105px" runat="server"
				Height="20px" Width="93px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">DSDate</asp:textbox>
			<DIV title="申請日期：" style="DISPLAY: inline; Z-INDEX: 120; LEFT: 4px; WIDTH: 88px; COLOR: #0000ff; POSITION: absolute; TOP: 106px; HEIGHT: 24px"
				ms_positioning="FlowLayout">申請日期：</DIV>
			<asp:button id="Go" style="Z-INDEX: 112; LEFT: 382px; POSITION: absolute; TOP: 103px" runat="server"
				Height="24px" Width="40px" ForeColor="Blue" BackColor="White" Text="Go"></asp:button>
            &nbsp;
            <asp:TextBox ID="DName" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="20px" Style="z-index: 113; left: 607px;
                position: absolute; top: 40px" Width="114px">擔當者</asp:TextBox>
            &nbsp;
            <asp:DropDownList ID="DType1" runat="server" BackColor="Yellow" Style="z-index: 115;
                left: 97px; position: absolute; top: 74px" Width="121px">
            </asp:DropDownList><asp:DropDownList ID="DType2" runat="server" BackColor="Yellow" Style="z-index: 116;
                left: 229px; position: absolute; top: 74px" Width="115px">
            </asp:DropDownList>
            <asp:DropDownList ID="DFormName" runat="server" BackColor="Yellow" Style="z-index: 121;
                left: 97px; position: absolute; top: 8px" Width="248px">
            </asp:DropDownList><asp:DropDownList ID="DProgress" runat="server" BackColor="Yellow" Style="z-index: 121;
                left: 354px; position: absolute; top: 6px" Width="119px">
                <asp:ListItem>ALL</asp:ListItem>
                <asp:ListItem Value="1">開發中</asp:ListItem>
                <asp:ListItem Value="2">開發完成</asp:ListItem>
            </asp:DropDownList><asp:DropDownList ID="DSts" runat="server" BackColor="Yellow" Style="z-index: 121;
                left: 482px; position: absolute; top: 8px" Width="119px">
                <asp:ListItem>ALL</asp:ListItem>
                <asp:ListItem Value=" 1">OK</asp:ListItem>
                <asp:ListItem Value="2">NG</asp:ListItem>
                <asp:ListItem Value="3">取消</asp:ListItem>
            </asp:DropDownList><asp:DropDownList ID="DDivision" runat="server" BackColor="Yellow" Style="z-index: 121;
                left: 227px; position: absolute; top: 41px" Width="119px">
            </asp:DropDownList><asp:DropDownList ID="DEngine" runat="server" BackColor="Yellow" Style="z-index: 121;
                left: 481px; position: absolute; top: 39px" Width="119px">
            </asp:DropDownList>
        </form>
	</body>
</HTML>
