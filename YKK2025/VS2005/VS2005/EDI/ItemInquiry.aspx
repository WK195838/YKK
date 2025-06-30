<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ItemInquiry.aspx.vb" Inherits="ItemInquiry" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>ITEM-查詢</title>
</head>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體"></FONT>
			<DIV title="篩選項目：" style="DISPLAY: inline; Z-INDEX: 100; LEFT: 7px; WIDTH: 44px; COLOR: #0000ff; POSITION: absolute; TOP: 34px; HEIGHT: 24px"
				ms_positioning="FlowLayout">
                Code</DIV>
			<asp:imagebutton id="BExcel" style="Z-INDEX: 122; LEFT: 337px; POSITION: absolute; TOP: 65px" runat="server"
				ImageUrl="Images\msexcel.gif" Height="21px" Width="21px"></asp:imagebutton>
				<asp:textbox id="DCode" style="Z-INDEX: 116; LEFT: 52px; POSITION: absolute; TOP: 35px" runat="server"
				Height="20px" Width="140px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove"></asp:textbox>
			<asp:button id="Go" style="Z-INDEX: 110; LEFT: 288px; POSITION: absolute; TOP: 63px" runat="server"
				Height="24px" Width="40px" ForeColor="Blue" BackColor="White" Text="Go"></asp:button>
            <asp:TextBox ID="DColor" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="20px" Style="z-index: 111;
                left: 341px; position: absolute; top: 35px" Width="140px"></asp:TextBox>
            <DIV title="篩選項目：" style="DISPLAY: inline; Z-INDEX: 100; LEFT: 247px; WIDTH: 92px; COLOR: #0000ff; POSITION: absolute; TOP: 36px; HEIGHT: 24px"
				ms_positioning="FlowLayout">
                Color1 / 2 / 3</div>
            <DIV title="篩選項目：" style="DISPLAY: inline; Z-INDEX: 100; LEFT: 7px; WIDTH: 44px; COLOR: #0000ff; POSITION: absolute; TOP: 7px; HEIGHT: 24px"
				ms_positioning="FlowLayout">
                Buyer</div>
            <asp:DropDownList ID="DBuyer" runat="server" BackColor="Yellow"
                Font-Size="10pt" ForeColor="Blue" Height="26px" Style="z-index: 103; left: 52px;
                position: absolute; top: 8px" Width="145px" AutoPostBack="True">
            </asp:DropDownList>
                
            <asp:DropDownList ID="DAction" runat="server" BackColor="Yellow"
                Font-Size="10pt" ForeColor="Blue" Height="26px" Style="z-index: 103; left: 200px;
                position: absolute; top: 37px" Width="44px">
                <asp:ListItem Selected="True">0</asp:ListItem>
                <asp:ListItem>1</asp:ListItem>
                <asp:ListItem>2</asp:ListItem>
                <asp:ListItem>3</asp:ListItem>
            </asp:DropDownList>
                
			<DIV title="抽出期間：" style="DISPLAY: inline; Z-INDEX: 100; LEFT: 7px; WIDTH: 74px; COLOR: #0000ff; POSITION: absolute; TOP: 64px; HEIGHT: 24px"
				ms_positioning="FlowLayout">
                抽出期間</DIV>
            <asp:textbox id="DStartTime" style="Z-INDEX: 116; LEFT: 84px; POSITION: absolute; TOP: 61px" runat="server"
				Height="20px" Width="80px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">2013/6/1</asp:textbox>
			<DIV title="抽出期間：" style="DISPLAY: inline; Z-INDEX: 100; LEFT: 172px; WIDTH: 23px; COLOR: #0000ff; POSITION: absolute; TOP: 64px; HEIGHT: 24px"
				ms_positioning="FlowLayout">
                ～</DIV>
				
            <asp:textbox id="DEndTime" style="Z-INDEX: 116; LEFT: 198px; POSITION: absolute; TOP: 62px" runat="server"
				Height="20px" Width="80px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove">2013/7/21</asp:textbox>
                
            <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" BorderColor="#CC9966"
                BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt" 
                Style="z-index: 114; left: 8px; position: absolute; top: 93px" Width="1000px">
                <Columns>
                    <asp:BoundField DataField="BuyerDesc" HeaderText="Buyer" ReadOnly="True" />
                    <asp:BoundField DataField="Action" HeaderText="Action" ReadOnly="True" />
                    <asp:BoundField DataField="Code" HeaderText="Code" ReadOnly="True" />
                    <asp:BoundField DataField="Description" HeaderText="Description" ReadOnly="True" />
                    <asp:BoundField DataField="Color1" HeaderText="Color-1" ReadOnly="True" />
                    <asp:BoundField DataField="Color2" HeaderText="Color-2" ReadOnly="True" />
                    <asp:BoundField DataField="Color3" HeaderText="Color-3" ReadOnly="True" />
                    <asp:BoundField DataField="SliderStatus" HeaderText="其他條件" ReadOnly="True" />
                    <asp:BoundField DataField="YCode" HeaderText="YKK-ITEM" ReadOnly="True" />
                    <asp:BoundField DataField="YName" HeaderText="YKK-NAME" ReadOnly="True" />
                </Columns>
                <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:GridView>
        </form>
	</body>
</html>
