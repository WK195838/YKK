<%@ Page Language="VB" AutoEventWireup="false" CodeFile="StockList.aspx.vb" Inherits="StockList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>在庫資料</title>
</head>
<body>
    <form id="form1" runat="server">
        尋找<asp:TextBox ID="DData" runat="server" style="z-index: 100; left: 120px; position: absolute;;text-transform : uppercase; top: 16px" Height="17px" Width="64px"></asp:TextBox><asp:Button ID="Button1" runat="server" Text="查詢" style="z-index: 101; left: 344px; position: absolute; top: 16px" />
        <asp:Button ID="BAdd" runat="server" Text="加入" style="z-index: 102; left: 456px; position: absolute; top: 16px" Visible="False" />
        ItemCode &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
        Color
        <asp:TextBox ID="DColor" runat="server" Height="17px" Style="z-index: 103; left: 248px;
            text-transform: uppercase; position: absolute; top: 16px" Width="64px"></asp:TextBox>
        <asp:Label ID="Label1" runat="server" ForeColor="Red" Style="z-index: 105; left: 504px;
            position: absolute; top: 24px"></asp:Label>
        <br />
        <br />
        <asp:Panel ID="Panel1" runat="server"  Height="688px" Width="760px" ScrollBars="Auto" BorderStyle="Solid" BorderWidth="1px">    &nbsp;
        
        <asp:GridView id="GridView1" runat="server" Height="32px"  CellPadding="4"  EmptyDataText="查無資料" Width="720px" Font-Size="Small" AutoGenerateColumns="False" AllowSorting="True" PageSize="15" DataKeyNames="ITMCXA,CLRCXA" Visible="False">
<RowStyle HorizontalAlign="Center" BackColor="#FFFF99"></RowStyle>
         
<Columns>
    <asp:CommandField ShowHeader="True" ShowSelectButton="True" />
<asp:BoundField DataField="ITMCXA" HeaderText="ITEM CODE"></asp:BoundField>
    <asp:BoundField HeaderText="Color" DataField="CLRCXA" />
    <asp:BoundField DataField="Type" HeaderText="Type" />
</Columns>
                <HeaderStyle BackColor="#FFCC99" />
                <AlternatingRowStyle BackColor="#CCFFCC" />
            </asp:GridView>
            <asp:GridView id="GridView2" runat="server" Height="32px"  CellPadding="4"  EmptyDataText="查無資料" Width="720px" Font-Size="Small" AutoGenerateColumns="False" AllowSorting="True" PageSize="15" DataKeyNames="WSHCXA" Visible="False">
                <RowStyle BackColor="#FFFF99" HorizontalAlign="Center" />
                <Columns>
    <asp:TemplateField HeaderText="複選">
        <ItemTemplate>
            &nbsp;
            <asp:CheckBox ID="CheckNo" runat="server" />
        </ItemTemplate>
    </asp:TemplateField>
<asp:BoundField DataField="ITMCXA" HeaderText="ITEM CODE"></asp:BoundField>
<asp:BoundField DataField="CLRCXA" HeaderText="COLOR"></asp:BoundField>
<asp:BoundField DataField="QUNCXA" HeaderText="UNIT"></asp:BoundField>
<asp:BoundField DataField="WQTYXA" HeaderText="入庫數量" DataFormatString="{0:N0}"></asp:BoundField>
    <asp:BoundField DataField="RCD1XA" HeaderText="入庫日期" />
    <asp:BoundField DataField="RCT1XA" HeaderText="入庫時間" />
    <asp:BoundField DataField="WSHCXA" HeaderText="棧板號" />
    <asp:BoundField DataField="REM1XA" HeaderText="庫位" />
</Columns>

<HeaderStyle BackColor="#FFCC99"></HeaderStyle>

<AlternatingRowStyle BackColor="#CCFFCC"></AlternatingRowStyle>
</asp:GridView>
 </asp:Panel> &nbsp; &nbsp; &nbsp;
        &nbsp; &nbsp;&nbsp;<br />
        &nbsp;&nbsp;<br />
        <br />
        &nbsp;
            <br />
    </form>
</body>
</html>
