<%@ Page Language="VB" AutoEventWireup="false" CodeFile="YKKGroupCopyList.aspx.vb" Inherits="YKKGroupCopyList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>姊妹社圖面複製履歷List</title>
</head>
<body>
    <form id="form1" runat="server">
        <asp:GridView ID="GridView1" runat="server" AllowSorting="True" AutoGenerateColumns="False"
            BackColor="White" BorderColor="#CC9966" BorderStyle="None" BorderWidth="1px"
            CellPadding="4" EmptyDataText="查無資料" Font-Size="Small" Height="34px" PageSize="15"
            Width="642px">
            <RowStyle BackColor="White" ForeColor="#330099" HorizontalAlign="Center" />
            <Columns>
                <asp:BoundField DataField="StsDesc" HeaderText="狀態" />
                <asp:BoundField DataField="No" HeaderText="No" />
                <asp:BoundField DataField="DateDesc" HeaderText="日期" />
                <asp:BoundField DataField="SliderCode" HeaderText="拉頭品名" />
                <asp:BoundField DataField="MapNo" HeaderText="圖號" />
                <asp:BoundField DataField="OFormDesc" HeaderText="原委託" />
                <asp:BoundField DataField="FormSno" HeaderText="FormSno" />
                <asp:BoundField DataField="YKKGroup" HeaderText="YKKGroup" />
                <asp:BoundField DataField="BUYER" HeaderText="BUYER" />
            </Columns>
            <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
            <PagerStyle BackColor="#FFFFCC" ForeColor="#330099" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#FFCC66" Font-Bold="True" ForeColor="#663399" />
            <HeaderStyle BackColor="#990000" BorderStyle="None" Font-Bold="True" ForeColor="#FFFFCC" Height="50px" />
        </asp:GridView>
    </form>
</body>
</html>
