<%@ Page Language="VB" AutoEventWireup="false" CodeFile="FCDataList.aspx.vb" Inherits="FCDataList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>細項內容</title>
</head>
<body>
    <form id="form1" runat="server">
        &nbsp;&nbsp;
        <asp:ImageButton ID="BExcel" runat="server" Height="21px" ImageUrl="~/Images/msexcel.gif"
            Style="z-index: 100; left: 13px; position: absolute; top: 18px" Width="21px" />
        <br />
        <asp:GridView id="GridView1" runat="server" Height="34px"  CellPadding="4"  EmptyDataText="查無資料" Width="554px" Font-Size="Small" AutoGenerateColumns="False" AllowSorting="True" PageSize="15" style="z-index: 102; left: 6px; position: absolute; top: 54px">
<RowStyle HorizontalAlign="Center" BackColor="#FFFF99"></RowStyle>

<HeaderStyle BackColor="#FFCC99"></HeaderStyle>

<AlternatingRowStyle BackColor="#CCFFCC"></AlternatingRowStyle>
            <Columns>
                <asp:BoundField DataField="SCheck" HeaderText="狀態" />
                <asp:BoundField DataField="CustomerCode" HeaderText="Customer Code" />
                <asp:BoundField DataField="CustomerName" HeaderText="Customer Name" />
                <asp:BoundField DataField="BUYERCODE" HeaderText="Buyer Code" />
                <asp:BoundField DataField="BUYERNAME" HeaderText="Buyer Name" />
                <asp:BoundField DataField="REQDATE" HeaderText="REQ. Date" />
                <asp:BoundField DataField="KEEPCODE" HeaderText="KEEP CODE" />
                <asp:BoundField DataField="BUYMONTH" HeaderText="BUY MONTH" />
                <asp:BoundField DataField="CORDERNO" HeaderText="C. ORDER NO" />
                <asp:BoundField DataField="SeqNo" HeaderText="*No" />
                <asp:BoundField DataField="AtStock" HeaderText="在庫" />
                <asp:BoundField DataField="ITEMCODE1" HeaderText="ITEM CODE" />
                <asp:BoundField DataField="ITEMNAME1" HeaderText="ITEM NAME" />
                <asp:BoundField DataField="COLOR1" HeaderText="COLOR" />
                <asp:BoundField DataField="LENGTH1" HeaderText="LENGTH" />
                <asp:BoundField DataField="LENGTHU1" HeaderText="LENGTH U." />
                <asp:BoundField DataField="ORDERQTYP1" HeaderText="ORDER QTY(P)" />
                <asp:BoundField DataField="ITEMCODE2" HeaderText="ITEM CODE" />
                <asp:BoundField DataField="ITEMNAME2" HeaderText="ITEM NAME" />
                <asp:BoundField DataField="COLOR2" HeaderText="COLOR" />
                <asp:BoundField DataField="LENGTH2" HeaderText="LENGTH" />
                <asp:BoundField DataField="LENGTHU2" HeaderText="LENGTH U." />
                <asp:BoundField AccessibleHeaderText="ORDERQTYP2" HeaderText="ORDER QTY(P)" DataField="ORDERQTYP2" />
                <asp:BoundField DataField="UnitPrice" DataFormatString="{0:N0}" HeaderText="金額" />
                <asp:BoundField DataField="INVQTY" HeaderText=" KEEP在庫數" DataFormatString="{0:N0}" />
                <asp:BoundField DataField="INVAMT" HeaderText="KEEP在庫金額" DataFormatString="{0:N0}" />
            </Columns>
</asp:GridView>
        <br />
        &nbsp;&nbsp; &nbsp; &nbsp;<br />
        &nbsp;&nbsp;<br />
        <br />
        &nbsp;
            <br />
    </form>
</body>
</html>
