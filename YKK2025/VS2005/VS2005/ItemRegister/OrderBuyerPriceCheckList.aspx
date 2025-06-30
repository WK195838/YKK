<%@ Page Language="VB" AutoEventWireup="false" CodeFile="OrderBuyerPriceCheckList.aspx.vb" Inherits="OrderBuyerPriceCheckList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Order Buyer Price</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" BorderColor="Black"
            BorderStyle="Groove" BorderWidth="1px" CellPadding="4" Font-Size="9pt" PageSize="20"
            Style="z-index: 114; left: 8px; position: absolute; top: 41px">
            <Columns>
                <asp:BoundField DataField="ORDERNO" HeaderText="ORDER-NO" />
                <asp:BoundField DataField="SEQ" HeaderText="SEQ" />
                <asp:BoundField DataField="CUST" HeaderText="CUST" />
                <asp:BoundField DataField="BUYER" HeaderText="BUYER" />
                <asp:BoundField DataField="VERSION" HeaderText="VERSION" />

                <asp:BoundField DataField="ORDERDATE" HeaderText="ORD.DATE" />
                <asp:BoundField DataField="CODE" HeaderText="CODE" />
                <asp:BoundField DataField="LENGTH" HeaderText="LENGTH" />
                <asp:BoundField DataField="U" HeaderText="U" />
                <asp:BoundField DataField="COLOR" HeaderText="COLOR" />

                <asp:BoundField DataField="QUANTITY" HeaderText="QUANTITY" />
                <asp:BoundField DataField="UNIT" HeaderText="UNIT" />
                <asp:BoundField DataField="SALESPRICE" HeaderText="SALES PRICE" />
                <asp:BoundField DataField="CURRENCY" HeaderText="CURRENCY" />
                <asp:BoundField DataField="PRICEA" HeaderText="PRICE A" />

                <asp:BoundField DataField="PRICEB" HeaderText="PRICE B" />
                <asp:BoundField DataField="REGISTER" HeaderText="REGISTER" />
                <asp:BoundField DataField="UPDATE" HeaderText="UPDATE" />

            </Columns>
            <HeaderStyle BackColor="Silver" Font-Bold="False" ForeColor="Black" HorizontalAlign="Center"
                VerticalAlign="Middle" />
        </asp:GridView>
        <asp:TextBox ID="DCustomer" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Height="20px" Style="z-index: 103; left: 37px; position: absolute; top: 10px"
            Width="85px"></asp:TextBox>
        <asp:ImageButton ID="BExcel" runat="server" Height="21px" ImageUrl="Images\msexcel.gif"
            Style="z-index: 103; left: 9px; position: absolute; top: 9px" Width="21px" />
        <asp:TextBox ID="DBuyer" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Height="20px" Style="z-index: 103; left: 129px; position: absolute; top: 10px"
            Width="85px"></asp:TextBox>
        <asp:Button ID="Go" runat="server" BackColor="White" ForeColor="Blue" Height="24px"
            Style="z-index: 110; left: 429px; position: absolute; top: 11px" Text="Go" Width="40px" />
        <asp:TextBox ID="DStartDate" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="20px" Style="z-index: 103; left: 222px; position: absolute;
            top: 10px" Width="85px">Start Date</asp:TextBox>
        <asp:TextBox ID="DEndDate" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="20px" Style="z-index: 103; left: 335px; position: absolute;
            top: 10px" Width="85px">End Date</asp:TextBox>
        <div ms_positioning="FlowLayout" style="display: inline; z-index: 106; left: 315px;
            width: 16px; color: #0000ff; position: absolute; top: 14px; height: 24px" title="">
            ～</div>
        <asp:HyperLink ID="LHelpPage" runat="server" ForeColor="Navy" Height="20px" NavigateUrl="~/images/OrderBuyerPrice.pdf"
            Style="z-index: 1; left: 490px; position: absolute; top: 14px" Target="_blank"
            Width="74px">使用說明</asp:HyperLink>
    
    </div>
    </form>
</body>
</html>
