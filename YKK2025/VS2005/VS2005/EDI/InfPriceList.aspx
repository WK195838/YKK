<%@ Page Language="VB" AutoEventWireup="false" CodeFile="InfPriceList.aspx.vb" Inherits="InfPriceList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Inf. PriceList</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
            <asp:GridView style="Z-INDEX: 114; LEFT: 8px; POSITION: absolute; TOP: 40px" id="GridView1" runat="server" BorderStyle="Groove" Font-Size="9pt" CellPadding="4" BorderWidth="1px" BorderColor="Black" AutoGenerateColumns="False">
                <Columns>
                    <asp:BoundField DataField="OrderNo" HeaderText="OrderNo" ></asp:BoundField>
                    <asp:BoundField DataField="CustBuyer" HeaderText="Customer/Buyer" ></asp:BoundField>
                    <asp:BoundField DataField="Item" HeaderText="Item" ></asp:BoundField>
                    <asp:BoundField DataField="Color" HeaderText="Color" ></asp:BoundField>

                    <asp:BoundField DataField="PriceVersion" HeaderText="單價版本" ></asp:BoundField>
                    <asp:BoundField DataField="PriceList" HeaderText="單價類型" ></asp:BoundField>
                    <asp:BoundField DataField="SalesPrice" HeaderText="單價" ></asp:BoundField>
                    <asp:BoundField DataField="RegisterTime" HeaderText="最終更新時間" ></asp:BoundField>
                    <asp:BoundField DataField="PriceListRemark" HeaderText="備註" ></asp:BoundField>

                </Columns>
                <HeaderStyle BackColor="Silver" ForeColor="Black" HorizontalAlign="Center" VerticalAlign="Middle"  />
            </asp:GridView>    
        <asp:TextBox ID="DBuyer" runat="server" BackColor="White" BorderStyle="None" ForeColor="Blue"
            Height="20px" Style="z-index: 103; left: 37px; position: absolute; top: 10px" Width="322px"></asp:TextBox>
        <asp:ImageButton ID="BExcel" runat="server" Height="21px" ImageUrl="Images\msexcel.gif"
            Style="z-index: 103; left: 9px; position: absolute; top: 9px" Width="21px" />
        &nbsp;
    </div>
    </form>
</body>
</html>
