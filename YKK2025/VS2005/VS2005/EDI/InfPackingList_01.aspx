<%@ Page Language="VB" AutoEventWireup="false" CodeFile="InfPackingList_01.aspx.vb" Inherits="InfPackingList_01" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Packing List</title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <asp:TextBox ID="DBuyer" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
                Height="20px" Style="z-index: 103; left: 37px; position: absolute; top: 10px"
                Width="120px"></asp:TextBox>
            <asp:ImageButton ID="BExcel" runat="server" Height="21px" ImageUrl="Images\msexcel.gif"
                Style="z-index: 103; left: 9px; position: absolute; top: 9px" Width="21px" />
            <asp:TextBox ID="DPO" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
                Height="20px" Style="z-index: 103; left: 163px; position: absolute; top: 10px"
                Width="140px"></asp:TextBox>
            &nbsp;&nbsp;
            <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" BorderColor="Black"
                BorderStyle="Groove" BorderWidth="1px" CellPadding="4" Font-Size="9pt" PageSize="20"
                Style="z-index: 114; left: 8px; position: absolute; top: 52px">
                <Columns>
                    <asp:BoundField DataField="PO" HeaderText="PO" />
                    <asp:BoundField DataField="SeqNo" HeaderText="SeqNo"><ItemStyle HorizontalAlign="Center" /></asp:BoundField>
                    <asp:BoundField DataField="Delivery" HeaderText="Delivery"><ItemStyle HorizontalAlign="Center" /></asp:BoundField>
                    <asp:BoundField DataField="OrderNo" HeaderText="OrderNo" />
                    <asp:BoundField DataField="OrderSubNo" HeaderText="SubNo" ><ItemStyle HorizontalAlign="Center" /></asp:BoundField>
                    <asp:BoundField DataField="Item" HeaderText="Item" />
                    <asp:BoundField DataField="ItemName" HeaderText="ItemName" />
                    <asp:BoundField DataField="CaseNo" HeaderText="Case" ><ItemStyle HorizontalAlign="Center" /></asp:BoundField>
                    <asp:BoundField DataField="Length" HeaderText="Length" ><ItemStyle HorizontalAlign="Center" /></asp:BoundField>
                    <asp:BoundField DataField="Color" HeaderText="Color" />
                    <asp:BoundField DataField="PackQty" HeaderText="Qty" ><ItemStyle HorizontalAlign="right" /></asp:BoundField>
                    <asp:BoundField DataField="Count" HeaderText="Count" ><ItemStyle HorizontalAlign="Center" /></asp:BoundField>
                    <asp:BoundField DataField="ItemNet" HeaderText="淨重(kg)" ><ItemStyle HorizontalAlign="right" /></asp:BoundField>
                    <asp:BoundField DataField="OuterNet" HeaderText="外箱淨重(kg)" ><ItemStyle HorizontalAlign="right" /></asp:BoundField>
                    <asp:BoundField DataField="Gross" HeaderText="毛重(kg)"><ItemStyle HorizontalAlign="right" /></asp:BoundField>
                    
                    <asp:BoundField DataField="EndFlag" HeaderText="EndFlag" />
                    
                </Columns>
                <HeaderStyle BackColor="Silver" Font-Bold="False" ForeColor="Black" HorizontalAlign="Center"
                    VerticalAlign="Middle" />
            </asp:GridView>
            <asp:Button ID="BRefresh" runat="server" BackColor="White" ForeColor="Blue" Height="24px"
                Style="z-index: 110; left: 464px; position: absolute; top: 11px" Text="更新" Width="40px" />
            <asp:TextBox ID="DLastUpdate" runat="server" BackColor="White" BorderStyle="None"
                ForeColor="Black" Height="20px" Style="z-index: 103; left: 797px; position: absolute;
                top: 27px" Width="191px"></asp:TextBox>
            <asp:TextBox ID="TextBox1" runat="server" BackColor="White" BorderStyle="None" ForeColor="Black"
                Height="20px" Style="z-index: 103; left: 797px; position: absolute; top: 5px"
                Width="153px">[★] 未完成 Delivery Note</asp:TextBox>
            &nbsp;&nbsp;&nbsp;
            <asp:TextBox ID="DOrderNo" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="20px" Style="z-index: 103; left: 310px; position: absolute;
                top: 10px" Width="140px"></asp:TextBox>
        </div>
    </form>
</body>
</html>
