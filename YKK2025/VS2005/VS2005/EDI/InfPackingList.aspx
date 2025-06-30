<%@ Page Language="VB" AutoEventWireup="false" CodeFile="InfPackingList.aspx.vb" Inherits="InfPackingList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Packing List</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:TextBox ID="DBuyer" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Height="20px" Style="z-index: 103; left: 37px; position: absolute; top: 10px" Width="120px"></asp:TextBox>
        <asp:ImageButton ID="BExcel" runat="server" Height="21px" ImageUrl="Images\msexcel.gif"
            Style="z-index: 103; left: 9px; position: absolute; top: 9px" Width="21px" />
        <asp:TextBox ID="DPO" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Height="20px" Style="z-index: 103; left: 163px; position: absolute; top: 10px"
            Width="140px"></asp:TextBox>
        <asp:Button ID="Go" runat="server" BackColor="White" ForeColor="Blue" Height="24px"
            Style="z-index: 110; left: 428px; position: absolute; top: 11px" Text="Go" Width="40px" />
        <asp:Button ID="BSelectPO" runat="server" Height="25px" Style="z-index: 111; left: 311px;
            position: absolute; top: 11px" Text="....." Width="25px" />

            <asp:GridView style="Z-INDEX: 114; LEFT: 8px; POSITION: absolute; TOP: 52px" id="GridView1" runat="server" BorderStyle="Groove" Font-Size="9pt" CellPadding="4" BorderWidth="1px" BorderColor="Black" AutoGenerateColumns="False" PageSize="20">
                <Columns>
                    <asp:BoundField DataField="PO" HeaderText="PO" ></asp:BoundField>
                    <asp:BoundField DataField="SeqNo" HeaderText="SeqNo" ></asp:BoundField>
                    <asp:BoundField DataField="Delivery" HeaderText="Delivery" ></asp:BoundField>

                <asp:HyperLinkField DataNavigateUrlFields="OrderURL" DataNavigateUrlFormatString="{0}"
                    DataTextField="OrderNo" HeaderText="OrderNo" Target="_blank">
                </asp:HyperLinkField>

                    <asp:BoundField DataField="OrderSubNo" HeaderText="SubNo" ></asp:BoundField>
                    
                    <asp:BoundField DataField="Item" HeaderText="Item" ></asp:BoundField>
                    <asp:BoundField DataField="ItemName" HeaderText="ItemName" ></asp:BoundField>
                    <asp:BoundField DataField="CaseNo" HeaderText="Case" ></asp:BoundField>
                    <asp:BoundField DataField="Length" HeaderText="Length" ></asp:BoundField>
                    <asp:BoundField DataField="Color" HeaderText="Color" ></asp:BoundField>

                    <asp:BoundField DataField="PackQty" HeaderText="Qty" />
                    <asp:BoundField DataField="Count" HeaderText="Count" />    
                    <asp:BoundField DataField="ItemNet" HeaderText="淨重(kg)" />
                                    
                </Columns>
                <HeaderStyle BackColor="Silver" ForeColor="Black" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="False"  />
            </asp:GridView>    &nbsp;
        <asp:TextBox ID="DLastUpdate" runat="server" BackColor="White" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 103; left: 660px; position: absolute;
            top: 27px" Width="191px"></asp:TextBox>
        <asp:TextBox ID="TextBox1" runat="server" BackColor="White" BorderStyle="None" ForeColor="Black"
            Height="20px" Style="z-index: 103; left: 479px; position: absolute; top: 14px"
            Width="153px">[★] 未完成 Delivery Note</asp:TextBox>

        <asp:HyperLink ID="LHelpPage" runat="server" ForeColor="Navy" Height="20px" NavigateUrl="~/images/PackingList.pdf"
            Style="z-index: 1; left: 763px; position: absolute; top: 5px" Target="_blank"
            Width="74px">使用說明</asp:HyperLink>
        <asp:TextBox ID="DInputDate" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="20px" Style="z-index: 103; left: 338px; position: absolute;
            top: 10px" Width="80px"></asp:TextBox>
    </div>
    </form>
</body>
</html>
