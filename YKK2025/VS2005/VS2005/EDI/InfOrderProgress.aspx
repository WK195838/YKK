<%@ Page Language="VB" AutoEventWireup="false" CodeFile="InfOrderProgress.aspx.vb" Inherits="InfOrderProgress" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Order Progress</title>
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

            <asp:GridView style="Z-INDEX: 114; LEFT: 8px; POSITION: absolute; TOP: 69px" id="GridView1" runat="server" BorderStyle="Groove" Font-Size="9pt" CellPadding="4" BorderWidth="1px" BorderColor="Black" AutoGenerateColumns="False" PageSize="20">
                <Columns>
                    <asp:BoundField DataField="PO" HeaderText="PO" ></asp:BoundField>
                    <asp:BoundField DataField="SeqNo" HeaderText="SeqNo" ></asp:BoundField>
                    <asp:BoundField DataField="OrderNo" HeaderText="OrderNo" ></asp:BoundField>
                    <asp:BoundField DataField="OrderSubNo" HeaderText="SubNo" ></asp:BoundField>
                    <asp:BoundField DataField="Item" HeaderText="Item" ></asp:BoundField>

                    <asp:BoundField DataField="Length" HeaderText="Length" ></asp:BoundField>
                    <asp:BoundField DataField="Unit" HeaderText="Unit" ></asp:BoundField>
                    <asp:BoundField DataField="Color" HeaderText="Color" ></asp:BoundField>
                    <asp:BoundField DataField="OrderDate" HeaderText="Order Date" />
                    <asp:BoundField DataField="ReqDate" HeaderText="Request Date" />  
                                      
                    <asp:BoundField DataField="PlanDate" HeaderText="Plan Date" />                    
                    <asp:BoundField DataField="PriceVersion" HeaderText="Price Version" ></asp:BoundField>
                    <asp:BoundField DataField="OrderQty" HeaderText="Order Qty" ></asp:BoundField>
                    <asp:BoundField DataField="SalesPrice" HeaderText="Price" ></asp:BoundField>
                    <asp:BoundField DataField="SalesAmount" HeaderText="Amount" ></asp:BoundField>

                </Columns>
                <HeaderStyle BackColor="Silver" ForeColor="Black" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="False"  />
            </asp:GridView>    <asp:Button ID="BRefresh" runat="server" BackColor="White" ForeColor="Blue" Height="24px"
            Style="z-index: 110; left: 471px; position: absolute; top: 11px" Text="更新" Width="40px" />
        <asp:TextBox ID="DLastUpdate" runat="server" BackColor="White" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 103; left: 751px; position: absolute;
            top: 27px" Width="191px"></asp:TextBox>
        <asp:TextBox ID="TextBox1" runat="server" BackColor="White" BorderStyle="None" ForeColor="Black"
            Height="20px" Style="z-index: 103; left: 533px; position: absolute; top: 5px"
            Width="213px">[●]Completed Date＞Request Date</asp:TextBox>

        <asp:HyperLink ID="LHelpPage" runat="server" ForeColor="Navy" Height="20px" NavigateUrl="~/images/OrderProgess.pdf"
            Style="z-index: 1; left: 854px; position: absolute; top: 5px" Target="_blank"
            Width="74px">使用說明</asp:HyperLink>

        <asp:HyperLink ID="LCustomerOrder" runat="server" ForeColor="Navy" Height="20px" NavigateUrl="InfCustomerOrder.aspx"
            Style="z-index: 1; left: 10px; position: absolute; top: 42px" Target="_blank"
            Width="107px">客戶訂單資料</asp:HyperLink>

        <asp:TextBox ID="DInputDate" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="20px" Style="z-index: 103; left: 338px; position: absolute;
            top: 10px" Width="80px"></asp:TextBox>
        <asp:TextBox ID="TextBox2" runat="server" BackColor="White" BorderStyle="None" ForeColor="Black"
            Height="20px" Style="z-index: 103; left: 533px; position: absolute; top: 25px"
            Width="213px">[★]Completed Date＞Plan Date                         </asp:TextBox>
    </div>
    </form>
</body>
</html>
