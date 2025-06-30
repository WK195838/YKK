<%@ Page Language="VB" AutoEventWireup="false" CodeFile="InfCustomerOrder.aspx.vb" Inherits="InfCustomerOrder" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>Customer Order</title>
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
        &nbsp;&nbsp;

            <asp:GridView style="Z-INDEX: 114; LEFT: 8px; POSITION: absolute; TOP: 42px" id="GridView1" runat="server" BorderStyle="Groove" Font-Size="9pt" CellPadding="4" BorderWidth="1px" BorderColor="Black" AutoGenerateColumns="False" PageSize="20">
                <Columns>
                    <asp:BoundField DataField="PO" HeaderText="PO" ></asp:BoundField>
                    <asp:BoundField DataField="SeqNo" HeaderText="SeqNo" ></asp:BoundField>
                    <asp:BoundField DataField="A1" HeaderText="Inf-1" ></asp:BoundField>
                    <asp:BoundField DataField="B1" HeaderText="Inf-2" ></asp:BoundField>
                    <asp:BoundField DataField="C1" HeaderText="Inf-3" ></asp:BoundField>
                    <asp:BoundField DataField="D1" HeaderText="Inf-4" ></asp:BoundField>
                    <asp:BoundField DataField="E1" HeaderText="Inf-5" ></asp:BoundField>
                    <asp:BoundField DataField="F1" HeaderText="Inf-6" ></asp:BoundField>
                    <asp:BoundField DataField="G1" HeaderText="Inf-7" />
                    <asp:BoundField DataField="H1" HeaderText="Inf-8" />  
                    <asp:BoundField DataField="I1" HeaderText="Inf-9" />                    
                    <asp:BoundField DataField="J1" HeaderText="Inf-10" ></asp:BoundField>
                    <asp:BoundField DataField="K1" HeaderText="Inf-11" ></asp:BoundField>
                    <asp:BoundField DataField="L1" HeaderText="Inf-12" ></asp:BoundField>
                    <asp:BoundField DataField="M1" HeaderText="Inf-13" ></asp:BoundField>
                    <asp:BoundField DataField="N1" HeaderText="Inf-14" ></asp:BoundField>
                    <asp:BoundField DataField="O1" HeaderText="Inf-15" ></asp:BoundField>
                    <asp:BoundField DataField="P1" HeaderText="Inf-16" ></asp:BoundField>
                    <asp:BoundField DataField="Q1" HeaderText="Inf-17" ></asp:BoundField>
                    <asp:BoundField DataField="R1" HeaderText="Inf-18" ></asp:BoundField>
                    <asp:BoundField DataField="S1" HeaderText="Inf-19" ></asp:BoundField>
                    <asp:BoundField DataField="T1" HeaderText="Inf-20" ></asp:BoundField>
                    <asp:BoundField DataField="U1" HeaderText="Inf-21" ></asp:BoundField>
                    <asp:BoundField DataField="V1" HeaderText="Inf-22" ></asp:BoundField>
                    <asp:BoundField DataField="W1" HeaderText="Inf-23" ></asp:BoundField>
                    <asp:BoundField DataField="X1" HeaderText="Inf-24" ></asp:BoundField>
                    <asp:BoundField DataField="Y1" HeaderText="Inf-25" ></asp:BoundField>
                    <asp:BoundField DataField="Z1" HeaderText="Inf-26" ></asp:BoundField>
                </Columns>
                <HeaderStyle BackColor="Silver" ForeColor="Black" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="False"  />
            </asp:GridView>  

        <asp:TextBox ID="DInputDate" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="20px" Style="z-index: 103; left: 310px; position: absolute;
            top: 10px" Width="80px"></asp:TextBox>
        &nbsp;
    </div>
    </form>
</body>
</html>
