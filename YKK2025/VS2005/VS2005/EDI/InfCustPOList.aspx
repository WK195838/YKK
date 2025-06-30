<%@ Page Language="VB" AutoEventWireup="false" CodeFile="InfCustPOList.aspx.vb" Inherits="InfCustPOList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>系統履歷-訂單</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" BorderColor="#CC9966"
            BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt"
            Style="z-index: 114; left: 8px; position: absolute; top: 2px" Width="1200px">
            <Columns>
                    <asp:BoundField DataField="PO" HeaderText="PO No." ></asp:BoundField>
                    <asp:BoundField DataField="Seq" HeaderText="Seq No" ></asp:BoundField>
                    <asp:BoundField DataField="OrderInf" HeaderText="客戶訂單 Inf." ></asp:BoundField>
            </Columns>
            <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
        </asp:GridView>
    </div>
    </form>
</body>
</html>
