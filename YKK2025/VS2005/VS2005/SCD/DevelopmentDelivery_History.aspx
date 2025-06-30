<%@ Page Language="VB" AutoEventWireup="false" CodeFile="DevelopmentDelivery_History.aspx.vb" Inherits="DevelopmentDelivery_History" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>開發履歷</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
            <asp:GridView style="Z-INDEX: 136; LEFT: 10px; POSITION: absolute; top: 11px" id="GridView1" runat="server" Width="800px" Height="100px" BorderStyle="Groove" BackColor="White" BorderColor="#CC9966" CellPadding="4" BorderWidth="1px" AutoGenerateColumns="False">

            <RowStyle BackColor="White" Font-Size="9pt" ForeColor="#330099"  />
            <Columns>
                <asp:BoundField DataField="STSD" />
                <asp:BoundField DataField="NO" HeaderText="NO." />
                <asp:BoundField DataField="StepNameDesc" HeaderText="工程">
                    <ItemStyle HorizontalAlign="Left"  />
                </asp:BoundField>
                <asp:BoundField DataField="DecideName" HeaderText="擔當" ></asp:BoundField>
                <asp:BoundField DataField="Description" HeaderText="核定時間">
                    <ItemStyle HorizontalAlign="Left"  />
                </asp:BoundField>
            </Columns>
            <HeaderStyle BackColor="#CC6600" Font-Size="9pt" ForeColor="#FFFFCC" HorizontalAlign="Center"
                VerticalAlign="Middle"  />
        </asp:GridView>                
   
    </div>
    </form>
</body>
</html>