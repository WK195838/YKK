<%@ Page Language="VB" Debug="true" CodeFile="FASSheet_02.aspx.vb" Inherits="FASSheet_02" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>備料申請單</title>
	 <script language="javascript"> 

    </script>
 
	 	
</head>
<body>
    <form id="form1" runat="server">
    <div>
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        <asp:Image ID="Image1" runat="server" ImageUrl="~/images/FASSheet_01.jpg" Style="z-index: 100;
            left: 5px; position: absolute; top: 7px" />
        <asp:TextBox ID="DTYPE" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="100" Style="z-index: 101; left: 218px;
            position: absolute; top: 118px; text-align: left" Width="204px"></asp:TextBox>
        &nbsp;
        &nbsp;&nbsp;
        <asp:TextBox ID="DRemark" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
            Height="20px" MaxLength="100" Style="z-index: 101; left: 219px; position: absolute;
            top: 143px; text-align: left" TextMode="MultiLine" Width="556px"></asp:TextBox>
        <asp:TextBox ID="DApper" runat="server" AutoPostBack="True" BackColor="#C0FFFF"
            BorderStyle="None" ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 102;
            left: 553px; position: absolute; top: 65px; text-align: left" Width="220px"></asp:TextBox>
        <asp:TextBox ID="DDivision" runat="server" AutoPostBack="True" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 103; left: 219px;
            position: absolute; top: 90px; text-align: left" Width="201px"></asp:TextBox>
        <asp:TextBox ID="DAppDate" runat="server" AutoPostBack="True" BackColor="#C0FFFF"
            BorderStyle="None" ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 104;
            left: 552px; position: absolute; top: 90px; text-align: left" Width="220px"></asp:TextBox>
        <asp:TextBox ID="DNo" runat="server" BackColor="White" BorderColor="White" BorderStyle="None"
            Font-Names="Times New Roman" Font-Size="11pt" Style="z-index: 105; left: 218px;
            position: absolute; top: 65px" Width="198px">DNo</asp:TextBox>
        &nbsp;&nbsp;&nbsp; &nbsp;&nbsp;
        <asp:HyperLink ID="LFCDataList" runat="server" Style="z-index: 107; left: 254px;
            position: absolute; top: 216px" Target="_blank">細項內容</asp:HyperLink>
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:GridView ID="GridView3" runat="server" AutoGenerateColumns="False" BorderColor="#CC9966"
            BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt" ShowFooter="True"
            Style="z-index: 118; left: 25px; position: absolute; top: 262px" Width="767px">
            <Columns>
                <asp:BoundField DataField="Buyer" HeaderText="客戶-BUYER">
                    <ItemStyle Width="100px" />
                </asp:BoundField>
                <asp:BoundField DataField="ItemCode" HeaderText="ITEM NAME">
                    <ItemStyle Width="100px" />
                </asp:BoundField>
                <asp:BoundField DataField="Qty" DataFormatString="{0:N0}" HeaderText="QTY">
                    <ItemStyle Width="100px" />
                </asp:BoundField>
                <asp:BoundField DataField="AMT" DataFormatString="{0:N0}" HeaderText="金額">
                    <ItemStyle Width="100px" />
                </asp:BoundField>
                <asp:BoundField DataField="INVQTY" DataFormatString="{0:N0}" HeaderText="KEEP在庫數">
                    <ItemStyle Width="100px" />
                </asp:BoundField>
                <asp:BoundField DataField="INVAMT" DataFormatString="{0:N0}" HeaderText="KEEP在庫金額">
                    <ItemStyle Width="120px" />
                </asp:BoundField>
                <asp:BoundField DataField="Attachfile" HeaderText="附檔">
                    <ItemStyle Width="50px" />
                </asp:BoundField>
            </Columns>
            <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
            <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
        </asp:GridView>
        </div>
    </form>
</body>
</html>
