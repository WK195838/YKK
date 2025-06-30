<%@ Page Language="VB" AutoEventWireup="false" CodeFile="AutoItemRequest.aspx.vb" Inherits="AutoItemRequest" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>自動配對處理</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:TextBox ID="DSProgressBar" runat="server" AutoPostBack="True" BackColor="#FFFF80"
            BorderStyle="None" Font-Bold="False" Font-Size="20pt" ForeColor="Red" Height="89px"
            MaxLength="7" Style="z-index: 126; left: 5px; position: absolute; top: 37px; text-align: left"
            Width="849px">自動配對處理中..........</asp:TextBox>
        <asp:TextBox ID="DEProgressBar" runat="server" AutoPostBack="True" BackColor="#FFFF80"
            BorderStyle="None" Font-Bold="False" Font-Size="20pt" ForeColor="Red" Height="90px"
            MaxLength="7" Style="z-index: 126; left: 5px; position: absolute; top: 207px;
            text-align: left" Width="849px">自動配對處理中..........</asp:TextBox>
        <asp:TextBox ID="DSystem" runat="server" AutoPostBack="True" BackColor="#80FF80"
            BorderStyle="None" Font-Bold="False" Font-Size="20pt" ForeColor="#0000C0" Height="28px"
            MaxLength="7" Style="z-index: 126; left: 5px; position: absolute; top: 7px; text-align: left"
            Width="849px">----------  IRW  ----------</asp:TextBox>
        <asp:TextBox ID="DCount" runat="server" AutoPostBack="True" BackColor="#FFFF80" BorderStyle="None"
            Font-Bold="False" Font-Size="20pt" ForeColor="Red" Height="78px" MaxLength="7"
            Style="z-index: 126; left: 5px; position: absolute; top: 128px; text-align: left"
            Width="849px">處理筆數</asp:TextBox>
    
    </div>
    </form>
</body>
</html>
