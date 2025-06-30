<%@ Page Language="VB" AutoEventWireup="false" CodeFile="YKKColorShow.aspx.vb" Inherits="YKKColorShow" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>色番資料</title>
</head>
<body>
    <form id="form1" runat="server">
        &nbsp;<asp:TextBox ID="DData" runat="server" style=" TEXT-TRANSFORM: uppercase; z-index: 100; left: 136px; position: absolute; top: 16px" Height="17px"></asp:TextBox><asp:Button ID="Button1" runat="server" Text="查詢" style="z-index: 101; left: 304px; position: absolute; top: 8px" Height="30px" BackColor="Yellow" />
        <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="DData"
            ErrorMessage="請輸入YKK COLOR" Style="left: 96px; position: relative; top: 32px; z-index: 102;" Width="152px"></asp:RequiredFieldValidator>
        <asp:Button ID="BCHECK" runat="server" Text="確認" style="z-index: 103; left: 304px; position: absolute; top: 40px" Height="30px" BackColor="#80FFFF" />&nbsp;<br />
        &nbsp;&nbsp; 
        <asp:Image ID="Image1" runat="server" ImageUrl="~/Images/YKKCOLORSHOW_01.jpg" style="z-index: 104; left: 16px; position: absolute; top: 80px" />
        <asp:TextBox ID="DDarkLight2" runat="server" Height="17px" Style="TEXT-TRANSFORM: uppercase; z-index: 105; left: 344px;
            position: absolute; top: 152px" Width="40px" ReadOnly="True" BackColor="#E0E0E0"></asp:TextBox>
        <asp:TextBox ID="DDarkLight1" runat="server" BackColor="#E0E0E0" Height="17px" ReadOnly="True"
            Style="z-index: 114; left: 136px; text-transform: uppercase; position: absolute;
            top: 152px" Width="40px"></asp:TextBox>
        <asp:TextBox ID="DCOLOR2" runat="server" BackColor="#E0E0E0" Height="17px" ReadOnly="True"
            Style="z-index: 105; left: 280px; text-transform: uppercase; position: absolute;
            top: 152px" Width="40px"></asp:TextBox>
        <asp:Panel ID="Panel2" runat="server" Height="129px" Style="z-index: 106; left: 232px;
            position: absolute; top: 248px" Width="193px">
        </asp:Panel>
        <asp:TextBox ID="DR2" runat="server" Style="z-index: 107; left: 248px; position: absolute;
            top: 216px" Width="40px" BackColor="#E0E0E0" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DG2" runat="server" Style="z-index: 108; left: 304px; position: absolute;
            top: 216px" Width="40px" BackColor="#E0E0E0" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DB2" runat="server" Style="z-index: 109; left: 368px; position: absolute;
            top: 216px" Width="40px" BackColor="#E0E0E0" ReadOnly="True"></asp:TextBox>
        &nbsp;&nbsp;
        &nbsp;&nbsp;
        <asp:TextBox ID="DR1" runat="server" Style="z-index: 110; left: 48px; position: absolute;
            top: 216px" Width="40px" BackColor="#E0E0E0" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DG1" runat="server" Style="z-index: 111; left: 112px; position: absolute;
            top: 216px" Width="40px" BackColor="#E0E0E0" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DB1" runat="server" Style="z-index: 112; left: 176px; position: absolute;
            top: 216px" Width="40px" BackColor="#E0E0E0" ReadOnly="True"></asp:TextBox>
        &nbsp;&nbsp;
        <asp:Label ID="Label2" runat="server" Style="z-index: 116; left: 32px; position: absolute;
            top: 16px" Text="COLOR CODE"></asp:Label>
        <br />
        <br />
        &nbsp;&nbsp;<br />
        <asp:Panel ID="Panel1" runat="server" Height="129px" Style="z-index: 113; left: 40px;
            position: absolute; top: 248px" Width="193px">
        </asp:Panel>
        <asp:TextBox ID="DCOLOR1" runat="server" Height="17px" Style="TEXT-TRANSFORM: uppercase; z-index: 114; left: 72px;
            position: absolute; top: 152px" Width="48px" ReadOnly="True" BackColor="#E0E0E0"></asp:TextBox>
        <asp:Label ID="Label1" runat="server" Style="z-index: 116; left: 32px; position: absolute;
            top: 48px" Font-Size="Medium" ForeColor="Red"></asp:Label>
    </form>
</body>
</html>
