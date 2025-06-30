<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ItemRegisterFSLDSheet_02.aspx.vb" Inherits="ItemRegisterFSLDSheet_02" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>Wave's Item登錄申請單(SLD)</title>
    <script language="javascript" src="RegisterItem.js"></script>   
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <img src="images/RegisterFSLDSheet_001.jpg" style="z-index: 1; left: 6px; position: absolute;top: 7px" />
        <asp:TextBox ID="DTYPE" runat="server" AutoPostBack="True" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126; left: 16px;
            position: absolute; top: 48px; text-align: left" Width="120px"></asp:TextBox>
        &nbsp;
        <asp:TextBox ID="DDate" runat="server" AutoPostBack="True" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126; left: 116px;
            position: absolute; top: 78px; text-align: left" Width="226px"></asp:TextBox>
        <asp:TextBox ID="DJobTitle" runat="server" AutoPostBack="True" BackColor="#C0FFFF"
            BorderStyle="None" ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126;
            left: 116px; position: absolute; top: 103px; text-align: left" Width="226px"></asp:TextBox>
        <asp:TextBox ID="DDivision" runat="server" AutoPostBack="True" BackColor="#C0FFFF"
            BorderStyle="None" ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126;
            left: 452px; position: absolute; top: 103px; text-align: left" Width="246px"></asp:TextBox>
        <asp:TextBox ID="DName" runat="server" AutoPostBack="True" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126; left: 452px;
            position: absolute; top: 77px; text-align: left" Width="246px"></asp:TextBox>
        <asp:TextBox ID="DNo" runat="server" BackColor="White" BorderColor="White" BorderStyle="None"
            Font-Names="Times New Roman" Font-Size="11pt" Style="z-index: 100; left: 18px;
            position: absolute; top: 14px">DNo</asp:TextBox>
        &nbsp; &nbsp;
        <asp:HyperLink ID="LAttachfile1" runat="server" Font-Size="12pt" Height="8px" NavigateUrl="BoardEdit.aspx"
            Style="z-index: 124; left: 122px; position: absolute; top: 183px" Target="_blank"
            Width="100px">參考文件</asp:HyperLink>
        <asp:TextBox ID="DRemark" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="43px" Style="z-index: 126; left: 115px;
            position: absolute; top: 129px; text-align: left" Width="583px" MaxLength="100" TextMode="MultiLine"></asp:TextBox>
        </div>
    </form>
</body>
</html>

