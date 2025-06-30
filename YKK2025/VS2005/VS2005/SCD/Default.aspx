<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Default.aspx.vb" Inherits="_Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>未命名頁面</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
<!-- FormNo-->
        FormNo:<asp:DropDownList ID="DropDownList1" runat="server">
            <asp:ListItem Selected="True" Value="002002">開發委託單</asp:ListItem>
            <asp:ListItem Value="002003">開發見本</asp:ListItem>
        </asp:DropDownList>
<!-- FormSno-->
        FormSNo<asp:TextBox ID="TextBox2" runat="server" Width="95px">5580</asp:TextBox>
<!-- Step-->
        Step<asp:TextBox ID="TextBox3" runat="server"  Width="43px">1</asp:TextBox>
<!-- SeqNo-->
        SeqNo<asp:TextBox ID="TextBox4" runat="server" Width="66px">1</asp:TextBox>
<!-- 申請者ID-->
        ApplyID<asp:TextBox ID="TextBox5" runat="server"  Width="67px">it003</asp:TextBox>
<!-- 核定者ID-->
        UserID<asp:TextBox ID="TextBox6" runat="server"  Width="67px">it003</asp:TextBox>
<!-- 開發委託單-->
        <br />
        <asp:Button ID="Button1" runat="server" Text="開發委託單" />
        <asp:Button ID="Button2" runat="server" Text="開發委託單_02" />
<!-- 開發見本-->
        <br />
        <asp:Button ID="Button4" runat="server" Text="開發見本" />
        <asp:Button ID="Button5" runat="server" Text="開發見本_02" />
        <br />
<!-- 開視窗-->
        <asp:Button ID="Button3" runat="server" Text="OPENNOPICKER" />
        <asp:TextBox ID="DAPPBUYER" runat="server" Width="95px">APPBUYER</asp:TextBox>
        <asp:TextBox ID="DSIZENO" runat="server" Width="95px" BackColor="Yellow">SIZENO</asp:TextBox></div>
    </form>
</body>
</html>
