<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Default.aspx.vb" Inherits="_Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>未命名頁面</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
<!-- FormNo-->
        FormNo:<asp:DropDownList ID="DropDownList1" runat="server">
            <asp:ListItem Selected="True" Value="003001">開發委託單</asp:ListItem>
            <asp:ListItem Value="003002">表面處理委託書</asp:ListItem>
        </asp:DropDownList>
<!-- FormSno-->
        FormSNo<asp:TextBox ID="TextBox2" runat="server" Width="95px">1</asp:TextBox>
<!-- Step-->
        Step<asp:TextBox ID="TextBox3" runat="server"  Width="43px" AutoCompleteType="Disabled">10</asp:TextBox>
<!-- SeqNo-->
        SeqNo<asp:TextBox ID="TextBox4" runat="server" Width="66px">1</asp:TextBox>
<!-- 申請者ID-->
        ApplyID<asp:TextBox ID="TextBox5" runat="server"  Width="67px">sb003</asp:TextBox>
<!-- 核定者ID-->
        UserID<asp:TextBox ID="TextBox6" runat="server"  Width="67px" AutoPostBack="True">sb005</asp:TextBox>
<!-- 開發委託單-->
        <br />
        <asp:Button ID="Button1" runat="server" Text="開發委託單01" />
        &nbsp;&nbsp;<asp:Button ID="Button2" runat="server" Text="開發委託單02" /><!-- 開發見本--><br />
        &nbsp;
        <br />
<!-- 開視窗--><asp:Button ID="Button3" runat="server" Text="表面處理託單01" />
        <asp:Button ID="Button4" runat="server" Text="表面處理託單02" /></div>
    </form>
</body>
</html>
