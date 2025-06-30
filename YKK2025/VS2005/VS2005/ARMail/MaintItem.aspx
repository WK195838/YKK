<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MaintItem.aspx.vb" Inherits="MaintItem" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>未命名頁面</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    <table>
    <tr>
    
    <td> <asp:TextBox ID="TextBox1" runat="server" SkinID="txt_Lightgray" Style="z-index: 100;
            left: 125px; position: absolute; top: 100px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="TextBox3" runat="server" SkinID="txt_Green" Style="z-index: 101;
            left: 125px; position: absolute; top: 199px" Width="546px" MaxLength="250"></asp:TextBox>
        <asp:TextBox ID="TextBox4" runat="server" SkinID="txt_Lightgray" Style="z-index: 102;
            left: 125px; position: absolute; top: 134px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="TextBox2" runat="server" SkinID="txt_Green" Style="z-index: 103;
            left: 125px; position: absolute; top: 166px" Width="266px" MaxLength="20"></asp:TextBox>
        <img src="images/ControlTable.jpg" />
        <asp:DropDownList ID="DropDownList1" runat="server" SkinID="DDL_Green" Style="z-index: 104;
            left: 288px; position: absolute; top: 133px" Width="111px">
            <asp:ListItem Selected="True">0</asp:ListItem>
            <asp:ListItem>1</asp:ListItem>
            <asp:ListItem>2</asp:ListItem>
        </asp:DropDownList></td>
    </tr>
    <tr>
    <td align="right">
        <asp:Button ID="Button2" runat="server"  Text="回上一頁"  /><asp:Button ID="Button1" runat="server"  Text="儲存" />
        </td>
    </tr>
    
    </table>
       
        
    
    </div>
    </form>
</body>
</html>
