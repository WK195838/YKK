<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ModifyDataSheet.aspx.vb" Inherits="ModifyDataSheet" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>未命名頁面</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:Image ID="Image1" runat="server" ImageUrl="~/Images/ModifyDataSheet_01.png"
             Style="z-index: 99; left: 2px; position: absolute; top: 2px" />
        <asp:TextBox ID="TextBox7" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="22px" MaxLength="20" Style="z-index: 126; left: 225px;
            position: absolute; top: 489px; text-align: left" Width="158px"></asp:TextBox>
        <asp:TextBox ID="TextBox6" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="237px" MaxLength="20" Style="z-index: 126; left: 60px;
            position: absolute; top: 243px; text-align: left" TextMode="MultiLine" Width="536px"></asp:TextBox>
        <asp:TextBox ID="TextBox5" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="237px" MaxLength="20" Style="z-index: 126; left: 60px;
            position: absolute; top: 515px; text-align: left" TextMode="MultiLine" Width="536px"></asp:TextBox>
        <asp:DropDownList ID="DropDownList3" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 59px; position: absolute; top: 490px"
            Width="163px">
        </asp:DropDownList>
        <asp:HyperLink ID="aa" runat="server" Height="22px" Style="z-index: 261; left: 312px;
            position: absolute; top: 154px" Target="_blank" Width="121px">[aa]</asp:HyperLink>
        <asp:DropDownList ID="DAPPBUYER" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 143px; position: absolute; top: 154px"
            Width="163px">
        </asp:DropDownList><asp:DropDownList ID="DropDownList1" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 143px; position: absolute; top: 187px"
            Width="163px">
        </asp:DropDownList>
        <asp:DropDownList ID="DropDownList2" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 59px; position: absolute; top: 218px"
            Width="163px">
        </asp:DropDownList>
        <asp:TextBox ID="DNO" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="22px" MaxLength="20" Style="z-index: 126; left: 143px; position: absolute;
            top: 84px; text-align: left" Width="158px"></asp:TextBox>
        <asp:TextBox ID="TextBox1" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="22px" MaxLength="20" Style="z-index: 126; left: 437px;
            position: absolute; top: 84px; text-align: left" Width="158px"></asp:TextBox>
        <asp:TextBox ID="TextBox2" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="22px" MaxLength="20" Style="z-index: 126; left: 143px;
            position: absolute; top: 118px; text-align: left" Width="158px"></asp:TextBox>
        <asp:TextBox ID="TextBox3" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="22px" MaxLength="20" Style="z-index: 126; left: 437px;
            position: absolute; top: 118px; text-align: left" Width="158px"></asp:TextBox>
        <asp:TextBox ID="TextBox4" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="22px" MaxLength="20" Style="z-index: 126; left: 311px;
            position: absolute; top: 152px; text-align: left" Width="158px"></asp:TextBox>
        <asp:Button ID="BEXPDEL" runat="server" Height="20px" Style="z-index: 111; left: 473px;
            position: absolute; top: 157px" Text="....." Width="20px" /><asp:Button ID="Button1" runat="server" Height="20px" Style="z-index: 111; left: 386px;
            position: absolute; top: 494px" Text="....." Width="20px" />
   
    </div>
    </form>
</body>
</html>
