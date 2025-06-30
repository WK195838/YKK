<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Commission_Structure.aspx.vb" Inherits="Commission_Structure" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>委託單結構</title>
</head>
<body >
    <form id="form1" runat="server">
    <div>
        <div style="display: inline; z-index: 113; left: 13px;
            width: 81px; color: blue; position: absolute; top: 8px; height: 24px" title="更新頻率：05:30／每日">
            篩選項目：</div>
        &nbsp;&nbsp;
        <asp:TextBox ID="DNo" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Height="20px" ReadOnly="True" Style="z-index: 100; left: 96px; position: absolute;
            top: 8px" Width="120px">DNo</asp:TextBox>
        <asp:DropDownList ID="DDivision" runat="server" BackColor="Yellow" ForeColor="Blue"
            Height="40px" Style="z-index: 101; left: 224px; position: absolute; top: 8px"
            Width="124px">
        </asp:DropDownList>
        <asp:TextBox ID="DPerson" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="20px" ReadOnly="True" Style="z-index: 102; left: 353px;
            position: absolute; top: 8px" Width="120px">DPerson</asp:TextBox>
        <asp:TextBox ID="DBuyer" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Height="20px" ReadOnly="True" Style="z-index: 103; left: 481px; position: absolute;
            top: 8px" Width="120px">DBuyer</asp:TextBox>
        <asp:TextBox ID="DCpsc" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Height="20px" ReadOnly="True" Style="z-index: 104; left: 609px; position: absolute;
            top: 37px" Width="120px">DCpsc</asp:TextBox>
        <asp:TextBox ID="DMsgBox" runat="server" BackColor="White" BorderStyle="None" Font-Size="10pt"
            ForeColor="Red" Height="16px" ReadOnly="True" Style="z-index: 105; left: 360px;
            position: absolute; top: 70px" Width="413px">DMsgBox</asp:TextBox>
        <asp:TextBox ID="DSpec" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Height="20px" ReadOnly="True" Style="z-index: 106; left: 481px; position: absolute;
            top: 37px" Width="120px">DSpec</asp:TextBox>
        &nbsp;
        <asp:TextBox ID="DMakeMap" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="20px" ReadOnly="True" Style="z-index: 107; left: 224px;
            position: absolute; top: 37px" Width="120px">DMakeMap</asp:TextBox>
        <asp:TextBox ID="DMapNo" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Height="20px" ReadOnly="True" Style="z-index: 108; left: 96px; position: absolute;
            top: 37px" Width="120px">DMapNo</asp:TextBox>
        <asp:Button ID="Go" runat="server" BackColor="White" ForeColor="Blue" Height="24px"
            Style="z-index: 109; left: 740px; position: absolute; top: 37px" Text="Go" Width="40px" />
    </div>
        &nbsp; &nbsp; &nbsp; &nbsp;
    
        <asp:TreeView ID="TreeView1" runat="server" style="z-index: 110; left: 97px; position: absolute; top: 91px" ImageSet="XPFileExplorer" Visible="False" ExpandDepth="2">
        </asp:TreeView>
        &nbsp;&nbsp;
        <div style="display: inline; z-index: 114; left: 97px;
            width: 259px; color: blue; position: absolute; top: 69px; height: 16px" title="篩選項目：">
            <span style="font-size: 10pt; color: darkblue">以下各項目中，需指定１個以上的篩選項目</span></div>
        <asp:TextBox ID="DLastUpdateTime" runat="server" BackColor="White" BorderStyle="None"
            Font-Size="10pt" ForeColor="Blue" Height="16px" ReadOnly="True" Style="z-index: 111;
            left: 614px; position: absolute; top: 9px" Width="174px">DLastUpdateTime</asp:TextBox>
        <asp:TextBox ID="DSliderCode" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="20px" ReadOnly="True" Style="z-index: 115; left: 353px;
            position: absolute; top: 38px" Width="120px">DSliderCode</asp:TextBox>
    </form>
</body>

</html>
