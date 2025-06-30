<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Commission_Sheet.aspx.vb" Inherits="Commission_Sheet" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>委託單</title>
</head>
<body >
    <form id="form1" runat="server">
    <div>
        <div style="display: inline; z-index: 116; left: 13px;
            width: 81px; color: blue; position: absolute; top: 8px; height: 24px" title="更新頻率：05:30／每日">
            篩選項目：</div>
        <asp:DropDownList ID="DFormName" runat="server" AutoPostBack="True" BackColor="Yellow"
            ForeColor="Blue" Height="40px" Style="z-index: 100; left: 96px; position: absolute;
            top: 9px" Width="253px">
        </asp:DropDownList>
        <asp:DropDownList ID="DProgress" runat="server" BackColor="Yellow" ForeColor="Blue"
            Height="40px" Style="z-index: 101; left: 354px; position: absolute; top: 9px"
            Width="120px">
            <asp:ListItem Selected="True" Value="ALL">ALL</asp:ListItem>
            <asp:ListItem Value="1">開發中</asp:ListItem>
            <asp:ListItem Value="2">開發完成</asp:ListItem>
        </asp:DropDownList>
        <asp:DropDownList ID="DSts" runat="server" BackColor="Yellow" ForeColor="Blue" Height="40px"
            Style="z-index: 102; left: 482px; position: absolute; top: 9px" Width="120px">
            <asp:ListItem Selected="True" Value="ALL">ALL</asp:ListItem>
            <asp:ListItem Value="1">OK</asp:ListItem>
            <asp:ListItem Value="2">NG</asp:ListItem>
            <asp:ListItem Value="3">取消</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DNo" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Height="20px" ReadOnly="True" Style="z-index: 103; left: 96px; position: absolute;
            top: 54px" Width="120px">DNo</asp:TextBox>
        <asp:DropDownList ID="DDivision" runat="server" BackColor="Yellow" ForeColor="Blue"
            Height="40px" Style="z-index: 104; left: 224px; position: absolute; top: 54px"
            Width="120px">
        </asp:DropDownList>
        <asp:TextBox ID="DPerson" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="20px" ReadOnly="True" Style="z-index: 105; left: 352px;
            position: absolute; top: 54px" Width="120px">DPerson</asp:TextBox>
        <asp:TextBox ID="DBuyer" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Height="20px" ReadOnly="True" Style="z-index: 106; left: 480px; position: absolute;
            top: 54px" Width="120px">DBuyer</asp:TextBox>
        <asp:TextBox ID="DCpsc" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Height="20px" ReadOnly="True" Style="z-index: 107; left: 608px; position: absolute;
            top: 78px" Width="120px">DCpsc</asp:TextBox>
        <asp:TextBox ID="DMsgBox" runat="server" BackColor="White" BorderStyle="None" Font-Size="10pt"
            ForeColor="Red" Height="16px" ReadOnly="True" Style="z-index: 108; left: 363px;
            position: absolute; top: 32px" Width="413px">DMsgBox</asp:TextBox>
        <asp:TextBox ID="DLastUpdateTime" runat="server" BackColor="White" BorderStyle="None" Font-Size="10pt"
            ForeColor="Blue" Height="16px" ReadOnly="True" Style="z-index: 118; left: 607px;
            position: absolute; top: 13px" Width="187px">DLastUpdateTime</asp:TextBox>
        <asp:TextBox ID="DSpec" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Height="20px" ReadOnly="True" Style="z-index: 110; left: 480px; position: absolute;
            top: 78px" Width="120px">DSpec</asp:TextBox>
        <asp:TextBox ID="DSliderCode" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="20px" ReadOnly="True" Style="z-index: 111; left: 352px;
            position: absolute; top: 78px" Width="120px">DSliderCode</asp:TextBox>
        <asp:TextBox ID="DMakeMap" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="20px" ReadOnly="True" Style="z-index: 112; left: 224px;
            position: absolute; top: 78px" Width="120px">DMakeMap</asp:TextBox>
        <asp:TextBox ID="DMapNo" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Height="20px" ReadOnly="True" Style="z-index: 113; left: 96px; position: absolute;
            top: 78px" Width="120px">DMapNo</asp:TextBox>
        <asp:Button ID="Go" runat="server" BackColor="White" ForeColor="Blue" Height="24px"
            Style="z-index: 114; left: 739px; position: absolute; top: 78px" Text="Go" Width="40px" />
    </div>
        &nbsp; &nbsp; &nbsp; &nbsp;
    
        <asp:TreeView ID="TreeView1" runat="server" style="z-index: 115; left: 97px; position: absolute; top: 111px" ImageSet="XPFileExplorer" Visible="False" ExpandDepth="2">
        </asp:TreeView>
        &nbsp;&nbsp;
        <div style="display: inline; z-index: 117; left: 98px;
            width: 259px; color: blue; position: absolute; top: 36px; height: 16px" title="篩選項目：">
            <span style="font-size: 10pt; color: darkblue">以下各項目中，需指定１個以上的篩選項目</span></div>
    </form>
</body>
</html>
