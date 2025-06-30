<%@ Page Language="VB" AutoEventWireup="false" CodeFile="HR_WorkRatioOtherInfor.aspx.vb" Inherits="HR_WorkRatioOtherInfor" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>出勤率-其他說明</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        &nbsp;&nbsp;
        <asp:Image ID="Image1" runat="server" ImageUrl="~/Images/other.PNG" Style="z-index: 99;
            left: 5px; position: absolute; top: 2px" />
        <asp:TextBox ID="DZZSum" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 104; left: 127px;
            position: absolute; top: 631px" Width="76px"></asp:TextBox>
        <asp:TextBox ID="DZZMin1" runat="server" BackColor="#FFFF80" BorderStyle="None" ForeColor="Black"
            Height="17px" ReadOnly="True" Style="z-index: 116; left: 127px; position: absolute;
            top: 464px" Width="76px"></asp:TextBox>
        <asp:TextBox ID="DZZMin2" runat="server" BackColor="#FFFF80" BorderStyle="None" ForeColor="Black"
            Height="17px" ReadOnly="True" Style="z-index: 117; left: 127px; position: absolute;
            top: 498px" Width="76px"></asp:TextBox>
        <asp:TextBox ID="DZZMin3" runat="server" BackColor="#FFFF80" BorderStyle="None" ForeColor="Black"
            Height="17px" ReadOnly="True" Style="z-index: 118; left: 127px; position: absolute;
            top: 528px" Width="76px"></asp:TextBox>
        <asp:TextBox ID="DZZMin4" runat="server" BackColor="#FFFF80" BorderStyle="None" ForeColor="Black"
            Height="17px" ReadOnly="True" Style="z-index: 119; left: 127px; position: absolute;
            top: 563px" Width="76px"></asp:TextBox>
        <asp:TextBox ID="DZZMin5" runat="server" BackColor="#FFFF80" BorderStyle="None" ForeColor="Black"
            Height="17px" ReadOnly="True" Style="z-index: 120; left: 127px; position: absolute;
            top: 597px" Width="76px"></asp:TextBox>
        <asp:TextBox ID="DZZDesc2" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 121; left: 228px;
            position: absolute; top: 498px" Width="485px"></asp:TextBox>
        <asp:TextBox ID="DZZDesc3" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 122; left: 228px;
            position: absolute; top: 528px" Width="485px"></asp:TextBox>
        <asp:TextBox ID="DZZDesc4" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 123; left: 228px;
            position: absolute; top: 563px" Width="485px"></asp:TextBox>
        <asp:TextBox ID="DZZDesc5" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 124; left: 228px;
            position: absolute; top: 597px" Width="485px"></asp:TextBox>
        <asp:TextBox ID="DZZDesc1" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 126; left: 228px;
            position: absolute; top: 464px" Width="485px"></asp:TextBox>
        <asp:TextBox ID="DX2Sum" runat="server" BackColor="#FFFF80" BorderStyle="None" ForeColor="Black"
            Height="17px" ReadOnly="True" Style="z-index: 104; left: 128px; position: absolute;
            top: 903px" Width="76px"></asp:TextBox>
        <asp:TextBox ID="DYearDay" runat="server" BackColor="#FFC080" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 105; left: 229px;
            position: absolute; top: 114px" Width="488px"></asp:TextBox>
        <asp:TextBox ID="DY3Desc1" runat="server" BackColor="#FFC080" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 107; left: 228px;
            position: absolute; top: 189px" Width="488px"></asp:TextBox>
        <asp:TextBox ID="DY3Min1" runat="server" BackColor="#FFC080" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 108; left: 126px;
            position: absolute; top: 189px" Width="77px"></asp:TextBox>
        <asp:TextBox ID="DY2Min" runat="server" BackColor="#FFC080" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 109; left: 130px;
            position: absolute; top: 83px" Width="77px"></asp:TextBox>
        <asp:TextBox ID="DY2Sum" runat="server" BackColor="#FFC080" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 110; left: 130px;
            position: absolute; top: 114px" Width="77px"></asp:TextBox>
        <asp:TextBox ID="DY3Min2" runat="server" BackColor="#FFC080" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 111; left: 126px;
            position: absolute; top: 223px" Width="77px"></asp:TextBox>
        <asp:TextBox ID="DY3Min3" runat="server" BackColor="#FFC080" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 112; left: 126px;
            position: absolute; top: 256px" Width="77px"></asp:TextBox>
        <asp:TextBox ID="DY3Min4" runat="server" BackColor="#FFC080" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 113; left: 126px;
            position: absolute; top: 289px" Width="77px"></asp:TextBox>
        <asp:TextBox ID="DY3Min5" runat="server" BackColor="#FFC080" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 114; left: 126px;
            position: absolute; top: 321px" Width="77px"></asp:TextBox>
        <asp:TextBox ID="DY3Sum" runat="server" BackColor="#FFC080" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 115; left: 126px;
            position: absolute; top: 355px" Width="77px"></asp:TextBox>
        <asp:TextBox ID="DX2Min1" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 116; left: 128px;
            position: absolute; top: 740px" Width="76px"></asp:TextBox>
        <asp:TextBox ID="DX2Min2" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 117; left: 128px;
            position: absolute; top: 770px" Width="76px"></asp:TextBox>
        <asp:TextBox ID="DX2Min3" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 118; left: 128px;
            position: absolute; top: 803px" Width="76px"></asp:TextBox>
        <asp:TextBox ID="DX2Min4" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 119; left: 128px;
            position: absolute; top: 835px" Width="76px"></asp:TextBox>
        <asp:TextBox ID="DX2Min5" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 120; left: 128px;
            position: absolute; top: 869px" Width="76px"></asp:TextBox>
        <asp:TextBox ID="DX2Desc2" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 121; left: 228px;
            position: absolute; top: 771px" Width="485px"></asp:TextBox>
        <asp:TextBox ID="DX2Desc3" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 122; left: 228px;
            position: absolute; top: 804px" Width="485px"></asp:TextBox>
        <asp:TextBox ID="DX2Desc4" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 123; left: 228px;
            position: absolute; top: 836px" Width="485px"></asp:TextBox>
        <asp:TextBox ID="DX2Desc5" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 124; left: 228px;
            position: absolute; top: 870px" Width="485px"></asp:TextBox>
        <asp:TextBox ID="DX2Desc1" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 126; left: 228px;
            position: absolute; top: 741px" Width="485px"></asp:TextBox>
        <asp:TextBox ID="DY3Desc2" runat="server" BackColor="#FFC080" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 127; left: 228px;
            position: absolute; top: 223px" Width="488px"></asp:TextBox>
        <asp:TextBox ID="DY3Desc3" runat="server" BackColor="#FFC080" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 128; left: 228px;
            position: absolute; top: 256px" Width="488px"></asp:TextBox>
        <asp:TextBox ID="DY3Desc4" runat="server" BackColor="#FFC080" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 129; left: 228px;
            position: absolute; top: 289px" Width="488px"></asp:TextBox>
        <asp:TextBox ID="DY3Desc5" runat="server" BackColor="#FFC080" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 130; left: 228px;
            position: absolute; top: 321px" Width="488px"></asp:TextBox>
        <asp:TextBox ID="DY2Desc" runat="server" BackColor="#FFC080" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 131; left: 229px;
            position: absolute; top: 82px" Width="488px"></asp:TextBox>
    </div>
    </form>
</body>
</html>
