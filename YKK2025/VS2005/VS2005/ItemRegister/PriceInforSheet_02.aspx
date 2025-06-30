<%@ Page Language="VB" AutoEventWireup="false" CodeFile="PriceInforSheet_02.aspx.vb" Inherits="PriceInforSheet_02" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>未命名頁面</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:TextBox ID="DNo" runat="server" BackColor="White" BorderColor="White" BorderStyle="None"
            Font-Names="Times New Roman" Font-Size="11pt" Style="z-index: 100; left: 24px;
            position: relative; top: 8px">DNo</asp:TextBox>
        <img src="images/WFS_PriceInfor.PNG" style="z-index: 99; left: 16px; position: absolute;
            top: 56px" id="IMG1" language="javascript" onclick="return IMG1_onclick()" />
        <asp:TextBox ID="DEmpID" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Style="z-index: 102; left: 618px; position: absolute; top: 128px" Text="EmpID"
            Width="93px"></asp:TextBox>

        <asp:TextBox ID="Dremark" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="32px" Style="z-index: 101; left: 144px; position: absolute;
            top: 184px" Width="567px"></asp:TextBox>
        <asp:TextBox ID="DDate" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Style="z-index: 102; left: 144px; position: absolute; top: 128px" Text="Date"
            Width="229px"></asp:TextBox>
        <asp:TextBox ID="DName" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Style="z-index: 103; left: 496px; position: absolute; top: 128px" Text="Name"
            Width="116px"></asp:TextBox>
        <asp:TextBox ID="DJobTitle" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Style="z-index: 104; left: 144px; position: absolute; top: 152px"
            Text="DJobTitle" Width="150px"></asp:TextBox>
        <asp:TextBox ID="DJobCode" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Style="z-index: 104; left: 301px; position: absolute; top: 152px"
            Text="JobCode" Width="72px"></asp:TextBox>
        &nbsp;&nbsp;
        <asp:TextBox ID="DDepoName" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Style="z-index: 106; left: 496px; position: absolute; top: 152px"
            Text="DepoName" Width="32px"></asp:TextBox>
        <asp:TextBox ID="DDepoCode" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Style="z-index: 107; left: 535px; position: absolute; top: 152px"
            Text="DepoCode" Width="24px"></asp:TextBox>
        <asp:TextBox ID="DDivision" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Style="z-index: 108; left: 566px; position: absolute; top: 152px"
            Text="Division" Width="88px"></asp:TextBox>
        <asp:TextBox ID="DDivisionCode" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Style="z-index: 109; left: 661px; position: absolute; top: 152px"
            Text="DivisionCode" Width="50px"></asp:TextBox>
        &nbsp; &nbsp; &nbsp; &nbsp;

        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <asp:GridView ID="grd" runat="server" AllowSorting="True" AutoGenerateColumns="False"
            BackColor="White" Font-Bold="False" Font-Names="Arial" Font-Size="9pt" Height="120px"
            ShowHeader="False" Style="z-index: 106; left: 24px; position: absolute; top: 296px"
            Width="696px">
            <Columns>
                <asp:TemplateField>
                    <ItemTemplate>
                        <table border="0">
                            <tr>
                                <td style="width: 38px;">
                                    <asp:TextBox ID="ItemNo" runat="server" ReadOnly="true" Text='<%# Eval("ItemNo") %>' Width="115px"></asp:TextBox></td>
                                <td style="width: 38px;">
                                    <asp:TextBox ID="Code" runat="server" ReadOnly="true" Text='<%# Eval("Code") %>'
                                        Width="55px"></asp:TextBox></td>
                                <td colspan="8" style="width: 500px;">
                                    <asp:TextBox ID="ItemName" runat="server" ReadOnly="true" Text='<%# Eval("ItemName") %>'
                                        Width="490px"></asp:TextBox>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <asp:TextBox ID="ApplyName" runat="server" ReadOnly="true" Text='<%# Eval("ApplyName") %>'
                                        Width="115px"></asp:TextBox></td>
                                <td>
                                    <asp:TextBox ID="Version" runat="server" ReadOnly="true" Text='<%# Eval("Version") %>'
                                        Width="55px"></asp:TextBox></td>
                                <td>
                                    <asp:TextBox ID="A1" runat="server" ReadOnly="true" Text='<%# Eval("A1") %>' Width="50px"></asp:TextBox></td>
                                <td>
                                    <asp:TextBox ID="B1" runat="server" ReadOnly="true" Text='<%# Eval("B1") %>' Width="50px"></asp:TextBox></td>
                                <td>
                                    <asp:TextBox ID="B2" runat="server" ReadOnly="true" Text='<%# Eval("B2") %>' Width="50px"></asp:TextBox></td>
                                <td>
                                    <asp:TextBox ID="B4" runat="server" ReadOnly="true" Text='<%# Eval("B4") %>' Width="50px"></asp:TextBox></td>
                                <td>
                                    <asp:TextBox ID="C1" runat="server" ReadOnly="true" Text='<%# Eval("C1") %>' Width="50px"></asp:TextBox></td>
                                <td>
                                    <asp:TextBox ID="C2" runat="server" ReadOnly="true" Text='<%# Eval("C2") %>' Width="50px"></asp:TextBox></td>
                                <td>
                                    <asp:TextBox ID="C3" runat="server" ReadOnly="true" Text='<%# Eval("C3") %>' Width="50px"></asp:TextBox></td>
                                <td>
                                    <asp:TextBox ID="D1" runat="server" ReadOnly="true" Text='<%# Eval("D1") %>' Width="50px"></asp:TextBox></td>
                            </tr>
                        </table>
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
            <HeaderStyle BackColor="White" />
            <AlternatingRowStyle BackColor="#FFFF80" />
        </asp:GridView>
        &nbsp;<br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        &nbsp;<br />
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
            
    
    </div>
        &nbsp;
    </form>
</body>
</html>
