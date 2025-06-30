<%@ Page Language="VB" AutoEventWireup="false" CodeFile="HR_OverTimeReport01.aspx.vb" Inherits="HR_OverTimeReport01" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>加班統計</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:DropDownList ID="DYear" runat="server" BackColor="Yellow" ForeColor="Blue" Height="40px"
            Style="z-index: 102; left: 99px; position: absolute; top: 9px" Width="92px">
        </asp:DropDownList>
        <asp:DropDownList ID="DMonth" runat="server" BackColor="Yellow" ForeColor="Blue"
            Height="40px" Style="z-index: 102; left: 196px; position: absolute; top: 9px"
            Width="92px">
        </asp:DropDownList>
        <div ms_positioning="FlowLayout" style="display: inline; z-index: 100; left: 8px;
            width: 88px; color: #0000ff; position: absolute; top: 9px; height: 24px" title="篩選項目：">
            加班年月：
        </div>
        <asp:Button ID="Go" runat="server" BackColor="White" ForeColor="Blue" Height="24px"
            Style="z-index: 110; left: 471px; position: absolute; top: 8px" Text="Go" Width="40px" />
        <asp:ImageButton ID="BExcel" runat="server" Height="21px" ImageUrl="Images\msexcel.gif"
            Style="z-index: 122; left: 519px; position: absolute; top: 10px" Width="21px" />
            
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False"
            BorderColor="#CC9966" BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt"
            Style="z-index: 114; left: 10px; position: absolute; top: 40px" Width="790px" ShowFooter="True">
            <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
            <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
            <Columns>
                <asp:HyperLinkField DataNavigateUrlFields="DivisionURL" DataNavigateUrlFormatString="{0}"
                    DataTextField="DivisionDesc" HeaderText="部門">
                </asp:HyperLinkField>
                <asp:HyperLinkField DataNavigateUrlFields="NameURL" DataNavigateUrlFormatString="{0}"
                    DataTextField="NameDesc" HeaderText="姓名" Target="_blank">
                </asp:HyperLinkField>
                <asp:BoundField DataField="YMDesc"  HeaderText="月份" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="FoodDesc"  HeaderText="伙食" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>
                <asp:BoundField DataField="TrafficDesc"  HeaderText="交通" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>
                <asp:BoundField DataField="Total"  HeaderText="核定時數" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>
                <asp:BoundField DataField="CVTime"  HeaderText="換休時數" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>
                <asp:BoundField DataField="TwoIn"  HeaderText="２內" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>
                <asp:BoundField DataField="TwoOut"  HeaderText="２外" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>
                <asp:BoundField DataField="Vacation"  HeaderText="假日" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>
                <asp:BoundField DataField="CVacation"  HeaderText="國假" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>
            </Columns>
        </asp:GridView><asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False"
            BorderColor="#CC9966" BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt"
            Style="z-index: 114; left: 11px; position: absolute; top: -10000px" Width="790px" ShowFooter="True">
            <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
            <Columns>
                <asp:BoundField DataField="DivisionDesc"  HeaderText="部門" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>

                <asp:BoundField DataField="NameDesc"  HeaderText="姓名" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>

                <asp:BoundField DataField="YMDesc"  HeaderText="月份" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>

                <asp:BoundField DataField="FoodDesc"  HeaderText="伙食" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>
                <asp:BoundField DataField="TrafficDesc"  HeaderText="交通" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>
                <asp:BoundField DataField="Total"  HeaderText="核定時數" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>
                <asp:BoundField DataField="CVTime"  HeaderText="換休時數" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>
                <asp:BoundField DataField="TwoIn"  HeaderText="２內" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>
                <asp:BoundField DataField="TwoOut"  HeaderText="２外" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>
                <asp:BoundField DataField="Vacation"  HeaderText="假日" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>
                <asp:BoundField DataField="CVacation"  HeaderText="國假" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>
            </Columns>
            <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
        </asp:GridView>
        <asp:DropDownList ID="DDivision" runat="server" BackColor="Yellow" ForeColor="Blue"
            Height="40px" Style="z-index: 100; left: 293px; position: absolute; top: 8px"
            Width="168px">
        </asp:DropDownList>
    </div>
    </form>
</body>
</html>
