<%@ Page Language="VB" AutoEventWireup="false" CodeFile="HR_OverTimeReport03.aspx.vb" Inherits="HR_OverTimeReport03" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>個人加班資料</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:ImageButton ID="BExcel" runat="server" Height="21px" ImageUrl="Images\msexcel.gif"
            Style="z-index: 122; left: 10px; position: absolute; top: 7px" Width="21px" Visible="False" />
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" BorderColor="#CC9966"
            BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt" ShowFooter="True"
            Style="z-index: 114; left: 10px; position: absolute; top: 33px" Width="790px">
            <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
            <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
            <Columns>
                <asp:BoundField DataField="DivisionDesc" HeaderText="部門" ReadOnly="True">
                </asp:BoundField>
                <asp:BoundField DataField="NameDesc" HeaderText="姓名" ReadOnly="True">
                </asp:BoundField>

                <asp:HyperLinkField DataNavigateUrlFields="YMDURL" DataNavigateUrlFormatString="{0}"
                    DataTextField="YMDDesc" HeaderText="日期" Target="_blank">
                </asp:HyperLinkField>

                <asp:BoundField DataField="FoodDesc" HeaderText="伙食" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>
                <asp:BoundField DataField="TrafficDesc" HeaderText="交通" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>
                <asp:BoundField DataField="Total" HeaderText="核定時數" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>
                <asp:BoundField DataField="CVTime"  HeaderText="換休時數" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>
                
                <asp:BoundField DataField="TwoIn" HeaderText="２內" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>
                <asp:BoundField DataField="TwoOut" HeaderText="２外" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>
                <asp:BoundField DataField="Vacation" HeaderText="假日" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>
                <asp:BoundField DataField="CVacation" HeaderText="國假" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>
            </Columns>

        </asp:GridView><asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False"
            BorderColor="#CC9966" BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt"
            Style="z-index: 114; left: 11px; position: absolute; top: -10000px" Width="790px" ShowFooter="True">
            <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
            <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
            <Columns>
                <asp:BoundField DataField="DivisionDesc" HeaderText="部門" ReadOnly="True" />
                <asp:BoundField DataField="NameDesc" HeaderText="姓氏" ReadOnly="True" />
                
                <asp:BoundField DataField="YMDDesc" HeaderText="日期" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                    
                <asp:BoundField DataField="FoodDesc" HeaderText="伙食" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>
                <asp:BoundField DataField="TrafficDesc" HeaderText="交通" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>
                <asp:BoundField DataField="Total" HeaderText="核定時數" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>
                <asp:BoundField DataField="CVTime"  HeaderText="換休時數" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>
                
                <asp:BoundField DataField="TwoIn" HeaderText="２內" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>
                <asp:BoundField DataField="TwoOut" HeaderText="２外" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>
                <asp:BoundField DataField="Vacation" HeaderText="假日" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>
                <asp:BoundField DataField="CVacation" HeaderText="國假" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Right" />
                    <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>
            </Columns>
        </asp:GridView>
   
    </div>
    </form>
</body>

</html>
