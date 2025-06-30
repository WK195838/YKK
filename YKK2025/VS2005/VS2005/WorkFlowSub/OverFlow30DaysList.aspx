<%@ Page Language="VB" AutoEventWireup="false" CodeFile="OverFlow30DaysList.aspx.vb" Inherits="OverFlow30DaysList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>逾期開發案件</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <div style="display: inline; z-index: 116; left: 13px;
            width: 81px; color: blue; position: absolute; top: 8px; height: 24px" >
            篩選項目：</div>
        <asp:DropDownList ID="DFormNo" runat="server" AutoPostBack="True" BackColor="Yellow"
            ForeColor="Blue" Height="40px" Style="z-index: 100; left: 96px; position: absolute;
            top: 9px" Width="174px">
            <asp:ListItem Selected="True" Value="000001">圖面委託書</asp:ListItem>
            <asp:ListItem Value="000002">圖面修改委託書</asp:ListItem>
            <asp:ListItem Value="000003">內製委託書</asp:ListItem>
            <asp:ListItem Value="000007">外注委託書</asp:ListItem>
            <asp:ListItem Value="000014">表面處理委託書</asp:ListItem>
        </asp:DropDownList>
        <asp:DropDownList ID="DDays" runat="server" AutoPostBack="True" BackColor="Yellow"
            ForeColor="Blue" Height="40px" Style="z-index: 100; left: 276px; position: absolute;
            top: 9px" Width="111px">
            <asp:ListItem>1</asp:ListItem>
            <asp:ListItem>2</asp:ListItem>
            <asp:ListItem>3</asp:ListItem>
            <asp:ListItem>4</asp:ListItem>
            <asp:ListItem>5</asp:ListItem>
            <asp:ListItem>6</asp:ListItem>
            <asp:ListItem>7</asp:ListItem>
            <asp:ListItem>8</asp:ListItem>
            <asp:ListItem>9</asp:ListItem>
            <asp:ListItem>10</asp:ListItem>
            <asp:ListItem>11</asp:ListItem>
            <asp:ListItem>12</asp:ListItem>
            <asp:ListItem>13</asp:ListItem>
            <asp:ListItem>14</asp:ListItem>
            <asp:ListItem Selected="True">15</asp:ListItem>
            <asp:ListItem>16</asp:ListItem>
            <asp:ListItem>17</asp:ListItem>
            <asp:ListItem>18</asp:ListItem>
            <asp:ListItem>19</asp:ListItem>
            <asp:ListItem>20</asp:ListItem>
            <asp:ListItem>21</asp:ListItem>
            <asp:ListItem>22</asp:ListItem>
            <asp:ListItem>23</asp:ListItem>
            <asp:ListItem>24</asp:ListItem>
            <asp:ListItem>25</asp:ListItem>
            <asp:ListItem>26</asp:ListItem>
            <asp:ListItem>27</asp:ListItem>
            <asp:ListItem>28</asp:ListItem>
            <asp:ListItem>29</asp:ListItem>
            <asp:ListItem>30</asp:ListItem>
        </asp:DropDownList>
            <asp:Button ID="Go" runat="server" BackColor="White" ForeColor="Blue" Height="24px"
            Style="z-index: 114; left: 511px; position: absolute; top: 8px" Text="Go" Width="40px" />
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" BorderColor="#CC9966"
            BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt" ShowFooter="True"
            Style="z-index: 114; left: 13px; position: absolute; top: 41px" Width="1000px">
            <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
            <Columns>
                <asp:BoundField DataField="Field1" HeaderText="Field1" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                    <HeaderStyle Width="40px" />
                </asp:BoundField>
                <asp:BoundField DataField="Field2" HeaderText="Field2" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Left" />
                    <HeaderStyle Width="150px" />
                </asp:BoundField>
                <asp:BoundField DataField="Field3" HeaderText="Field3" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Left" />
                    <HeaderStyle Width="150px" />
                </asp:BoundField>
                <asp:BoundField DataField="Field4" HeaderText="Field4" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                    <HeaderStyle Width="50px" />
                </asp:BoundField>
                <asp:BoundField DataField="Field5" HeaderText="Field5" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                    <HeaderStyle Width="50px" />
                </asp:BoundField>
                <asp:BoundField DataField="Field6" HeaderText="Field6" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center" />
                    <HeaderStyle Width="50px" />
                </asp:BoundField>                
                <asp:BoundField DataField="Field7" HeaderText="Field7" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Left" />
                    <HeaderStyle Width="80px" />
                </asp:BoundField>
                <asp:BoundField DataField="Field8" HeaderText="Field8" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Left" />
                    <HeaderStyle Width="40px" />
                </asp:BoundField>
                <asp:BoundField DataField="Field9" HeaderText="Field9" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Left" />
                    <HeaderStyle Width="60px" />
                </asp:BoundField>
                <asp:BoundField DataField="Field10" HeaderText="Field10" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Left" />
                    <HeaderStyle Width="300px" />
                </asp:BoundField>
            </Columns>
            <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
        </asp:GridView>
        <asp:ImageButton ID="BExcel" runat="server" Height="21px" ImageUrl="Images\msexcel.gif"
            Style="z-index: 122; left: 561px; position: absolute; top: 11px" Width="21px" /><asp:DropDownList ID="DVType" runat="server" AutoPostBack="True" BackColor="Yellow"
            ForeColor="Blue" Height="40px" Style="z-index: 100; left: 392px; position: absolute;
            top: 9px" Width="111px">
                <asp:ListItem Selected="True" Value="1">不含假日</asp:ListItem>
                <asp:ListItem Value="0">含假日</asp:ListItem>
            </asp:DropDownList>

   
    </div>
    </form>
</body>
</html>
