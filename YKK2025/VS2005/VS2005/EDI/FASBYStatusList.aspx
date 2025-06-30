<%@ Page Language="VB" AutoEventWireup="false" CodeFile="FASBYStatusList.aspx.vb" Inherits="FASBYStatusList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Status List</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  底圖                                                                                ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <img src="iMages/BYStatusList-1.png" style="z-index: 1; left: 0px; position: absolute;top: 5px; width: 846px; height: 152px;"/>
        <img src="iMages/BYStatusList.png" style="z-index: 1; left: 0px; position: absolute;top: 155px; width: 846px; height: 557px;"/>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Exce                                                                                ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:ImageButton ID="BExcel" runat="server" Height="21px" ImageUrl="Images\msexcel.gif"
            Style="z-index: 122; left: 694px; position: absolute; top: 163px" Width="21px" />
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Gridview (Excel)                                                                    ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="GridView1" runat="server" AllowSorting="True" AutoGenerateColumns="False" BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt"
            PageSize="25" Style="z-index: 114; left: 4px; position: absolute;
            top: 10px" Width="806px">
            <RowStyle BackColor="White" ForeColor="Blue" />
            <Columns>
                <asp:BoundField DataField="CheckSeq" HeaderText="檢測順位" />
                <asp:BoundField DataField="EZ1" HeaderText="Status" />
                <asp:BoundField DataField="STSDESCR" HeaderText="說明" />
                <asp:BoundField DataField="REMARK" HeaderText="理由" />
                <asp:BoundField DataField="RCOUNT" HeaderText="件數"/>
            </Columns>
            <PagerStyle ForeColor="#8C4510" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#738A9C" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="DarkGray" BorderStyle="Groove" Font-Bold="True" />
            <FooterStyle BackColor="#FFFFCC" BorderStyle="Groove" HorizontalAlign="Right" />
        </asp:GridView>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  隱藏欄位                                                                       ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:TextBox ID="DBYCount" runat="server" Height="16px" Style="z-index: 318; left: 762px;text-align:right;
            position: absolute; top: 13px;font-size:16px;background:transparent" Width="72px"   BorderStyle="None" BorderWidth="0px" Font-Bold="False" Font-Overline="False" Font-Size="22pt" Font-Strikeout="False" Font-Underline="True"></asp:TextBox>

        <asp:TextBox ID="DSUMCount" runat="server" Height="16px" Style="z-index: 318; left: 762px;text-align:right;
            position: absolute; top: 129px;font-size:16px;background:transparent" Width="72px"   BorderStyle="None" BorderWidth="0px" Font-Bold="False" Font-Overline="False" Font-Size="22pt" Font-Strikeout="False" Font-Underline="True"></asp:TextBox>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->

        <asp:TextBox ID="D10Reason" runat="server" Height="16px" Style="z-index: 318; left: 362px;text-align:left;
            position: absolute; top: 223px;font-size:16px;background:transparent" Width="388px"   BorderStyle="None" BorderWidth="0px" Font-Bold="False" Font-Overline="False" Font-Size="22pt" Font-Strikeout="False" ></asp:TextBox>

        <asp:TextBox ID="D10" runat="server" Height="16px" Style="z-index: 318; left: 762px;text-align:right;
            position: absolute; top: 223px;font-size:16px;background:transparent" Width="72px"   BorderStyle="None" BorderWidth="0px" Font-Bold="False" Font-Overline="False" Font-Size="22pt" Font-Strikeout="False" Font-Underline="True"></asp:TextBox>

        <asp:TextBox ID="D20" runat="server" Height="16px" Style="z-index: 318; left: 762px;text-align:right;
            position: absolute; top: 252px;font-size:16px;background:transparent" Width="72px"   BorderStyle="None" BorderWidth="0px" Font-Bold="False" Font-Overline="False" Font-Size="22pt" Font-Strikeout="False" Font-Underline="True"></asp:TextBox>

        <asp:TextBox ID="D40" runat="server" Height="16px" Style="z-index: 318; left: 762px;text-align:right;
            position: absolute; top: 281px;font-size:16px;background:transparent" Width="72px"   BorderStyle="None" BorderWidth="0px" Font-Bold="False" Font-Overline="False" Font-Size="22pt" Font-Strikeout="False" Font-Underline="True"></asp:TextBox>

        <asp:TextBox ID="D41" runat="server" Height="16px" Style="z-index: 318; left: 762px;text-align:right;
            position: absolute; top: 310px;font-size:16px;background:transparent" Width="72px"   BorderStyle="None" BorderWidth="0px" Font-Bold="False" Font-Overline="False" Font-Size="22pt" Font-Strikeout="False" Font-Underline="True"></asp:TextBox>

        <asp:TextBox ID="D42" runat="server" Height="16px" Style="z-index: 318; left: 762px;text-align:right;
            position: absolute; top: 339px;font-size:16px;background:transparent" Width="72px"   BorderStyle="None" BorderWidth="0px" Font-Bold="False" Font-Overline="False" Font-Size="22pt" Font-Strikeout="False" Font-Underline="True"></asp:TextBox>

        <asp:TextBox ID="D50" runat="server" Height="16px" Style="z-index: 318; left: 762px;text-align:right;
            position: absolute; top: 368px;font-size:16px;background:transparent" Width="72px"   BorderStyle="None" BorderWidth="0px" Font-Bold="False" Font-Overline="False" Font-Size="22pt" Font-Strikeout="False" Font-Underline="True"></asp:TextBox>

        <asp:TextBox ID="D52" runat="server" Height="16px" Style="z-index: 318; left: 762px;text-align:right;
            position: absolute; top: 397px;font-size:16px;background:transparent" Width="72px"   BorderStyle="None" BorderWidth="0px" Font-Bold="False" Font-Overline="False" Font-Size="22pt" Font-Strikeout="False" Font-Underline="True"></asp:TextBox>

        <asp:TextBox ID="D30" runat="server" Height="16px" Style="z-index: 318; left: 762px;text-align:right;
            position: absolute; top: 426px;font-size:16px;background:transparent" Width="72px"   BorderStyle="None" BorderWidth="0px" Font-Bold="False" Font-Overline="False" Font-Size="22pt" Font-Strikeout="False" Font-Underline="True"></asp:TextBox>

        <asp:TextBox ID="D80" runat="server" Height="16px" Style="z-index: 318; left: 762px;text-align:right;
            position: absolute; top: 455px;font-size:16px;background:transparent" Width="72px"   BorderStyle="None" BorderWidth="0px" Font-Bold="False" Font-Overline="False" Font-Size="22pt" Font-Strikeout="False" Font-Underline="True"></asp:TextBox>

        <asp:TextBox ID="D81" runat="server" Height="16px" Style="z-index: 318; left: 762px;text-align:right;
            position: absolute; top: 484px;font-size:16px;background:transparent" Width="72px"   BorderStyle="None" BorderWidth="0px" Font-Bold="False" Font-Overline="False" Font-Size="22pt" Font-Strikeout="False" Font-Underline="True"></asp:TextBox>

        <asp:TextBox ID="D60" runat="server" Height="16px" Style="z-index: 318; left: 762px;text-align:right;
            position: absolute; top: 513px;font-size:16px;background:transparent" Width="72px"   BorderStyle="None" BorderWidth="0px" Font-Bold="False" Font-Overline="False" Font-Size="22pt" Font-Strikeout="False" Font-Underline="True"></asp:TextBox>

        <asp:TextBox ID="D61" runat="server" Height="16px" Style="z-index: 318; left: 762px;text-align:right;
            position: absolute; top: 542px;font-size:16px;background:transparent" Width="72px"   BorderStyle="None" BorderWidth="0px" Font-Bold="False" Font-Overline="False" Font-Size="22pt" Font-Strikeout="False" Font-Underline="True"></asp:TextBox>

        <asp:TextBox ID="D70" runat="server" Height="16px" Style="z-index: 318; left: 762px;text-align:right;
            position: absolute; top: 571px;font-size:16px;background:transparent" Width="72px"   BorderStyle="None" BorderWidth="0px" Font-Bold="False" Font-Overline="False" Font-Size="22pt" Font-Strikeout="False" Font-Underline="True"></asp:TextBox>

        <asp:TextBox ID="D71" runat="server" Height="16px" Style="z-index: 318; left: 762px;text-align:right;
            position: absolute; top: 600px;font-size:16px;background:transparent" Width="72px"   BorderStyle="None" BorderWidth="0px" Font-Bold="False" Font-Overline="False" Font-Size="22pt" Font-Strikeout="False" Font-Underline="True"></asp:TextBox>

        <asp:TextBox ID="D01" runat="server" Height="16px" Style="z-index: 318; left: 762px;text-align:right;
            position: absolute; top: 629px;font-size:16px;background:transparent" Width="72px"   BorderStyle="None" BorderWidth="0px" Font-Bold="False" Font-Overline="False" Font-Size="22pt" Font-Strikeout="False" Font-Underline="True"></asp:TextBox>

        <asp:TextBox ID="D00" runat="server" Height="16px" Style="z-index: 318; left: 762px;text-align:right;
            position: absolute; top: 658px;font-size:16px;background:transparent" Width="72px"   BorderStyle="None" BorderWidth="0px" Font-Bold="False" Font-Overline="False" Font-Size="22pt" Font-Strikeout="False" Font-Underline="True"></asp:TextBox>

        <asp:TextBox ID="DTTL" runat="server" Height="16px" Style="z-index: 318; left: 762px;text-align:right;
            position: absolute; top: 687px;font-size:16px;background:transparent" Width="72px"   BorderStyle="None" BorderWidth="0px" Font-Bold="False" Font-Overline="False" Font-Size="22pt" Font-Strikeout="False" Font-Underline="True"></asp:TextBox>
    </div>
    </form>
</body>
</html>
