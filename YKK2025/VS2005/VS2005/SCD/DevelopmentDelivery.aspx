<%@ Page Language="VB" AutoEventWireup="false" CodeFile="DevelopmentDelivery.aspx.vb" Inherits="DevelopmentDelivery" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>開發進度</title>
</head>
<body>

    <form id="form1" runat="server">
    <div>
<!-- *************************************************************************************************************** -->
<!-- #[狀態] -->
        <asp:TextBox ID="TextBox4" runat="server" BackColor="White" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 20px; position: absolute;
            top: 8px; text-align: left" Width="40px">狀態</asp:TextBox>

        <asp:DropDownList ID="DSts" runat="server" BackColor="Yellow" ForeColor="Blue" Height="26px"
            Style="z-index: 120; left: 67px; position: absolute; top: 8px" Width="170px">
            <asp:ListItem Selected="True" Value="ALL">ALL</asp:ListItem>
            <asp:ListItem Value="0">開發中</asp:ListItem>
            <asp:ListItem Value="1">OK</asp:ListItem>
            <asp:ListItem Value="2">NG</asp:ListItem>
            <asp:ListItem Value="3">取消</asp:ListItem>
        </asp:DropDownList>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- #[No] -->
        <asp:TextBox ID="TextBox1" runat="server" BackColor="White" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 244px; position: absolute;
            top: 8px; text-align: left" Width="40px">Ｎｏ</asp:TextBox>
        <asp:TextBox ID="DNo" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Height="18px" MaxLength="20" Style="z-index: 126; left: 288px; position: absolute;
            top: 8px; text-align: left" Width="166px"></asp:TextBox>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- #[開發No.] -->
        <asp:TextBox ID="TextBox3" runat="server" BackColor="White" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 465px; position: absolute;
            top: 8px; text-align: left" Width="51px">開發No.</asp:TextBox>
        <asp:TextBox ID="DDevNo" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" MaxLength="10" Style="z-index: 126; left: 524px;
            position: absolute; top: 8px; text-align: left" Width="114px"></asp:TextBox>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- #[Code] -->
        <asp:TextBox ID="TextBox2" runat="server" BackColor="White" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 651px; position: absolute;
            top: 8px; text-align: left" Width="34px">Code</asp:TextBox>
        <asp:TextBox ID="DCode" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" MaxLength="10" Style="z-index: 126; left: 704px;
            position: absolute; top: 8px; text-align: left" Width="93px"></asp:TextBox>
        <asp:TextBox ID="DORNO" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" MaxLength="10" Style="z-index: 126; left: 704px;
            position: absolute; top: 40px; text-align: left" Width="93px"></asp:TextBox>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- #[Buyer] -->
        <asp:TextBox ID="TextBox5" runat="server" BackColor="White" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 20px; position: absolute;
            top: 38px; text-align: left" Width="39px">Buyer</asp:TextBox>
        <asp:TextBox ID="DBuyer" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" MaxLength="10" Style="z-index: 126; left: 66px;
            position: absolute; top: 38px; text-align: left" Width="166px"></asp:TextBox>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- #[型別] -->
        <asp:TextBox ID="TextBox7" runat="server" BackColor="White" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 244px; position: absolute;
            top: 38px; text-align: left" Width="34px">型別</asp:TextBox>
        <asp:DropDownList ID="DSizeNo" runat="server" BackColor="Yellow" ForeColor="Blue" Height="26px"
            Style="z-index: 101; left: 288px; position: absolute; top: 38px" Width="171px">
        </asp:DropDownList>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- #[鏈條型式] -->
        <asp:TextBox ID="TextBox8" runat="server" BackColor="White" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 465px; position: absolute;
            top: 38px; text-align: left" Width="57px">鏈條型式</asp:TextBox>
        <asp:TextBox ID="TextBox11" runat="server" BackColor="White" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 656px; position: absolute;
            top: 40px; text-align: left" Width="40px">ORNO</asp:TextBox>
        <asp:TextBox ID="TextBox12" runat="server" BackColor="White" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 813px; position: absolute;
            top: 12px; text-align: left" Width="54px">委託廠商</asp:TextBox>
        <asp:DropDownList ID="DItem" runat="server" BackColor="Yellow" ForeColor="Blue"
            Height="26px" Style="z-index: 101; left: 525px; position: absolute; top: 38px"
            Width="120px">
        </asp:DropDownList>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- #[委託日] -->
        <asp:TextBox ID="TextBox6" runat="server" BackColor="White" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 20px; position: absolute;
            top: 68px; text-align: left" Width="39px">委託日</asp:TextBox>
        <asp:TextBox ID="DOrderStart" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" MaxLength="10" Style="z-index: 126; left: 66px;
            position: absolute; top: 68px; text-align: left" Width="67px"></asp:TextBox>
        <asp:TextBox ID="DOrderEnd" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" MaxLength="10" Style="z-index: 126; left: 165px;
            position: absolute; top: 68px; text-align: left" Width="67px"></asp:TextBox>
        <asp:TextBox ID="TextBox9" runat="server" BackColor="White" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 142px; position: absolute;
            top: 68px; text-align: left" Width="16px">～</asp:TextBox>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- #[希望完成] -->
        <asp:TextBox ID="TextBox10" runat="server" BackColor="White" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 244px; position: absolute;
            top: 68px; text-align: left" Width="55px">希望完成</asp:TextBox>
        <asp:TextBox ID="DFinishStart" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" MaxLength="10" Style="z-index: 126; left: 301px;
            position: absolute; top: 68px; text-align: left" Width="67px"></asp:TextBox>
        <asp:TextBox ID="DFinishEnd" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" MaxLength="10" Style="z-index: 126; left: 400px;
            position: absolute; top: 68px; text-align: left" Width="67px"></asp:TextBox>
        <asp:TextBox ID="TextBox13" runat="server" BackColor="White" BorderStyle="None" ForeColor="Black"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 377px; position: absolute;
            top: 68px; text-align: left" Width="16px">～</asp:TextBox>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- #[*] -->
        <asp:TextBox ID="DOP130" runat="server" BorderStyle="None" ForeColor="Black" BackColor="White"
            Height="18px" MaxLength="10" Style="z-index: 126; left: 800px; position: absolute;
            top: 68px; text-align: right" Width="144px">* 至130工程</asp:TextBox>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- #[現在] -->
        <asp:TextBox ID="DNowDateTime" runat="server" BorderStyle="None" ForeColor="Black" BackColor="White"
            Height="18px" MaxLength="10" Style="z-index: 126; left: 964px; position: absolute;
            top: 68px; text-align: right" Width="144px">2011/12/12 12:12 現在</asp:TextBox>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- #[GO] -->
        <asp:Button ID="BOK" runat="server" Height="25px" Style="z-index: 104; left: 526px;
            position: absolute; top: 68px" Text="GO" Width="62px" />
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- #[Excel] -->
        <asp:ImageButton ID="BExcel" runat="server" Height="21px" ImageUrl="Images\msexcel.gif"
            Style="z-index: 122; left: 596px; position: absolute; top: 68px" Width="21px" />
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- *************************************************************************************************************** -->
<!-- #[GRIDVIEW1] -->
        <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AutoGenerateColumns="False"
            BorderColor="#CC9966" BorderStyle="Groove" BorderWidth="1px" CellPadding="4" Font-Size="9pt"
            PageSize="15" Style="z-index: 114; left: 16px; position: absolute; top: 98px"
            Width="1300px">
            <RowStyle BackColor="White" ForeColor="Blue" />
            <Columns>
                <asp:BoundField DataField="STSD" HeaderText="狀態" />
                <asp:HyperLinkField DataNavigateUrlFields="URL" DataNavigateUrlFormatString="{0}"
                    DataTextField="NO" HeaderText="NO." Target="_blank">
                </asp:HyperLinkField>
                <asp:BoundField DataField="APPPER" HeaderText="委託人" />
                <asp:BoundField DataField="DEVNO" HeaderText="開發NO." />
                <asp:BoundField DataField="CODENO" HeaderText="CODE" />
                <asp:BoundField DataField="ORNO" HeaderText="ORNO" />
                <asp:BoundField DataField="BUYER" HeaderText="BUYER" />
                <asp:BoundField DataField="SIZE" HeaderText="型別" />
                <asp:BoundField DataField="ITEM" HeaderText="鏈條" />
                <asp:BoundField DataField="DATE" HeaderText="委託日" />
                <asp:HyperLinkField DataNavigateUrlFields="URL1" DataNavigateUrlFormatString="{0}"
                    DataTextField="ViewOP" Target="_blank">
                </asp:HyperLinkField>
                <asp:BoundField DataField="EDATE" HeaderText="希望完成" />
                <asp:BoundField DataField="CDATE" HeaderText="開發完成" />
                <asp:BoundField DataField="BDATE" HeaderText="*樣品預定完成" />
                <asp:BoundField DataField="ADATE" HeaderText="*樣品實際完成" />
                <asp:BoundField DataField="BDAYS" HeaderText="*預定日數" >
                    <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>
                <asp:BoundField DataField="ADAYS" HeaderText="*實際日數" >
                    <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>
                <asp:BoundField DataField="OP" HeaderText="進-工程" />
                <asp:BoundField DataField="OPBDATE" HeaderText="進-預定完成" />
                <asp:BoundField DataField="Sellvendor" HeaderText="委託廠商" />
            </Columns>
            <PagerStyle ForeColor="#8C4510" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#738A9C" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#FF8000" BorderStyle="Groove" Font-Bold="True" ForeColor="White" />
            <FooterStyle BorderStyle="Groove" />
        </asp:GridView>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- #[GRIDVIEW2] -->
        <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False"
            BorderColor="#CC9966" BorderStyle="Groove" BorderWidth="1px" CellPadding="4" Font-Size="9pt"
            PageSize="20" Style="z-index: 114; left: -5000px; position: absolute; top: 98px"
            Width="1300px">
            <RowStyle BackColor="White" ForeColor="Blue" />
            <Columns>
                <asp:BoundField DataField="STSD" HeaderText="狀態" />
                <asp:BoundField DataField="NO" HeaderText="NO." />
                <asp:BoundField DataField="APPPER" HeaderText="委託人" />
                <asp:BoundField DataField="DEVNO" HeaderText="開發NO." />
                <asp:BoundField DataField="CODENO" HeaderText="CODE" />
                <asp:BoundField DataField="BUYER" HeaderText="BUYER" />
                <asp:BoundField DataField="SIZE" HeaderText="型別" />
                <asp:BoundField DataField="ITEM" HeaderText="鏈條" />
                <asp:BoundField DataField="DATE" HeaderText="委託日" />
                <asp:BoundField DataField="EDATE" HeaderText="希望完成" />
                <asp:BoundField DataField="CDATE" HeaderText="開發完成" />               
                <asp:BoundField DataField="BDATE" HeaderText="*樣品預定完成" />
                <asp:BoundField DataField="ADATE" HeaderText="*樣品實際完成" />
                <asp:BoundField DataField="BDAYS" HeaderText="*預定日數" >
                    <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>
                <asp:BoundField DataField="ADAYS" HeaderText="*實際日數" >
                    <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>
                <asp:BoundField DataField="OP" HeaderText="進-工程" />
                <asp:BoundField DataField="OPBDATE" HeaderText="進-預定完成" />
                <asp:BoundField DataField="Sellvendor" HeaderText="委託廠商" />
            </Columns>
            <PagerStyle ForeColor="#8C4510" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#738A9C" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#FF8000" BorderStyle="Groove" Font-Bold="True" ForeColor="White" />
            <FooterStyle BorderStyle="Groove" />
        </asp:GridView>
        <asp:TextBox ID="DSellvendor" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" MaxLength="10" Style="z-index: 126; left: 871px;
            position: absolute; top: 9px; text-align: left" Width="78px"></asp:TextBox>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- *************************************************************************************************************** -->
    </div>
    </form>

</body>
</html>
