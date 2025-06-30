<%@ Page Language="VB" AutoEventWireup="false" CodeFile="DevelopmentTimeList.aspx.vb" Inherits="DevelopmentTimeList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
		<title>開發納期一覽</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <div ms_positioning="FlowLayout" style="display: inline; z-index: 105; left: 8px;
            width: 96px; color: #0000ff; position: absolute; top: 8px; height: 24px" title="委託年月：">
            委託年月：</div>
        <asp:DropDownList ID="DSYY" runat="server" BackColor="Yellow"
            ForeColor="Blue" Height="40px" Style="z-index: 104; left: 104px; position: absolute;
            top: 8px" Width="72px">
        </asp:DropDownList>
        <div ms_positioning="FlowLayout" style="display: inline; z-index: 103; left: 176px;
            width: 1px; color: #0000ff; position: absolute; top: 8px; height: 24px" title="">
            年</div>
        <asp:DropDownList ID="DSMM" runat="server" BackColor="Yellow"
            ForeColor="Blue" Height="40px" Style="z-index: 102; left: 200px; position: absolute;
            top: 8px" Width="48px">
        </asp:DropDownList>
        <div ms_positioning="FlowLayout" style="display: inline; z-index: 101; left: 248px;
            width: 8px; color: #0000ff; position: absolute; top: 8px; height: 24px" title="">
            <p>月</p>
        </div>
        <div id="DIV1" ms_positioning="FlowLayout" style="display: inline; z-index: 115;
            left: 280px; width: 8px; color: #0000ff; position: absolute; top: 8px; height: 24px"
            title="">
            <p> ～</p>
        </div>
        <asp:DropDownList ID="DEYY" runat="server" BackColor="Yellow"
            ForeColor="Blue" Height="40px" Style="z-index: 111; left: 312px; position: absolute;
            top: 8px" Width="72px">
        </asp:DropDownList>
        <div ms_positioning="FlowLayout" style="display: inline; z-index: 112; left: 384px;
            width: 1px; color: #0000ff; position: absolute; top: 8px; height: 24px" title="">
            年</div>
        <asp:DropDownList ID="DEMM" runat="server" BackColor="Yellow"
            ForeColor="Blue" Height="40px" Style="z-index: 113; left: 408px; position: absolute;
            top: 8px" Width="48px">
        </asp:DropDownList>
        <div ms_positioning="FlowLayout" style="display: inline; z-index: 114; left: 456px;
            width: 16px; color: #0000ff; position: absolute; top: 8px; height: 24px" title="">
            <p>
                月</p>
        </div>
        <div ms_positioning="FlowLayout" style="display: inline; z-index: 107; left: 8px;
            width: 80px; color: #0000ff; position: absolute; top: 36px; height: 24px" title="BUYER：">
            <p>BUYER：</p>
        </div>
        <asp:DropDownList ID="DBuyer" runat="server" BackColor="Yellow" ForeColor="Blue"
            Height="40px" Style="z-index: 106; left: 104px; position: absolute; top: 36px"
            Width="144px">
        </asp:DropDownList>
        <asp:Button ID="Go" runat="server" BackColor="White" ForeColor="Blue" Height="24px"
            Style="z-index: 108; left: 272px; position: absolute; top: 36px" Text="Go" Width="40px" />
        <asp:ImageButton ID="BExcel" runat="server" Height="21px" ImageUrl="Images\msexcel.gif"
            Style="z-index: 110; left: 336px; position: absolute; top: 36px" Width="21px" />
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" BorderColor="#CC9966"
            BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt" Style="z-index: 114;
            left: 8px; position: absolute; top: 65px">
            <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
            <Columns>
                <asp:BoundField DataField="MNo" HeaderText="No" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="Buyer" HeaderText="Buyer" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="MapNo" HeaderText="圖號" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="MLevel" HeaderText="難易度" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                    <HeaderStyle BackColor="Gray" />
                </asp:BoundField>
                <asp:BoundField DataField="MapCreateTime" HeaderText="起始日" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                    <HeaderStyle BackColor="Gray" />
                </asp:BoundField>
                <asp:BoundField DataField="MapFinishTime" HeaderText="完成日" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                    <HeaderStyle BackColor="Gray" />
                </asp:BoundField>
                <asp:BoundField DataField="MapModTime" HeaderText="開發時間" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                    <HeaderStyle BackColor="Gray" />
                </asp:BoundField>
                <asp:BoundField DataField="MapModCount" HeaderText="次數" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                    <HeaderStyle BackColor="Gray" />
                </asp:BoundField>
                <asp:BoundField DataField="MapSts" HeaderText="狀態" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                    <HeaderStyle BackColor="Gray" />
                </asp:BoundField>

                <asp:HyperLinkField DataNavigateUrlFields="MISURL" DataNavigateUrlFormatString="{0}"
                    DataTextField="ManufInCreateTime" HeaderText="起始日" Target="_blank">
                    <FooterStyle HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#00C0C0" />
                </asp:HyperLinkField>
                <asp:BoundField DataField="ManufInFinishTime" HeaderText="完成日" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#00C0C0" />
                </asp:BoundField>
                <asp:BoundField DataField="ManufInModTime" HeaderText="開發時間" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#00C0C0" />
                </asp:BoundField>
                <asp:BoundField DataField="ManufInModCount" HeaderText="次數" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#00C0C0" />
                </asp:BoundField>
                <asp:BoundField DataField="ManufInSts" HeaderText="狀態" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#00C0C0" />
                </asp:BoundField>
                <asp:BoundField DataField="MapToInET" HeaderText="繪內-開發時間" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#00C0C0" />
                </asp:BoundField>

                <asp:HyperLinkField DataNavigateUrlFields="MOSURL" DataNavigateUrlFormatString="{0}"
                    DataTextField="ManufOutCreateTime" HeaderText="起始日" Target="_blank">
                    <FooterStyle HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#C0C000" />
                </asp:HyperLinkField>
                <asp:BoundField DataField="ManufOutFinishTime" HeaderText="完成日" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#C0C000" />
                </asp:BoundField>
                <asp:BoundField DataField="ManufOutModTime" HeaderText="開發時間" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#C0C000" />
                </asp:BoundField>
                <asp:BoundField DataField="ManufOutModCount" HeaderText="次數" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#C0C000" />
                </asp:BoundField>
                <asp:BoundField DataField="ManufOutSts" HeaderText="狀態" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#C0C000" />
                </asp:BoundField>
                <asp:BoundField DataField="MapToOutET" HeaderText="繪外-開發時間" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#C0C000" />
                </asp:BoundField>
            </Columns>
            <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
        </asp:GridView>
        &nbsp; &nbsp;
        <asp:TextBox ID="TextBox2" runat="server" BackColor="White" BorderStyle="None" Font-Size="10pt"
            ForeColor="#C00000" Height="16px" ReadOnly="True" Style="z-index: 111; left: 503px;
            position: absolute; top: 38px" Width="40px">共通：</asp:TextBox>
        <asp:TextBox ID="TextBox4" runat="server" BackColor="White" BorderStyle="None" Font-Size="10pt"
            ForeColor="#C00000" Height="16px" ReadOnly="True" Style="z-index: 111; left: 581px;
            position: absolute; top: 38px" Width="40px">繪圖：</asp:TextBox>
        <asp:TextBox ID="TextBox5" runat="server" BackColor="White" BorderStyle="None" Font-Size="10pt"
            ForeColor="#C00000" Height="16px" ReadOnly="True" Style="z-index: 111; left: 660px;
            position: absolute; top: 38px" Width="40px">內製：</asp:TextBox>
        <asp:TextBox ID="TextBox6" runat="server" BackColor="White" BorderStyle="None" Font-Size="10pt"
            ForeColor="#C00000" Height="16px" ReadOnly="True" Style="z-index: 111; left: 739px;
            position: absolute; top: 38px" Width="40px">外注：</asp:TextBox>
        <asp:TextBox ID="TextBox3" runat="server" BackColor="#C04000" BorderStyle="None"
            Font-Size="10pt" Height="16px" ReadOnly="True" Style="z-index: 111; left: 546px;
            position: absolute; top: 38px" Width="20px"></asp:TextBox>
        <asp:TextBox ID="TextBox7" runat="server" BackColor="Gray" BorderStyle="None" Font-Size="10pt"
            Height="16px" ReadOnly="True" Style="z-index: 111; left: 624px; position: absolute;
            top: 38px" Width="20px"></asp:TextBox>
        <asp:TextBox ID="TextBox8" runat="server" BackColor="#00C0C0" BorderStyle="None"
            Font-Size="10pt" Height="16px" ReadOnly="True" Style="z-index: 111; left: 703px;
            position: absolute; top: 38px" Width="20px"></asp:TextBox>
        <asp:TextBox ID="TextBox9" runat="server" BackColor="#C0C000" BorderStyle="None"
            Font-Size="10pt" Height="16px" ReadOnly="True" Style="z-index: 111; left: 782px;
            position: absolute; top: 38px" Width="20px"></asp:TextBox>
        <asp:TextBox ID="TextBox1" runat="server" BackColor="White" BorderStyle="None" Font-Size="10pt"
            ForeColor="#C00000" Height="16px" ReadOnly="True" Style="z-index: 111; left: 502px;
            position: absolute; top: 13px" Width="103px">時間單位：H</asp:TextBox>
        <asp:TextBox ID="DLastUpdateTime" runat="server" BackColor="White" BorderStyle="None"
            Font-Size="10pt" ForeColor="#C00000" Height="16px" ReadOnly="True" Style="z-index: 111;
            left: 660px; position: absolute; top: 13px" Width="296px">時間單位：H</asp:TextBox>
        <asp:HyperLink ID="LFormHelp" runat="server" Font-Size="12pt" Height="8px" NavigateUrl="~/Images/DevelopmentTimeList_001.jpg"
            Style="z-index: 281; left: 844px; position: absolute; top: 38px" Target="_blank"
            Width="80px">表格說明</asp:HyperLink>
    </div>
    </form>
</body>
</html>
