<%@ Page Language="VB" Debug="true" CodeFile="FASSheet_01.aspx.vb" Inherits="FASSheet_01" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>備料申請單</title>
	 <script language="javascript"> 

    </script>
 
	 	
</head>
<body>
    <form id="form1" runat="server">
    <div>
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        <asp:Button ID="BSAVE" runat="server" Height="23px" Style="z-index: 100;
            left: 354px; position: absolute; top: 1144px" Text="儲存" Width="80px" />
        <asp:Button ID="BNG2" runat="server" Height="23px" Style="z-index: 101;
            left: 445px; position: absolute; top: 1144px" Text="NG2" Width="80px" />
        <asp:Button ID="BNG1" runat="server" Height="23px" Style="z-index: 102;
            left: 537px; position: absolute; top: 1144px" Text="NG1" Width="80px" />
        <asp:Button ID="BOK" runat="server" Height="23px" Style="z-index: 103;
            left: 629px; position: absolute; top: 1144px" Text="OK" Width="80px" />
        <img id="DDelay" runat="server" src="images/Sheet_Delay.jpg" style="z-index: 121; left: 12px;
            position: absolute; top: 1034px" visible="false" />
        <asp:TextBox ID="DDecideDesc" runat="server" BackColor="Yellow" ForeColor="Blue"
            Height="56px" Style="z-index: 104; left: 59px; position: absolute; top: 965px"
            TextMode="MultiLine" Width="536px"></asp:TextBox>
        <asp:TextBox ID="DReason" runat="server" BackColor="Yellow" ForeColor="Blue" Height="20px"
            Style="z-index: 105; left: 243px; position: absolute; top: 1039px" Visible="False"
            Width="352px"></asp:TextBox>
        <asp:DropDownList ID="DReasonCode" runat="server" Height="20px" Style="z-index: 106;
            left: 167px; position: absolute; top: 1042px" Visible="False" Width="64px" BackColor="Yellow">
            <asp:ListItem>01</asp:ListItem>
            <asp:ListItem>02</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DReasonDesc" runat="server" BackColor="Yellow" ForeColor="Blue"
            Height="56px" Style="z-index: 107; left: 171px; position: absolute; top: 1072px"
            Visible="False" Width="424px"></asp:TextBox>
        <asp:Label ID="DHistoryLabel" runat="server" Font-Size="11pt" ForeColor="Blue" Height="20px"
            Style="z-index: 108; left: 14px; position: absolute; top: 1157px" Text="核定履歷"></asp:Label>
        <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" BackColor="White"
            BorderColor="#CC9966" BorderStyle="None" BorderWidth="1px" CellPadding="4" Height="100px"
            Style="z-index: 109; left: 13px; position: absolute; top: 1179px" Width="780px">
            <RowStyle BackColor="White" Font-Size="9pt" ForeColor="#330099" />
            <Columns>
                <asp:BoundField DataField="StepNameDesc" HeaderText="工程">
                    <ItemStyle HorizontalAlign="Left" />
                </asp:BoundField>
                <asp:BoundField DataField="DecideName" HeaderText="擔當" />
                <asp:BoundField DataField="AgentName" HeaderText="代理/兼職" />
                <asp:BoundField DataField="FlowTypeDesc" HeaderText="類別" />
                <asp:BoundField DataField="StsDesc" HeaderText="處理結果" />
                <asp:BoundField DataField="DecideDescA" HeaderText="說明">
                    <ItemStyle HorizontalAlign="Left" />
                </asp:BoundField>
                <asp:BoundField DataField="Description" HeaderText="核定時間">
                    <ItemStyle HorizontalAlign="Left" />
                </asp:BoundField>
            </Columns>
            <HeaderStyle BackColor="#CC6600" Font-Size="9pt" ForeColor="#FFFFCC" HorizontalAlign="Center"
                VerticalAlign="Middle" />
        </asp:GridView>
        <img id="DDescSheet" runat="server" src="images/Sheet_Remark.jpg" style="z-index: 99;
            left: 13px; position: absolute; top: 959px" />
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        <asp:Image ID="Image1" runat="server" ImageUrl="~/images/FASSheet_01.jpg" Style="z-index: 110;
            left: 5px; position: absolute; top: 7px" />
        <asp:TextBox ID="DRemark" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
            Height="20px" MaxLength="100" Style="z-index: 111; left: 217px; position: absolute;
            top: 142px; text-align: left" TextMode="MultiLine" Width="556px"></asp:TextBox>
        &nbsp;&nbsp;&nbsp; &nbsp;
        <asp:DropDownList ID="DTYPE" runat="server" BackColor="#C0FFFF"
            ForeColor="Blue" Height="20px" Style="z-index: 112; left: 220px; position: absolute;
            top: 115px" Width="206px">
        </asp:DropDownList>
        &nbsp; &nbsp;
        &nbsp; &nbsp;&nbsp;
        <asp:TextBox ID="DApper" runat="server" AutoPostBack="True" BackColor="#C0FFFF"
            BorderStyle="None" ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 113;
            left: 553px; position: absolute; top: 65px; text-align: left" Width="220px"></asp:TextBox>
        <asp:TextBox ID="DDivision" runat="server" AutoPostBack="True" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 114; left: 219px;
            position: absolute; top: 90px; text-align: left" Width="201px"></asp:TextBox>
        <asp:TextBox ID="DAppDate" runat="server" AutoPostBack="True" BackColor="#C0FFFF"
            BorderStyle="None" ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 115;
            left: 552px; position: absolute; top: 90px; text-align: left" Width="220px"></asp:TextBox>
        <asp:TextBox ID="DNo" runat="server" BackColor="White" BorderColor="White" BorderStyle="None"
            Font-Names="Times New Roman" Font-Size="11pt" Style="z-index: 116; left: 218px;
            position: absolute; top: 65px" Width="198px">DNo</asp:TextBox>
        &nbsp;
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False"
            BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt" ShowFooter="True"
            Style="z-index: 117; left: 24px; position: absolute; top: 262px" Width="767px">
            <Columns>
                <asp:BoundField DataField="BuyerCode" HeaderText="客戶-BUYER" >
                    <ItemStyle Width="100px" />
                </asp:BoundField>
                <asp:BoundField DataField="ItemCode" HeaderText="ITEM CODE" >
                    <ItemStyle Width="100px" />
                </asp:BoundField>
                <asp:BoundField DataField="Qty" HeaderText="QTY" DataFormatString="{0:N0}" >
                    <ItemStyle Width="100px" />
                </asp:BoundField>
                <asp:BoundField DataField="AMT" DataFormatString="{0:N0}" HeaderText="金額" >
                    <ItemStyle Width="100px" />
                </asp:BoundField>
                <asp:BoundField DataField="INVQTY" HeaderText="KEEP在庫數" DataFormatString="{0:N0}" >
                    <ItemStyle Width="100px" />
                </asp:BoundField>
                <asp:BoundField DataField="INVAMT" HeaderText="KEEP在庫金額" DataFormatString="{0:N0}" >
                    <ItemStyle Width="120px" />
                </asp:BoundField>
                <asp:TemplateField HeaderText="附檔">
                    <ItemTemplate>
                        <asp:FileUpload ID="FileUpload1" runat="server" Width="222px"/>
                    </ItemTemplate>
                    <ItemStyle Width="10px" />
                </asp:TemplateField>
            </Columns>
            <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
            <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
        </asp:GridView><asp:GridView ID="GridView3" runat="server" AutoGenerateColumns="False" BorderColor="#CC9966"
            BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt" ShowFooter="True"
            Style="z-index: 118; left: 24px; position: absolute; top: 263px" Width="767px">
            <Columns>
                <asp:BoundField DataField="Buyer" HeaderText="客戶-BUYER" >
                    <ItemStyle Width="100px" />
                </asp:BoundField>
                <asp:BoundField DataField="ItemCode" HeaderText="ITEM NAME" >
                    <ItemStyle Width="100px" />
                </asp:BoundField>
                <asp:BoundField DataField="Qty" HeaderText="QTY" DataFormatString="{0:N0}" >
                    <ItemStyle Width="100px" />
                </asp:BoundField>
                <asp:BoundField DataField="AMT" DataFormatString="{0:N0}" HeaderText="金額" >
                    <ItemStyle Width="100px" />
                </asp:BoundField>
                <asp:BoundField DataField="INVQTY" HeaderText="KEEP在庫數" DataFormatString="{0:N0}" >
                    <ItemStyle Width="100px" />
                </asp:BoundField>
                <asp:BoundField DataField="INVAMT" HeaderText="KEEP在庫金額" DataFormatString="{0:N0}" >
                    <ItemStyle Width="120px" />
                </asp:BoundField>
                <asp:BoundField DataField="Attachfile" HeaderText="附檔" >
                    <ItemStyle Width="50px" />
                </asp:BoundField>
                <asp:BoundField DataField="Attachfile" HeaderText="URL" />
            </Columns>
            <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
            <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
        </asp:GridView>
        &nbsp;
        <asp:HyperLink ID="LFCDataList" runat="server" Style="z-index: 119; left: 254px;
            position: absolute; top: 216px" Target="_blank">細項內容</asp:HyperLink>
        </div>
    </form>
</body>
</html>
