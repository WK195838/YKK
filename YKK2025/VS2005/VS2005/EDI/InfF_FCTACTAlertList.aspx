<%@ Page Language="VB" AutoEventWireup="false" CodeFile="InfF_FCTACTAlertList.aspx.vb" Inherits="InfF_FCTACTAlertList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>ACT-FCT Alert List</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  篩選欄位                                                                            ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:Label ID="Label1" runat="server" Style="z-index: 103; left: 10px; position: absolute; top: 8px">Buyer：</asp:Label>
        <asp:Label ID="Label6" runat="server" Style="z-index: 103; left: 149px; position: absolute; top: 9px">Season：</asp:Label>
        <asp:Label ID="Label3" runat="server" Style="z-index: 103; left: 296px; position: absolute; top: 9px">Type：</asp:Label>
        &nbsp;
        <asp:Label ID="Label2" runat="server" Style="z-index: 103; left: 630px; position: absolute; top: 8px">Ratio：</asp:Label>
        
        <asp:Label ID="Label5" runat="server" Style="z-index: 103; left: 476px; position: absolute; top: 8px">Alert：</asp:Label>
        &nbsp;

        <asp:TextBox ID="DBuyer" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="16px"  Width="74px" Style="z-index: 103; left: 64px; position: absolute; top: 8px"></asp:TextBox>

        <asp:DropDownList id="DSeason" runat="server" Width="80px" ForeColor="Blue" BackColor="Yellow" Style="z-index: 103; left: 211px; position: absolute; top: 9px">
            <asp:ListItem>SS14</asp:ListItem>
        </asp:DropDownList>
        &nbsp;&nbsp;

        <asp:DropDownList id="DRatio" runat="server" Width="80px" ForeColor="Blue" BackColor="Yellow" Style="z-index: 103; left: 680px; position: absolute; top: 7px">
            <asp:ListItem>SS14</asp:ListItem>
        </asp:DropDownList>

        <asp:DropDownList id="DLevel" runat="server" Width="124px" ForeColor="Blue" BackColor="Yellow" Style="z-index: 103; left: 345px; position: absolute; top: 8px">
            <asp:ListItem Selected="True">ITEM</asp:ListItem>
            <asp:ListItem>ITEM+COLOR</asp:ListItem>
        </asp:DropDownList>

        <asp:DropDownList id="DAlertType" runat="server" Width="98px" ForeColor="Blue" BackColor="Yellow" Style="z-index: 103; left: 525px; position: absolute; top: 7px">
            <asp:ListItem Selected="True">ACT / FCT</asp:ListItem>
            <asp:ListItem>FCT / ACT</asp:ListItem>
        </asp:DropDownList>
        
        <asp:TextBox ID="DBlank" runat="server" BackColor="White" BorderStyle="None"
            ForeColor="#C00000" Height="16px" Style="z-index: 103; left: 16px; position: absolute;
            top: 39px" Width="276px">實際Order資料時點=當日7:00 (受注日基準)</asp:TextBox>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  按鈕                                                                                ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
    	<asp:button id="BInq" runat="server" Height="24px" Width="40px" Style="z-index: 103; left: 767px; position: absolute; top: 5px" ForeColor="Black" BackColor="Silver" Text="Go"></asp:button>
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- ++  ITEM GridView                                                                   ++ -->
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="ITEMGridView" runat="server" Style="z-index: 103; left: 13px; position: absolute; top: 62px" AutoGenerateColumns="False" CellPadding="3" Font-Size="9pt"  GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="CustItem" HeaderText="ITEM"  />
                <asp:BoundField DataField="Color" HeaderText="COLOR"  />
                <asp:BoundField DataField="Customer" HeaderText="成衣廠"  />
                <asp:BoundField DataField="Season" HeaderText="季"  />
                <asp:BoundField DataField="Month" HeaderText="年月"  />
                <asp:BoundField DataField="Version" HeaderText="Version"  />

                <asp:BoundField DataField="ACTQty" HeaderText="A" > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="FCTQty" HeaderText="F" > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>
                <asp:BoundField DataField="ACTFCTRatio" HeaderText="R" > <ItemStyle HorizontalAlign="Right" /> </asp:BoundField>

            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>

    </div>
    </form>
</body></html>
