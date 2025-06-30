<%@ Page Language="VB" AutoEventWireup="false" CodeFile="FindSPDPage.aspx.vb" Inherits="FindSPDPage" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>搜尋SPD</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  slider                                                                             ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<asp:TextBox style="Z-INDEX: 126; LEFT: 0px; POSITION: absolute; TOP: 8px; TEXT-ALIGN: left" id="DSLDKey1" runat="server" MaxLength="35" Width="189px" Height="18px" ForeColor="Blue" BorderStyle="Groove" BackColor="Yellow"></asp:TextBox>
<asp:TextBox style="Z-INDEX: 126; LEFT: 200px; POSITION: absolute; TOP: 8px; TEXT-ALIGN: left" id="DSLDKey2" runat="server" MaxLength="35" Width="184px" Height="18px" ForeColor="Blue" BorderStyle="Groove" BackColor="Yellow"></asp:TextBox>
<asp:Button ID="BSEARCH1" runat="server" Height="25px" Style="z-index: 104; left: 392px;
    position: absolute; top: 8px" Text="搜尋" Width="45px" />

<asp:TextBox style="Z-INDEX: 126; LEFT: 610px; POSITION: absolute; TOP: 8px; TEXT-ALIGN: left" id="DFINKey1" runat="server" MaxLength="35" Width="189px" Height="18px" ForeColor="Blue" BorderStyle="Groove" BackColor="Yellow"></asp:TextBox>
<asp:TextBox style="Z-INDEX: 126; LEFT: 808px; POSITION: absolute; TOP: 8px; TEXT-ALIGN: left" id="DFINKey2" runat="server" MaxLength="35" Width="184px" Height="18px" ForeColor="Blue" BorderStyle="Groove" BackColor="Yellow"></asp:TextBox>
<asp:Button ID="BSEARCH2" runat="server" Height="25px" Style="z-index: 104; left: 1000px;
    position: absolute; top: 8px" Text="搜尋" Width="45px" />
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Customer GridView                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="GridView1" runat="server" WIDTH=600 AutoGenerateColumns="False" Style="z-index: 103; left: 0px; position: absolute; top: 32px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:CommandField ShowSelectButton="True" />

                <asp:BoundField DataField="SYS" HeaderText="類型"  />
                <asp:BoundField DataField="STATUS" HeaderText="狀態"  />

                <asp:HyperLinkField DataNavigateUrlFields="FURL" DataNavigateUrlFormatString="{0}" 
                DataTextField="NO" HeaderText="NO" Target="_blank"  >
                </asp:HyperLinkField>

                <asp:BoundField DataField="BUYER" HeaderText="BUYER"  />
                <asp:BoundField DataField="SPEC" HeaderText="SPEC"  />
                <asp:BoundField DataField="SLIDERGRCODE" HeaderText="SLIDER GR."  />

                <asp:BoundField DataField="PERSON" HeaderText="申請者"  />
                <asp:BoundField DataField="APPLYDATE" HeaderText="申請日"  />

                <asp:BoundField DataField="NO" HeaderText=""  />
            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>
    
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Customer GridView                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="GridView2" runat="server"  AutoGenerateColumns="False" Style="z-index: 103; left: 610px; position: absolute; top: 32px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:CommandField ShowSelectButton="True" />

                <asp:BoundField DataField="SYS" HeaderText="類型"  />
                <asp:BoundField DataField="STATUS" HeaderText="狀態"  />

                <asp:HyperLinkField DataNavigateUrlFields="FURL" DataNavigateUrlFormatString="{0}" 
                DataTextField="NO" HeaderText="NO" Target="_blank"  >
                </asp:HyperLinkField>

                <asp:BoundField DataField="BUYER" HeaderText="BUYER"  />
                <asp:BoundField DataField="SPEC" HeaderText="SPEC"  />
                <asp:BoundField DataField="SLIDERGRCODE" HeaderText="SLIDER GR."  />

                <asp:BoundField DataField="PERSON" HeaderText="申請者"  />
                <asp:BoundField DataField="APPLYDATE" HeaderText="申請日"  />
                
                <asp:BoundField DataField="NO" HeaderText=""  />
            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>

    </div>
    </form>
</body>
</html>
