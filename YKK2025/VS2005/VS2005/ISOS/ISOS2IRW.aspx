<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ISOS2IRW.aspx.vb" Inherits="ISOS2IRW" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >




<head id="Head1" runat="server">
    <title>ISOS2IRW Inf.</title>
	    <script language="javascript" src="ForProject.js"></script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  底圖                                                                                ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <img src="iMages/Download.jpg" style="z-index: 1; left: 16px; position: absolute;top: 195px; width: 40px; height: 40px;" />
        <img src="iMages/Download.jpg" style="z-index: 1; left: 16px; position: absolute;top: 395px; width: 40px; height: 40px;" />
        <img src="iMages/Download.jpg" style="z-index: 1; left: 16px; position: absolute;top: 695px; width: 40px; height: 40px;" />
    
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  欄位                                                                            ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:TextBox ID="DBUYER" runat="server" BackColor="#80FF80" BorderStyle="Groove"
            ForeColor="Blue" Height="14px" Style="z-index: 103; left: 16px; position: absolute;
            top: 16px" Width="112px"></asp:TextBox>
        <asp:TextBox ID="DSlider1" runat="server" BackColor="#80FF80" BorderStyle="Groove"
            ForeColor="Blue" Height="14px" Style="z-index: 103; left: 136px; position: absolute;
            top: 16px" Width="112px"></asp:TextBox>
        <asp:TextBox ID="DCust" runat="server" BackColor="yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="14px" Style="z-index: 103; left: 254px; position: absolute;
            top: 16px" Width="112px"></asp:TextBox>
            
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  按鈕                                                                                ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
   
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  隱藏欄位                                                                       ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ------------------------------------------------------------------------------------ -->
<!-- System Parameter  -->
        <asp:TextBox ID="HBuyerCode" runat="server" Height="16px" Style="z-index: 318; left: 8px;
            position: absolute; top: 192px;font-size:10px;background:transparent" Width="600px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>

        <asp:TextBox ID="HBuyerColor" runat="server" Height="16px" Style="z-index: 318; left: 24px;
            position: absolute; top: 48px;font-size:16px;background:transparent" Width="200px"   BorderStyle="None" BorderWidth="0px" ForeColor="Red">搜尋不到資料,請確認!</asp:TextBox>
        <asp:TextBox ID="HIRW" runat="server" Height="16px" Style="z-index: 318; left: 24px;
            position: absolute; top: 248px;font-size:16px;background:transparent" Width="200px"   BorderStyle="None" BorderWidth="0px" ForeColor="Red">搜尋不到資料,請確認!</asp:TextBox>
        <asp:TextBox ID="HBuyerItem" runat="server" Height="16px" Style="z-index: 318; left: 24px;
            position: absolute; top: 448px;font-size:16px;background:transparent" Width="200px"   BorderStyle="None" BorderWidth="0px" ForeColor="Red">搜尋不到資料,請確認!</asp:TextBox>
        <asp:TextBox ID="HSales" runat="server" Height="16px" Style="z-index: 318; left: 24px;
            position: absolute; top: 744px;font-size:16px;background:transparent" Width="200px"   BorderStyle="None" BorderWidth="0px" ForeColor="Red">搜尋不到資料,請確認!</asp:TextBox>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  BUYER PULLER GridView                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
      <asp:Panel ID="Panel1" runat="server" Height="150px" Width="1200px"  ScrollBars="Auto"  BorderWidth="1px" style="left: 16px; position: absolute; top: 40px">
        <asp:GridView ID="GridView4" runat="server" AutoGenerateColumns="False"   Style="z-index: 103; position: absolute" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="10" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="Z1" HeaderText=""  />
                <asp:BoundField DataField="A1" HeaderText=""  />
                <asp:BoundField DataField="B1" HeaderText=""  />
                <asp:BoundField DataField="C1" HeaderText=""  />
                <asp:BoundField DataField="D1" HeaderText=""  />
                <asp:BoundField DataField="E1" HeaderText=""  />
            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>
      </asp:Panel>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  IRW GridView                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
      <asp:Panel ID="Panel2" runat="server" Height="150px" Width="1200px"  ScrollBars="Auto"  BorderWidth="1px" style="left: 16px; position: absolute; top: 240px">
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False"   Style="z-index: 103; position: absolute" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="10" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="Z1" HeaderText=""  />
                <asp:CommandField ShowSelectButton="True" SelectText="展" />
            
                <asp:BoundField DataField="A1" HeaderText=""  />
                <asp:BoundField DataField="B1" HeaderText=""  />
                <asp:BoundField DataField="C1" HeaderText=""  />
                <asp:BoundField DataField="D1" HeaderText=""  />
                <asp:BoundField DataField="E1" HeaderText=""  />
                <asp:BoundField DataField="F1" HeaderText=""  />
                <asp:BoundField DataField="G1" HeaderText=""  />
                <asp:BoundField DataField="H1" HeaderText=""  />                
                <asp:BoundField DataField="I1" HeaderText=""  />
                <asp:BoundField DataField="J1" HeaderText=""  />                

                <asp:HyperLinkField DataNavigateUrlFields="FormURL" DataNavigateUrlFormatString="{0}" 
                DataTextField="FormMark" HeaderText="Select" Target="_blank">
                </asp:HyperLinkField>

            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>
      </asp:Panel>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  BUYER ITEM GridView                                                                 ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
      <asp:Panel ID="Panel3" runat="server" Height="250px" Width="1200px"  ScrollBars="Auto"  BorderWidth="1px" style="left: 16px; position: absolute; top: 440px">
        <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" Style="z-index: 103; position: absolute" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="10" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="ZZ" HeaderText=""  />
                <asp:BoundField DataField="ZZ1" HeaderText=""  />
                <asp:BoundField DataField="ZZ2" HeaderText=""  />
                <asp:BoundField DataField="ZZ3" HeaderText=""  />
                <asp:BoundField DataField="ZZ4" HeaderText=""  />
                <asp:BoundField DataField="ZZ5" HeaderText=""  />
                <asp:BoundField DataField="ZZ6" HeaderText=""  />
                
                <asp:HyperLinkField DataNavigateUrlFields="URL" DataNavigateUrlFormatString="{0}" 
                DataTextField="Mark" HeaderText="Select" Target="_blank"  >
                </asp:HyperLinkField>

            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>
      </asp:Panel>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  SALES GridView                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
      <asp:Panel ID="Panel4" runat="server" Height="150px" Width="1200px"  ScrollBars="Auto"  BorderWidth="1px" style="left: 16px; position: absolute; top: 740px">      
        <asp:GridView ID="GridView3" runat="server" AutoGenerateColumns="False"  Style="z-index: 103 position: absolute" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="10" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="ZZ" HeaderText=""  />
                <asp:BoundField DataField="A1" HeaderText=""  />
                <asp:BoundField DataField="B1" HeaderText=""  />
                <asp:BoundField DataField="C1" HeaderText=""  />
                
                <asp:HyperLinkField DataNavigateUrlFields="URL" DataNavigateUrlFormatString="{0}" 
                DataTextField="Mark" HeaderText="Select" Target="_blank"  >
                </asp:HyperLinkField>

            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>
      </asp:Panel>
        &nbsp;


    </div>
    </form>
</body>
</html>