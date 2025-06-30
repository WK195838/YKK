<%@ Page Language="VB" AutoEventWireup="false" CodeFile="IRWNoticePage.aspx.vb" Inherits="IRWNoticePage" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>NOTICE PAGE</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  hyperlink                                                      ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
    
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Customer GridView                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="GridView1" runat="server" WIDTH=200 AutoGenerateColumns="False" Style="z-index: 103; left:0px; position: absolute; top: 30px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>

                <asp:BoundField DataField="SEQ" HeaderText="TOP" >  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  
                <asp:BoundField DataField="NAME" HeaderText="NAME" />  
                <asp:BoundField DataField="COUNT" HeaderText="COUNT"  >  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  
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
        <asp:GridView ID="GridView2" runat="server" WIDTH=200 AutoGenerateColumns="False" Style="z-index: 103; left: 210px; position: absolute; top: 30px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="SEQ" HeaderText="TOP"   >  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  
                <asp:BoundField DataField="NAME" HeaderText="NAME"  /> 
                <asp:BoundField DataField="COUNT" HeaderText="COUNT"   >  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  
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
        <asp:GridView ID="GridView3" runat="server" WIDTH=200 AutoGenerateColumns="False" Style="z-index: 103; left: 420px; position: absolute; top: 30px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="SEQ" HeaderText="TOP"   >  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  
                <asp:BoundField DataField="NAME" HeaderText="NAME"  />  
                <asp:BoundField DataField="COUNT" HeaderText="COUNT"   >  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  
            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Customer GridView8                                                          ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="GridView8" runat="server" WIDTH=600 AutoGenerateColumns="False" Style="z-index: 103; left:0px; position: absolute; top: 320px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>

                <asp:HyperLinkField DataNavigateUrlFields="NoURL" DataNavigateUrlFormatString="{0}"
                    DataTextField="No" HeaderText="" Target="_blank">
                       <ItemStyle HorizontalAlign="CENTER" />
                </asp:HyperLinkField>
                
                <asp:BoundField DataField="Status" HeaderText="1"   >  <ItemStyle HorizontalAlign="left" />   </asp:BoundField>  
                <asp:BoundField DataField="AppName" HeaderText="2"   >  <ItemStyle HorizontalAlign="left" />   </asp:BoundField>  
                <asp:BoundField DataField="YYMM" HeaderText="3" >  <ItemStyle HorizontalAlign="LEFT" />   </asp:BoundField>  
                <asp:BoundField DataField="ALLCount" HeaderText="4"   >  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  
                <asp:BoundField DataField="NG" HeaderText="5"   >  <ItemStyle HorizontalAlign="right"   />   </asp:BoundField>  
                <asp:BoundField DataField="NGPer" HeaderText="6"   >  <ItemStyle HorizontalAlign="right"   />   </asp:BoundField>  
                <asp:BoundField DataField="NoWorkHour" HeaderText="7"   >  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  

            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>
            
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Customer GridView   4                                                                ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="GridView4" runat="server" WIDTH=600 AutoGenerateColumns="False" Style="z-index: 103; left: 0px; position: absolute; top: 520px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="DepName" HeaderText="1"   >  <ItemStyle HorizontalAlign="left" />   </asp:BoundField>  
                <asp:BoundField DataField="A_YES" HeaderText="2" >  <ItemStyle HorizontalAlign="right" width=50 />   </asp:BoundField>  
                <asp:BoundField DataField="A_NO" HeaderText="3"   >  <ItemStyle HorizontalAlign="right" width=50  />   </asp:BoundField>  
                <asp:BoundField DataField="A_TOTAL" HeaderText="4"   >  <ItemStyle HorizontalAlign="right" width=50  />   </asp:BoundField>  
                <asp:BoundField DataField="A_PERCENT" HeaderText="5"   >  <ItemStyle HorizontalAlign="right" width=50  />   </asp:BoundField>  

                <asp:BoundField DataField="S_PERCENT" HeaderText="6"   >  <ItemStyle HorizontalAlign="right" width=50  />   </asp:BoundField>  
                <asp:BoundField DataField="S_NO" HeaderText="7"   >  <ItemStyle HorizontalAlign="right" width=50  />   </asp:BoundField>  
                <asp:BoundField DataField="H_NO" HeaderText="8"   >  <ItemStyle HorizontalAlign="right" width=60  />   </asp:BoundField>  
                <asp:BoundField DataField="L_NO" HeaderText="9"   >  <ItemStyle HorizontalAlign="right" width=50  />   </asp:BoundField>  
                <asp:BoundField DataField="B_NO" HeaderText="10"   >  <ItemStyle HorizontalAlign="right" width=50  />   </asp:BoundField>  
            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>


<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Customer GridView5                                                            ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="GridView5" runat="server" WIDTH=600 AutoGenerateColumns="False" Style="z-index: 103; left:0px; position: absolute; top: 720px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="DepName" HeaderText="1"   >  <ItemStyle HorizontalAlign="left" />   </asp:BoundField>  
                <asp:BoundField DataField="EmpName" HeaderText="2"   >  <ItemStyle HorizontalAlign="left" />   </asp:BoundField>  
                <asp:BoundField DataField="A_YES" HeaderText="3" >  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  
                <asp:BoundField DataField="A_NO" HeaderText="4"   >  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  
                <asp:BoundField DataField="A_TOTAL" HeaderText="5"   >  <ItemStyle HorizontalAlign="right"   />   </asp:BoundField>  
                <asp:BoundField DataField="A_PERCENT" HeaderText="6"   >  <ItemStyle HorizontalAlign="right"   />   </asp:BoundField>  
                <asp:BoundField DataField="REMARK" HeaderText="7"   >  <ItemStyle HorizontalAlign="left" />   </asp:BoundField>  

                <asp:HyperLinkField DataNavigateUrlFields="V2URL" DataNavigateUrlFormatString="{0}"
                    DataTextField="V2" HeaderText="" Target="_blank">
                       <ItemStyle HorizontalAlign="CENTER" />
                </asp:HyperLinkField>

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
        <asp:GridView ID="GridView9" runat="server" WIDTH=600 AutoGenerateColumns="False" Style="z-index: 103; left: 0px; position: absolute; top: 1150px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="Rank" HeaderText="1"   >  <ItemStyle HorizontalAlign="center" />   </asp:BoundField>  
                <asp:BoundField DataField="DepName" HeaderText="1"   >  <ItemStyle HorizontalAlign="left" />   </asp:BoundField>  
                <asp:BoundField DataField="Hour" HeaderText="10"   >  <ItemStyle HorizontalAlign="right"   />   </asp:BoundField>  
                <asp:BoundField DataField="HourTotal" HeaderText="10"   >  <ItemStyle HorizontalAlign="right"   />   </asp:BoundField>  
                <asp:BoundField DataField="HourPer" HeaderText="10"   >  <ItemStyle HorizontalAlign="right"   />   </asp:BoundField>  
                
                <asp:HyperLinkField DataNavigateUrlFields="V2URL" DataNavigateUrlFormatString="{0}"
                    DataTextField="V2" HeaderText="" Target="_blank">
                       <ItemStyle HorizontalAlign="CENTER" />
                </asp:HyperLinkField>

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
        <asp:GridView ID="GridView6" runat="server" WIDTH=600 AutoGenerateColumns="False" Style="z-index: 103; left: 0px; position: absolute; top: 950px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="DepName" HeaderText="1"   >  <ItemStyle HorizontalAlign="left" />   </asp:BoundField>  
                <asp:BoundField DataField="A_YES" HeaderText="2" >  <ItemStyle HorizontalAlign="right" width=50 />   </asp:BoundField>  
                <asp:BoundField DataField="A_NO" HeaderText="3"   >  <ItemStyle HorizontalAlign="right" width=50  />   </asp:BoundField>  
                <asp:BoundField DataField="A_TOTAL" HeaderText="4"   >  <ItemStyle HorizontalAlign="right" width=50  />   </asp:BoundField>  
                <asp:BoundField DataField="A_PERCENT" HeaderText="5"   >  <ItemStyle HorizontalAlign="right" width=50  />   </asp:BoundField>  

                <asp:BoundField DataField="S_PERCENT" HeaderText="6"   >  <ItemStyle HorizontalAlign="right" width=50  />   </asp:BoundField>  
                <asp:BoundField DataField="S_NO" HeaderText="7"   >  <ItemStyle HorizontalAlign="right" width=50  />   </asp:BoundField>  
                <asp:BoundField DataField="H_NO" HeaderText="8"   >  <ItemStyle HorizontalAlign="right" width=60  />   </asp:BoundField>  
                <asp:BoundField DataField="L_NO" HeaderText="9"   >  <ItemStyle HorizontalAlign="right" width=50  />   </asp:BoundField>  
                <asp:BoundField DataField="B_NO" HeaderText="10"   >  <ItemStyle HorizontalAlign="right" width=50  />   </asp:BoundField>  
            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Customer GridView5                                                            ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="GridView7" runat="server" WIDTH=600 AutoGenerateColumns="False" Style="z-index: 103; left:0px; position: absolute; top: 1080px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="DepName" HeaderText="1"   >  <ItemStyle HorizontalAlign="left" />   </asp:BoundField>  
                <asp:BoundField DataField="EmpName" HeaderText="2"   >  <ItemStyle HorizontalAlign="left" />   </asp:BoundField>  
                <asp:BoundField DataField="A_YES" HeaderText="3" >  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  
                <asp:BoundField DataField="A_NO" HeaderText="4"   >  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  
                <asp:BoundField DataField="A_TOTAL" HeaderText="5"   >  <ItemStyle HorizontalAlign="right"   />   </asp:BoundField>  
                <asp:BoundField DataField="A_PERCENT" HeaderText="6"   >  <ItemStyle HorizontalAlign="right"   />   </asp:BoundField>  
                <asp:BoundField DataField="REMARK" HeaderText="7"   >  <ItemStyle HorizontalAlign="left" />   </asp:BoundField>  

                <asp:HyperLinkField DataNavigateUrlFields="V2URL" DataNavigateUrlFormatString="{0}"
                    DataTextField="V2" HeaderText="" Target="_blank">
                       <ItemStyle HorizontalAlign="CENTER" />
                </asp:HyperLinkField>

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
        <asp:GridView ID="GridView10" runat="server" WIDTH=600 AutoGenerateColumns="False" Style="z-index: 103; left: 0px; position: absolute; top: 30px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="DataType" HeaderText="1"   >  <ItemStyle HorizontalAlign="left" />   </asp:BoundField>  
                <asp:BoundField DataField="MM1_PER" HeaderText="2" >  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  
                <asp:BoundField DataField="MM2_PER" HeaderText="2" >  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  
                <asp:BoundField DataField="MM3_PER" HeaderText="2" >  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  
                <asp:BoundField DataField="MM4_PER" HeaderText="2" >  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  
                <asp:BoundField DataField="MM5_PER" HeaderText="2" >  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  
                <asp:BoundField DataField="MM6_PER" HeaderText="2" >  <ItemStyle HorizontalAlign="right" />   </asp:BoundField>  
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
        <asp:GridView ID="GridView11" runat="server" WIDTH=600 AutoGenerateColumns="False" Style="z-index: 103; left: 0px; position: absolute; top: 110px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="Rank" HeaderText="1"   >  <ItemStyle HorizontalAlign="center" />   </asp:BoundField>  
                <asp:BoundField DataField="DataType" HeaderText="1"   >  <ItemStyle HorizontalAlign="left" />   </asp:BoundField>  
                <asp:BoundField DataField="MM1_PER" HeaderText="2" >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>  
                <asp:BoundField DataField="MM2_PER" HeaderText="2" >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>  
                <asp:BoundField DataField="MM3_PER" HeaderText="2" >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>  
                <asp:BoundField DataField="MM4_PER" HeaderText="2" >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>  
                <asp:BoundField DataField="MM5_PER" HeaderText="2" >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>  
                <asp:BoundField DataField="MM6_PER" HeaderText="2" >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>  
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
