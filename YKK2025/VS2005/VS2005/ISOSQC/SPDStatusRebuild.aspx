<%@ Page Language="VB" AutoEventWireup="false" CodeFile="SPDStatusRebuild.aspx.vb" Inherits="SPDStatusRebuild" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Magement-Refresh SPD Information</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- ++  底圖                                                                                ++ -->
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <img src="iMages/SPDStatusRebuild.png" style="z-index: 1; left: 16px; position: absolute;top: 8px; width: 135px; height: 45px;" />
		<!-- ****************************************************************************************** -->
		<!-- **                                                                             -->
		<!-- ****************************************************************************************** -->
        <asp:TextBox ID="TextBox1" runat="server" BackColor="Black" BorderStyle="None" ForeColor="White"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 48px; position: absolute;
            top: 56px; text-align: left" Width="80px">Puller Code</asp:TextBox>
        <asp:TextBox ID="DKPullerCode" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" MaxLength="7" Style="z-index: 126; left: 128px;
            position: absolute; top: 56px; text-align: left" Width="100px"></asp:TextBox>
        <asp:TextBox ID="TextBox11" runat="server" BackColor="Black" BorderStyle="None" ForeColor="White"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 232px; position: absolute;
            top: 56px; text-align: left" Width="80px">Color</asp:TextBox>
        <asp:TextBox ID="DKColor" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" MaxLength="7" Style="z-index: 126; left: 312px;
            position: absolute; top: 56px; text-align: left" Width="48px"></asp:TextBox>
        &nbsp; &nbsp;&nbsp;
		<!-- ****************************************************************************************** -->
		<!-- **                                                                             -->
		<!-- ****************************************************************************************** -->
        <asp:Button ID="BFind" runat="server" Height="25px" Style="z-index: 104; left: 376px;
            position: absolute; top: 56px" Text="Go" Width="45px" />
    
		<!-- ****************************************************************************************** -->
		<!-- **                                                                             -->
		<!-- ****************************************************************************************** -->
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" Style="z-index: 103; left: 48px; position: absolute; top: 88px" CellPadding="3" Font-Size="10pt" GridLines="Vertical"  ForeColor="Black" BackColor="White" >
            <Columns>
                <asp:TemplateField ShowHeader="False">
                    <ItemTemplate>
                        <asp:LinkButton ID="BRefresh" runat="server" CausesValidation="False" CommandName="Refresh"
                            Text="Refresh" onclientclick="javascript:return confirm('確定重新整理？');" CommandArgument='<%# Container.DataItemIndex%>' >                           
                        </asp:LinkButton>
                    </ItemTemplate>
                </asp:TemplateField>

                <asp:BoundField DataField="RD"  />          
                <asp:BoundField DataField="DTM"  />    
                <asp:BoundField DataField="EDX"  />          
                <asp:BoundField DataField="IRW"  />            
                <asp:BoundField DataField="ORDERS"  />            


                <asp:HyperLinkField DataNavigateUrlFields="PullerURL" DataNavigateUrlFormatString="{0}" 
                DataTextField="Puller" HeaderText="色" Target="_blank"  >
                </asp:HyperLinkField>


                <asp:BoundField DataField="Color"  />          
                <asp:BoundField DataField="ColorName"  />    
                <asp:BoundField DataField="BYColorCode"  />          
                <asp:BoundField DataField="Buyer"  />            
                    
                <asp:BoundField DataField="BuyerName"  />            
                <asp:BoundField DataField="Remark"  />            
                <asp:BoundField DataField="DTM_YOBI1"  />            
                <asp:BoundField DataField="DTM_YOBI2"  />            
                <asp:BoundField DataField="IRW_YOBI1"  />            
                <asp:BoundField DataField="IRW_YOBI2"  />            
                
                <asp:BoundField DataField="OR_YOBI2"  />            
                <asp:BoundField DataField="Yobi1"  />            
                <asp:BoundField DataField="Yobi2"  />            
           
            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="Chartreuse" Font-Bold="False" ForeColor="Black" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>     
        <asp:TextBox ID="DOther" runat="server" BackColor="white" BorderStyle="None" Height="18px"
            MaxLength="7" Style="z-index: 126; left: 440px; position: absolute; top: 56px;
            text-align: left" Width="312px">重整[SPD Information]需處理時間...</asp:TextBox>
    </form>
</body>
</html>
