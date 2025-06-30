<%@ Page Language="VB" AutoEventWireup="false" CodeFile="BackupList.aspx.vb" Inherits="BackupList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Magement-Backup List</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- ++  底圖                                                                                ++ -->
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <img src="iMages/BackupList.png" style="z-index: 1; left: 16px; position: absolute;top: 8px; width: 135px; height: 45px;" />
		<!-- ****************************************************************************************** -->
		<!-- ** CheckBox                                                                              -->
		<!-- ****************************************************************************************** -->
        <asp:CheckBox ID="AtCloseIMGW" runat="server" style="z-index: 174; left: 1064px; position: absolute; top: 56px" Font-Size="9pt" Text="Close IMG Windows" Width="130px" AutoPostBack="True" Font-Bold="true" />
        
		<!-- ****************************************************************************************** -->
		<!-- **                                                                             -->
		<!-- ****************************************************************************************** -->
        <asp:TextBox ID="TextBox2" runat="server" BackColor="Black" BorderStyle="None" ForeColor="White"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 24px; position: absolute;
            top: 56px; text-align: left" Width="96px">Backup Version</asp:TextBox>
        <asp:DropDownList ID="DBackupVersion" runat="server" BackColor="Yellow" ForeColor="Blue"
            Height="24px" Style="z-index: 112; left: 128px; position: absolute; top: 56px"
            Width="176px">
            <asp:ListItem Selected="True" Value="PULLER">PULLER</asp:ListItem>
        </asp:DropDownList>

        <asp:TextBox ID="TextBox1" runat="server" BackColor="Black" BorderStyle="None" ForeColor="White"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 312px; position: absolute;
            top: 56px; text-align: left" Width="80px">Puller Code</asp:TextBox>
        <asp:TextBox ID="DKPullerCode" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" MaxLength="7" Style="z-index: 126; left: 392px;
            position: absolute; top: 56px; text-align: left" Width="100px"></asp:TextBox>
        <asp:TextBox ID="TextBox11" runat="server" BackColor="Black" BorderStyle="None" ForeColor="White"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 496px; position: absolute;
            top: 56px; text-align: left" Width="80px">Color</asp:TextBox>
        <asp:TextBox ID="DKColor" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" MaxLength="7" Style="z-index: 126; left: 576px;
            position: absolute; top: 56px; text-align: left" Width="48px"></asp:TextBox>
        <asp:TextBox ID="TextBox8" runat="server" BackColor="Black" BorderStyle="None" ForeColor="White"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 624px; position: absolute;
            top: 56px; text-align: left" Width="80px">Buyer</asp:TextBox>
        <asp:TextBox ID="DKBuyer" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" MaxLength="7" Style="z-index: 126; left: 704px;
            position: absolute; top: 56px; text-align: left" Width="100px"></asp:TextBox>
        <asp:TextBox ID="TextBox12" runat="server" BackColor="Black" BorderStyle="None" ForeColor="White"
            Height="18px" ReadOnly="True" Style="z-index: 126; left: 808px; position: absolute;
            top: 56px; text-align: left" Width="80px">Search</asp:TextBox>
        <asp:TextBox ID="DKOther" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" MaxLength="50" Style="z-index: 126; left: 888px;
            position: absolute; top: 56px; text-align: left" Width="100px"></asp:TextBox>
		<!-- ****************************************************************************************** -->
		<!-- **                                                                             -->
		<!-- ****************************************************************************************** -->
        <asp:Button ID="BFind" runat="server" Height="25px" Style="z-index: 104; left: 1008px;
            position: absolute; top: 56px" Text="Go" Width="45px" />
    
		<!-- ****************************************************************************************** -->
		<!-- **                                                                             -->
		<!-- ****************************************************************************************** -->
		<!-- ****************************************************************************************** -->
		<!-- ** Puller List (GridView1)         table-layout:fixed;        width="1000px"                                                 -->
		<!-- ****************************************************************************************** -->
        <!-- ** RD & EDX  -->
            <asp:GridView ID="GridView1" runat="server" WIDTH="2000" AutoGenerateColumns="False"  Style="z-index: 103; left: 24px; position: absolute; top: 88px" CellPadding="3" Font-Size="10pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" DataKeyNames="Puller">
                <Columns>

                    <asp:BoundField DataField="HSCODE"  />
                    <asp:BoundField DataField=""  />

                    <asp:BoundField DataField="Puller"  />
                    <asp:BoundField DataField="Color"  />
                    <asp:BoundField DataField="ColorName"  />
                    <asp:BoundField DataField="BYColorCode" />
                    <asp:BoundField DataField="Buyer"  />
                    <asp:BoundField DataField="BuyerName"  />
                    <asp:BoundField DataField="Remark"  />
                    <asp:BoundField DataField="DTM_YOBI1"  />

                    
                    <asp:HyperLinkField DataNavigateUrlFields="RDMUrl" DataNavigateUrlFormatString="{0}" 
                    DataTextField="RD" HeaderText="Link" Target="_blank"  >
                    </asp:HyperLinkField>
                    
                    <asp:BoundField DataField="DTM"  />

                    <asp:HyperLinkField DataNavigateUrlFields="EDXMUrl" DataNavigateUrlFormatString="{0}" 
                    DataTextField="EDX" HeaderText="Link" Target="_blank"  >
                    </asp:HyperLinkField>

                    <asp:BoundField DataField="IRW"  />
                    <asp:BoundField DataField="ORDERS"  />
 
                    <asp:HyperLinkField DataNavigateUrlFields="RDUrl" DataNavigateUrlFormatString="{0}" 
                    DataTextField="RD_No" HeaderText="Link" Target="_blank"  >
                    </asp:HyperLinkField>

                    <asp:BoundField DataField="RD_AppDate"  />                
                    <asp:BoundField DataField="RD_Supplier"  />                

                    <asp:CommandField ShowDeleteButton="true" SelectText="..." DeleteText="IMG" />

                    <asp:HyperLinkField DataNavigateUrlFields="EDXUrl" DataNavigateUrlFormatString="{0}" 
                    DataTextField="EDX_No" HeaderText="Link" Target="_blank"  >
                    </asp:HyperLinkField>
                    <asp:BoundField DataField="EDX_AppDate"  />                
                    <asp:BoundField DataField="EDX_Supplier"  />    

                    <asp:HyperLinkField DataNavigateUrlFields="EDXIMGUrl" DataNavigateUrlFormatString="{0}" 
                    DataTextField="EDX_IMG" HeaderText="Link" Target="_blank"  >
                    </asp:HyperLinkField>
                        
                    <asp:HyperLinkField DataNavigateUrlFields="IRWUrl" DataNavigateUrlFormatString="{0}" 
                    DataTextField="IRW_No" HeaderText="Link" Target="_blank"  >
                    </asp:HyperLinkField>
                    <asp:BoundField DataField="IRW_AppDate"  />                

                    <asp:HyperLinkField DataNavigateUrlFields="ORUrl" DataNavigateUrlFormatString="{0}" 
                    DataTextField="OR_No" HeaderText="Link" Target="_blank"  >
                    </asp:HyperLinkField>
                    <asp:BoundField DataField="OR_CfmDate"  />                

                    <asp:HyperLinkField DataNavigateUrlFields="ORIMGUrl" DataNavigateUrlFormatString="{0}" 
                    DataTextField="OR_IMG" HeaderText="Link" Target="_blank"  >
                    </asp:HyperLinkField>

                    <asp:BoundField DataField="TapeColor"  />                

                </Columns>
                <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
                <FooterStyle BackColor="#CCCCCC" />
                <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
                <SelectedRowStyle BackColor="Chartreuse" Font-Bold="False" ForeColor="Black" />
                <AlternatingRowStyle BackColor="#CCCCCC" />
            </asp:GridView>

		<!-- ****************************************************************************************** -->
		<!-- ** R&D IMAGES
		<!-- ****************************************************************************************** -->
        <asp:Image ID="DRDImage" runat="server" ImageUrl="iMages/ISIPLOGO.png" Style="z-index: 103; left: 600px; position: absolute; top: 336px" Width="200" Height="230"  BorderColor="Red" BorderStyle="Solid" BorderWidth="5px" TabIndex="99"/>
        <asp:TextBox ID="DOther" runat="server" BackColor="white" BorderStyle="None" Height="18px"
            MaxLength="7" Style="z-index: 126; left: 1208px; position: absolute; top: 56px;
            text-align: left" Width="312px">顯示150件(件數過多)，請使用篩選</asp:TextBox>

    </form>
</body>
</html>
