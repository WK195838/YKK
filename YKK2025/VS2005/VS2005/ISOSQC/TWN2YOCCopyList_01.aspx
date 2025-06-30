<%@ Page Language="VB" AutoEventWireup="false" CodeFile="TWN2YOCCopyList_01.aspx.vb" Inherits="TWN2YOCCopyList_01" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Place of production(TWN-YOC)</title>
</head>

<body>
    <form id="form1" runat="server">
    <div>
		<!-- ****************************************************************************************** -->
		<!-- ** System
		<!-- ****************************************************************************************** -->

		<!-- ****************************************************************************************** -->
		<!-- ** Button                                                                                  -->
		<!-- ****************************************************************************************** -->
		    <!-- Find -->
                <asp:Button ID="BFind" runat="server" Height="25px" Style="z-index: 104; left: 688px;
                    position: absolute; top: 16px" Text="Go" Width="45px" />
		<!-- ****************************************************************************************** -->
		<!-- ** Puller Key                                                                              -->
		<!-- ****************************************************************************************** -->
            <asp:TextBox ID="TextBox1" runat="server" BackColor="Black" BorderStyle="None" ForeColor="White"
                Height="18px" ReadOnly="True" Style="z-index: 126; left: 16px; position: absolute;
                top: 16px; text-align: left" Width="80px">Slider Code</asp:TextBox>
            <asp:TextBox ID="DKSliderCode" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
                Height="18px" Style="z-index: 126; left: 96px; position: absolute;
                top: 16px; text-align: left" Width="152px" MaxLength="7"></asp:TextBox>

            <asp:TextBox ID="TextBox8" runat="server" BackColor="Black" BorderStyle="None" ForeColor="White"
                Height="18px" ReadOnly="True" Style="z-index: 126; left:256px; position: absolute;
                top: 16px; text-align: left" Width="80px">Buyer</asp:TextBox>
            <asp:TextBox ID="DKBuyer" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
                Height="18px" Style="z-index: 126; left: 336px; position: absolute;
                top: 16px; text-align: left" Width="100px" MaxLength="7"></asp:TextBox>

            <asp:TextBox ID="TextBox12" runat="server" BackColor="Black" BorderStyle="None" ForeColor="White"
                Height="18px" ReadOnly="True" Style="z-index: 126; left:440px; position: absolute;
                top: 16px; text-align: left" Width="80px">Search</asp:TextBox>
            <asp:TextBox ID="DKOther" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
                Height="18px" Style="z-index: 126; left: 520px; position: absolute;
                top: 16px; text-align: left" Width="152px" MaxLength="50"></asp:TextBox>
        <!-- ****************************************************************************************** -->
        <!-- ** CheckBox                                                                                -->
        <!-- ****************************************************************************************** -->
        <asp:CheckBox ID="AtCloseIMGW" runat="server" style="z-index: 174; left: 752px; position: absolute; top: 48px" Font-Size="9pt" Text="Close IMG Windows" Width="130px" AutoPostBack="True" Font-Bold="true" />
        <asp:CheckBox ID="AtCloseWIMGW" runat="server" style="z-index: 174; left: 952px; position: absolute; top: 48px" Font-Size="9pt" Text="Close W.IMG Windows" Width="160px" AutoPostBack="True" Font-Bold="true" />
        <!-- ******************************************************************************************
        <!-- ** Puller List (GridView1)   
        <!-- ****************************************************************************************** -->
            <asp:GridView ID="GridView1" runat="server" WIDTH="2000px" AutoGenerateColumns="False"  Style="z-index: 103; left: 16px; position: absolute; top: 72px" CellPadding="3" Font-Size="10pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" DataKeyNames="SliderCode">
                <Columns>
                    <asp:BoundField DataField="SliderCode"  />
                    <asp:BoundField DataField="YOC"  />
                    <asp:BoundField DataField="Buyer"  />
                    <asp:BoundField DataField="ActualDate" />
                    <asp:BoundField DataField="ActualTarget"  />
                    <asp:BoundField DataField="MapNo"  />
                    <asp:BoundField DataField="Forcast"  />
                    <asp:BoundField DataField="Remark"  />                

                                        
                    <asp:HyperLinkField DataNavigateUrlFields="RDNooUrl" DataNavigateUrlFormatString="{0}" 
                    DataTextField="RDNo" HeaderText="Link" Target="_blank"  >
                    </asp:HyperLinkField>
                    <asp:BoundField DataField="RDDate"  />
                    <asp:BoundField DataField="RDSuppiler"  />
                    <asp:BoundField DataField="RDSliderGRCode"  />
            		<asp:CommandField ShowDeleteButton="True" SelectText="..." DeleteText="IMG" />
            		
                    <asp:HyperLinkField DataNavigateUrlFields="CopyNoUrl" DataNavigateUrlFormatString="{0}" 
                    DataTextField="CopyNo" HeaderText="Link" Target="_blank"  >
                    </asp:HyperLinkField>
                    <asp:BoundField DataField="CopyDate"  />

                    <asp:HyperLinkField DataNavigateUrlFields="RDMapNooUrl" DataNavigateUrlFormatString="{0}" 
                    DataTextField="RDMapNo" HeaderText="Link" Target="_blank"  >
                    </asp:HyperLinkField>
                    <asp:BoundField DataField="RDMapDate"  />
                    <asp:CommandField ShowSelectButton="True" SelectText="IMG" />


                    <asp:HyperLinkField DataNavigateUrlFields="ORNOUrl" DataNavigateUrlFormatString="{0}" 
                    DataTextField="OR_NO" HeaderText="Link" Target="_blank"  >
                    </asp:HyperLinkField>
                    <asp:BoundField DataField="OR_SalesDate"  />
                    <asp:BoundField DataField="OR_Item"  />

                    <asp:HyperLinkField DataNavigateUrlFields="ORIMGURL" DataNavigateUrlFormatString="{0}" 
                    DataTextField="OR_IMG" HeaderText="Link" Target="_blank"  >
                    </asp:HyperLinkField>

                </Columns>
                <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
                <FooterStyle BackColor="#CCCCCC" />
                <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
                <SelectedRowStyle BackColor="Chartreuse" Font-Bold="False" ForeColor="Black" />
                <AlternatingRowStyle BackColor="#CCCCCC" />
            </asp:GridView>
		<!-- ***  ORIMGURL *** -->
		<!-- ***  <asp:CommandField ShowEditButton="True" SelectText="..." EditText="IMG" /> *** -->

		<!-- ****************************************************************************************** -->
		<!-- ** R&D & PR IMAGES
		<!-- ****************************************************************************************** -->
        <asp:Image ID="DRDImage" runat="server" ImageUrl="iMages/ISIPLOGO.png" Style="z-index: 103; left: 600px; position: absolute; top: 232px" Width="200" Height="230"  BorderColor="Red" BorderStyle="Solid" BorderWidth="5px" TabIndex="99"/>
        <asp:Image ID="DWImage" runat="server" ImageUrl="iMages/ISIPLOGO.png" Style="z-index: 103; left: 600px; position: absolute; top: 232px" Width="200" Height="230"  BorderColor="Red" BorderStyle="Solid" BorderWidth="5px" TabIndex="99"/>
    </div>
    </form>
</body>
</html>
