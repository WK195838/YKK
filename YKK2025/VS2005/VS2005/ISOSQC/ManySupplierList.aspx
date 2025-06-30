<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ManySupplierList.aspx.vb" Inherits="ManySupplierList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Many Supplier</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
		<!-- ****************************************************************************************** -->
		<!-- ** MANY DATA                                                                             -->
		<!-- ****************************************************************************************** -->
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" Style="z-index: 103; left: 224px; position: absolute; top: 16px" CellPadding="3" Font-Size="10pt" GridLines="Vertical"  ForeColor="Black" BackColor="White" >
            <Columns>
                <asp:BoundField DataField="MMSDESC" HeaderText=""  />          
                <asp:BoundField DataField="Cat" HeaderText=""  />          
                <asp:BoundField DataField="Puller" HeaderText=""  />          
                <asp:BoundField DataField="COLOR" HeaderText=""  />          
                <asp:BoundField DataField="Remark" HeaderText=""  />          
                <asp:BoundField DataField="Supplier"  />                
                
                <asp:BoundField DataField="Sts"  />                
                <asp:HyperLinkField DataNavigateUrlFields="NOUrl" DataNavigateUrlFormatString="{0}" 
                DataTextField="NO" HeaderText="No" Target="_blank"  >
                </asp:HyperLinkField>

                <asp:BoundField DataField="AppDate"  /> 

                <asp:CommandField ShowDeleteButton="true" SelectText="..." DeleteText="IMG" />
                
                    <asp:HyperLinkField DataNavigateUrlFields="EDXIMGUrl" DataNavigateUrlFormatString="{0}" 
                    DataTextField="EDXIMAGES" HeaderText="Link" Target="_blank"  >
                    </asp:HyperLinkField>

            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#7FFF00" Font-Bold="false" ForeColor="black" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>
		<!-- ****************************************************************************************** -->
		<!-- ** R&D IMAGES
		<!-- ****************************************************************************************** -->
        <asp:Image ID="DRDImage" runat="server" ImageUrl="iMages/No_Images.png" Style="z-index: 103; left: 8px; position: absolute; top: 8px" Width="200" Height="230"  BorderColor="Red" BorderStyle="Solid" BorderWidth="5px" TabIndex="99"/>
    
    </div>
    </form>
</body>
</html>
