<%@ Page Language="VB" AutoEventWireup="false" CodeFile="IRWNoticePageZip01.aspx.vb" Inherits="IRWNoticePageZip01" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>NOTICE PAGE(DETAIL)</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Customer GridView                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" Style="z-index: 103; left:0px; position: absolute; top: 10px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="YYMM"  HeaderText="YYMM"> </asp:BoundField>   
                <asp:BoundField DataField="DataTypeDesc" HeaderText="日"> </asp:BoundField>    
                <asp:BoundField DataField="Cat"  HeaderText="類別"> </asp:BoundField>            

                <asp:HyperLinkField DataNavigateUrlFields="NOURL" DataNavigateUrlFormatString="{0}"
                    DataTextField="NO" HeaderText="委託書NO" Target="_blank">
                </asp:HyperLinkField>

                <asp:HyperLinkField DataNavigateUrlFields="OPURL" DataNavigateUrlFormatString="{0}"
                    DataTextField="OP" HeaderText="承認明細" Target="_blank">
                </asp:HyperLinkField>

                <asp:BoundField DataField="REJECT"  HeaderText=""> </asp:BoundField>            
                <asp:BoundField DataField="RHH"  HeaderText="時間(H)"> </asp:BoundField>            
                <asp:BoundField DataField="RHHSales"  HeaderText="[營]處理(H)"> </asp:BoundField>            

                <asp:BoundField DataField="SendTime"  HeaderText="營業業務室送出時間"> </asp:BoundField>            
                <asp:BoundField DataField="CodeTime"  HeaderText="Code出時間"> </asp:BoundField>            
                <asp:BoundField DataField="StartTime"  HeaderText="開始時間"> </asp:BoundField>            
                <asp:BoundField DataField="EndTime"  HeaderText="完成時間"> </asp:BoundField>  
                <asp:BoundField DataField="HH"  HeaderText="實績(H)"> </asp:BoundField> 
                <asp:BoundField DataField="StdHH"  HeaderText="目標(H)"> </asp:BoundField> 
                <asp:BoundField DataField="Balance"  HeaderText="差(H)"> </asp:BoundField> 

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
