<%@ Page Language="VB" AutoEventWireup="false" CodeFile="INQ_ReqDateReport.aspx.vb" Inherits="INQ_ReqDateReport" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>REQ Date Report</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  篩選條件                                                                            ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        客戶代碼　<asp:TextBox ID="DCust" runat="server" BackColor="#FFFF80"></asp:TextBox>　ex. A4760 or ALL=空白<br />
<!--  -->
        業務員代號　<asp:TextBox ID="DSalesCode" runat="server" BackColor="#FFFF80"></asp:TextBox>　ex. M or ALL=空白<br />
<!--  -->
        Entry Date　<asp:TextBox ID="DSStartDate" runat="server" BackColor="#FFFF80"></asp:TextBox>～<asp:TextBox ID="DSEndDate" runat="server" BackColor="#FFFF80"></asp:TextBox>　ex. 20210801 ~ 200815
<!--  -->

        <asp:CheckBox ID="AtLongDate" runat="server" style="z-index: 174; left: 432px; position: absolute; top: 16px" Font-Size="12pt" Text="＞標準REQ.DATE" Width="160px" Height="48px" />
        <asp:imagebutton id="BExcel" style="Z-INDEX: 100; LEFT: 688px; POSITION: absolute; TOP: 64px" runat="server"
				ImageUrl="~/Images/Excel.jpg" Height="21px" Width="21px"></asp:imagebutton>
        <asp:Button ID="BSearch" runat="server" Text="搜尋" Width="80px" style="Z-INDEX: 100; LEFT: 600px; POSITION: absolute; TOP: 16px" Height="72px" /></div>          
				
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  顧客資料                                                                                ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:datagrid id="DataGrid1" style="Z-INDEX: 101; LEFT: 16px; POSITION: absolute; TOP: 100px" runat="server"
				Height="100px" Width="1500px" BackColor="White" BorderStyle="None" AutoGenerateColumns="False" CellPadding="4" BorderWidth="1px" BorderColor="#CC9966" Font-Size="9pt">
				<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
				<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
				<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
				<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
				<Columns>
                    <asp:BoundColumn DataField="OrderNo" HeaderText="OrderNo" />
                    <asp:BoundColumn DataField="Salesman" HeaderText="Sales" />
                    <asp:BoundColumn DataField="SalesName" HeaderText="" />
                    <asp:BoundColumn DataField="Customer" HeaderText="Customer" />
                    <asp:BoundColumn DataField="CustomerName" HeaderText="" />
                    <asp:BoundColumn DataField="Buyer1" HeaderText="Buyer" />
                    <asp:BoundColumn DataField="BuyerName" HeaderText="" />
                    <asp:BoundColumn DataField="Sample" HeaderText="Sample" />

                    <asp:BoundColumn DataField="EntryDate" HeaderText="ConfirmDate(a)" />
                    <asp:BoundColumn DataField="RequestDate" HeaderText="REQ.Date(b)" />
                    <asp:BoundColumn DataField="Entry2RequestDays" HeaderText="Days(b-a)" />
                    <asp:BoundColumn DataField="StdDays" HeaderText="Rule Days" />
                    <asp:BoundColumn DataField="StdDate" HeaderText="Rule Date" />

                    <asp:BoundColumn DataField="EDIInf" HeaderText="EDI Inf." />


						<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="URL" DataNavigateUrlFormatString="{0}" DataTextField="MODInf"
							HeaderText="Modify Inf.">
						</asp:HyperLinkColumn>




				</Columns>
			</asp:datagrid>

     </form>
</body>
</html>
