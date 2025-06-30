<%@ Page Language="VB" EnableEventValidation = "false" AutoEventWireup="false" CodeFile="AutoXml2Excel.aspx.vb" Inherits="AutoXml2Excel" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>XML to Excel</title>
</head>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">

            <asp:HyperLink ID="CheckData" runat="server" NavigateUrl="OrderDataList.aspx" Style="z-index: 110; left: 20px; position: absolute; top: 15px" >連結資料確認</asp:HyperLink>

            <asp:Button ID="BImport" runat="server" BackColor="White" ForeColor="Blue" Height="24px"
                Style="z-index: 110; left: 141px; position: absolute; top: 15px" Text="接收資料" Width="100px" />
               
               
            <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" BorderColor="Black"
                BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt" 
                Style="z-index: 114; left: 8px; position: absolute; top: 72px">
                <Columns>
                    <asp:BoundField DataField="purd_no" HeaderText="採購單號" />
                    <asp:BoundField DataField="purd_item" HeaderText="NO" />
                    <asp:BoundField DataField="purno_no" HeaderText="彙總單號" />
                    <asp:BoundField DataField="pre_insp_date" HeaderText="交貨日期" />
                    <asp:BoundField DataField="vend_no" HeaderText="供應商代號" />
                    <asp:BoundField DataField="vend_f_name" HeaderText="供應商名稱" />
                    <asp:BoundField DataField="vend_tel" HeaderText="電話" />
                    <asp:BoundField DataField="purd_ver" HeaderText="版次" />
                    <asp:BoundField DataField="ref_no" HeaderText="REF_NO" />
                    <asp:BoundField DataField="mat_desc" HeaderText="材料描述" />
                    
                    <asp:BoundField DataField="size_spec" HeaderText="規格" />
                    <asp:BoundField DataField="ad_color_id" HeaderText="顏色" />
                    <asp:BoundField DataField="color_desc" HeaderText="顏色描述" />
                    <asp:BoundField DataField="purd_qty" HeaderText="訂購數量" />
                    <asp:BoundField DataField="mat_unit" HeaderText="單位" />
                    <asp:BoundField DataField="purd_price" HeaderText="單價" />
                    <asp:BoundField DataField="coin_kind" HeaderText="幣別" />
                    <asp:BoundField DataField="purd_note" HeaderText="備註" />
                    <asp:BoundField DataField="cm_odr" HeaderText="訂單號碼" />
                    <asp:BoundField DataField="cm_season" HeaderText="款式代號_季" />
                    
                    <asp:BoundField DataField="cm_style" HeaderText="款式代號_Style" />
                    <asp:BoundField DataField="cm_style_desc" HeaderText="款式代號_StyleDesc" />
                    <asp:BoundField DataField="style_range" HeaderText="款式代號_StyleRange" />
                    <asp:BoundField DataField="operate_mk" HeaderText="加工" />
                    <asp:BoundField DataField="vend_memo1" HeaderText="表尾說明-1" />
                    <asp:BoundField DataField="vend_memo2" HeaderText="表尾說明-2" />
                    <asp:BoundField DataField="vend_memo3" HeaderText="表尾說明-3" />
                    <asp:BoundField DataField="vend_memo4" HeaderText="表尾說明-4" />
                    <asp:BoundField DataField="vend_memo5" HeaderText="表尾說明-5" />
                    <asp:BoundField DataField="vend_memo6" HeaderText="表尾說明-6" />
                    
                    <asp:BoundField DataField="vend_memo7" HeaderText="表尾說明-7" />
                    <asp:BoundField DataField="vend_memo8" HeaderText="表尾說明-8" />
                    <asp:BoundField DataField="vend_memo9" HeaderText="表尾說明-9" />
                    <asp:BoundField DataField="vend_memo10" HeaderText="表尾說明-10" />
                    <asp:BoundField DataField="vend_memo11" HeaderText="表尾說明-11" />
                    <asp:BoundField DataField="vend_memo12" HeaderText="表尾說明-12" />
                    <asp:BoundField DataField="usr_name" HeaderText="經辦人員" />
                    <asp:BoundField DataField="shipping_mark" HeaderText="shipping mark" />
                    <asp:BoundField DataField="msid" HeaderText="msid" />
                    <asp:BoundField DataField="gsid" HeaderText="gsid" />
                    
                    <asp:BoundField DataField="gmt" HeaderText="gmt" />
                    <asp:BoundField DataField="mdp" HeaderText="mdp" />
                    <asp:BoundField DataField="cm_type" HeaderText="cm_type" />
                    <asp:BoundField DataField="cfm_date" HeaderText="cfm_date" />
                    <asp:BoundField DataField="buy_month" HeaderText="Buy_Month" />
                    <asp:BoundField DataField="sample_mrk" HeaderText="訂單種類" />
                    <asp:BoundField DataField="fact_no" HeaderText="廠區" />                    
                    
                </Columns>
                <HeaderStyle BackColor="Silver" ForeColor="Black" HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:GridView>
            
            <asp:GridView ID="TNFGridView" runat="server" AutoGenerateColumns="False" BorderColor="Black"
                BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt" 
                Style="z-index: 114; left: 8px; position: absolute; top: 72px">
                <Columns>
                    <asp:BoundField DataField="purd_no" HeaderText="採購單號" />
                    <asp:BoundField DataField="purd_item" HeaderText="NO" />
                    <asp:BoundField DataField="purno_no" HeaderText="彙總單號" />
                    <asp:BoundField DataField="purd_ver" HeaderText="版次" />
                    <asp:BoundField DataField="ref_no" HeaderText="REF_NO" />
                    <asp:BoundField DataField="mat_desc" HeaderText="材料描述" />
                    <asp:BoundField DataField="size_spec" HeaderText="規格" />
                    <asp:BoundField DataField="color_desc" HeaderText="顏色描述" />
                    <asp:BoundField DataField="purd_qty" HeaderText="訂購數量" />
                    <asp:BoundField DataField="coin_kind" HeaderText="幣別" />
                    <asp:BoundField DataField="cm_season" HeaderText="款式代號_季" />
                    <asp:BoundField DataField="cm_style" HeaderText="款式代號_Style" />
                    <asp:BoundField DataField="style_range" HeaderText="款式代號_StyleRange" />
                    <asp:BoundField DataField="sample_mrk" HeaderText="訂單種類" />
                    <asp:BoundField DataField="buy_month" HeaderText="Buy_Month" />
                    <asp:BoundField DataField="fact_no" HeaderText="廠區" />                    
                </Columns>
                <HeaderStyle BackColor="Silver" ForeColor="Black" HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:GridView>
           
            <asp:GridView ID="ADIDASGridView" runat="server" AutoGenerateColumns="False" BorderColor="Black"
                BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt" 
                Style="z-index: 114; left: 8px; position: absolute; top: 72px">
                <Columns>
                    <asp:BoundField DataField="purd_no" HeaderText="採購單號" />
                    <asp:BoundField DataField="purd_item" HeaderText="NO" />
                    <asp:BoundField DataField="purno_no" HeaderText="彙總單號" />
                    <asp:BoundField DataField="purd_ver" HeaderText="版次" />
                    <asp:BoundField DataField="ref_no" HeaderText="REF_NO" />
                    <asp:BoundField DataField="mat_desc" HeaderText="材料描述" />
                    <asp:BoundField DataField="size_spec" HeaderText="規格" />
                    <asp:BoundField DataField="ad_color_id" HeaderText="顏色" />
                    <asp:BoundField DataField="purd_qty" HeaderText="訂購數量" />
                    <asp:BoundField DataField="coin_kind" HeaderText="幣別" />
                    <asp:BoundField DataField="cm_season" HeaderText="款式代號_季" />
                    <asp:BoundField DataField="cm_style" HeaderText="款式代號_Style" />
                    <asp:BoundField DataField="style_range" HeaderText="款式代號_StyleRange" />
                    <asp:BoundField DataField="sample_mrk" HeaderText="訂單種類" />
                    <asp:BoundField DataField="buy_month" HeaderText="Buy_Month" />
                    <asp:BoundField DataField="fact_no" HeaderText="廠區" />                    
                </Columns>
                <HeaderStyle BackColor="Silver" ForeColor="Black" HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:GridView>
           
            <asp:GridView ID="REEBOKGridView" runat="server" AutoGenerateColumns="False" BorderColor="Black"
                BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt" 
                Style="z-index: 114; left: 8px; position: absolute; top: 72px">
                <Columns>
                    <asp:BoundField DataField="purd_no" HeaderText="採購單號" />
                    <asp:BoundField DataField="purd_item" HeaderText="NO" />
                    <asp:BoundField DataField="purno_no" HeaderText="彙總單號" />
                    <asp:BoundField DataField="purd_ver" HeaderText="版次" />
                    <asp:BoundField DataField="ref_no" HeaderText="REF_NO" />
                    <asp:BoundField DataField="mat_desc" HeaderText="材料描述" />
                    <asp:BoundField DataField="size_spec" HeaderText="規格" />
                    <asp:BoundField DataField="ad_color_id" HeaderText="顏色" />
                    <asp:BoundField DataField="purd_qty" HeaderText="訂購數量" />
                    <asp:BoundField DataField="coin_kind" HeaderText="幣別" />
                    <asp:BoundField DataField="cm_season" HeaderText="款式代號_季" />
                    <asp:BoundField DataField="cm_style" HeaderText="款式代號_Style" />
                    <asp:BoundField DataField="style_range" HeaderText="款式代號_StyleRange" />
                    <asp:BoundField DataField="sample_mrk" HeaderText="訂單種類" />
                    <asp:BoundField DataField="buy_month" HeaderText="Buy_Month" />
                    <asp:BoundField DataField="fact_no" HeaderText="廠區" />                    
                </Columns>
                <HeaderStyle BackColor="Silver" ForeColor="Black" HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:GridView>

            <asp:TextBox ID="DMsg" runat="server" BackColor="White" BorderStyle="Groove" ForeColor="Black"
                Height="20px" Style="z-index: 103; left: 21px; position: absolute; top: 48px"
                Width="252px"></asp:TextBox>

        </form>
	</body>
</html>
