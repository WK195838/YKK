<%@ Page Language="VB" AutoEventWireup="false" CodeFile="AddItemList.aspx.vb" Inherits="AddItemList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>登錄</title>

    <script language="javascript" type="text/javascript">
		function calendarPicker(strField)
		{
			window.open('DatePicker1.aspx?field=' + strField,'calendarPopup','width=250,height=190,resizable=yes');
		}
        			
       function GetExpItem()
       {
            window.open('ExpItemList.aspx','Popup','status=0,toolbar=0,width=580,height=350,top=10,resizable=yes');
       }
        			
		function GetRemark(strField,strField1)
		{
			window.open('RemarkList.aspx?field=' + strField +'&field1='+ strField1,'Popup','width=600,height=400,top=10,resizable=yes');
		}
        			
       function GetReference()
       {
            window.open('ReferenceList.aspx','Popup','status=0,toolbar=0,width=700,height=450,top=10,resizable=yes');
       }
       
 

	</script>
	
	<style type="text/css">
        .TextUpper
        {
            text-transform:uppercase;
        }
</style>

</head>
<body>
    <form id="Form1" runat="server">
    <div>
        <!-- -->
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- ++  底圖                                                                                ++ -->
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        &nbsp;<!-- --><!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ --><!-- ++  Date                                                                            ++ -->
        <input id="BADate" runat="server" style="z-index: 123; left: 136px; width: 24px;
            position: absolute; top: 48px" type="button" value="..." />
        <asp:TextBox ID="DADate" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Height="24px" Style="z-index: 100; left: 160px; position: absolute; top: 48px" Width="136px"></asp:TextBox>
        <asp:Image ID="Image1" runat="server" ImageUrl="~/images/AddItemList_04.jpg" Style="z-index: 99;
            left: 8px; position: absolute; top: 24px" />
        <asp:TextBox ID="DGUINo" runat="server" onkeyup="if(isNaN(value))execCommand('undo')"  AutoPostBack="True" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 104; left: 136px; position: absolute;
            top: 184px; text-align: right" Width="160px" MaxLength="8"></asp:TextBox>
        <asp:TextBox ID="DTaxNo" runat="server" CssClass="TextUpper" AutoPostBack="True" BackColor="Yellow"
            BorderStyle="Groove" ForeColor="Blue" Height="24px" Style="z-index: 104; left: 136px;
            position: absolute; top: 152px; text-align: right" Width="160px" MaxLength="14"></asp:TextBox>
        <asp:TextBox ID="DTaxBase" runat="server" onkeyup="if(isNaN(value))execCommand('undo')" BackColor="LightGray"
            BorderStyle="Groove" ForeColor="Blue" Height="24px" Style="z-index: 104; left: 136px;
            position: absolute; top: 312px; text-align: right" Width="160px">0</asp:TextBox>
        <asp:TextBox ID="DInvoiceNo" runat="server"  CssClass="TextUpper" AutoPostBack="True" BackColor="Yellow"
            BorderStyle="Groove" ForeColor="Blue" Height="24px" Style="z-index: 104; left: 136px;
            position: absolute; top: 120px; text-align: right" Width="160px" MaxLength="15"></asp:TextBox>
        <!-- ++  ExpItem                                                                             ++ -->
        <input id="BExpItemList" runat="server" style="z-index: 121; left: 424px; width: 24px;
            position: absolute; top: 48px" type="button" value="..." disabled="disabled" visible="true" />
        <asp:TextBox ID="DExpItem" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Height="24px" Style="z-index: 102; left: 448px; position: absolute; top: 48px" Width="248px" ReadOnly="True"></asp:TextBox>

        <!-- ++  稅                                                                       ++ -->
        <asp:DropDownList style="Z-INDEX: 103; LEFT: 144px; POSITION: absolute; TOP: 88px" id="DTaxType" runat="server" Width="160px"  BackColor="Yellow" AutoPostBack="true" Height="24px">
        </asp:DropDownList>
        &nbsp;
        <!-- ++  總額                                                                            ++ -->
        <asp:TextBox ID="DAmt" runat="server" onkeyup="if(isNaN(value))execCommand('undo')"  BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Height="24px" Style="z-index: 104; left: 136px; position: absolute; top: 216px; text-align: right" Width="160px" AutoPostBack="True">0</asp:TextBox>
        <!-- ++  內容                                                                            ++ -->
        <input id="BContent" runat="server" style="z-index: 120; left: 424px; width: 24px;
            position: absolute; top: 80px" type="button" value="..." disabled="disabled" />
        &nbsp;
        <asp:TextBox ID="DContent" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue" TextMode="MultiLine"
            Height="160px" Style="z-index: 105; left: 448px; position: absolute; top: 80px" Width="248px"></asp:TextBox>

        <!-- ++  稅額                                                                            ++ -->
        <asp:TextBox ID="DTaxAmt" runat="server" onkeyup="if(isNaN(value))execCommand('undo')"  BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Height="24px" Style="z-index: 106; left: 136px; position: absolute; top: 248px; text-align: right" Width="160px">0</asp:TextBox>
        <!-- ++  淨額                                                                            ++ -->
        <asp:TextBox ID="DNetAmt" runat="server" onkeyup="if(isNaN(value))execCommand('undo')"  BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Height="24px" Style="z-index: 107; left: 136px; position: absolute; top: 280px ; text-align: right" Width="160px" AutoPostBack="True">0</asp:TextBox>
        <!-- ++  備註                                                                            ++ -->
        <input id="BRemark" runat="server" style="z-index: 119; left: 424px; width: 24px;
            position: absolute; top: 248px" type="button" value="..." disabled="disabled" />
        <asp:TextBox ID="DRemark" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue" TextMode="MultiLine"
            Height="88px" Style="z-index: 108; left: 448px; position: absolute; top: 248px" Width="248px"></asp:TextBox>
        &nbsp;<!-- ------------------------------------------------------------------------------------------ --><!-- ++  表格 ++ -->
        <input id="BReference" runat="server" style="z-index: 118; left: 744px; width: 136px;
            position: absolute; top: 232px; height: 40px " type="button" value="經費參照" />
        <input id="BADD" runat="server" style="z-index: 117; left: 744px; width: 136px;
            position: absolute; top: 280px; height: 56px " type="button" value="登錄" />
        <input id="BClose" runat="server" style="z-index: 116; left: 1048px; width: 136px;
            position: absolute; top: 280px; height: 56px " type="button" value="完成" />
        &nbsp;<!-- ------------------------------------------------------------------------------------------ --><!-- ++  連結 ++ -->
        <asp:TextBox ID="DACID" runat="server" BackColor="White" BorderStyle="None" ForeColor="White"
            Height="24px" Style="z-index: 110; left: 1608px; position: absolute; top: 176px"
            Width="142px"></asp:TextBox>
        &nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;
        <asp:TextBox ID="DExpItem1" runat="server" BackColor="White" BorderStyle="None" ForeColor="White"
            Height="24px" Style="z-index: 111; left: 1608px; position: absolute; top: 128px"
            Width="142px"></asp:TextBox>
        <asp:TextBox ID="DContent1" runat="server" BackColor="White" BorderStyle="None" ForeColor="White"
            Height="24px" Style="z-index: 112; left: 1608px; position: absolute; top: 72px"
            Width="448px"></asp:TextBox>
        <asp:TextBox ID="DRemark1" runat="server" BackColor="White" BorderStyle="None" ForeColor="White"
            Height="24px" Style="z-index: 113; left: 1608px; position: absolute; top: 32px"
            Width="142px"></asp:TextBox>
        &nbsp;
        <asp:Panel ID="Panel2" runat="server" BorderWidth="1px" Height="496px" ScrollBars="Auto"
            Style="left: 16px; position: absolute; top: 352px; z-index: 114;" Width="1768px">
            &nbsp;&nbsp;&nbsp;
            <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" AutoGenerateDeleteButton="True"
                AutoGenerateSelectButton="True" BackColor="White" BorderColor="#999999" BorderStyle="Solid"
                BorderWidth="1px" CellPadding="3" DataKeyNames="Unique_ID" Font-Size="9pt" ForeColor="Black"
                GridLines="Vertical" PageSize="100" Style="z-index: 103; left: 8px; position: absolute;
                top: 16px" Width="1344px">
                <Columns>
                    <asp:BoundField DataField="Unique_ID" HeaderText="ID">
                        <HeaderStyle Height="20px" />
                        <ItemStyle Width="50px" />
                    </asp:BoundField>
                    <asp:BoundField DataField="ExpItem" HeaderText="類別用途">
                        <HeaderStyle Height="20px" />
                        <ItemStyle Width="100px" />
                    </asp:BoundField>
                    <asp:BoundField DataField="ADate" DataFormatString="{0:d}" HeaderText="日期">
                        <HeaderStyle Height="20px" />
                        <ItemStyle Width="80px" />
                    </asp:BoundField>
                    <asp:BoundField DataField="TaxType" HeaderText="發票類型(稅別)">
                        <HeaderStyle Height="20px" />
                        <ItemStyle Width="100px" />
                    </asp:BoundField>
                    <asp:BoundField DataField="InvoiceNo" HeaderText="發票號碼＼稅單號碼">
                        <ItemStyle Width="120px" />
                    </asp:BoundField>
                    <asp:BoundField DataField="GUINo" HeaderText="賣方統一編號">
                        <ItemStyle Width="80px" />
                    </asp:BoundField>
                    <asp:BoundField DataField="TaxBase" HeaderText="稅基">
                        <ItemStyle Width="80px" />
                    </asp:BoundField>
                    <asp:BoundField DataField="NetAmt" DataFormatString="{0:N0}" HeaderText="淨額">
                        <HeaderStyle Height="20px" />
                        <ItemStyle Width="80px" />
                    </asp:BoundField>
                    <asp:BoundField DataField="TaxAmt" DataFormatString="{0:N0}" HeaderText="稅額">
                        <HeaderStyle Height="20px" />
                        <ItemStyle Width="80px" />
                    </asp:BoundField>
                    <asp:BoundField DataField="Amt" DataFormatString="{0:N0}" HeaderText="總額">
                        <HeaderStyle Height="20px" />
                        <ItemStyle Width="80px" />
                    </asp:BoundField>
                    <asp:BoundField DataField="Content" HeaderText="內容">
                        <HeaderStyle Height="20px" />
                        <ItemStyle Width="200px" />
                    </asp:BoundField>
                    <asp:BoundField DataField="Remark" HeaderText="備註">
                        <HeaderStyle Height="20px" />
                        <ItemStyle Width="200px" />
                    </asp:BoundField>
                </Columns>
                <FooterStyle BackColor="#CCCCCC" />
                <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
                <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />
                <HeaderStyle BackColor="Black" Font-Bold="True" ForeColor="White" HorizontalAlign="Center"
                    VerticalAlign="Middle" />
                <AlternatingRowStyle BackColor="#CCCCCC" />
            </asp:GridView>
        </asp:Panel>
        &nbsp;&nbsp;
        <input id="BEDIT" runat="server" style="z-index: 122; left: 896px; width: 136px;
            position: absolute; top: 280px; height: 56px " type="button" value="修改" />
        <asp:TextBox ID="DID" runat="server" BackColor="White" BorderStyle="None" ForeColor="Blue"
            Height="18px" onkeyup="if(isNaN(value))execCommand('undo')" Style="z-index: 115;
            left: 1624px; position: absolute; top: 248px" Width="72px"></asp:TextBox>
        <asp:HyperLink ID="HyperLink1" runat="server" BorderStyle="None" Height="1px" ImageUrl="~/images/Invoice.jpg"
            NavigateUrl="~/InvoiceExample.html" Style="z-index: 139; left: 896px; position: absolute;
            top: 48px" Target="_blank" Width="88px">HyperLink</asp:HyperLink>
        <asp:HyperLink ID="HyperLink3" runat="server" BorderStyle="None" Height="1px" ImageUrl="~/images/trip.jpg"
            NavigateUrl="~/MealsList.aspx" Style="z-index: 139; left: 744px; position: absolute;
            top: 48px" Target="_blank" Width="88px">HyperLink</asp:HyperLink>
        <asp:HyperLink ID="HyperLink2" runat="server" BorderStyle="None" Height="1px" ImageUrl="~/images/Import.jpg"
            NavigateUrl="~/ImportTaxEample.aspx" Style="z-index: 139; left: 1048px; position: absolute;
            top: 48px" Target="_blank" Width="88px">HyperLink</asp:HyperLink>
    </div>
    </form>
</body>
</html>
