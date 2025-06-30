<%@ Page Language="VB" AutoEventWireup="false" CodeFile="AddItemListAcc.aspx.vb" Inherits="AddItemListACC" %>

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
			window.open('RemarkList.aspx?field=' + strField +'&field1='+ strField1,'Popup','width=600,height=350,top=10,resizable=yes');
		}
        			
       function GetReference()
       {
            window.open('ReferenceListAcc.aspx','Popup','status=0,toolbar=0,width=1200,height=400,top=10,resizable=yes');
       }
                			
       function AddExpense()
       {
            window.open('ExpenseList.aspx','_blank','status=0,toolbar=0,width=600,height=400,top=10,resizable=yes');
       }
       
       

	</script>
</head>
<body>
    <form id="Form1" runat="server">
    <div>
        <!-- -->
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- ++  底圖                                                                                ++ -->
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        &nbsp;<!-- --><!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ --><!-- ++  Date                                                                            ++ -->
        <input id="BSDate" runat="server" style="z-index: 126; left: 688px; width: 24px;
            position: absolute; top: 72px" type="button" value="..." />
        <asp:TextBox ID="DSDate" runat="server" BackColor="GreenYellow" BorderStyle="Groove" ForeColor="Blue"
            Height="18px" Style="z-index: 100; left: 568px; position: absolute; top: 72px" Width="112px"></asp:TextBox>
        &nbsp;<!-- ++  ExpItem                                                                             ++ -->
        &nbsp;

        <!-- ++  稅                                                                       ++ -->
        <asp:DropDownList style="Z-INDEX: 101; LEFT: 136px; POSITION: absolute; TOP: 72px" id="DType" runat="server" Width="128px"  BackColor="GreenYellow" AutoPostBack="True">
        </asp:DropDownList>
        &nbsp;&nbsp;&nbsp;<!-- ++  總額                                                                            ++ -->
        <!-- ++  內容                                                                            ++ -->
        &nbsp;<!-- ++  稅額                                                                            ++ -->&nbsp;<!-- ++  淨額                                                                            ++ -->
        <!-- ++  備註                                                                            ++ -->
        &nbsp; &nbsp;<!-- ------------------------------------------------------------------------------------------ --><!-- ++  表格 ++ -->
        <input id="BADD" runat="server" style="z-index: 125; left: 768px; width: 88px;
            position: absolute; top: 144px; height: 56px " type="button" value="登錄" />
        <input id="BClose" runat="server" style="z-index: 122; left: 992px; width: 88px;
            position: absolute; top: 144px; height: 56px " type="button" value="完成" />
        &nbsp;<!-- ------------------------------------------------------------------------------------------ --><!-- ++  連結 ++ -->
        <asp:TextBox ID="DACID" runat="server" BackColor="White" BorderStyle="None" ForeColor="White"
            Height="24px" Style="z-index: 102; left: 1608px; position: absolute; top: 176px"
            Width="142px"></asp:TextBox>
        &nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;
        <asp:TextBox ID="DExpItem1" runat="server" BackColor="White" BorderStyle="None" ForeColor="White"
            Height="24px" Style="z-index: 103; left: 1608px; position: absolute; top: 128px"
            Width="142px"></asp:TextBox>
        <asp:TextBox ID="DContent1" runat="server" BackColor="White" BorderStyle="None" ForeColor="White"
            Height="24px" Style="z-index: 104; left: 1608px; position: absolute; top: 72px"
            Width="448px"></asp:TextBox>
        <asp:TextBox ID="DRemark1" runat="server" BackColor="White" BorderStyle="None" ForeColor="White"
            Height="24px" Style="z-index: 105; left: 1608px; position: absolute; top: 32px"
            Width="142px"></asp:TextBox>
        &nbsp;
        <asp:Panel ID="Panel2" runat="server" BorderWidth="1px" Height="320px" ScrollBars="Auto"
            Style="left: 16px; position: absolute; top: 264px; z-index: 106;" Width="1304px">
            &nbsp;&nbsp;
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" AutoGenerateDeleteButton="True"
            BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px"
            CellPadding="3" DataKeyNames="Unique_ID" Font-Size="9pt" ForeColor="Black" GridLines="Vertical"
            Style="z-index: 103; left: 8px; position: absolute; top: 0px" PageSize="100" AutoGenerateSelectButton="True">
            <Columns>
                <asp:BoundField DataField="Unique_ID" HeaderText="ID">
                    <HeaderStyle Height="20px" />
                    <ItemStyle Width="100px" />
                </asp:BoundField>
                <asp:BoundField DataField="Type" HeaderText="類別用途">
                    <HeaderStyle Height="20px" />
                    <ItemStyle Width="80px" />
                </asp:BoundField>
                <asp:BoundField DataField="Pay" HeaderText="支付" />
                <asp:BoundField DataField="SDate" HeaderText="日期">
                    <HeaderStyle Height="20px" />
                    <ItemStyle Width="50px" />
                </asp:BoundField>
                <asp:BoundField DataField="Currency" HeaderText="幣別">
                    <ItemStyle Width="100px" />
                </asp:BoundField>
                <asp:BoundField DataField="Money" HeaderText="金額" DataFormatString="{0:N2}">
                    <HeaderStyle Height="20px" />
                    <ItemStyle Width="100px" HorizontalAlign="Right" />
                </asp:BoundField>
                <asp:BoundField DataField="Days" HeaderText="數量/次數">
                    <HeaderStyle Height="20px" />
                    <ItemStyle Width="100px" />
                </asp:BoundField>
                <asp:BoundField DataField="Rate" HeaderText="匯率">
                    <HeaderStyle Height="20px" />
                    <ItemStyle Width="50px" HorizontalAlign="Right" />
                </asp:BoundField>
                <asp:BoundField DataField="SumAmt" HeaderText="小計金額(台幣)" DataFormatString="{0:N0}" >
                    <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>
                <asp:BoundField DataField="ExpenseNo" HeaderText="交際費單號" />
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
        <asp:Image ID="Image1" runat="server" ImageUrl="~/images/AddItemACC.jpg" Style="z-index: 99;
            left: 16px; position: absolute; top: 48px" />
        <asp:TextBox ID="DRate1" runat="server" BackColor="GreenYellow" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" onkeyup="if(isNaN(value))execCommand('undo')"
            Style="z-index: 108; left: 232px; position: absolute; top: 208px" Width="72px"></asp:TextBox>
        <asp:DropDownList style="Z-INDEX: 109; LEFT: 144px; POSITION: absolute; TOP: 208px" id="DCurrency1" runat="server" Width="72px"  BackColor="GreenYellow">
            </asp:DropDownList>
        <asp:DropDownList style="Z-INDEX: 110; LEFT: 272px; POSITION: absolute; TOP: 72px" id="DDayMoney" runat="server" Width="168px"  BackColor="GreenYellow" AutoPostBack="True" Visible="False">
            </asp:DropDownList>
        <input id="BExpense" runat="server" style="z-index: 128; left: 696px; width: 24px;
            position: absolute; top: 144px" type="button" value="..." visible="false" />
        <asp:TextBox ID="DExpenseNo" runat="server" BackColor="LightGray" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" Style="z-index: 111; left: 480px; position: absolute;
            top: 144px; text-align: left" Width="200px" Enabled="False"></asp:TextBox>
        <asp:TextBox ID="DRemark" runat="server" BackColor="GreenYellow" BorderStyle="Groove"
            ForeColor="Blue" Height="18px"
            Style="z-index: 112; left: 136px; position: absolute; top: 176px" Width="408px"></asp:TextBox>
        &nbsp;&nbsp;&nbsp; &nbsp; &nbsp;
        &nbsp; &nbsp; &nbsp;&nbsp;
        <asp:TextBox ID="DDays" runat="server"  onkeyup="if(isNaN(value))execCommand('undo')"  BackColor="GreenYellow" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" Style="z-index: 113; left: 352px;
            position: absolute; top: 144px" Width="72px"></asp:TextBox>
        <asp:DropDownList style="Z-INDEX: 114; LEFT: 144px; POSITION: absolute; TOP: 144px" id="DCurrency" runat="server" Width="72px"  BackColor="GreenYellow">
        </asp:DropDownList>
        <asp:TextBox ID="DMoney" runat="server"  onkeyup="if(isNaN(value))execCommand('undo')"  BackColor="GreenYellow" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" Style="z-index: 115; left: 248px;
            position: absolute; top: 144px" Width="72px"></asp:TextBox>
        <asp:DropDownList style="Z-INDEX: 116; LEFT: 32px; POSITION: absolute; TOP: 144px" id="DPay" runat="server" Width="88px"  BackColor="GreenYellow" AutoPostBack="True">
        </asp:DropDownList>
        &nbsp;&nbsp;
        <input id="BReference" runat="server" style="z-index: 127; left: 992px; width: 88px;
            position: absolute; top: 72px; height: 56px" type="button" value="消算參照" visible="false" />
        &nbsp;
        <asp:TextBox ID="DID" runat="server" BackColor="White" BorderStyle="None" ForeColor="Blue"
            Height="18px" onkeyup="if(isNaN(value))execCommand('undo')" Style="z-index: 117;
            left: 1624px; position: absolute; top: 248px" Width="72px"></asp:TextBox>
        <input id="BEDIT" runat="server" style="z-index: 123; left: 880px; width: 88px;
            position: absolute; top: 144px; height: 56px " type="button" value="修改" /><input id="BChange" runat="server" style="z-index: 124; left: 320px; width: 88px;
            position: absolute; top: 208px; height: 24px " type="button" value="變更" />
        <asp:TextBox ID="DRate" runat="server" BackColor="LightGray" BorderStyle="Groove"
            Enabled="False" ForeColor="Blue" Height="18px" onkeyup="if(isNaN(value))execCommand('undo')"
            Style="z-index: 118; left: 776px; position: absolute; top: 224px" Visible="False"
            Width="72px"></asp:TextBox>
        <asp:TextBox ID="DENo" runat="server" BackColor="White" BorderStyle="None" ForeColor="White"
            Height="24px" Style="z-index: 119; left: 1280px; position: absolute; top: 160px"
            Width="142px"></asp:TextBox>
        <asp:TextBox ID="DDTNO" runat="server" BackColor="White" BorderStyle="None" ForeColor="White"
            Height="24px" Style="z-index: 120; left: 1272px; position: absolute; top: 112px"
            Width="142px"></asp:TextBox>
        <asp:Label ID="Label2" runat="server" Font-Bold="True" ForeColor="Red" Style="z-index: 121;
            left: 552px; position: absolute; top: 176px" Text="※請再次確認日當天數"></asp:Label>
        &nbsp;
        <asp:HyperLink ID="HyperLink1" runat="server" Height="40px" ImageUrl="~/images/Ratejpg.jpg"
            Style="z-index: 139; left: 768px; position: absolute; top: 72px" Width="1px" NavigateUrl="~/RateList.aspx" Target="_blank" BorderStyle="None">HyperLink</asp:HyperLink>
        <asp:ImageButton ID="ImageButton1" runat="server" /></div>
    </form>
</body>
</html>
