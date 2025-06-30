<%@ Page Language="VB" AutoEventWireup="false" CodeFile="AddItemListBC.aspx.vb" Inherits="AddItemListBC" %>

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
            window.open('ReferenceList.aspx','Popup','status=0,toolbar=0,width=700,height=400,top=10,resizable=yes');
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
        <input id="BSDate" runat="server" style="z-index: 121; left: 240px; width: 24px;
            position: absolute; top: 128px" type="button" value="..." />
        <asp:TextBox ID="DSDate" runat="server" BackColor="GreenYellow" BorderStyle="Groove" ForeColor="Blue"
            Height="18px" Style="z-index: 100; left: 144px; position: absolute; top: 128px" Width="88px" AutoPostBack="True"></asp:TextBox>
        &nbsp;<!-- ++  ExpItem                                                                             ++ -->
        &nbsp;

        <!-- ++  稅                                                                       ++ -->
        <asp:DropDownList style="Z-INDEX: 102; LEFT: 144px; POSITION: absolute; TOP: 64px" id="DType" runat="server" Width="128px"  BackColor="GreenYellow" AutoPostBack="True">
        </asp:DropDownList>
        &nbsp;&nbsp;&nbsp;&nbsp;<!-- ++  總額                                                                            ++ -->
        <!-- ++  內容                                                                            ++ -->
        &nbsp;<!-- ++  稅額                                                                            ++ -->&nbsp;<!-- ++  淨額                                                                            ++ -->
        <!-- ++  備註                                                                            ++ -->
        &nbsp; 
        <asp:Label ID="Label1" runat="server" Font-Bold="True" ForeColor="Red" Style="z-index: 121;
            left: 32px; position: absolute; top: 16px" Text="※航班及住宿都要填寫 "></asp:Label>
        &nbsp;<!-- ------------------------------------------------------------------------------------------ --><!-- ++  表格 ++ -->
        <input id="BADD" runat="server" style="z-index: 116; left: 664px; width: 136px;
            position: absolute; top: 192px; height: 56px " type="button" value="登錄" />
        <input id="BClose" runat="server" style="z-index: 115; left: 808px; width: 136px;
            position: absolute; top: 192px; height: 56px " type="button" value="完成" />
        &nbsp;<!-- ------------------------------------------------------------------------------------------ --><!-- ++  連結 ++ -->
        <asp:TextBox ID="DACID" runat="server" BackColor="White" BorderStyle="None" ForeColor="White"
            Height="24px" Style="z-index: 109; left: 1608px; position: absolute; top: 176px"
            Width="142px"></asp:TextBox>
        &nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;
        <asp:TextBox ID="DExpItem1" runat="server" BackColor="White" BorderStyle="None" ForeColor="White"
            Height="24px" Style="z-index: 110; left: 1608px; position: absolute; top: 128px"
            Width="142px"></asp:TextBox>
        <asp:TextBox ID="DContent1" runat="server" BackColor="White" BorderStyle="None" ForeColor="White"
            Height="24px" Style="z-index: 111; left: 1608px; position: absolute; top: 72px"
            Width="448px"></asp:TextBox>
        <asp:TextBox ID="DRemark1" runat="server" BackColor="White" BorderStyle="None" ForeColor="White"
            Height="24px" Style="z-index: 112; left: 1608px; position: absolute; top: 32px"
            Width="142px"></asp:TextBox>
        &nbsp;
        <asp:Panel ID="Panel2" runat="server" BorderWidth="1px" Height="320px" ScrollBars="Auto"
            Style="left: 24px; position: absolute; top: 288px; z-index: 113;" Width="1304px">
            &nbsp;&nbsp;
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" AutoGenerateDeleteButton="True"
            BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px"
            CellPadding="3" DataKeyNames="Unique_ID" Font-Size="9pt" ForeColor="Black" GridLines="Vertical"
            Style="z-index: 103; left: 8px; position: absolute; top: 16px" PageSize="100">
            <Columns>
                <asp:BoundField DataField="Unique_ID" HeaderText="ID">
                    <HeaderStyle Height="20px" />
                    <ItemStyle Width="100px" />
                </asp:BoundField>
                <asp:BoundField DataField="Type" HeaderText="類別">
                    <HeaderStyle Height="20px" />
                    <ItemStyle Width="50px" />
                </asp:BoundField>
                <asp:BoundField DataField="Appoint" HeaderText="預約">
                    <HeaderStyle Height="20px" />
                    <ItemStyle Width="50px" />
                </asp:BoundField>
                <asp:BoundField DataField="SDate" HeaderText="出發/入住">
                    <HeaderStyle Height="20px" />
                    <ItemStyle Width="100px" />
                </asp:BoundField>
                <asp:BoundField DataField="EDate" HeaderText="到達/退房 ">
                    <HeaderStyle Height="20px" />
                    <ItemStyle Width="100px" />
                </asp:BoundField>
                <asp:BoundField DataField="Days" HeaderText="天數">
                    <HeaderStyle Height="20px" />
                    <ItemStyle Width="50px" />
                </asp:BoundField>
                <asp:BoundField DataField="Currency" HeaderText="幣別" />
                <asp:BoundField DataField="Money" HeaderText="金額">
                    <HeaderStyle Height="20px" />
                    <ItemStyle Width="50px" />
                </asp:BoundField>
                <asp:BoundField DataField="FlyInf" HeaderText="航班資訊">
                    <ItemStyle Width="100px" />
                </asp:BoundField>
                <asp:BoundField DataField="SFly" HeaderText="起點" />
                <asp:BoundField DataField="EFly" HeaderText="迄點" />
                <asp:BoundField DataField="HotelInf" HeaderText="飯店資訊" />
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
        <asp:Image ID="Image1" runat="server" ImageUrl="~/images/AddBC02.jpg" Style="z-index: 99;
            left: 16px; position: absolute; top: 40px" />
        <asp:TextBox ID="DSFly" runat="server" BackColor="GreenYellow" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" Style="z-index: 104; left: 504px; position: absolute;
            top: 96px" Visible="False" Width="112px"></asp:TextBox>
        <asp:TextBox ID="DFlyInf" runat="server" BackColor="GreenYellow" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" Style="z-index: 104; left: 504px; position: absolute;
            top: 64px" Visible="False" Width="112px"></asp:TextBox>
        &nbsp;
        <asp:TextBox ID="DEFly" runat="server" BackColor="GreenYellow" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" Style="z-index: 104; left: 504px; position: absolute;
            top: 128px" Visible="False" Width="112px"></asp:TextBox>
        <asp:DropDownList style="Z-INDEX: 102; LEFT: 272px; POSITION: absolute; TOP: 160px" id="DETime1" runat="server" Width="56px"  BackColor="Yellow" AutoPostBack="true">
            </asp:DropDownList>
        <asp:DropDownList style="Z-INDEX: 102; LEFT: 336px; POSITION: absolute; TOP: 160px" id="DETime2" runat="server" Width="56px"  BackColor="Yellow" AutoPostBack="true">
        </asp:DropDownList>
        <asp:DropDownList style="Z-INDEX: 102; LEFT: 336px; POSITION: absolute; TOP: 128px" id="DSTime2" runat="server" Width="56px"  BackColor="Yellow" AutoPostBack="true">
        </asp:DropDownList>
        <asp:DropDownList style="Z-INDEX: 102; LEFT: 272px; POSITION: absolute; TOP: 128px" id="DSTime1" runat="server" Width="56px"  BackColor="Yellow" AutoPostBack="true">
        </asp:DropDownList>
        <asp:TextBox ID="DRemark" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue" TextMode="MultiLine"
            Height="56px" Style="z-index: 107; left: 504px; position: absolute; top: 192px" Width="120px"></asp:TextBox>
        &nbsp;
        <asp:TextBox ID="DDays" runat="server"  onkeyup="if(isNaN(value))execCommand('undo')"  BackColor="LightGray" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" Style="z-index: 101; left: 144px;
            position: absolute; top: 192px" Width="32px" Enabled="False"></asp:TextBox>
        <asp:DropDownList style="Z-INDEX: 102; LEFT: 256px; POSITION: absolute; TOP: 192px" id="DCurrency" runat="server" Width="96px"  BackColor="GreenYellow" AutoPostBack="true">
        </asp:DropDownList>
        <asp:TextBox ID="DMoney" runat="server"  onkeyup="if(isNaN(value))execCommand('undo')"  BackColor="GreenYellow" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" Style="z-index: 101; left: 144px;
            position: absolute; top: 224px" Width="72px"></asp:TextBox>
        <asp:DropDownList style="Z-INDEX: 102; LEFT: 144px; POSITION: absolute; TOP: 96px" id="DAppoint" runat="server" Width="128px"  BackColor="GreenYellow">
        </asp:DropDownList>
        <input id="BEDate" runat="server" style="z-index: 121; left: 240px; width: 24px;
            position: absolute; top: 160px" type="button" value="..." />
        <asp:TextBox ID="DEDate" runat="server" BackColor="GreenYellow" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" Style="z-index: 100; left: 144px; position: absolute;
            top: 160px" Width="88px" AutoPostBack="True"></asp:TextBox>
        <input id="BReference" runat="server" style="z-index: 133; left: 664px; width: 136px;
            position: absolute; top: 64px; height: 56px" type="button" value="出差參照" visible="false" />
        <asp:TextBox ID="DHotelInf" runat="server" BackColor="GreenYellow" BorderStyle="Groove"
            ForeColor="Blue" Height="18px" Style="z-index: 104; left: 504px; position: absolute;
            top: 160px" Visible="False" Width="112px"></asp:TextBox>
    </div>
    </form>
</body>
</html>
