<%@ Page Language="VB" AutoEventWireup="false" CodeFile="PS_INQ_Item03.aspx.vb" Inherits="PS_INQ_Item03" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Item Combination Inf.</title>
	    <script language="javascript" src="ForProject.js"></script>

</head>
<body>
    <form id="form1" runat="server">
    <div>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  底圖                                                                                ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <img src="iMages/ITEMSOURCE_01.png" style="z-index: 1; left: 6px; position: absolute;top: 0px; width: 334px; height: 80px;"/>
		
		<asp:hyperlink id="LSearch" style="Z-INDEX: 102; LEFT: 584px; POSITION: absolute; TOP: 8px" runat="server"
					Width="80px" Height="16px" Font-Size="12pt" NavigateUrl="iMages/SEARCHREMARK.pdf" Target="_blank">Search說明</asp:hyperlink>

        <asp:hyperlink style="Z-INDEX: 102; LEFT: 832px; POSITION: absolute; TOP: 40px" id="LSliderBodyReplacement" runat="server" Target="_blank" NavigateUrl="http://10.245.0.3/ISOS/Document/000151/SliderBody_Replacement.xlsx" Font-Size="12pt" Height="16px" Width="120px">胴體切替一覽</asp:hyperlink>
        <asp:hyperlink style="Z-INDEX: 102; LEFT: 832px; POSITION: absolute; TOP: 64px" id="LFinishList" runat="server" Target="_blank" NavigateUrl="http://10.245.0.3/ISOS/Document/000151/FinishList.png" Font-Size="12pt" Height="16px" Width="120px">表面處理一覽</asp:hyperlink>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Action (CheckBox)                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:CheckBox ID="AtSPD" runat="server" style="z-index: 174; left: 30px; position: absolute; top: 22px" Font-Size="12pt" Text="" Width="20px" AutoPostBack="True" />
        <asp:CheckBox ID="AtIRW" runat="server" style="z-index: 174; left: 96px; position: absolute; top: 22px" Font-Size="12pt" Text="" Width="20px" AutoPostBack="True" />
        <asp:CheckBox ID="AtBuyer" runat="server" style="z-index: 174; left: 166px; position: absolute; top: 12px" Font-Size="12pt" Text="" Width="20px" AutoPostBack="True"  />
        <asp:CheckBox ID="AtFAS" runat="server" style="z-index: 174; left: 260px; position: absolute; top: 12px" Font-Size="12pt" Text="" Width="20px" AutoPostBack="True" />

        <asp:CheckBox ID="AtReplace" runat="server" style="z-index: 174; left: 60px; position: absolute; top: 48px" Font-Size="12pt" Text="" Width="20px" AutoPostBack="True" />
        <asp:CheckBox ID="AtSales" runat="server" style="z-index: 174; left: 166px; position: absolute; top: 48px" Font-Size="12pt" Text="" Width="20px" AutoPostBack="True" />
    
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  篩選欄位                                                                            ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:Label ID="Label1" runat="server" Style="z-index: 103; left: 348px; position: absolute; top: 36px"  Font-Size="10pt" >SEARCH-F</asp:Label>
        <asp:Label ID="Label2" runat="server" Style="z-index: 103; left: 348px; position: absolute; top: 56px"  Font-Size="10pt" >SEARCH-1</asp:Label>
        <asp:Label ID="Label3" runat="server" Style="z-index: 103; left: 348px; position: absolute; top: 76px"  Font-Size="10pt" >SEARCH-2</asp:Label>
        <asp:Label ID="Label4" runat="server" Style="z-index: 103; left: 280px; position: absolute; top: 14px"  Font-Size="10pt" BackColor="Black" ForeColor="White" 
        Width="40px" Height="16px">_自學_</asp:Label>

        <asp:Label ID="Label5" runat="server" Style="z-index: 103; left: 32px; position: absolute; top: 184px"  Font-Size="10pt" Width="848px" >-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------</asp:Label>

        <asp:TextBox ID="DBUYER" runat="server" BackColor="#80FF80" BorderStyle="Groove"
                ForeColor="Blue" Height="14px"  Width="100px" Style="z-index: 103; left: 348px; position: absolute; top: 8px"></asp:TextBox>

        <asp:TextBox ID="DFKEY" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="14px"  Width="248px" Style="z-index: 103; left: 412px; position: absolute; top: 34px"  Font-Size="10pt"></asp:TextBox>

        <asp:TextBox ID="DKEY1" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="14px"  Width="248px" Style="z-index: 103; left: 412px; position: absolute; top: 54px"  Font-Size="10pt"></asp:TextBox>

        <asp:TextBox ID="DKEY2" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="14px"  Width="248px" Style="z-index: 103; left: 412px; position: absolute; top: 74px"  Font-Size="10pt"></asp:TextBox>

        <asp:TextBox ID="DRemark" runat="server" BackColor="White" BorderColor="Black" Font-Size="16pt"
            Height="24px" ReadOnly="true" Style="z-index: 100; left: 832px; position: absolute;
            top: 8px" Width="240px"> BULK &amp; 20210601 ~ </asp:TextBox>

        <asp:TextBox ID="DZKEY" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="16px"  Width="200px" Style="z-index: 103; left: 24px; position: absolute; top: 205px"  Font-Size="10pt"></asp:TextBox>
        <asp:TextBox ID="DPKEY" runat="server" BackColor="Yellow" BorderStyle="Groove"
                ForeColor="Blue" Height="16px"  Width="200px" Style="z-index: 103; left: 334px; position: absolute; top: 205px"  Font-Size="10pt"></asp:TextBox>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  按鈕                                                                                ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
    	<asp:button id="BInq" runat="server" Height="64px" Width="64px" Style="z-index: 103; left: 676px; position: absolute; top: 8px" ForeColor="Black" BackColor="Silver" Text="Go" Font-Size="Larger"></asp:button>
        <asp:ImageButton ID="BExcel" runat="server" Height="24px" ImageUrl="Images\msexcel.gif" Style="z-index: 103; left: 16px; position: absolute; top: 140px" Width="24px" />
    
    	<asp:button id="BSPInq" runat="server" Height="64px" Width="80px" Style="z-index: 103; left: 744px; position: absolute; top: 8px" ForeColor="Black" BackColor="Silver" Text="特別篩選(B)" ></asp:button>

    	<asp:button id="BGetItem" runat="server" Height="40px" Width="750px" Style="z-index: 103; left: 24px; position: absolute; top: 136px" ForeColor="Black" BackColor="Silver" Text="搜尋ITEM" ></asp:button>
    	<asp:button id="BReset" runat="server" Height="40px" Width="104px" Style="z-index: 103; left: 776px; position: absolute; top: 136px" ForeColor="Black" BackColor="Silver" Text="RESET" ></asp:button>

    	<asp:button id="BZKEY" runat="server" Height="28px" Width="48px" Style="z-index: 103; left: 234px; position: absolute; top: 200px" ForeColor="Black" BackColor="Silver" Text="Go" ></asp:button>
    	<asp:button id="BPKEY" runat="server" Height="28px" Width="48px" Style="z-index: 103; left: 544px; position: absolute; top: 200px" ForeColor="Black" BackColor="Silver" Text="Go" ></asp:button>
   
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  隱藏欄位                                                                       ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- ------------------------------------------------------------------------------------ -->
        <!-- System Parameter  -->
        <asp:TextBox ID="HBuyerCode" runat="server" Height="16px" Style="z-index: 318; left: 8px;
            position: absolute; top: 192px;font-size:10px;background:transparent" Width="600px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>

        <asp:TextBox ID="DINQBUYERITEMFile" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  ITEM KEY 欄位                                                                       ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:TextBox ID="LItemKey1" runat="server" BackColor="Black" BorderStyle="Groove"
                ForeColor="White" Height="14px"  Width="250px" Style="z-index: 103; left: 24px; position: absolute; top: 96px" Font-Size="10pt">Z#</asp:TextBox>
        <asp:TextBox ID="DItemKey1" runat="server" BackColor="#80FFFF" BorderStyle="Groove"
                ForeColor="Red" Height="14px"  Width="250px" Style="z-index: 103; left: 24px; position: absolute; top: 116px" Font-Size="10pt"></asp:TextBox>

        <asp:TextBox ID="LItemKey2" runat="server" BackColor="Black" BorderStyle="Groove"
                ForeColor="White" Height="14px"  Width="150px" Style="z-index: 103; left: 284px; position: absolute; top: 96px" Font-Size="10pt">P#</asp:TextBox>
        <asp:TextBox ID="DItemKey2" runat="server" BackColor="#80FFFF" BorderStyle="Groove"
                ForeColor="Red" Height="14px"  Width="150px" Style="z-index: 103; left: 284px; position: absolute; top: 116px" Font-Size="10pt"></asp:TextBox>

        <asp:TextBox ID="LItemKey3" runat="server" BackColor="Black" BorderStyle="Groove"
                ForeColor="White" Height="14px"  Width="100px" Style="z-index: 103; left: 444px; position: absolute; top: 96px" Font-Size="10pt">F#</asp:TextBox>
        <asp:TextBox ID="DItemKey3" runat="server" BackColor="#80FFFF" BorderStyle="Groove"
                ForeColor="red" Height="14px"  Width="100px" Style="z-index: 103; left: 444px; position: absolute; top: 116px" Font-Size="10pt"></asp:TextBox>

        <asp:TextBox ID="LItemKey4" runat="server" BackColor="Black" BorderStyle="Groove"
                ForeColor="White" Height="14px"  Width="100px" Style="z-index: 103; left: 554px; position: absolute; top: 96px" Font-Size="10pt">T# / OTHER</asp:TextBox>
        <asp:TextBox ID="DItemKey4" runat="server" BackColor="#80FFFF" BorderStyle="Groove"
                ForeColor="red" Height="14px"  Width="100px" Style="z-index: 103; left: 554px; position: absolute; top: 116px" Font-Size="10pt"></asp:TextBox>

        <asp:TextBox ID="LItemKey5" runat="server" BackColor="Black" BorderStyle="Groove"
                ForeColor="White" Height="14px"  Width="100px" Style="z-index: 103; left: 664px; position: absolute; top: 96px" Font-Size="10pt">SPC</asp:TextBox>
        <asp:TextBox ID="DItemKey5" runat="server" BackColor="#80FFFF" BorderStyle="Groove"
                ForeColor="red" Height="14px"  Width="100px" Style="z-index: 103; left: 664px; position: absolute; top: 116px" Font-Size="10pt"></asp:TextBox>

        <asp:TextBox ID="LLeadTime" runat="server" BackColor="Black" BorderStyle="Groove"
                ForeColor="White" Height="14px"  Width="100px" Style="z-index: 103; left: 774px; position: absolute; top: 96px" Font-Size="10pt">L/T[Z#]</asp:TextBox>
        <asp:TextBox ID="DLeadTime" runat="server" BackColor="#80FFFF" BorderStyle="Groove"
                ForeColor="red" Height="14px"  Width="100px" Style="z-index: 103; left: 774px; position: absolute; top: 116px" Font-Size="10pt">0</asp:TextBox>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Customer GridView-1                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!--            <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />  -->

        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" Width="300px" Style="z-index: 103; left: 24px; position: absolute; top: 232px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" DataKeyNames="J1" >
            <Columns>
                <asp:BoundField DataField="A1" HeaderText=""  />
                <asp:BoundField DataField="B1" HeaderText=""  />
                <asp:BoundField DataField="C1" HeaderText=""  />
                <asp:BoundField DataField="D1" HeaderText=""  />
                <asp:BoundField DataField="E1" HeaderText=""  />
                <asp:BoundField DataField="F1" HeaderText=""  />
                <asp:BoundField DataField="G1" HeaderText=""  />
                <asp:BoundField DataField="H1" HeaderText=""  />
                <asp:BoundField DataField="I1" HeaderText=""  />
                <asp:BoundField DataField="J1" HeaderText=""  />
                <asp:BoundField DataField="K1" HeaderText=""  />
                <asp:BoundField DataField="L1" HeaderText=""  />
                <asp:BoundField DataField="M1" HeaderText=""  />
                <asp:BoundField DataField="N1" HeaderText=""  />
                <asp:BoundField DataField="O1" HeaderText=""  />
                <asp:BoundField DataField="P1" HeaderText=""  />
                <asp:BoundField DataField="Q1" HeaderText=""  />
                <asp:BoundField DataField="R1" HeaderText=""  />
                <asp:BoundField DataField="S1" HeaderText=""  />
                <asp:BoundField DataField="T1" HeaderText=""  />
                <asp:BoundField DataField="U1" HeaderText=""  />
                <asp:BoundField DataField="V1" HeaderText=""  />
                <asp:BoundField DataField="W1" HeaderText=""  />
                <asp:BoundField DataField="X1" HeaderText=""  />
                <asp:BoundField DataField="Y1" HeaderText=""  />
                <asp:BoundField DataField="Z1" HeaderText=""  />
                
                <asp:HyperLinkField DataNavigateUrlFields="URL" DataNavigateUrlFormatString="{0}" 
                DataTextField="Mark" HeaderText="Select" Target="_blank"  >
                </asp:HyperLinkField>

                <asp:HyperLinkField DataNavigateUrlFields="FormURL" DataNavigateUrlFormatString="{0}" 
                DataTextField="FormMark" HeaderText="Select" Target="_blank">
                </asp:HyperLinkField>

                <asp:CommandField ShowSelectButton="True" SelectText="Ｓ" />

                <asp:TemplateField HeaderText="" ShowHeader="False">
                <ItemTemplate>
                    <asp:LinkButton ID="LinkButton1" runat="server" CausesValidation="False" 
                        CommandName="Delete" Text="Ｄ" OnClientClick="return confirm('你確定要刪除嗎？')"></asp:LinkButton>
                </ItemTemplate>
            </asp:TemplateField>

            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Customer GridView-2                                                               ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!--            <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />  -->

        <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" Width="300px" Style="z-index: 103; left: 334px; position: absolute; top: 232px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="A1" HeaderText=""  />
                <asp:BoundField DataField="B1" HeaderText=""  />
                <asp:BoundField DataField="C1" HeaderText=""  />
                <asp:BoundField DataField="D1" HeaderText=""  />
                <asp:BoundField DataField="E1" HeaderText=""  />
                <asp:BoundField DataField="F1" HeaderText=""  />
                <asp:BoundField DataField="G1" HeaderText=""  />
                <asp:BoundField DataField="H1" HeaderText=""  />
                <asp:BoundField DataField="I1" HeaderText=""  />
                <asp:BoundField DataField="J1" HeaderText=""  />
                <asp:BoundField DataField="K1" HeaderText=""  />
                <asp:BoundField DataField="L1" HeaderText=""  />
                <asp:BoundField DataField="M1" HeaderText=""  />
                <asp:BoundField DataField="N1" HeaderText=""  />
                <asp:BoundField DataField="O1" HeaderText=""  />
                <asp:BoundField DataField="P1" HeaderText=""  />
                <asp:BoundField DataField="Q1" HeaderText=""  />
                <asp:BoundField DataField="R1" HeaderText=""  />
                <asp:BoundField DataField="S1" HeaderText=""  />
                <asp:BoundField DataField="T1" HeaderText=""  />
                <asp:BoundField DataField="U1" HeaderText=""  />
                <asp:BoundField DataField="V1" HeaderText=""  />
                <asp:BoundField DataField="W1" HeaderText=""  />
                <asp:BoundField DataField="X1" HeaderText=""  />
                <asp:BoundField DataField="Y1" HeaderText=""  />
                <asp:BoundField DataField="Z1" HeaderText=""  />
                
                <asp:HyperLinkField DataNavigateUrlFields="URL" DataNavigateUrlFormatString="{0}" 
                DataTextField="Mark" HeaderText="Select" Target="_blank"  >
                </asp:HyperLinkField>

                <asp:HyperLinkField DataNavigateUrlFields="FormURL" DataNavigateUrlFormatString="{0}" 
                DataTextField="FormMark" HeaderText="Select" Target="_blank">
                </asp:HyperLinkField>

                <asp:CommandField ShowSelectButton="True" SelectText="Ｓ" />

            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>        
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Customer GridView-3                                                              ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!--            <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />  -->

        <asp:GridView ID="GridView3" runat="server" AutoGenerateColumns="False" Width="50px" Style="z-index: 103; left: 644px; position: absolute; top: 232px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="C1" HeaderText=""  />
                <asp:CommandField ShowSelectButton="True" SelectText="Ｓ" />
            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>        
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Customer GridView-4                                                              ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!--            <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />  -->

        <asp:GridView ID="GridView4" runat="server" AutoGenerateColumns="False" Width="50px" Style="z-index: 103; left: 704px; position: absolute; top: 232px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="C1" HeaderText=""  />
                <asp:CommandField ShowSelectButton="True" SelectText="Ｓ" />
            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>  
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Customer GridView-6                                                              ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!--            <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />  -->
        <asp:GridView ID="GridView6" runat="server" AutoGenerateColumns="False" Width="80px" Style="z-index: 103; left: 784px; position: absolute; top: 232px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:BoundField DataField="B1" HeaderText=""  />
                <asp:CommandField ShowSelectButton="True" SelectText="Ｓ" />
            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>  
              
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  自學GridView-5 (ITEM)                                                              ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!--            <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />  -->

        <asp:GridView ID="GridView5" runat="server" AutoGenerateColumns="False" Width="700px" Style="z-index: 103; left: 874px; position: absolute; top: 232px" CellPadding="3" Font-Size="9pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
            <Columns>
                <asp:CommandField ShowSelectButton="True" SelectText="確定" />
                <asp:BoundField DataField="ItemInf" HeaderText="ItemInf"  />
                <asp:BoundField DataField="Code" HeaderText=""  />
                <asp:BoundField DataField="Name1" HeaderText=""  />
                <asp:BoundField DataField="Name2" HeaderText=""  />
                <asp:BoundField DataField="Name3" HeaderText=""  />
            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <AlternatingRowStyle BackColor="#CCCCCC" />
        </asp:GridView>

    
    </div>
    </form>
</body>
</html>
