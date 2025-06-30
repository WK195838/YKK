<%@ Page Language="VB" AutoEventWireup="false" CodeFile="SP_MaterialMenu.aspx.vb" Inherits="SP_MaterialMenu" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>Material Plan</title>
	    <script language="javascript" src="ForProject_SP.js"></script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  底圖                                                                                ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <img src="iMages/SP_Material.png" style="z-index: 1; left: 30px; position: absolute;top: 34px; width: 1310px;" height="630"/>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  各Action按鈕                                                                        ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->

<!-- ****************************************************************************************** -->
<!-- ** Automate                                                                         -->
<!-- ****************************************************************************************** -->
        <asp:Button ID="BO365" runat="server" Height="32px" Style="z-index: 104; left: 344px;
            position: absolute; top: 16px" Text="New Customer" Width="112px" onclientclick="javascript:return confirm('確定執行？');" />

        <asp:Button ID="BToolBox" runat="server" Height="32px" Style="z-index: 131;
            left: 472px; position: absolute; top: 16px" Text="Shopping Sheet" Width="112px"  />
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  HYPER LINK
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:HyperLink ID="LUpdateWINGS" runat="server" Font-Bold="True" Font-Size="Larger"
            ForeColor="Blue" Height="24px"
            Style="z-index: 104; left: 1100px; position: absolute; top: 568px; text-align: center"
            Target="_blank" Width="80px">Update WINGS</asp:HyperLink>

        <asp:HyperLink ID="LPendingFinal" runat="server" Font-Bold="True" Font-Size="Medium"
            ForeColor="Blue" Height="24px"
            Style="z-index: 104; left: 344px; position: absolute; top: 56px; text-align: LEFT"
            Target="_blank" Width="240px">Pending Final Count=(1)</asp:HyperLink>
            
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  各按鈕                                                                              ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- 1列  -->
        <asp:ImageButton ID="BCustomer" runat="server" style="z-index: 1; left: 176px; position: absolute;top: 40px; width: 50px; height: 50px;" ImageUrl="iMages/Go.jpg" />
        <asp:ImageButton ID="BReset" runat="server" style="z-index: 1; left: 120px; position: absolute;top: 16px; width: 50px; height: 50px;" ImageUrl="iMages/Reset.jpg" />
        <asp:ImageButton ID="BImport" runat="server" style="z-index: 1; left: 184px; position: absolute;top: 248px; width: 50px; height: 50px;" ImageUrl="iMages/Go.jpg"  />
        <asp:ImageButton ID="BDemand" runat="server" style="z-index: 1; left: 184px; position: absolute;top: 392px; width: 50px; height: 50px;" ImageUrl="iMages/Go.jpg" />
        <asp:ImageButton ID="BActionPlan" runat="server" style="z-index: 1; left: 184px; position: absolute;top:552px; width: 50px; height: 50px;" ImageUrl="iMages/Go.jpg" />
        <!-- 2列  -->
        <asp:ImageButton ID="BActionPlanImport" runat="server" style="z-index: 1; left: 632px; position: absolute;top: 104px; width: 50px; height: 50px;" ImageUrl="iMages/Go.jpg" />
        <asp:ImageButton ID="BActionPlanYOC" runat="server" style="z-index: 1; left: 632px; position: absolute;top: 240px; width: 50px; height: 50px;" ImageUrl="iMages/Go.jpg" />
        <asp:ImageButton ID="BKPIFSheet" runat="server" style="z-index: 1; left: 632px; position: absolute;top:488px; width: 50px; height: 50px;" ImageUrl="iMages/Go.jpg" />
        <!-- 3列  -->
        <asp:ImageButton ID="BWINGS" runat="server" style="z-index: 1; left: 1080px; position: absolute;top: 224px; width: 50px; height: 50px;" ImageUrl="iMages/Go.jpg" />
        <asp:ImageButton ID="BPIL" runat="server" style="z-index: 1; left: 1192px; position: absolute;top: 304px; width: 50px; height: 50px;" ImageUrl="iMages/Go.jpg" />
        <asp:ImageButton ID="BFinal" runat="server" style="z-index: 1; left: 1008px; position: absolute;top: 480px; width: 50px; height: 50px;" ImageUrl="iMages/Go.jpg" />
        <asp:ImageButton ID="BChangeFinal" runat="server" style="z-index: 1; left: 1184px; position: absolute;top: 496px; width: 50px; height: 50px;" ImageUrl="iMages/Go.jpg" />
        <asp:ImageButton ID="BProgress" runat="server" style="z-index: 1; left: 1040px; position: absolute;top: 560px; width: 50px; height: 50px;" ImageUrl="iMages/Go.jpg" />
        <!-- 4列  -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  檢查資料-狀態                                                                       ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- 1列  -->
        <asp:HyperLink ID="StsCustomer" runat="server" style="z-index: 1; left: 176px; position: absolute;top: 88px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
        <asp:HyperLink ID="StsImport" runat="server" style="z-index: 1; left: 184px; position: absolute;top: 296px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
        <asp:HyperLink ID="StsDemand" runat="server" style="z-index: 1; left: 184px; position: absolute;top: 448px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
        <asp:HyperLink ID="StsActionPlan" runat="server" style="z-index: 1; left: 184px; position: absolute;top: 600px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
        <!-- 2列  -->
        <asp:HyperLink ID="StsActionPlanImport" runat="server" style="z-index: 1; left: 632px; position: absolute;top: 152px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
        <asp:HyperLink ID="StsActionPlanYOC" runat="server" style="z-index: 1; left: 632px; position: absolute;top: 288px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
        <asp:HyperLink ID="StsKPIFSheet" runat="server" style="z-index: 1; left: 632px; position: absolute;top: 536px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
        <!-- 3列  -->
        <asp:HyperLink ID="StsWINGS" runat="server" style="z-index: 1; left: 1136px; position: absolute;top: 224px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
        <asp:HyperLink ID="StsPIL" runat="server" style="z-index: 1; left: 1248px; position: absolute;top: 304px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
        <asp:HyperLink ID="StsFinal" runat="server" style="z-index: 1; left: 1064px; position: absolute;top: 480px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
        <asp:HyperLink ID="StsChangeFinal" runat="server" style="z-index: 1; left: 1240px; position: absolute;top: 496px; width: 40px; height: 40px;" ImageUrl="iMages/OK.jpg"/>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  執行跑馬燈                                                                       ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- 1列  -->
        <marquee id="ProImport" runat="server" style="z-index: 1; left: 240px; position: absolute;top: 264px; width: 80px; height: 20px; color: #000099;" >■■■■■■■■■■■■■■■</marquee>
        <marquee id="ProDemand" runat="server" style="z-index: 1; left: 240px; position: absolute;top: 408px; width: 80px; height: 20px; color: #000099;" >■■■■■■■■■■■■■■■</marquee>
        <marquee id="ProActionPlan" runat="server" style="z-index: 1; left: 240px; position: absolute;top: 568px; width: 80px; height: 20px; color: #000099;" >■■■■■■■■■■■■■■■</marquee>
        <!-- 2列  -->
        <marquee id="ProActionPlanImport" runat="server" style="z-index: 1; left: 688px; position: absolute;top: 120px; width: 80px; height: 20px; color: #000099;" >■■■■■■■■■■■■■■■</marquee>
        <marquee id="ProActionPlanYOC" runat="server" style="z-index: 1; left: 688px; position: absolute;top: 256px; width: 80px; height: 20px; color: #000099;" >■■■■■■■■■■■■■■■</marquee>
        <marquee id="ProKPIFSheet" runat="server" style="z-index: 1; left: 688px; position: absolute;top: 504px; width: 80px; height: 20px; color: #000099;" >■■■■■■■■■■■■■■■</marquee>
        <!-- 3列  -->
        <marquee id="ProWINGS" runat="server" style="z-index: 1; left: 1192px; position: absolute;top: 240px; width: 80px; height: 20px; color: #000099;" >■■■■■■■■■■■■■■■</marquee>
        <marquee id="ProPIL" runat="server" style="z-index: 1; left: 1208px; position: absolute;top: 360px; width: 80px; height: 20px; color: #000099;" >■■■■■■■■■■■■■■■</marquee>
        <marquee id="ProFinal" runat="server" style="z-index: 1; left: 1016px; position: absolute;top: 536px; width: 80px; height: 20px; color: #000099;" >■■■■■■■■■■■■■■■</marquee>
        <marquee id="ProChangeFinal" runat="server" style="z-index: 1; left: 1192px; position: absolute;top: 552px; width: 80px; height: 20px; color: #000099;" >■■■■■■■■■■■■■■■</marquee>
        <!-- 4列  -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  客戶選擇                                                                       ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- 客戶  -->
        <asp:DropDownList ID="DCustomerBuyer" runat="server" BackColor="Yellow" ForeColor="Blue"
            Style="z-index: 112; left: 56px; position: absolute; top: 168px" Width="200px">
            <asp:ListItem>12345678901234567890</asp:ListItem>
        </asp:DropDownList>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Action (CheckBox)                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- 客戶  -->
        <asp:CheckBox ID="AtDemand" runat="server" style="z-index: 174; left: 232px; position: absolute; top: 40px" Font-Size="9pt" Text="材料/需求量" Width="100px" AutoPostBack="True" />
<!-- ++  AtActionPlan top=60 / AtActionPlanImport top=80     ++ -->
        <asp:CheckBox ID="AtActionPlan" runat="server" style="z-index: 174; left: 232px; position: absolute; top: -60px" Font-Size="9pt" Text="材料Action" Width="100px" AutoPostBack="True"/>
        <asp:CheckBox ID="AtActionPlanImport" runat="server" style="z-index: 174; left: 232px; position: absolute; top: 60px" Font-Size="9pt" Text="Action Import" Width="100px" AutoPostBack="True"/>
<!-- ++  AtActionPlanYOC top=100     ++ -->
        <asp:CheckBox ID="AtActionPlanYOC" runat="server" style="z-index: 174; left: 232px; position: absolute; top: -100px" Font-Size="9pt" Text="材料計画表" Width="96px" AutoPostBack="True"/>

<!-- ++  AtKPIFSheet top=120  / AtWINGS top=140   ++ -->
        <asp:CheckBox ID="AtKPIFSheet" runat="server" style="z-index: 174; left: 232px; position: absolute; top: 80px" Font-Size="9pt" Text="KP I/F" Width="100px" AutoPostBack="True"/>
        <asp:CheckBox ID="AtWINGS" runat="server" style="z-index: 174; left: 232px; position: absolute; top: 100px" Font-Size="9pt" Text="回報WINGS" Width="100px" AutoPostBack="True"/>

<!-- ++  AtPIL top=160 / AtFinal top=180 / AtChangeFinal top=200  ++ -->
        <asp:CheckBox ID="AtPIL" runat="server" style="z-index: 174; left: 232px; position: absolute; top: -160px" Font-Size="9pt" Text="Purchase Infor" Width="96px" AutoPostBack="True"/>
        <asp:CheckBox ID="AtFinal" runat="server" style="z-index: 174; left: 232px; position: absolute; top: 120px" Font-Size="9pt" Text="Final" Width="100px" AutoPostBack="True" />
        <asp:CheckBox ID="AtChangeFinal" runat="server" style="z-index: 174; left: 232px; position: absolute; top: 140px" Font-Size="9pt" Text="Change Final" Width="100px" AutoPostBack="True"/>

<!-- ++  AtPending top=160 / AtFinal top=180 / AtChangeFinal top=200  ++ -->
        <asp:CheckBox ID="AtPending" runat="server" style="z-index: 1; left: 1048px; position: absolute; top: 464px" Font-Size="9pt" Text="Pending.." Width="72px" AutoPostBack="True"/>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Log                                                                       ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" Style="z-index: 103; left: 150px; position: absolute; top: 200px" CellPadding="3" Font-Size="10pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="Red" BorderStyle="Solid" BorderWidth="5px" >
            <Columns>
                    <asp:BoundField DataField="User"  />
                    <asp:BoundField DataField="AccessTime"  />
                    <asp:BoundField DataField="Cat"  />
                    <asp:BoundField DataField="Active" />
                    <asp:BoundField DataField="Code" />
                    <asp:BoundField DataField="Name"  />

                    <asp:BoundField DataField="Status" />
                   
                    <asp:BoundField DataField="Customer" >  <ItemStyle HorizontalAlign="CENTER"  />   </asp:BoundField>  
                    <asp:BoundField DataField="Import"  >  <ItemStyle HorizontalAlign="CENTER"  />   </asp:BoundField>  
                    <asp:BoundField DataField="Demand"  >  <ItemStyle HorizontalAlign="CENTER"  />   </asp:BoundField>  
                    <asp:BoundField DataField="ActPlan"   >  <ItemStyle HorizontalAlign="CENTER"  />   </asp:BoundField>  
                    <asp:BoundField DataField="ImpActPlan"  >  <ItemStyle HorizontalAlign="CENTER"  />   </asp:BoundField>  
                    <asp:BoundField DataField="RspActPlan"   >  <ItemStyle HorizontalAlign="CENTER"  />   </asp:BoundField>  
                    <asp:BoundField DataField="KPInterface"  >  <ItemStyle HorizontalAlign="CENTER"  />   </asp:BoundField>  

                    <asp:BoundField DataField="RspWINGS"  >  <ItemStyle HorizontalAlign="CENTER"  />   </asp:BoundField>  
                    <asp:BoundField DataField="PILSheet"  >  <ItemStyle HorizontalAlign="CENTER"  />   </asp:BoundField>  
                    <asp:BoundField DataField="Final"   >  <ItemStyle HorizontalAlign="CENTER"  />   </asp:BoundField>  
                    <asp:BoundField DataField="ChgFinal"  >  <ItemStyle HorizontalAlign="CENTER"  />   </asp:BoundField>  
                    <asp:BoundField DataField="Progress"  >  <ItemStyle HorizontalAlign="CENTER"  />   </asp:BoundField>                    

            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="Chartreuse" Font-Bold="False" ForeColor="Black" />
            <AlternatingRowStyle BackColor="#FFF3DB" />
        </asp:GridView>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Pending Final                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" Style="z-index: 103; left: 210px; position: absolute; top: 200px" CellPadding="3" Font-Size="10pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="Red" BorderStyle="Solid" BorderWidth="5px" >
            <Columns>
                    <asp:BoundField DataField="SP_Code"  />
                    <asp:BoundField DataField="SP_Name"  />
                    <asp:BoundField DataField="SP_No"  />
                    <asp:BoundField DataField="SP_Date" />
                    <asp:BoundField DataField="Now_Date" />

                    <asp:BoundField DataField="Diff_Date"  >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>  


                    <asp:HyperLinkField DataNavigateUrlFields="PDLUrl" DataNavigateUrlFormatString="{0}" 
                    DataTextField="PDL" HeaderText="Link" Target="_blank"  >
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:HyperLinkField>

            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="Chartreuse" Font-Bold="False" ForeColor="Black" />
            <AlternatingRowStyle BackColor="#FFF3DB" />
        </asp:GridView>   
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  Order Final                                                                   ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:GridView ID="GridView3" runat="server" AutoGenerateColumns="False" Style="z-index: 103; left: 210px; position: absolute; top: 200px" CellPadding="3" Font-Size="10pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="Red" BorderStyle="Solid" BorderWidth="5px" >
            <Columns>
                    <asp:BoundField DataField="SPCode"  />
                    <asp:BoundField DataField="SPName"  />
                    <asp:BoundField DataField="SPTime" />
                    <asp:BoundField DataField="SPNo"  />
                    
                    <asp:BoundField DataField="F_ORNo" />
                    <asp:BoundField DataField="F_COrderNo" />
                    <asp:BoundField DataField="F_Time" />
                    <asp:BoundField DataField="F_User" />
                    <asp:BoundField DataField="F_ItemInf" />

                    <asp:BoundField DataField="DiffDaysDesc"  >  <ItemStyle HorizontalAlign="right"  />   </asp:BoundField>  

                    <asp:HyperLinkField DataNavigateUrlFields="ChgFinalUrl" DataNavigateUrlFormatString="{0}" 
                    DataTextField="ChgFinal" HeaderText="Link" Target="_blank"  >
                        <HeaderStyle HorizontalAlign="Center" />
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:HyperLinkField>

            </Columns>
            <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="Chartreuse" Font-Bold="False" ForeColor="Black" />
            <AlternatingRowStyle BackColor="#FFF3DB" />
        </asp:GridView>        
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  隱藏欄位                                                                       ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- ------------------------------------------------------------------------------------ -->
        <!-- Control  -->
        <asp:TextBox ID="DLogID" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DCode" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DCustomerGr" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DFunList" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DUserID" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <!-- ------------------------------------------------------------------------------------ -->
        <!-- FileData  -->
        <asp:TextBox ID="DFileName" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>

        <asp:TextBox ID="DPathImport" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DPathActionPlan" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DPathActionPlanImport" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DPathActionPlanYOC" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DPathKPIFSheet" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DPathWINGS" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DPathPIL" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DPathChangeFinal" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DPathPILFinal" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <!-- ------------------------------------------------------------------------------------ -->
    </div>
    </form>
</body>
</html>
    