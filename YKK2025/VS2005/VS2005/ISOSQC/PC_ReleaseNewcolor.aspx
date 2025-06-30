<%@ Page Language="VB" AutoEventWireup="false" CodeFile="PC_ReleaseNewcolor.aspx.vb" Inherits="PC_ReleaseNewcolor" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>ReleaseNewcolor</title>
	    <script language="javascript" src="ForProject.js"></script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
		<!-- ****************************************************************************************** -->
		<!-- ** System
		<!-- ****************************************************************************************** -->
        <img src="iMages/ISIPLOGO.png" style="z-index: 1; left: 6px; position: absolute;top: 6px; width: 500px; height: 32px;"/>

		<!-- ****************************************************************************************** -->
		<!-- ** Button                                                                                  -->
		<!-- ****************************************************************************************** -->
		    <!-- Find -->
                <asp:Button ID="BFind" runat="server" Height="25px" Style="z-index: 104; left: 704px;
                    position: absolute; top: 48px" Text="Go" Width="45px" />
		    <!-- Register -->
                <asp:Button ID="BRegister" runat="server" Height="44px" Style="z-index: 104; left: 1256px;
                    position: absolute; top: 80px" Text="Register" Width="64px" onclientclick="javascript:return confirm('確定執行？');" />
		    <!-- Advanced Image -->
                <asp:Button ID="BADVImage" runat="server" Height="32px" Style="z-index: 104; left: 920px;
                    position: absolute; top: 40px" Text="Advenced Images" Width="114px" onclientclick="javascript:return confirm('確定執行？');" />
		    <!-- Advanced Report -->
    	        <asp:button id="BADVReport" runat="server" Height="32px" Width="114px" Style="z-index: 103; left: 920px; position: absolute; top: 4px" Text="Advenced Report"></asp:button>
		    <!-- Advanced Sales -->
    	        <asp:button id="BADVSales" runat="server" Height="32px" Width="114px" Style="z-index: 103; left: 1040px; position: absolute; top: 4px" Text="Sales Report"></asp:button>
			<!-- Puller List -->
    	        <asp:button id="BPullerList" runat="server" Height="32px" Width="104px" Style="z-index: 103; left: 1160px; position: absolute; top: 4px" Text="Puller List"></asp:button>
		    <!-- Admin -->
    	        <asp:button id="BAdmin" runat="server" Height="32px" Width="104px" Style="z-index: 103; left: 1270px; position: absolute; top: 4px" Text="Admin. Report"></asp:button>
		<!-- ****************************************************************************************** -->
		<!-- ** Puller Key                                                                              -->
		<!-- ****************************************************************************************** -->
            <asp:TextBox ID="TextBox1" runat="server" BackColor="Black" BorderStyle="None" ForeColor="White"
                Height="18px" ReadOnly="True" Style="z-index: 126; left: 8px; position: absolute;
                top: 48px; text-align: left" Width="80px">Puller Code</asp:TextBox>
            <asp:TextBox ID="DKPullerCode" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
                Height="18px" Style="z-index: 126; left: 88px; position: absolute;
                top: 48px; text-align: left" Width="100px" MaxLength="7"></asp:TextBox>

            <asp:TextBox ID="TextBox11" runat="server" BackColor="Black" BorderStyle="None" ForeColor="White"
                Height="18px" ReadOnly="True" Style="z-index: 126; left:192px; position: absolute;
                top: 48px; text-align: left" Width="80px">Color</asp:TextBox>
            <asp:TextBox ID="DKColor" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
                Height="18px" Style="z-index: 126; left: 272px; position: absolute;
                top: 48px; text-align: left" Width="48px" MaxLength="7"></asp:TextBox>

            <asp:TextBox ID="TextBox8" runat="server" BackColor="Black" BorderStyle="None" ForeColor="White"
                Height="18px" ReadOnly="True" Style="z-index: 126; left:320px; position: absolute;
                top: 48px; text-align: left" Width="80px">Buyer</asp:TextBox>
            <asp:TextBox ID="DKBuyer" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
                Height="18px" Style="z-index: 126; left: 400px; position: absolute;
                top: 48px; text-align: left" Width="100px" MaxLength="7"></asp:TextBox>

            <asp:TextBox ID="TextBox12" runat="server" BackColor="Black" BorderStyle="None" ForeColor="White"
                Height="18px" ReadOnly="True" Style="z-index: 126; left:504px; position: absolute;
                top: 48px; text-align: left" Width="80px">Search</asp:TextBox>
            <asp:TextBox ID="DKOther" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
                Height="18px" Style="z-index: 126; left: 584px; position: absolute;
                top: 48px; text-align: left" Width="100px" MaxLength="50"></asp:TextBox>

            <asp:TextBox ID="DKSource" runat="server" BackColor="white" BorderStyle="Groove" ForeColor="Blue"
                Height="18px" Style="z-index: 126; left: 760px; position: absolute;
                top: 48px; text-align: left" Width="48px" MaxLength="7"></asp:TextBox>

            <asp:TextBox ID="DKPuller" runat="server" BackColor="white" BorderStyle="Groove" ForeColor="Blue"
                Height="18px" Style="z-index: 126; left: 816px; position: absolute;
                top: 48px; text-align: left" Width="72px" MaxLength="7"></asp:TextBox>


            <asp:TextBox ID="DOther" runat="server" BackColor="white" BorderStyle="None" ForeColor="Red"
                Height="18px" Style="z-index: 126; left: 816px; position: absolute;
                top: 128px; text-align: left" Width="312px" MaxLength="7">OK*代表多個外注廠開發，請留意各型別打色狀況</asp:TextBox>

		<!-- ****************************************************************************************** -->
		<!-- ** CheckBox                                                                              -->
		<!-- ****************************************************************************************** -->
        <asp:CheckBox ID="AtBColor" runat="server" style="z-index: 174; left: 1144px; position: absolute; top: 80px" Font-Size="9pt" Text="Transparent (B)" Width="112px" AutoPostBack="True" Font-Bold="true" />
        <asp:CheckBox ID="AtCColor" runat="server" style="z-index: 174; left: 1144px; position: absolute; top: 104px" Font-Size="9pt" Text="Translucent (C)" Width="112px" AutoPostBack="True" Font-Bold="true" />
        <asp:CheckBox ID="AtCloseRDW" runat="server" style="z-index: 174; left: 1120px; position: absolute; top: 128px" Font-Size="9pt" Text="Close RD Windows" Width="130px" AutoPostBack="True" Font-Bold="true" />
        <asp:CheckBox ID="AtCloseIMGW" runat="server" style="z-index: 174; left: 600px; position: absolute; top: 128px" Font-Size="9pt" Text="Close IMG Windows" Width="130px" AutoPostBack="True" Font-Bold="true" />
        <asp:CheckBox ID="AtCloseEDXW" runat="server" style="z-index: 174; left: 600px; position: absolute; top: 128px" Font-Size="9pt" Text="Close EDX Windows" Width="130px" AutoPostBack="True" Font-Bold="true" />

        <asp:CheckBox ID="AtSPColor" runat="server" style="z-index: 174; left: 584px; position: absolute; top: 24px" Font-Size="9pt" Text="Color=空白" Width="104px" Font-Bold="False" />
		<!-- ****************************************************************************************** -->
		<!-- ** Automate                                                                         -->
		<!-- ****************************************************************************************** -->
        <asp:TextBox style="FONT-SIZE: 12px; Z-INDEX: 318; BACKGROUND: none transparent scroll repeat 0% 0%; LEFT: 1176px; POSITION: absolute; TOP: 40px" id="DO365Comment" runat="server" Height="32px" Width="176px" BorderWidth="0px" BorderStyle="None" TextMode="MultiLine">無法開啟時請複製以下連結至其它BROWSER使用</asp:TextBox>
        <asp:TextBox style="FONT-SIZE: 12px; Z-INDEX: 318; BACKGROUND: none transparent scroll repeat 0% 0%; LEFT: 1360px; POSITION: absolute; TOP: 40px" id="DO365Comment2" runat="server" Height="24px" Width="300px" BorderWidth="0px" BorderStyle="None">https://forms.office.com/r/4uvpdHxgaH</asp:TextBox>
        <asp:Button ID="BO365" runat="server" Height="32px" Style="z-index: 104; left: 1040px;
            position: absolute; top: 40px" Text="Puller Maint." Width="112px" onclientclick="javascript:return confirm('確定執行？');" />
		
		<!-- ****************************************************************************************** -->
		<!-- ** Register Color                                                                          -->
		<!-- ****************************************************************************************** -->
		    <!-- Release Color -->
                <asp:TextBox ID="TextBox2" runat="server" BackColor="Black" BorderStyle="None" ForeColor="White"
                    Height="44px" ReadOnly="True" Style="z-index: 126; left: 8px; position: absolute;
                    top: 80px; text-align: left" Width="80px">Register</asp:TextBox>
		        <!-- Puller -->
                    <asp:TextBox ID="TextBox3" runat="server" BackColor="Black" BorderStyle="Groove" ForeColor="White"
                        Height="18px" ReadOnly="True" Style="z-index: 126; left: 88px; position: absolute;
                        top: 80px; text-align: left" Width="80px">Puller Code</asp:TextBox>
                    <asp:TextBox ID="DPullerCode" runat="server" BackColor="#C0FFC0" BorderStyle="Groove" ForeColor="Red"
                        Height="18px" Style="z-index: 126; left: 88px; position: absolute;
                        top: 104px; text-align: left" Width="80px" MaxLength="10" ReadOnly="True"></asp:TextBox>
	  		    <!-- Color -->
                <asp:TextBox ID="TextBox4" runat="server" BackColor="Black" BorderStyle="Groove" ForeColor="White"
                    Height="18px" ReadOnly="True" Style="z-index: 126; left: 176px; position: absolute;
                    top: 80px; text-align: left" Width="80px">Color Code</asp:TextBox>
                <asp:TextBox ID="DColor" runat="server" BackColor="#C0FFC0" BorderStyle="Groove" ForeColor="Red"
                    Height="18px" Style="z-index: 126; left: 176px; position: absolute;
                    top: 104px; text-align: left" Width="80px" MaxLength="5" ReadOnly="TRUE"></asp:TextBox>
		        <!-- ColorName -->
                <asp:TextBox ID="textbox5" runat="server" BackColor="Black" BorderStyle="Groove" ForeColor="White"
                    Height="18px" ReadOnly="True" Style="z-index: 126; left: 264px; position: absolute;
                    top: 80px; text-align: left" Width="160px">Buyer Color Name</asp:TextBox>
                <asp:TextBox ID="DColorName" runat="server" BackColor="Lime" BorderStyle="Groove" ForeColor="Red"
                    Height="18px" Style="z-index: 126; left: 264px; position: absolute;
                    top: 104px; text-align: left" Width="160px" MaxLength="100"></asp:TextBox>
		        <!-- Buyer Color Code -->
                <asp:TextBox ID="textbox6" runat="server" BackColor="Black" BorderStyle="Groove" ForeColor="White"
                    Height="18px" ReadOnly="True" Style="z-index: 126; left: 432px; position: absolute;
                    top: 80px; text-align: left" Width="110px">Buyer Color Code</asp:TextBox>
                <asp:TextBox ID="DBYColorCode" runat="server" BackColor="Lime" BorderStyle="Groove" ForeColor="Red"
                    Height="18px" Style="z-index: 126; left: 432px; position: absolute;
                    top: 104px; text-align: left" Width="110px" MaxLength="50"></asp:TextBox>
		        <!-- Buyer -->
                <asp:TextBox ID="textbox7" runat="server" BackColor="Black" BorderStyle="Groove" ForeColor="White"
                    Height="18px" ReadOnly="True" Style="z-index: 126; left: 552px; position: absolute;
                    top: 80px; text-align: left" Width="80px">Buyer</asp:TextBox>
                <asp:TextBox ID="DBuyer" runat="server" BackColor="Lime" BorderStyle="Groove" ForeColor="Red"
                    Height="18px" Style="z-index: 126; left: 552px; position: absolute;
                    top: 104px; text-align: left" Width="80px" MaxLength="6" ReadOnly="False"></asp:TextBox>
                    
		        <!-- BuyerName -->
                <asp:TextBox ID="textbox9" runat="server" BackColor="Black" BorderStyle="Groove" ForeColor="White"
                    Height="18px" ReadOnly="True" Style="z-index: 126; left: 640px; position: absolute;
                    top: 80px; text-align: left" Width="80px">Buyer Name</asp:TextBox>
                <asp:TextBox ID="DBuyerName" runat="server" BackColor="Lime" BorderStyle="Groove" ForeColor="Red"
                    Height="18px" Style="z-index: 126; left: 640px; position: absolute;
                    top: 104px; text-align: left" Width="80px" MaxLength="20" ReadOnly="False"></asp:TextBox>
                    
		        <!-- TapeColor -->
                <asp:TextBox ID="textbox13" runat="server" BackColor="Black" BorderStyle="Groove" ForeColor="White"
                    Height="18px" ReadOnly="True" Style="z-index: 126; left: 728px; position: absolute;
                    top: 80px; text-align: left" Width="80px">Tape Color</asp:TextBox>
                <asp:TextBox ID="DTapeColor" runat="server" BackColor="Lime" BorderStyle="Groove" ForeColor="Red"
                    Height="18px" Style="z-index: 126; left: 728px; position: absolute;
                    top: 104px; text-align: left" Width="80px" MaxLength="20" ReadOnly="False"></asp:TextBox>

		        <!-- Remark -->
                <asp:TextBox ID="textbox10" runat="server" BackColor="Black" BorderStyle="Groove" ForeColor="White"
                    Height="18px" ReadOnly="True" Style="z-index: 126; left: 816px; position: absolute;
                    top: 80px; text-align: left" Width="320px">Remark (有限定說明需填入)</asp:TextBox>
                <asp:TextBox ID="DRemark" runat="server" BackColor="Lime" BorderStyle="Groove" ForeColor="Red"
                    Height="18px" Style="z-index: 126; left: 816px; position: absolute;
                    top: 104px; text-align: left" Width="320px" MaxLength="250"></asp:TextBox>
		<!-- ****************************************************************************************** -->
		<!-- ** Puller List (GridView1)         table-layout:fixed;        width="1000px"                                                 -->
		<!-- ****************************************************************************************** -->
            <asp:GridView ID="GridView1" runat="server" WIDTH="2000px" AutoGenerateColumns="False"  Style="z-index: 103; left: 8px; position: absolute; top: 152px" CellPadding="3" Font-Size="10pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" DataKeyNames="PullerKey">
                <Columns>
                    <asp:CommandField ShowSelectButton="True" SelectText="New" />
                    
                    <asp:CommandField ShowEditButton="True" SelectText="..." EditText="R&D" />

                    <asp:BoundField DataField="Puller"  />
                    <asp:BoundField DataField="Color"  />
                    <asp:BoundField DataField="ColorName"  />
                    <asp:BoundField DataField="BYColorCode" />
                    <asp:BoundField DataField="Buyer"  />
                    <asp:BoundField DataField="BuyerName"  />

                    <asp:BoundField DataField="TapeColor"  />                
                    <asp:BoundField DataField="Remark"  />
                    <asp:BoundField DataField="DTM_YOBI1"  />
                    
                    <asp:HyperLinkField DataNavigateUrlFields="RDMUrl" DataNavigateUrlFormatString="{0}" 
                    DataTextField="RD" HeaderText="Link" Target="_blank"  >
                    </asp:HyperLinkField>
                    
                    <asp:BoundField DataField="DTM"  />

                    <asp:HyperLinkField DataNavigateUrlFields="EDXMUrl" DataNavigateUrlFormatString="{0}" 
                    DataTextField="EDX" HeaderText="Link" Target="_blank"  >
                    </asp:HyperLinkField>

                    <asp:BoundField DataField="IRW"  />
                    <asp:BoundField DataField="ORDERS"  />
 
                    <asp:HyperLinkField DataNavigateUrlFields="RDUrl" DataNavigateUrlFormatString="{0}" 
                    DataTextField="RD_No" HeaderText="Link" Target="_blank"  >
                    </asp:HyperLinkField>

                    <asp:BoundField DataField="RD_AppDate"  />                
                    <asp:BoundField DataField="RD_Supplier"  />                

                    <asp:CommandField ShowDeleteButton="True" SelectText="..." DeleteText="IMG" />

                    <asp:HyperLinkField DataNavigateUrlFields="EDXUrl" DataNavigateUrlFormatString="{0}" 
                    DataTextField="EDX_No" HeaderText="Link" Target="_blank"  >
                    </asp:HyperLinkField>
                    <asp:BoundField DataField="EDX_AppDate"  />                
                    <asp:BoundField DataField="EDX_Supplier"  />    

                    <asp:ButtonField CommandName="EDXIMG" Text="IMG" />
                        
                    <asp:HyperLinkField DataNavigateUrlFields="IRWUrl" DataNavigateUrlFormatString="{0}" 
                    DataTextField="IRW_No" HeaderText="Link" Target="_blank"  >
                    </asp:HyperLinkField>
                    <asp:BoundField DataField="IRW_AppDate"  />                

                    <asp:HyperLinkField DataNavigateUrlFields="ORUrl" DataNavigateUrlFormatString="{0}" 
                    DataTextField="OR_No" HeaderText="Link" Target="_blank"  >
                    </asp:HyperLinkField>
                    <asp:BoundField DataField="OR_CfmDate"  />                

                    <asp:HyperLinkField DataNavigateUrlFields="ORIMGUrl" DataNavigateUrlFormatString="{0}" 
                    DataTextField="OR_IMG" HeaderText="Link" Target="_blank"  >
                    </asp:HyperLinkField>

                </Columns>
                <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
                <FooterStyle BackColor="#CCCCCC" />
                <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
                <SelectedRowStyle BackColor="Chartreuse" Font-Bold="False" ForeColor="Black" />
                <AlternatingRowStyle BackColor="#CCCCCC" />
            </asp:GridView>
		<!-- ****************************************************************************************** -->
		<!-- ** SPD List                                                                                -->
		<!-- ****************************************************************************************** -->
        <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" Style="z-index: 103; left: 900px; position: absolute; top: 336px" CellPadding="3" Font-Size="10pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="Red" BorderStyle="Solid" BorderWidth="5px" >
            <Columns>
                <asp:BoundField DataField="MMSDESC" HeaderText=""  />          
                <asp:BoundField DataField="Puller" HeaderText=""  />          
                <asp:BoundField DataField="Status" HeaderText=""  />          
                
                <asp:HyperLinkField DataNavigateUrlFields="RDNOUrl" DataNavigateUrlFormatString="{0}" 
                DataTextField="RDNOM" HeaderText="RDNo" Target="_blank"  >
                </asp:HyperLinkField>


                <asp:BoundField DataField="SliderCode" HeaderText=""  /> 
                <asp:BoundField DataField="Spec" HeaderText=""  />          
                <asp:BoundField DataField="SUPPLIER" HeaderText=""  />          

                <asp:HyperLinkField DataNavigateUrlFields="OPContactL" DataNavigateUrlFormatString="{0}" 
                DataTextField="OPContactM" HeaderText="型" Target="_blank"  >
                </asp:HyperLinkField>
                
                <asp:HyperLinkField DataNavigateUrlFields="ContactL" DataNavigateUrlFormatString="{0}" 
                DataTextField="ContactM" HeaderText="連" Target="_blank"  >
                </asp:HyperLinkField>

                <asp:HyperLinkField DataNavigateUrlFields="SliderDetailL" DataNavigateUrlFormatString="{0}" 
                DataTextField="SliderDetailM" HeaderText="細" Target="_blank"  >
                </asp:HyperLinkField>

                <asp:HyperLinkField DataNavigateUrlFields="SurfaceL" DataNavigateUrlFormatString="{0}" 
                DataTextField="SurfaceM" HeaderText="表" Target="_blank"  >
                </asp:HyperLinkField>

                <asp:HyperLinkField DataNavigateUrlFields="ColorAppendL" DataNavigateUrlFormatString="{0}" 
                DataTextField="ColorAppendM" HeaderText="色" Target="_blank"  >
                </asp:HyperLinkField>

                <asp:HyperLinkField DataNavigateUrlFields="YKKGroupCopyL" DataNavigateUrlFormatString="{0}" 
                DataTextField="YKKGroupCopyM" HeaderText="Y" Target="_blank"  >
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
        <asp:Image ID="DRDImage" runat="server" ImageUrl="iMages/ISIPLOGO.png" Style="z-index: 103; left: 600px; position: absolute; top: 336px" Width="200" Height="230"  BorderColor="Red" BorderStyle="Solid" BorderWidth="5px" TabIndex="99"/>

		<!-- ****************************************************************************************** -->
		<!-- ** EDX IMAGES
		<!-- ****************************************************************************************** -->
        <asp:Image ID="DEDXImage" runat="server" ImageUrl="iMages/ISIPLOGO.png" Style="z-index: 103; left: 600px; position: absolute; top: 336px" Width="520" Height="630"  BorderColor="Red" BorderStyle="Solid" BorderWidth="5px" TabIndex="99"/>
     
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- ++  隱藏欄位                                                                       ++ -->
        <!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:TextBox ID="DADVREPORTFile" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>

        <asp:TextBox ID="DADVSALESFile" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <br />

        <asp:TextBox ID="DPULLERLISTFile" runat="server" Height="16px" Style="z-index: 318; left: 202px;
			position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
    </div>
    </form>
</body>
</html>
