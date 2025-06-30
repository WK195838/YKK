<%@ Page Language="VB" AutoEventWireup="false" CodeFile="SampleSheet_01.aspx.vb" Inherits="SampleSheet_01" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>未命名頁面</title>

		<script language="javascript" type="text/javascript">
            function ConfirmMe(btn) {
                if(Page_ClientValidate())   {
                    btn.disabled="disabled";
				    var answer = confirm("是否執行作業嗎？ 請確認....");
				    if (answer) {
                        document.forms[0].__EVENTTARGET.value = btn.name;
                        document.forms[0].__EVENTARGUMENT.value = '';
                        document.forms[0].submit();
				    }                    
                    else    {
                        btn.disabled="";
                        return false;   
                    }				    
                }
                else    {
                    return false;
                }
            }
		</script>

</head>
<body>
    <form id="form1" runat="server">
    <div>
         <asp:Image ID="Image4" runat="server" ImageUrl="~/Images/DevelopmentSample_A1.png" Style="z-index: 99;
                    left: 5px; position: absolute; top: 2px" />
<!-- No -->
        <asp:TextBox ID="DNO" runat="server" BackColor="White" BorderStyle="None" ForeColor="Black"
            Height="20px"   Style="z-index: 126; left: 16px; position: absolute;
            top: 59px; text-align: left" Width="159px" MaxLength="20"  ></asp:TextBox>
<!--  -->            
<!-- 原No -->
        <asp:TextBox ID="DONO" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="20px"   Style="z-index: 126; left: -200px; position: absolute;
            top: 80px; text-align: left" Width="120px" MaxLength="20"  ></asp:TextBox>
<!--  -->            
<!-- 客戶 -->
		 <asp:textbox ID="DAPPBUYER" style="Z-INDEX: 104; POSITION: absolute; TOP: 87px; LEFT: 141px"
		    runat="server" Width="390px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="False">
		 </asp:textbox>
		<asp:button id="BDataPump" style="Z-INDEX: 170; POSITION: absolute; TOP: 93px; LEFT: 537px" runat="server"
			Width="20px" Height="20px" CausesValidation="False" Text=".....">
		</asp:button>
<!--  -->            
<!-- 發行日 -->
		 <asp:textbox ID="DDATE" style="Z-INDEX: 164; POSITION: absolute; TOP: 87px; LEFT: 666px" runat="server"
			Width="124px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="False">
		 </asp:textbox>
<!--  -->            
<!-- Size -->
		<asp:textbox ID="DSIZENO" style="Z-INDEX: 118; POSITION: absolute; TOP: 140px; LEFT: 140px" runat="server"
		    Width="201px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="False">
		</asp:textbox>
<!--  -->            
<!-- ITEM -->
		<asp:textbox ID="DITEM" style="Z-INDEX: 119; POSITION: absolute; TOP: 140px; LEFT: 350px" runat="server"
			Width="201px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="False">
        </asp:textbox>
<!--  -->            
<!-- TAPE -->
	    <asp:textbox ID="DCODENO" style="Z-INDEX: 120; POSITION: absolute; TOP: 140px; LEFT: 560px" runat="server"
			Width="230px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="False">
		</asp:textbox>
<!--  -->            
<!-- 樣品圖-表 -->
        <asp:FileUpload ID="DSAMPLEFILE1" runat="server" BackColor="LightGray" Height="20px"
            Style="z-index: 121; left: 89px; position: absolute; top: 222px; background-color: #c0ffff"
            Width="300px" />
		<asp:image id="LSAMPLEFILE1" style="Z-INDEX: 121; POSITION: absolute; TOP: 243px; LEFT: 16px"
			runat="server" Width="370px" Height="215px" BorderStyle="Groove" ImageUrl="~/Images/SampleSheet_01.jpg">
		</asp:image>
<!--  -->            
<!-- 樣品圖-裏 -->
        <asp:FileUpload ID="DSAMPLEFILE2" runat="server" BackColor="LightGray" Height="20px"
            Style="z-index: 121; left: 392px; position: absolute; top: 222px; background-color: #c0ffff"
            Width="403px" />
        <asp:Image ID="LSAMPLEFILE2" runat="server" BorderStyle="Groove" Height="215px" ImageUrl="~/Images/SampleSheet_01.jpg"
            Style="z-index: 121; position: absolute; top: 243px; left: 397px" Width="370px" />
<!--  -->            
<!-- 布帶寬度 -->
        <asp:TextBox ID="DTAWIDTH" runat="server" BackColor="LightGray" BorderStyle="Groove"
            ForeColor="Blue" Height="20px" ReadOnly="False" Style="z-index: 123; left: 141px;
            position: absolute; top: 464px" Width="82px">
        </asp:TextBox>
<!--  -->            
<!-- 開發NO -->
		<asp:textbox ID="DDEVNO" style="Z-INDEX: 124; POSITION: absolute; TOP: 464px; LEFT: 372px" runat="server"
			Width="116px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="False">
		</asp:textbox>
<!--  -->            
<!-- 開發期間 -->
		<asp:textbox ID="DDEVPRD" style="Z-INDEX: 125; POSITION: absolute; TOP: 464px; LEFT: 603px" runat="server"
			Width="186px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="False">
		</asp:textbox>
<!--  -->            
<!-- 布帶 -->
		<asp:textbox ID="DTACOL" style="Z-INDEX: 126; POSITION: absolute; TOP: 493px; LEFT: 141px" runat="server"
			Width="648px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" MaxLength="240" TextMode="MultiLine" ReadOnly="False">
		</asp:textbox>
<!--  -->            
<!-- 條紋線 -->
		<asp:textbox ID="DTALINE" style="Z-INDEX: 103; POSITION: absolute; TOP: 523px; LEFT: 141px" runat="server"
			Width="648px" Height="31px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" MaxLength="240" TextMode="MultiLine" ReadOnly="False">
		</asp:textbox>
<!--  -->            
<!-- 務齒 -->
		<asp:textbox ID="DECOL" style="Z-INDEX: 127; POSITION: absolute; TOP: 567px; LEFT: 141px" runat="server"
			Width="284px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="False">
		</asp:textbox>
<!--  -->            
<!-- 丸紐 -->
		<asp:textbox ID="DCCOL" style="Z-INDEX: 128; POSITION: absolute; TOP: 567px; LEFT: 518px" runat="server"
			Width="271px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="False">
		</asp:textbox>
<!--  -->            
<!-- 縫工上線 -->
		<asp:textbox ID="DTHUPCOL" style="Z-INDEX: 129; POSITION: absolute; TOP: 601px; LEFT: 176px" runat="server"
    		Width="250px" Height="22px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" MaxLength="240" TextMode="MultiLine" ReadOnly="False">
        </asp:textbox>
<!--  -->            
<!-- 縫工下線 -->
        <asp:TextBox ID="DTHLOCOL" runat="server" BackColor="LightGray" BorderStyle="Groove"
            ForeColor="Blue" Height="22px" MaxLength="240" ReadOnly="False" Style="z-index: 129;
            left: 470px; position: absolute; top: 601px" TextMode="MultiLine" Width="319px">
        </asp:TextBox>
<!--  -->            
<!-- 生産注意事項 -->
		<asp:textbox ID="DMOP1" style="Z-INDEX: 123; POSITION: absolute; TOP: 635px; LEFT: 140px"
			runat="server" Width="160px" Height="20px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="False" MaxLength="20"></asp:textbox>
        <asp:TextBox ID="DMOP2" runat="server" BackColor="LightGray" BorderStyle="Groove"
            ForeColor="Blue" Height="20px" ReadOnly="False" Style="z-index: 123; left: 308px;
            position: absolute; top: 635px" Width="160px" MaxLength="20"></asp:TextBox>
        <asp:TextBox ID="DMOP3" runat="server" BackColor="LightGray" BorderStyle="Groove"
            ForeColor="Blue" Height="20px" ReadOnly="False" Style="z-index: 123; left: 476px;
            position: absolute; top: 635px" Width="160px" MaxLength="20"></asp:TextBox>
        <asp:TextBox ID="DMOP4" runat="server" BackColor="LightGray" BorderStyle="Groove"
            ForeColor="Blue" Height="20px" ReadOnly="False" Style="z-index: 123; left: 644px;
            position: absolute; top: 635px" Width="146px" MaxLength="20"></asp:TextBox>

		<asp:textbox ID="DMNote1" style="Z-INDEX: 130; POSITION: absolute; TOP: 663px; LEFT: 140px" runat="server"
			Width="160px" Height="195px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" MaxLength="240"
			TextMode="MultiLine" ReadOnly="False"></asp:textbox>
        <asp:TextBox ID="DMNote2" runat="server" BackColor="LightGray" BorderStyle="Groove"
            ForeColor="Blue" Height="195px" MaxLength="240" ReadOnly="False" Style="z-index: 130;
            left: 308px; position: absolute; top: 663px" TextMode="MultiLine" Width="160px"></asp:TextBox>
        <asp:TextBox ID="DMNote3" runat="server" BackColor="LightGray" BorderStyle="Groove"
            ForeColor="Blue" Height="195px" MaxLength="240" ReadOnly="False" Style="z-index: 130;
            left: 476px; position: absolute; top: 663px" TextMode="MultiLine" Width="160px"></asp:TextBox>
        <asp:TextBox ID="DMNote4" runat="server" BackColor="LightGray" BorderStyle="Groove"
            ForeColor="Blue" Height="195px" MaxLength="240" ReadOnly="False" Style="z-index: 130;
            left: 644px; position: absolute; top: 663px" TextMode="MultiLine" Width="146px"></asp:TextBox>
<!--  -->
<!-- WAVE'S ITEM CODE -->            
<!--    TAPE NAT-左     TAPE NAT-右																		
		TAPE SET-左		TAPE SET-右																		
		TAPE DYED-左	TAPE DYED-右																		
		CHAIN NAT		CHAIN SET																		
		CHAIN DYED																													
		其他1																													
		其他2																													
		CORD																													
-->
		<asp:textbox ID="DTNLITEM" style="Z-INDEX: 136; POSITION: absolute; TOP: 872px; LEFT: 289px"
			runat="server" Width="92px" Height="16px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="False">
		</asp:textbox>
    	<asp:textbox ID="DTNRITEM" style="Z-INDEX: 140; POSITION: absolute; TOP: 872px; LEFT: 519px"
			runat="server" Width="94px" Height="16px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="False">
		</asp:textbox>
		<asp:textbox ID="DTSLITEM" style="Z-INDEX: 137; POSITION: absolute; TOP: 896px; LEFT: 289px"
			runat="server" Width="92px" Height="16px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="False">
		</asp:textbox>
		<asp:textbox ID="DTSRITEM" style="Z-INDEX: 141; POSITION: absolute; TOP: 896px; LEFT: 519px"
			runat="server" Width="94px" Height="16px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="False">
	    </asp:textbox>
		<asp:textbox ID="DTDLITEM" style="Z-INDEX: 138; POSITION: absolute; TOP: 920px; LEFT: 289px"
			runat="server" Width="92px" Height="16px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="False">
		</asp:textbox>
		<asp:textbox ID="DTDRITEM" style="Z-INDEX: 142; POSITION: absolute; TOP: 920px; LEFT: 519px"
			runat="server" Width="94px" Height="16px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="False"> 
		</asp:textbox>
		<asp:textbox ID="DCNITEM" style="Z-INDEX: 139; POSITION: absolute; TOP: 944px; LEFT: 289px" runat="server"
			Width="92px" Height="16px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="False">
		</asp:textbox>
		<asp:textbox ID="DCSITEM" style="Z-INDEX: 143; POSITION: absolute; TOP: 944px; LEFT: 519px" runat="server"
			Width="94px" Height="16px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="False">
		</asp:textbox>
		<asp:textbox ID="DCDITEM" style="Z-INDEX: 144; POSITION: absolute; TOP: 968px; LEFT: 289px" runat="server"
			Width="324px" Height="16px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="False">
		</asp:textbox>
		
		<asp:textbox ID="DODESCR1" style="Z-INDEX: 144; POSITION: absolute; TOP: 992px; LEFT: 160px" runat="server"
			Width="120px" Height="18px" ForeColor="Black" BorderStyle="None" BackColor="White" ReadOnly="False"></asp:textbox>
		
		<asp:textbox ID="DO1ITEM" style="Z-INDEX: 146; POSITION: absolute; TOP: 992px; LEFT: 289px" runat="server"
			Width="324px" Height="16px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="False">
		</asp:textbox>

		<asp:textbox ID="DODESCR2" style="Z-INDEX: 144; POSITION: absolute; TOP: 1016px; LEFT: 160px" runat="server"
			Width="120px" Height="18px" ForeColor="Black" BorderStyle="None" BackColor="White" ReadOnly="False"></asp:textbox>

		<asp:textbox ID="DO2ITEM" style="Z-INDEX: 145; POSITION: absolute; TOP: 1016px; LEFT: 289px" runat="server"
			Width="324px" Height="16px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="False"> 
		</asp:textbox>
		<asp:textbox ID="DCITEM" style="Z-INDEX: 147; POSITION: absolute; TOP: 1040px; LEFT: 289px" runat="server"
			Width="324px" Height="16px" ForeColor="Blue" BorderStyle="Groove" BackColor="LightGray" ReadOnly="False">
		</asp:textbox>
<!--  -->
<!-- 工程 -->            
        <asp:TextBox ID="DOP1" runat="server" BackColor="LightGray" BorderStyle="Groove"
            ForeColor="Blue" Height="16px" ReadOnly="False" Style="z-index: 136; left: 161px;
            position: absolute; top: 1077px" Width="78px"></asp:TextBox>
        <asp:TextBox ID="DOP2" runat="server" BackColor="LightGray" BorderStyle="Groove"
            ForeColor="Blue" Height="16px" ReadOnly="False" Style="z-index: 136; left: 266px;
            position: absolute; top: 1077px" Width="78px"></asp:TextBox>
        <asp:TextBox ID="DOP3" runat="server" BackColor="LightGray" BorderStyle="Groove"
            ForeColor="Blue" Height="16px" ReadOnly="False" Style="z-index: 136; left: 370px;
            position: absolute; top: 1077px" Width="78px"></asp:TextBox>
        <asp:TextBox ID="DOP4" runat="server" BackColor="LightGray" BorderStyle="Groove"
            ForeColor="Blue" Height="16px" ReadOnly="False" Style="z-index: 136; left: 476px;
            position: absolute; top: 1077px" Width="78px"></asp:TextBox>
        <asp:TextBox ID="DOP5" runat="server" BackColor="LightGray" BorderStyle="Groove"
            ForeColor="Blue" Height="16px" ReadOnly="False" Style="z-index: 136; left: 582px;
            position: absolute; top: 1077px" Width="78px"></asp:TextBox>
        <asp:TextBox ID="DOP6" runat="server" BackColor="LightGray" BorderStyle="Groove"
            ForeColor="Blue" Height="16px" ReadOnly="False" Style="z-index: 136; left: 686px;
            position: absolute; top: 1077px" Width="78px"></asp:TextBox>
<!--  -->
        <!-- +++++++++++++++++++++++++++++++++++++++++++++
             ++ 處理說明及Button Start                  ++ 
             +++++++++++++++++++++++++++++++++++++++++++++ -->
        <img id="DDescSheet" runat="server" src="images/Sheet_Remark.jpg" style="z-index: 1;
            left: 10px; position: absolute; top: 1133px" />
        <asp:TextBox ID="DDecideDesc" runat="server" BackColor="Yellow" ForeColor="Blue"
            Height="56px" Style="z-index: 132; left: 56px; position: absolute; top: 1138px"
            TextMode="MultiLine" Width="536px" Visible="False"></asp:TextBox>
        <asp:Button ID="BSAVE" runat="server" Height="23px" Style="z-index: 128; left: 416px;
            position: absolute; top: 1212px" Text="儲存" Width="80px" OnClientClick="return ConfirmMe(this)" UseSubmitBehavior="false"/>
        <asp:Button ID="BNG2" runat="server" Height="23px" Style="z-index: 129; left: 507px;
            position: absolute; top: 1212px" Text="NG2" Width="80px" OnClientClick="return ConfirmMe(this)" UseSubmitBehavior="false"/>
        <asp:Button ID="BNG1" runat="server" Height="23px" Style="z-index: 130; left: 599px;
            position: absolute; top: 1212px" Text="NG1" Width="80px" OnClientClick="return ConfirmMe(this)" UseSubmitBehavior="false"/>
        <asp:Button ID="BOK" runat="server" Height="23px" Style="z-index: 131; left: 691px;
            position: absolute; top: 1212px" Text="OK" Width="80px" OnClientClick="return ConfirmMe(this)" UseSubmitBehavior="false"/>
        <!-- 處理說明及Button End -->
            <!-- +++++++++++++++++++++++++++++++++++++++++++++
                 ++ 核定履歷(View5) Start                     ++ 
                 +++++++++++++++++++++++++++++++++++++++++++++ -->

            <asp:GridView style="Z-INDEX: 136; LEFT: 10px; POSITION: absolute; top: 1257px" id="GridView1" runat="server" Width="950px" Height="100px" BorderStyle="Groove" BackColor="White" BorderColor="#CC9966" CellPadding="4" BorderWidth="1px" AutoGenerateColumns="False">

            <RowStyle BackColor="White" Font-Size="9pt" ForeColor="#330099"  />
            <Columns>
                <asp:BoundField DataField="StepNameDesc" HeaderText="工程">
                    <ItemStyle HorizontalAlign="Left"  />
                </asp:BoundField>
                <asp:BoundField DataField="DecideName" HeaderText="擔當" ></asp:BoundField>
                <asp:BoundField DataField="AgentName" HeaderText="代理/兼職" ></asp:BoundField>
                <asp:BoundField DataField="FlowTypeDesc" HeaderText="類別" ></asp:BoundField>
                <asp:BoundField DataField="DelaySts" HeaderText="處理進度" ></asp:BoundField>
                <asp:BoundField DataField="StsDesc" HeaderText="處理結果" ></asp:BoundField>
                <asp:BoundField DataField="DecideDescA" HeaderText="說明">
                    <ItemStyle HorizontalAlign="Left"  />
                </asp:BoundField>
                <asp:BoundField DataField="Description" HeaderText="核定時間">
                    <ItemStyle HorizontalAlign="Left"  />
                </asp:BoundField>
            </Columns>
            <HeaderStyle BackColor="#CC6600" Font-Size="9pt" ForeColor="#FFFFCC" HorizontalAlign="Center"
                VerticalAlign="Middle"  />
        </asp:GridView>                

        <!--************************************************************ 
            ** 防止重覆按 [BUTTON 2 次]                               **     
            ************************************************************ -->
            <asp:TextBox ID="TextBox1" runat="server" Style="z-index: 126; left: -500px;position: absolute; top: 100px; text-align: left">AAA</asp:TextBox>
            <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="TextBox1" ErrorMessage="不可為空白"></asp:RequiredFieldValidator>            

        
    </div>
    </form>
</body>
</html>
