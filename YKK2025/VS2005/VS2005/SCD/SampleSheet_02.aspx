<%@ Page Language="VB" AutoEventWireup="false" CodeFile="SampleSheet_02.aspx.vb" Inherits="SampleSheet_02" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>未命名頁面</title>
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
		 <asp:textbox ID="DAPPBUYER" style="Z-INDEX: 104; POSITION: absolute; TOP: 89px; LEFT: 141px"
		    runat="server" Width="410px" Height="20px" ForeColor="Black" BorderStyle="None" BackColor="White" ReadOnly="False"></asp:textbox>
        &nbsp;
<!--  -->            
<!-- 發行日 -->
		 <asp:textbox ID="DDATE" style="Z-INDEX: 164; POSITION: absolute; TOP: 89px; LEFT: 666px" runat="server"
			Width="124px" Height="20px" ForeColor="Black" BorderStyle="None" BackColor="White" ReadOnly="False"></asp:textbox>
<!--  -->            
<!-- Size -->
		<asp:textbox ID="DSIZENO" style="Z-INDEX: 118; POSITION: absolute; TOP: 142px; LEFT: 140px" runat="server"
		    Width="201px" Height="20px" ForeColor="Black" BorderStyle="None" BackColor="White" ReadOnly="False"></asp:textbox>
<!--  -->            
<!-- ITEM -->
		<asp:textbox ID="DITEM" style="Z-INDEX: 119; POSITION: absolute; TOP: 142px; LEFT: 350px" runat="server"
			Width="201px" Height="20px" ForeColor="Black" BorderStyle="None" BackColor="White" ReadOnly="False"></asp:textbox>
<!--  -->            
<!-- TAPE -->
	    <asp:textbox ID="DCODENO" style="Z-INDEX: 120; POSITION: absolute; TOP: 142px; LEFT: 560px" runat="server"
			Width="230px" Height="20px" ForeColor="Black" BorderStyle="None" BackColor="White" ReadOnly="False"></asp:textbox>
        &nbsp;<!--  --><!-- 樣品圖-表 -->
		<asp:image id="LSAMPLEFILE1" style="Z-INDEX: 121; POSITION: absolute; TOP: 243px; LEFT: 16px"
			runat="server" Width="370px" Height="215px" BorderStyle="Groove" ImageUrl="~/Images/SampleSheet_01.jpg">
		</asp:image>
        &nbsp;<!--  --><!-- 樣品圖-裏 -->
        <asp:Image ID="LSAMPLEFILE2" runat="server" BorderStyle="Groove" Height="215px" ImageUrl="~/Images/SampleSheet_01.jpg"
            Style="z-index: 121; position: absolute; top: 243px; left: 397px" Width="370px" />
<!--  -->            
<!-- 布帶寬度 -->
        <asp:TextBox ID="DTAWIDTH" runat="server" BackColor="White" BorderStyle="None"
            ForeColor="Black" Height="20px" ReadOnly="False" Style="z-index: 123; left: 141px;
            position: absolute; top: 466px" Width="82px"></asp:TextBox>
<!--  -->            
<!-- 開發NO -->
		<asp:textbox ID="DDEVNO" style="Z-INDEX: 124; POSITION: absolute; TOP: 466px; LEFT: 372px" runat="server"
			Width="116px" Height="20px" ForeColor="Black" BorderStyle="None" BackColor="White" ReadOnly="False"></asp:textbox>
<!--  -->            
<!-- 開發期間 -->
		<asp:textbox ID="DDEVPRD" style="Z-INDEX: 125; POSITION: absolute; TOP: 466px; LEFT: 603px" runat="server"
			Width="186px" Height="20px" ForeColor="Black" BorderStyle="None" BackColor="White" ReadOnly="False"></asp:textbox>
<!--  -->            
<!-- 布帶 -->
		<asp:textbox ID="DTACOL" style="Z-INDEX: 126; POSITION: absolute; TOP: 496px; LEFT: 141px" runat="server"
			Width="648px" Height="20px" ForeColor="Black" BorderStyle="None" BackColor="White" MaxLength="240" TextMode="MultiLine" ReadOnly="False"></asp:textbox>
<!--  -->            
<!-- 條紋線 -->
		<asp:textbox ID="DTALINE" style="Z-INDEX: 103; POSITION: absolute; TOP: 526px; LEFT: 141px" runat="server"
			Width="648px" Height="31px" ForeColor="Black" BorderStyle="None" BackColor="White" MaxLength="240" TextMode="MultiLine" ReadOnly="False"></asp:textbox>
<!--  -->            
<!-- 務齒 -->
		<asp:textbox ID="DECOL" style="Z-INDEX: 127; POSITION: absolute; TOP: 571px; LEFT: 141px" runat="server"
			Width="284px" Height="20px" ForeColor="Black" BorderStyle="None" BackColor="White" ReadOnly="False"></asp:textbox>
<!--  -->            
<!-- 丸紐 -->
		<asp:textbox ID="DCCOL" style="Z-INDEX: 128; POSITION: absolute; TOP: 571px; LEFT: 518px" runat="server"
			Width="271px" Height="20px" ForeColor="Black" BorderStyle="None" BackColor="White" ReadOnly="False"></asp:textbox>
<!--  -->            
<!-- 縫工上線 -->
		<asp:textbox ID="DTHUPCOL" style="Z-INDEX: 129; POSITION: absolute; TOP: 605px; LEFT: 176px" runat="server"
    		Width="250px" Height="22px" ForeColor="Black" BorderStyle="None" BackColor="White" MaxLength="240" TextMode="MultiLine" ReadOnly="False"></asp:textbox>
<!--  -->            
<!-- 縫工下線 -->
        <asp:TextBox ID="DTHLOCOL" runat="server" BackColor="White" BorderStyle="None"
            ForeColor="Black" Height="22px" MaxLength="240" ReadOnly="False" Style="z-index: 129;
            left: 470px; position: absolute; top: 605px" TextMode="MultiLine" Width="319px"></asp:TextBox>
<!--  -->            
<!-- 生産注意事項 -->
		<asp:textbox ID="DMOP1" style="Z-INDEX: 123; POSITION: absolute; TOP: 635px; LEFT: 140px"
			runat="server" Width="160px" Height="20px" ForeColor="Black" BorderStyle="None" BackColor="White" ReadOnly="False" MaxLength="20"></asp:textbox>
        <asp:TextBox ID="DMOP2" runat="server" BackColor="White" BorderStyle="None"
            ForeColor="Black" Height="20px" ReadOnly="False" Style="z-index: 123; left: 308px;
            position: absolute; top: 635px" Width="160px" MaxLength="20"></asp:TextBox>
        <asp:TextBox ID="DMOP3" runat="server" BackColor="White" BorderStyle="None"
            ForeColor="Black" Height="20px" ReadOnly="False" Style="z-index: 123; left: 476px;
            position: absolute; top: 635px" Width="160px" MaxLength="20"></asp:TextBox>
        <asp:TextBox ID="DMOP4" runat="server" BackColor="White" BorderStyle="None"
            ForeColor="Black" Height="20px" ReadOnly="False" Style="z-index: 123; left: 644px;
            position: absolute; top: 635px" Width="146px" MaxLength="20"></asp:TextBox>

		<asp:textbox ID="DMNote1" style="Z-INDEX: 130; POSITION: absolute; TOP: 663px; LEFT: 140px" runat="server"
			Width="160px" Height="195px" ForeColor="Black" BorderStyle="None" BackColor="White" MaxLength="240"
			TextMode="MultiLine" ReadOnly="False"></asp:textbox>
        <asp:TextBox ID="DMNote2" runat="server" BackColor="White" BorderStyle="None"
            ForeColor="Black" Height="195px" MaxLength="240" ReadOnly="False" Style="z-index: 130;
            left: 308px; position: absolute; top: 663px" TextMode="MultiLine" Width="160px"></asp:TextBox>
        <asp:TextBox ID="DMNote3" runat="server" BackColor="White" BorderStyle="None"
            ForeColor="Black" Height="195px" MaxLength="240" ReadOnly="False" Style="z-index: 130;
            left: 476px; position: absolute; top: 663px" TextMode="MultiLine" Width="160px"></asp:TextBox>
        <asp:TextBox ID="DMNote4" runat="server" BackColor="White" BorderStyle="None"
            ForeColor="Black" Height="195px" MaxLength="240" ReadOnly="False" Style="z-index: 130;
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
		<asp:textbox ID="DTNLITEM" style="Z-INDEX: 136; POSITION: absolute; TOP: 873px; LEFT: 289px"
			runat="server" Width="92px" Height="16px" ForeColor="Black" BorderStyle="None" BackColor="White" ReadOnly="False"></asp:textbox>
    	<asp:textbox ID="DTNRITEM" style="Z-INDEX: 140; POSITION: absolute; TOP: 873px; LEFT: 519px"
			runat="server" Width="94px" Height="16px" ForeColor="Black" BorderStyle="None" BackColor="White" ReadOnly="False"></asp:textbox>
		<asp:textbox ID="DTSLITEM" style="Z-INDEX: 137; POSITION: absolute; TOP: 897px; LEFT: 289px"
			runat="server" Width="92px" Height="16px" ForeColor="Black" BorderStyle="None" BackColor="White" ReadOnly="False"></asp:textbox>
		<asp:textbox ID="DTSRITEM" style="Z-INDEX: 141; POSITION: absolute; TOP: 897px; LEFT: 519px"
			runat="server" Width="94px" Height="16px" ForeColor="Black" BorderStyle="None" BackColor="White" ReadOnly="False"></asp:textbox>
		<asp:textbox ID="DTDLITEM" style="Z-INDEX: 138; POSITION: absolute; TOP: 921px; LEFT: 289px"
			runat="server" Width="92px" Height="16px" ForeColor="Black" BorderStyle="None" BackColor="White" ReadOnly="False"></asp:textbox>
		<asp:textbox ID="DTDRITEM" style="Z-INDEX: 142; POSITION: absolute; TOP: 921px; LEFT: 519px"
			runat="server" Width="94px" Height="16px" ForeColor="Black" BorderStyle="None" BackColor="White" ReadOnly="False"></asp:textbox>
		<asp:textbox ID="DCNITEM" style="Z-INDEX: 139; POSITION: absolute; TOP: 945px; LEFT: 289px" runat="server"
			Width="92px" Height="16px" ForeColor="Black" BorderStyle="None" BackColor="White" ReadOnly="False"></asp:textbox>
		<asp:textbox ID="DCSITEM" style="Z-INDEX: 143; POSITION: absolute; TOP: 945px; LEFT: 519px" runat="server"
			Width="94px" Height="16px" ForeColor="Black" BorderStyle="None" BackColor="White" ReadOnly="False"></asp:textbox>
		<asp:textbox ID="DCDITEM" style="Z-INDEX: 144; POSITION: absolute; TOP: 969px; LEFT: 289px" runat="server"
			Width="324px" Height="16px" ForeColor="Black" BorderStyle="None" BackColor="White" ReadOnly="False"></asp:textbox>
		
		<asp:textbox ID="DODESCR1" style="Z-INDEX: 144; POSITION: absolute; TOP: 992px; LEFT: 160px" runat="server"
			Width="120px" Height="18px" ForeColor="Black" BorderStyle="None" BackColor="White" ReadOnly="False"></asp:textbox>
		
		<asp:textbox ID="DO1ITEM" style="Z-INDEX: 146; POSITION: absolute; TOP: 993px; LEFT: 289px" runat="server"
			Width="324px" Height="16px" ForeColor="Black" BorderStyle="None" BackColor="White" ReadOnly="False"></asp:textbox>

		<asp:textbox ID="DODESCR2" style="Z-INDEX: 144; POSITION: absolute; TOP: 1016px; LEFT: 160px" runat="server"
			Width="120px" Height="18px" ForeColor="Black" BorderStyle="None" BackColor="White" ReadOnly="False"></asp:textbox>

		<asp:textbox ID="DO2ITEM" style="Z-INDEX: 145; POSITION: absolute; TOP: 1017px; LEFT: 289px" runat="server"
			Width="324px" Height="16px" ForeColor="Black" BorderStyle="None" BackColor="White" ReadOnly="False"></asp:textbox>
		<asp:textbox ID="DCITEM" style="Z-INDEX: 147; POSITION: absolute; TOP: 1041px; LEFT: 289px" runat="server"
			Width="324px" Height="16px" ForeColor="Black" BorderStyle="None" BackColor="White" ReadOnly="False"></asp:textbox>
<!--  -->
<!-- 工程 -->            
        <asp:TextBox ID="DOP1" runat="server" BackColor="White" BorderStyle="None"
            ForeColor="Black" Height="16px" ReadOnly="False" Style="z-index: 136; left: 161px;
            position: absolute; top: 1079px" Width="78px"></asp:TextBox>
        <asp:TextBox ID="DOP2" runat="server" BackColor="White" BorderStyle="None"
            ForeColor="Black" Height="16px" ReadOnly="False" Style="z-index: 136; left: 266px;
            position: absolute; top: 1079px" Width="78px"></asp:TextBox>
        <asp:TextBox ID="DOP3" runat="server" BackColor="White" BorderStyle="None"
            ForeColor="Black" Height="16px" ReadOnly="False" Style="z-index: 136; left: 370px;
            position: absolute; top: 1079px" Width="78px"></asp:TextBox>
        <asp:TextBox ID="DOP4" runat="server" BackColor="White" BorderStyle="None"
            ForeColor="Black" Height="16px" ReadOnly="False" Style="z-index: 136; left: 476px;
            position: absolute; top: 1079px" Width="78px"></asp:TextBox>
        <asp:TextBox ID="DOP5" runat="server" BackColor="White" BorderStyle="None"
            ForeColor="Black" Height="16px" ReadOnly="False" Style="z-index: 136; left: 582px;
            position: absolute; top: 1079px" Width="78px"></asp:TextBox>
        <asp:TextBox ID="DOP6" runat="server" BackColor="White" BorderStyle="None"
            ForeColor="Black" Height="16px" ReadOnly="False" Style="z-index: 136; left: 686px;
            position: absolute; top: 1079px" Width="78px"></asp:TextBox>
        &nbsp;<!--  --><!-- +++++++++++++++++++++++++++++++++++++++++++++
             ++ 處理說明及Button Start                  ++ 
             +++++++++++++++++++++++++++++++++++++++++++++ -->&nbsp; &nbsp;<!-- 處理說明及Button End --><!-- +++++++++++++++++++++++++++++++++++++++++++++
                 ++ 核定履歷(View5) Start                     ++ 
                 +++++++++++++++++++++++++++++++++++++++++++++ -->
        
    </div>
    </form>
</body>
</html>
