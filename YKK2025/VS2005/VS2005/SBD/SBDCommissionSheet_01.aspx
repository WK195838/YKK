<%@ Page Language="vb" AutoEventWireup="false" Inherits="SBDCommissionSheet_01" aspCompat="True" CodeFile="SBDCommissionSheet_01.aspx.vb" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>S & B開發委託書</title>
    	<script language="javascript" type="text/javascript">
 
			//--Calendar------------------------------------
			function calendarPicker(field1)
			{
				window.open('DatePicker.aspx?field1=' + field1,'calendarPopup','width=250,height=190,resizable=yes');
			}
			
				function openNoPicker(field1)
			{
				window.open('OpenNoPicker.aspx?field1=' + field1,'calendarPopup','width=300,height=600,resizable=yes');
			}


		 
		    function Button(F, MSG) {
				//alert(F);
				document.cookie="RunBOK=False";
				document.cookie="RunBNG1=False";
				document.cookie="RunBNG2=False";
				document.cookie="RunBSAVE=False";

				answer = confirm("是否執行[" + MSG + "]作業嗎？ 請確認....");
				if (answer) {
					//OK Button
					//FOK=document.getElementById("BOK");
					//if(FOK!=null) document.Form1.BOK.disabled=true;  	
					//NG-1 Button
					//FNG1=document.getElementById("BNG1");
					//if(FNG1!=null) document.Form1.BNG1.disabled=true;  	
					//NG-2 Button
					//FNG2=document.getElementById("BNG2");
					//if(FNG2!=null) document.Form1.BNG2.disabled=true;  	
					//Save Button
					//FSAVE=document.getElementById("BSAVE");
					//if(FSAVE!=null) document.Form1.BSAVE.disabled=true;  	
						
					if (F=="OK")   document.cookie="RunBOK=True";
					if (F=="NG1")  document.cookie="RunBNG1=True";
					if (F=="NG2")  document.cookie="RunBNG2=True";
					if (F=="SAVE") document.cookie="RunBSAVE=True";
				}
			}
		   
		</script>
		
</head>
<body>

  	<form id="Form1" runat="server">
  	 <div>
        <asp:Image ID="Image1" runat="server" ImageUrl="~/Images/SBD-CommissionSheet05.jpg" style="z-index: 100; left: 10px; position: absolute; top: 15px" />&nbsp;
           <asp:TextBox ID="DSurfcolor1" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 101; left: 352px; position: absolute;
               top: 1088px" Width="80px"></asp:TextBox>
           &nbsp;&nbsp;
 
        <input id="DQCReqFile" runat="server" name="UploadFile" style="z-index: 175; left: 232px;
            width: 440px; position: absolute; top: 1128px; height: 20px; background-color: #ffff00"
            type="file" visible="true" />
           <asp:TextBox ID="DSurfFormNo" runat="server" BackColor="Transparent" BorderStyle="None"
               ForeColor="White" Height="24px" Style="z-index: 102; left: 688px; position: absolute;
               top: 1088px" Width="80px"></asp:TextBox>
           <asp:TextBox ID="DSurfFormSno" runat="server" BackColor="Transparent" BorderStyle="None"
               ForeColor="White" Height="24px" Style="z-index: 103; left: 688px; position: absolute;
               top: 1120px" Width="80px"></asp:TextBox>
      
        <input id="DForcastFile" runat="server" name="UploadFile" style="z-index: 187; left: 232px;
            width: 144px; position: absolute; top: 1224px; height: 20px; background-color: #ffff00"
            type="file" visible="true" />
    
           &nbsp;&nbsp;
           <asp:HyperLink ID="LRefMapFile" runat="server" Style="z-index: 104; left: 232px; position: absolute;
               top: 448px" Target="_blank">草圖</asp:HyperLink>
           <asp:HyperLink ID="LMapFile" runat="server" Style="z-index: 105; left: 443px; position: absolute;
               top: 548px" Target="_blank">圖檔</asp:HyperLink>
           <asp:HyperLink ID="LContentFile1" runat="server" Style="z-index: 106; left: 528px; position: absolute;
               top: 648px" Target="_blank">修改內容附檔1</asp:HyperLink>
           <asp:HyperLink ID="LContentFile2" runat="server" Style="z-index: 107; left: 528px; position: absolute;
               top: 680px" Target="_blank">修改內容附檔2</asp:HyperLink>
           <asp:HyperLink ID="LContentFile3" runat="server" Style="z-index: 108; left: 528px; position: absolute;
               top: 714px" Target="_blank">修改內容附檔3</asp:HyperLink>
           <asp:HyperLink ID="LContentFile4" runat="server" Style="z-index: 109; left: 528px; position: absolute;
               top: 752px" Target="_blank">修改內容附檔4</asp:HyperLink>
           <asp:HyperLink ID="LForcastFile" runat="server" Style="z-index: 110; left: 232px; position: absolute;
               top: 1224px" Target="_blank">報價單</asp:HyperLink>
           <asp:HyperLink ID="LQAFile" runat="server" Style="z-index: 111; left: 512px; position: absolute;
               top: 1224px" Target="_blank">品測報告書</asp:HyperLink>
           <asp:HyperLink ID="LAuthorizeFile" runat="server" Style="z-index: 112; left: 232px; position: absolute;
               top: 1256px" Target="_blank">確認書</asp:HyperLink>
           <asp:HyperLink ID="LSampleFile" runat="server" Style="z-index: 113; left: 512px; position: absolute;
               top: 1256px" Target="_blank">最終樣品圖</asp:HyperLink>
           <asp:HyperLink ID="LRefFile" runat="server" Style="z-index: 114; left: 512px; position: absolute;
               top: 1296px" Target="_blank">其它文件</asp:HyperLink>
           <asp:HyperLink ID="LContactFile" runat="server" Style="z-index: 115; left: 232px; position: absolute;
               top: 1296px" Target="_blank">切結書</asp:HyperLink>
           <asp:HyperLink ID="LContentFile5" runat="server" Style="z-index: 116; left: 528px; position: absolute;
               top: 784px" Target="_blank">修改內容附檔5</asp:HyperLink>
           <asp:HyperLink ID="LContentFile6" runat="server" Style="z-index: 117; left: 528px; position: absolute;
               top: 816px" Target="_blank">修改內容附檔6</asp:HyperLink>
        
          &nbsp; &nbsp;&nbsp; &nbsp; &nbsp;
           <asp:HyperLink ID="LMap1" runat="server" Style="z-index: 118; left: 256px; position: absolute;
               top: 856px" Target="_blank">圖檔1</asp:HyperLink>
           <asp:HyperLink ID="LMap4" runat="server" Style="z-index: 119; left: 256px; position: absolute;
               top: 928px" Target="_blank">圖檔4</asp:HyperLink>
           <asp:HyperLink ID="LQCReqFile" runat="server" Style="z-index: 120; left: 232px; position: absolute;
               top: 1128px" Target="_blank">分析依賴書</asp:HyperLink>
           <asp:HyperLink ID="LSurfColor" runat="server" Style="z-index: 121; left: 232px; position: absolute;
               top: 1088px" Target="_blank" Width="80px" Height="16px">表面處理</asp:HyperLink>
           <asp:HyperLink ID="LMap5" runat="server" Style="z-index: 122; left: 408px; position: absolute;
               top: 928px" Target="_blank">圖檔5</asp:HyperLink>
           <asp:HyperLink ID="LMap6" runat="server" Style="z-index: 123; left: 552px; position: absolute;
               top: 928px" Target="_blank">圖檔6</asp:HyperLink>
           <asp:HyperLink ID="LMap3" runat="server" Style="z-index: 124; left: 552px; position: absolute;
               top: 856px" Target="_blank">圖檔3</asp:HyperLink>
           <asp:HyperLink ID="LMap2" runat="server" Style="z-index: 125; left: 408px; position: absolute;
               top: 856px" Target="_blank">圖檔2</asp:HyperLink>
           &nbsp; &nbsp; &nbsp;&nbsp;
        <input id="DQAFile" runat="server" name="UploadFile" style="z-index: 189; left: 512px;
            width: 160px; position: absolute; top: 1224px; height: 20px; background-color: #ffff00"
            type="file" visible="true" />
        <input id="DAuthorizeFile" runat="server" name="UploadFile" style="z-index: 176; left: 232px;
            width: 144px; position: absolute; top: 1256px; height: 20px; background-color: #ffff00"
            type="file" visible="true" />
        <input id="DSampleFile" runat="server" name="UploadFile" style="z-index: 195; left: 512px;
            width: 160px; position: absolute; top: 1256px; height: 20px; background-color: #ffff00"
            type="file" visible="true" />
           &nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
          
        <input id="DContactFile" runat="server" name="UploadFile" style="z-index: 194; left: 232px;
            width: 144px; position: absolute; top: 1296px; height: 20px; background-color: #ffff00"
            type="file" visible="true" />
          &nbsp;
        <input id="DRefFile" runat="server" name="UploadFile" style="z-index: 193; left: 512px;
            width: 160px; position: absolute; top: 1296px; height: 20px; background-color: #ffff00"
            type="file" visible="true" />
   
        <input id="DContentFile2" runat="server" name="UploadFile" style="z-index: 191; left: 528px;
            width: 144px; position: absolute; top: 680px; height: 20px; background-color: #ffff00"
            type="file" visible="true" />
          &nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
          &nbsp; &nbsp;&nbsp;&nbsp;
        <input id="BSAVE" runat="server" name="Button1" onclick="Button('SAVE');" style="z-index: 199;
            left: 312px; width: 80px; position: absolute; top: 1368px; height: 28px" type="button"
            value="儲存" />
        <input id="BNG2" runat="server" name="Button1" onclick="Button('NG2');" style="z-index: 198;
            left: 400px; width: 80px; position: absolute; top: 1368px; height: 28px" type="button"
            value="NG2" />
        <input id="BNG1" runat="server" name="Button1" onclick="Button('NG1');" style="z-index: 197;
            left: 488px; width: 80px; position: absolute; top: 1368px; height: 28px" type="button"
            value="NG1" />
        &nbsp;&nbsp;
        <input id="BOK" runat="server" name="BOK" onclick="Button('OK');" style="z-index: 196;
            left: 576px; width: 80px; position: absolute; top: 1368px; height: 28px" type="button"
            value="OK" />
        <input id="DContentFile3" runat="server" name="UploadFile" style="z-index: 181; left: 528px;
            width: 144px; position: absolute; top: 714px; height: 20px; background-color: #ffff00"
            type="file" visible="true" />
        <input id="DContentFile4" runat="server" name="UploadFile" style="z-index: 190; left: 528px;
            width: 144px; position: absolute; top: 752px; height: 20px; background-color: #ffff00"
            type="file" visible="true" />
        <input id="DContentFile5" runat="server" name="UploadFile" style="z-index: 192; left: 528px;
            width: 144px; position: absolute; top: 816px; height: 20px; background-color: #ffff00"
            type="file" visible="true" />
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        <input id="DContentFile6" runat="server" name="UploadFile" style="z-index: 188; left: 528px;
            width: 144px; position: absolute; top: 784px; height: 20px; background-color: #ffff00"
            type="file" visible="true" />
          &nbsp;
        <asp:TextBox ID="DContent2" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 126; left: 128px; position: absolute;
            top: 680px" Width="392px"></asp:TextBox>
        <asp:TextBox ID="DContent3" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 127; left: 128px; position: absolute;
            top: 714px" Width="392px"></asp:TextBox>
        <asp:TextBox ID="DContent4" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 128; left: 128px; position: absolute;
            top: 752px" Width="392px"></asp:TextBox>
        <asp:TextBox ID="DContent5" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 129; left: 128px; position: absolute;
            top: 784px" Width="392px"></asp:TextBox>
        <asp:TextBox ID="DContent6" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 130; left: 128px; position: absolute;
            top: 816px" Width="392px"></asp:TextBox>
        <asp:TextBox ID="DHalfFinishOrderNo" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 131; left: 232px; position: absolute;
            top: 1024px" Width="144px"></asp:TextBox>
        <asp:TextBox ID="DMold" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 132; left: 232px; position: absolute;
            top: 1056px" Width="96px"></asp:TextBox>
           &nbsp;
        <asp:TextBox ID="DMoldPoint" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 133; left: 448px; position: absolute;
            top: 1056px" Width="88px"></asp:TextBox>
        &nbsp;
        <input id="BDate1" runat="server" style="z-index: 182; left: 352px; width: 20px;
            position: absolute; top: 272px" type="button" value="..." /><input id="BColor" runat="server" style="z-index: 183; left: 328px; width: 20px;
            position: absolute; top: 1088px" type="button" value="..." />
           <input id="DSurfcolor" runat="server" style="z-index: 179; left: 232px; width: 96px;
            color: blue; border-top-style: groove; border-right-style: groove; border-left-style: groove;
            position: absolute; top: 1088px; background-color: yellow; border-bottom-style: groove"
            type="text" />
        <input id="BDate2" runat="server" style="z-index: 184; left: 648px; width: 20px;
            position: absolute; top: 272px" type="button" value="..." />
        <input id="BDate3" runat="server" style="z-index: 185; left: 648px; width: 20px;
            position: absolute; top: 1024px" type="button" value="..." />
        &nbsp; &nbsp;
        <asp:TextBox ID="DQARemark" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 134; left: 232px; position: absolute;
            top: 1188px" Width="440px"></asp:TextBox>
        <asp:TextBox ID="DFQA" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 135; left: 232px; position: absolute;
            top: 1155px" Width="440px"></asp:TextBox>
           &nbsp;&nbsp;
           <asp:DropDownList ID="DSampleQty" runat="server" BackColor="Yellow" ForeColor="Blue"
               Height="20px" Style="z-index: 136; left: 560px; position: absolute; top: 1088px"
               Width="104px">
           </asp:DropDownList>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp;&nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
        <input id="DContentFile1" runat="server" name="UploadFile" style="z-index: 186; left: 528px;
            width: 144px; position: absolute; top: 648px; height: 20px; background-color: #ffff00"
            type="file" visible="true" />
          &nbsp; &nbsp;&nbsp; &nbsp;&nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp;&nbsp;&nbsp;
          &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;&nbsp;
          &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp;&nbsp; &nbsp;&nbsp;
        <input id="DRefMapFile" runat="server" name="UploadFile" style="z-index: 172; left: 232px;
            width: 304px; position: absolute; top: 448px; height: 20px; background-color: #ffff00"
            type="file" visible="false" />
           &nbsp;
        <input id="DMapFile" runat="server"  name="UploadFile" style="z-index: 173; left: 238px;
            width: 108px; position: absolute; top: 514px; height: 20px; background-color: #ffff00"
            type="file" visible="false" />
        &nbsp;&nbsp;&nbsp;&nbsp;
        <asp:DropDownList ID="DLight" runat="server" BackColor="Yellow"
            ForeColor="Blue" Height="20px" Style="z-index: 137; left: 232px; position: absolute;
            top: 304px" Width="144px">
            <asp:ListItem Value="3"></asp:ListItem>
        </asp:DropDownList>
        <asp:DropDownList ID="DMaterial" runat="server" AutoPostBack="True" BackColor="Yellow"
            ForeColor="Blue" Height="20px" Style="z-index: 138; left: 232px; position: absolute;
            top: 376px" Width="96px">
            <asp:ListItem Value="3"></asp:ListItem>
        </asp:DropDownList><asp:DropDownList ID="DBuyer" runat="server" BackColor="Yellow"
            ForeColor="Blue" Height="20px" Style="z-index: 139; left: 232px; position: absolute;
            top: 168px" Width="96px">
            <asp:ListItem Value="3"></asp:ListItem>
        </asp:DropDownList>
        <asp:DropDownList ID="DSample" runat="server" BackColor="Yellow"
            ForeColor="Blue" Height="20px" Style="z-index: 140; left: 616px; position: absolute;
            top: 448px" Width="56px">
            <asp:ListItem Value="3"></asp:ListItem>
        </asp:DropDownList>
        &nbsp;
        <asp:DropDownList ID="DLevel" runat="server" BackColor="Yellow"
            ForeColor="Blue" Height="20px" Style="z-index: 141; left: 616px; position: absolute;
            top: 512px" Width="56px">
            <asp:ListItem Value="3"></asp:ListItem>
        </asp:DropDownList>
        <asp:DropDownList ID="DMakeMap" runat="server" BackColor="Yellow"
            ForeColor="Blue" Height="20px" Style="z-index: 142; left: 437px; position: absolute;
            top: 513px" Width="102px">
            <asp:ListItem Value="3"></asp:ListItem>
        </asp:DropDownList><asp:DropDownList ID="DSupplier" runat="server" BackColor="Yellow"
            ForeColor="Blue" Height="20px" Style="z-index: 143; left: 232px; position: absolute;
            top: 992px" Width="440px">
            <asp:ListItem Value="3"></asp:ListItem>
        </asp:DropDownList>
           &nbsp; &nbsp;&nbsp;
        <asp:DropDownList ID="DMaterialDetail" runat="server" BackColor="Yellow"
            ForeColor="Blue" Height="20px" Style="z-index: 144; left: 344px; position: absolute;
            top: 376px" Width="328px">
            <asp:ListItem Value="3"></asp:ListItem>
        </asp:DropDownList><asp:DropDownList ID="DReason1" runat="server" BackColor="Yellow"
            ForeColor="Blue" Height="20px" Style="z-index: 145; left: 232px; position: absolute;
            top: 580px" Width="440px">
            <asp:ListItem Value="3"></asp:ListItem>
        </asp:DropDownList>
        &nbsp; &nbsp;
        <input id="DMapDate" runat="server" style="z-index: 178; left: 232px; width: 120px;
            color: blue; border-top-style: groove; border-right-style: groove; border-left-style: groove;
            position: absolute; top: 272px; background-color: yellow; border-bottom-style: groove"
            type="text" value="2009/12/12" />
        &nbsp;&nbsp;
        <input id="DSampleDate" runat="server" style="z-index: 180; left: 528px; width: 120px;
            color: blue; border-top-style: groove; border-right-style: groove; border-left-style: groove;
            position: absolute; top: 272px; background-color: yellow; border-bottom-style: groove"
            type="text" value="2009/12/12" />
        <input id="DHalfFinishDate" runat="server" style="z-index: 177; left: 528px; width: 120px;
            color: blue; border-top-style: groove; border-right-style: groove; border-left-style: groove;
            position: absolute; top: 1024px; background-color: yellow; border-bottom-style: groove"
            type="text" value="2009/12/12" />
        &nbsp;
        <asp:TextBox ID="DVendor" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 146; left: 528px; position: absolute;
            top: 168px" Width="144px"></asp:TextBox>
        <asp:TextBox ID="DDivision" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 147; left: 232px; position: absolute;
            top: 200px" Width="144px"></asp:TextBox>
        <asp:TextBox ID="DCode" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Height="24px" Style="z-index: 148; left: 528px; position: absolute; top: 232px"
            Width="144px"></asp:TextBox>
        <asp:TextBox ID="DBackground" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 149; left: 232px; position: absolute;
            top: 232px" Width="144px"></asp:TextBox>
        <asp:TextBox ID="DHalfFinishNo" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 150; left: 232px; position: absolute;
            top: 344px" Width="440px"></asp:TextBox>
        <asp:TextBox ID="DRemark" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 151; left: 232px; position: absolute;
            top: 480px" Width="440px"></asp:TextBox>
        <asp:TextBox ID="DContent1" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 152; left: 128px; position: absolute;
            top: 648px" Width="392px"></asp:TextBox>
        <asp:TextBox ID="DMapNo" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Height="24px" Style="z-index: 153; left: 237px; position: absolute; top: 545px"
            Width="106px"></asp:TextBox>
           <asp:TextBox ID="DMakeMapU" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 154; left: 577px; position: absolute;
               top: 544px" Width="85px"></asp:TextBox>
           <asp:TextBox ID="DMakeMap1" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 155; left: 257px; position: absolute;
               top: 884px" Width="109px"></asp:TextBox>
           <asp:TextBox ID="DMakeMap2" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 156; left: 408px; position: absolute;
               top: 885px" Width="109px"></asp:TextBox>
           <asp:TextBox ID="DMakeMap3" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 157; left: 554px; position: absolute;
               top: 885px" Width="109px"></asp:TextBox>
           <asp:TextBox ID="DMakeMap4" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 158; left: 257px; position: absolute;
               top: 955px" Width="109px"></asp:TextBox>
           <asp:TextBox ID="DMakeMap5" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 159; left: 407px; position: absolute;
               top: 956px" Width="109px"></asp:TextBox>
           <asp:TextBox ID="DMakeMap6" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 160; left: 557px; position: absolute;
               top: 955px" Width="109px"></asp:TextBox>
           &nbsp;
        <asp:TextBox ID="DMaterialDetail_1" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 161; left: 232px; position: absolute;
            top: 408px" Width="440px"></asp:TextBox>
        <asp:TextBox ID="DAppper" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 162; left: 528px; position: absolute;
            top: 200px" Width="144px"></asp:TextBox>
        &nbsp; &nbsp;
        <asp:TextBox ID="DAppDate" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 163; left: 528px; position: absolute;
            top: 128px" Width="144px"></asp:TextBox>
        <asp:TextBox ID="DNo" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 164; left: 232px; position: absolute;
            top: 128px" Width="144px"></asp:TextBox>
           &nbsp;<br />
           <asp:TextBox ID="DDecideDesc" runat="server" BackColor="Yellow" ForeColor="Blue"
               Height="58px" Style="z-index: 165; left: 72px; position: absolute; top: 1432px"
               TextMode="MultiLine" Width="536px"></asp:TextBox>
        <br />
        <br />
        <br />
        <br />
           &nbsp;
           <img id="DDescSheet" runat="server" src="images/MapSheet_004.jpg" style="z-index: 99;
               left: 24px; position: absolute; top: 1424px" />
        <br />
        <br />
        <br />
        <br />
        <br />
        <img id="DDelay" runat="server" src="images/Sheet_Delay.jpg" style="z-index: 174; left: 24px;
            width: 592px; position: absolute; top: 1504px; height: 112px" visible="false" />
        &nbsp;<br />
        <br />
        <br />
        <br />
        <asp:Label ID="DHistoryLabel" runat="server" Font-Names="新細明體" Font-Size="11pt" ForeColor="Blue"
            Height="16px" Style="z-index: 166; left: 16px; position: absolute; top: 1704px"
            Width="64px">核定履歷</asp:Label>
        <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" BackColor="White"
            BorderColor="#CC9966" BorderStyle="None" BorderWidth="1px" CellPadding="4" Height="100px"
            Style="z-index: 167; left: 16px; position: absolute; top: 1720px" Width="780px">
            <RowStyle BackColor="White" Font-Size="9pt" ForeColor="#330099" />
            <Columns>
                <asp:BoundField DataField="StepNameDesc" HeaderText="工程">
                    <ItemStyle HorizontalAlign="Left" />
                </asp:BoundField>
                <asp:BoundField DataField="DecideName" HeaderText="擔當" />
                <asp:BoundField DataField="AgentName" HeaderText="代理/兼職" />
                <asp:BoundField DataField="FlowTypeDesc" HeaderText="類別" />
                <asp:BoundField DataField="StsDesc" HeaderText="處理結果" />
                <asp:BoundField DataField="DecideDescA" HeaderText="說明">
                    <ItemStyle HorizontalAlign="Left" />
                </asp:BoundField>
                <asp:BoundField DataField="Description" HeaderText="核定時間">
                    <ItemStyle HorizontalAlign="Left" />
                </asp:BoundField>
            </Columns>
            <HeaderStyle BackColor="#CC6600" Font-Size="9pt" ForeColor="#FFFFCC" HorizontalAlign="Center"
                VerticalAlign="Middle" />
        </asp:GridView>
        <br />
        <asp:TextBox ID="DReason" runat="server" BackColor="Yellow" ForeColor="Blue" Height="20px"
            Style="z-index: 168; left: 248px; position: absolute; top: 1504px" Visible="False"
            Width="360px"></asp:TextBox>
        <asp:DropDownList ID="DReasonCode" runat="server" Height="20px" Style="z-index: 169;
            left: 184px; position: absolute; top: 1504px" Visible="False" Width="64px">
            <asp:ListItem>01</asp:ListItem>
            <asp:ListItem>02</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DReasonDesc" runat="server" BackColor="Yellow" ForeColor="Blue"
            Height="56px" Style="z-index: 170; left: 176px; position: absolute; top: 1544px"
            Visible="False" Width="432px"></asp:TextBox>
        &nbsp;&nbsp;<br />
         <asp:CustomValidator  ID="CustomValidator1" runat="server" ErrorMessage="CustomValidator"></asp:CustomValidator>
        <asp:ValidationSummary ID="ValidationSummary1" runat="server" ShowMessageBox="true" ShowSummary="false" />
        </div>
    </form>
</body>
</html>
