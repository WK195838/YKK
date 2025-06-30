<%@ Page Language="vb" AutoEventWireup="false" Inherits="SBDCommissionSheet_02" aspCompat="True" CodeFile="SBDCommissionSheet_02.aspx.vb" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>S & B開發委託書</title>
    	<script language="javascript" type="text/javascript">
			var wPop;
			//--Calendar------------------------------------
			function calendarPicker(field1)
			{
				window.open('DatePicker.aspx?field1=' + field1,'calendarPopup','width=250,height=190,resizable=yes');
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
        <asp:Image ID="Image1" runat="server" ImageUrl="~/Images/SBD-CommissionSheet05.jpg" style="z-index: 99; left: 10px; position: absolute; top: 15px" />&nbsp;&nbsp;
           &nbsp; &nbsp;
          &nbsp; &nbsp; &nbsp;&nbsp;
           <asp:HyperLink ID="LRefMapFile" runat="server" Style="z-index: 101; left: 232px; position: absolute;
               top: 448px" Target="_blank">草圖</asp:HyperLink>
           <asp:HyperLink ID="LMapFile" runat="server" Style="z-index: 102; left: 448px; position: absolute;
               top: 550px" Target="_blank">圖檔</asp:HyperLink>
           <asp:HyperLink ID="LContentFile1" runat="server" Style="z-index: 103; left: 528px; position: absolute;
               top: 648px" Target="_blank">修改內容附檔1</asp:HyperLink>
           <asp:HyperLink ID="LContentFile2" runat="server" Style="z-index: 104; left: 528px; position: absolute;
               top: 680px" Target="_blank">修改內容附檔2</asp:HyperLink>
           <asp:HyperLink ID="LContentFile3" runat="server" Style="z-index: 105; left: 528px; position: absolute;
               top: 714px" Target="_blank">修改內容附檔3</asp:HyperLink>
           <asp:HyperLink ID="LContentFile4" runat="server" Style="z-index: 106; left: 528px; position: absolute;
               top: 752px" Target="_blank">修改內容附檔4</asp:HyperLink>
           <asp:HyperLink ID="LForcastFile" runat="server" Style="z-index: 107; left: 232px; position: absolute;
               top: 1224px" Target="_blank">報價單</asp:HyperLink>
           <asp:HyperLink ID="LQAFile" runat="server" Style="z-index: 108; left: 512px; position: absolute;
               top: 1224px" Target="_blank">品測報告書</asp:HyperLink>
           <asp:HyperLink ID="LAuthorizeFile" runat="server" Style="z-index: 109; left: 232px; position: absolute;
               top: 1256px" Target="_blank">確認書</asp:HyperLink>
           <asp:HyperLink ID="LSampleFile" runat="server" Style="z-index: 110; left: 512px; position: absolute;
               top: 1256px" Target="_blank">最終樣品圖</asp:HyperLink>
           <asp:HyperLink ID="LRefFile" runat="server" Style="z-index: 111; left: 512px; position: absolute;
               top: 1296px" Target="_blank">其它文件</asp:HyperLink>
           <asp:HyperLink ID="LContactFile" runat="server" Style="z-index: 112; left: 232px; position: absolute;
               top: 1296px" Target="_blank">切結書</asp:HyperLink>
           <asp:HyperLink ID="LContentFile5" runat="server" Style="z-index: 113; left: 528px; position: absolute;
               top: 784px" Target="_blank">修改內容附檔5</asp:HyperLink>
           <asp:HyperLink ID="LContentFile6" runat="server" Style="z-index: 114; left: 528px; position: absolute;
               top: 816px" Target="_blank">修改內容附檔6</asp:HyperLink>
        
          &nbsp; &nbsp;&nbsp; &nbsp; &nbsp;
           <asp:HyperLink ID="LMap1" runat="server" Style="z-index: 115; left: 256px; position: absolute;
               top: 856px" Target="_blank">圖檔1</asp:HyperLink>
           <asp:HyperLink ID="LMap4" runat="server" Style="z-index: 116; left: 256px; position: absolute;
               top: 928px" Target="_blank">圖檔4</asp:HyperLink>
           <asp:HyperLink ID="LQCReqFile" runat="server" Style="z-index: 117; left: 232px; position: absolute;
               top: 1128px" Target="_blank">分析依賴書</asp:HyperLink>
           <asp:HyperLink ID="LMap5" runat="server" Style="z-index: 118; left: 408px; position: absolute;
               top: 928px" Target="_blank">圖檔5</asp:HyperLink>
           <asp:HyperLink ID="LMap6" runat="server" Style="z-index: 119; left: 552px; position: absolute;
               top: 928px" Target="_blank">圖檔6</asp:HyperLink>
           <asp:HyperLink ID="LMap3" runat="server" Style="z-index: 120; left: 552px; position: absolute;
               top: 856px" Target="_blank">圖檔3</asp:HyperLink>
           <asp:HyperLink ID="LMap2" runat="server" Style="z-index: 121; left: 408px; position: absolute;
               top: 856px" Target="_blank">圖檔2</asp:HyperLink>
           &nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
          &nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
          &nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
        <asp:TextBox ID="DContent2" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 122; left: 128px; position: absolute;
            top: 680px" Width="392px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DContent3" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 123; left: 128px; position: absolute;
            top: 714px" Width="392px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DContent4" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 124; left: 128px; position: absolute;
            top: 752px" Width="392px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DContent5" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 125; left: 128px; position: absolute;
            top: 784px" Width="392px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DContent6" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 126; left: 128px; position: absolute;
            top: 816px" Width="392px" ReadOnly="True"></asp:TextBox>
           <asp:TextBox ID="DSupplier" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 127; left: 240px; position: absolute;
               top: 992px" Width="424px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DHalfFinishOrderNo" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 128; left: 240px; position: absolute;
            top: 1024px" Width="136px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DMold" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 129; left: 240px; position: absolute;
            top: 1056px" Width="88px" ReadOnly="True"></asp:TextBox>
           &nbsp;
        <asp:TextBox ID="DMoldPoint" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 130; left: 456px; position: absolute;
            top: 1056px" Width="88px" ReadOnly="True"></asp:TextBox>
           <asp:TextBox ID="DHalfFinishDate" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 131; left: 536px; position: absolute;
               top: 1024px" Width="128px" ReadOnly="True"></asp:TextBox>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
        <asp:TextBox ID="DSurfcolor" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 132; left: 240px; position: absolute;
            top: 1088px" Width="136px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DQARemark" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 133; left: 232px; position: absolute;
            top: 1188px" Width="440px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DFQA" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 134; left: 232px; position: absolute;
            top: 1155px" Width="440px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DSampleqty" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 135; left: 536px; position: absolute;
            top: 1088px" Width="136px" ReadOnly="True"></asp:TextBox>
        &nbsp; &nbsp;&nbsp; &nbsp;&nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;&nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           <asp:TextBox ID="DMapDate" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" ReadOnly="True" Style="z-index: 136; left: 232px;
               position: absolute; top: 270px" Width="144px"></asp:TextBox>
           <asp:TextBox ID="DSampleDate" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" ReadOnly="True" Style="z-index: 137; left: 528px;
               position: absolute; top: 270px" Width="144px"></asp:TextBox>
           &nbsp; &nbsp;
        &nbsp;&nbsp;&nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
        <asp:TextBox ID="DVendor" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 138; left: 528px; position: absolute;
            top: 168px" Width="144px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DDivision" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 139; left: 232px; position: absolute;
            top: 200px" Width="144px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DCode" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Height="24px" Style="z-index: 140; left: 528px; position: absolute; top: 232px"
            Width="144px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DBackground" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 141; left: 232px; position: absolute;
            top: 232px" Width="144px" ReadOnly="True"></asp:TextBox>
           <asp:TextBox ID="DLight" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
               Height="24px" Style="z-index: 142; left: 232px; position: absolute; top: 304px"
               Width="144px" ReadOnly="True"></asp:TextBox>
           <asp:TextBox ID="DMaterial" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 143; left: 232px; position: absolute;
               top: 376px" Width="96px" ReadOnly="True"></asp:TextBox>
           <asp:TextBox ID="DSample" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 144; left: 616px; position: absolute;
               top: 445px" Width="56px" ReadOnly="True"></asp:TextBox>
           <asp:TextBox ID="DMakeMap" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 145; left: 580px; position: absolute;
               top: 546px" Width="80px" ReadOnly="True"></asp:TextBox>
           <asp:TextBox ID="DMakemap1" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 146; left: 256px; position: absolute;
               top: 888px" Width="112px" ReadOnly="True"></asp:TextBox>
           <asp:TextBox ID="DMakemap2" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 147; left: 408px; position: absolute;
               top: 888px" Width="112px" ReadOnly="True"></asp:TextBox>
           <asp:TextBox ID="DMakemap3" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 148; left: 552px; position: absolute;
               top: 888px" Width="112px" ReadOnly="True"></asp:TextBox>
           <asp:TextBox ID="DMakemap4" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 149; left: 256px; position: absolute;
               top: 958px" Width="112px" ReadOnly="True"></asp:TextBox>
           <asp:TextBox ID="DMakemap6" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 150; left: 552px; position: absolute;
               top: 958px" Width="112px" ReadOnly="True"></asp:TextBox>
           <asp:TextBox ID="DMakemap5" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 151; left: 408px; position: absolute;
               top: 958px" Width="112px" ReadOnly="True"></asp:TextBox>
           &nbsp;
           &nbsp;&nbsp;
           <asp:TextBox ID="DLevel" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
               Height="24px" Style="z-index: 152; left: 616px; position: absolute; top: 512px"
               Width="56px" ReadOnly="True"></asp:TextBox>
           <asp:TextBox ID="DHalfFinishNo" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 153; left: 232px; position: absolute;
               top: 344px" Width="440px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DMaterialDetail" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 154; left: 344px; position: absolute;
            top: 376px" Width="328px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DRemark" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 155; left: 232px; position: absolute;
            top: 480px" Width="440px" ReadOnly="True"></asp:TextBox>
           <asp:TextBox ID="DReason1" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 156; left: 232px; position: absolute;
               top: 580px" Width="440px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DContent1" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 157; left: 128px; position: absolute;
            top: 648px" Width="392px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DMapNo" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Height="24px" Style="z-index: 158; left: 240px; position: absolute; top: 542px"
            Width="92px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DMaterialDetail_1" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 159; left: 232px; position: absolute;
            top: 408px" Width="440px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DAppper" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 160; left: 528px; position: absolute;
            top: 200px" Width="144px" ReadOnly="True"></asp:TextBox>
        &nbsp;
        <asp:TextBox ID="DBuyer" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Style="z-index: 161; left: 232px; position: absolute; top: 168px" Width="144px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DAppDate" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 162; left: 528px; position: absolute;
            top: 128px" Width="144px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="DNo" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 163; left: 232px; position: absolute;
            top: 128px" Width="144px" ReadOnly="True"></asp:TextBox>
           &nbsp;&nbsp;<br />
        <br />
        <br />
        <br />
        <br />
           &nbsp;
        <br />
        <br />
        <br />
        <br />
        <br />
        &nbsp;<br />
        <br />
        <br />
        <br />
           &nbsp;<br />
           &nbsp; &nbsp;&nbsp;<br />
           &nbsp;
        </div>
    </form>
</body>
</html>
