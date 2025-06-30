<%@ Page Language="VB" AutoEventWireup="false" CodeFile="HR_WorkRatioInformation.aspx.vb" Inherits="HR_WorkRatioInformation" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head>
    <title>個人出勤率</title>
    <script>
        function ChangeImage(btn)
        {  

            __doPostBack(btn.name, "");
           
            var btnname = btn.id.substr(11,1);
            
             if (btnname == "1" )
             {
        alert(btn.id);
               document.getElementById('ImageButton1').src ="Images/WorkRatioInfor_B1_Blank.jpg";
               document.getElementById('ImageButton2').src ="Images/WorkRatioInfor_B2.jpg";
               document.getElementById('ImageButton3').src ="Images/WorkRatioInfor_B3.jpg";
             }           
              if (btnname == "2" )
             {
               document.getElementById('ImageButton1').src ="Images/WorkRatioInfor_B1.jpg";
               document.getElementById('ImageButton2').src ="Images/WorkRatioInfor_B2_Blank.jpg";
               document.getElementById('ImageButton3').src ="Images/WorkRatioInfor_B3.jpg";
             }
           
             if (btnname == "3" )
             {
               document.getElementById('ImageButton1').src ="Images/WorkRatioInfor_B1.jpg";
               document.getElementById('ImageButton2').src ="Images/WorkRatioInfor_B2.jpg";
               document.getElementById('ImageButton3').src ="Images/WorkRatioInfor_B3_Blank.jpg";
             }       
        } 
    </script>    
</head>
<body>
  <form id="form1" runat="server">
    <div>
        <!-- Panel Start -->
        <asp:Panel ID="Panel1" runat="server" Style="left: 0px; position: relative; top: 0px" Width="416px">
          <asp:ImageButton ID="ImageButton1" runat="server" ImageUrl="~/Images/WorkRatioInfor_B1.jpg" />
          <asp:ImageButton ID="ImageButton2" runat="server" ImageUrl="~/Images/WorkRatioInfor_B2.jpg" style="left: -6px; position: relative; top: 0px" />
          <asp:ImageButton ID="ImageButton3" runat="server" ImageUrl="~/Images/WorkRatioInfor_B3.jpg" style="left: -14px; position: relative; top: 0px" />
        </asp:Panel>
        <!-- Panel End -->
        <!-- MultiView Start -->
            <asp:MultiView ID="MultiView1" runat="server" ActiveViewIndex="0" >
            <!-- ********************************************* 
                 ***                出勤率                 *** 
                 ********************************************* -->
        <!-- View-1 Start -->
              <asp:View ID="View1" runat="server">
                <table border="0" style="width: 300px; height: 100px; position: relative;" cellpadding="0" cellspacing="0">
        <asp:DropDownList ID="DDivision" runat="server" AutoPostBack="True" BackColor="Yellow"
            ForeColor="Blue" Height="24px" Style="z-index: 100; left: 466px; position: absolute;
            top: 43px" Width="197px">
        </asp:DropDownList>
        <asp:DropDownList ID="DName" runat="server" BackColor="Yellow" ForeColor="Blue"
            Height="24px" Style="z-index: 101; left: 124px; position: absolute; top: 75px"
            Width="137px" AutoPostBack="True">
        </asp:DropDownList>
        <asp:TextBox ID="DJobTitle" runat="server" BackColor="#C0FFC0" BorderStyle="None"
            ForeColor="Black" Height="16px" ReadOnly="True" Style="z-index: 102; left: 466px;
            position: absolute; top: 80px" Width="93px">DJobTitle</asp:TextBox>
        <asp:TextBox ID="DYear" runat="server" BackColor="#C0FFC0" BorderStyle="None"
            ForeColor="Black" Height="20px" ReadOnly="True" Style="z-index: 103; left: 466px;
            position: absolute; top: 11px" Width="93px"></asp:TextBox>
        <asp:TextBox ID="DBaseDate" runat="server" BackColor="#C0FFC0" BorderStyle="None"
            ForeColor="Black" Height="20px" ReadOnly="True" Style="z-index: 104; left: 124px;
            position: absolute; top: 44px" Width="76px"></asp:TextBox>
        <asp:TextBox ID="DJobCode" runat="server" BackColor="#C0FFC0" BorderStyle="None"
            ForeColor="Black" Height="16px" ReadOnly="True" Style="z-index: 105; left: 577px;
            position: absolute; top: 78px" Width="81px">DJobCode</asp:TextBox>
        <asp:Image ID="Image1" runat="server" ImageUrl="~/Images/HR_WorkRatioInfor_01.PNG"
            Style="z-index: 99; left: 2px; position: absolute; top: -1px" />
        <asp:TextBox ID="DEmpID" runat="server" BackColor="#C0FFC0" BorderStyle="None" ForeColor="Black"
            Height="20px" ReadOnly="True" Style="z-index: 107; left: 274px; position: absolute;
            top: 75px" Width="61px">DEmpID</asp:TextBox>
        <asp:TextBox ID="DDate" runat="server" BackColor="#C0FFC0" BorderStyle="None" ForeColor="Black"
            Height="20px" ReadOnly="True" Style="z-index: 108; left: 124px; position: absolute;
            top: 11px" Width="213px">DDate</asp:TextBox>
        <asp:TextBox ID="DYVcation" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 110; left: 129px;
            position: absolute; top: 208px" Width="46px"></asp:TextBox>
        <asp:TextBox ID="DSVcation" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 111; left: 129px;
            position: absolute; top: 242px" Width="46px"></asp:TextBox>
        <asp:TextBox ID="DZVcation" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 112; left: 129px;
            position: absolute; top: 273px" Width="46px"></asp:TextBox>
        <asp:Label ID="DSum" runat="server" ForeColor="Blue" Style="z-index: 113; left: 129px;
            position: absolute; top: 508px" Width="335px"></asp:Label>
        <asp:TextBox ID="DMVcation" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 114; left: 129px;
            position: absolute; top: 307px" Width="46px"></asp:TextBox>
        <asp:TextBox ID="DXVcation" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 115; left: 129px;
            position: absolute; top: 341px" Width="46px"></asp:TextBox>
        <asp:TextBox ID="DPVcation" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 116; left: 129px;
            position: absolute; top: 375px" Width="46px"></asp:TextBox>
        <asp:TextBox ID="DQVcation" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 117; left: 129px;
            position: absolute; top: 406px" Width="46px"></asp:TextBox>
        <asp:TextBox ID="DIVcationM" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 118; left: 231px;
            position: absolute; top: 178px" Width="46px"></asp:TextBox>
        <asp:TextBox ID="DYVcationM" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 119; left: 231px;
            position: absolute; top: 208px" Width="46px"></asp:TextBox>
        <asp:TextBox ID="DSVcationM" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 120; left: 231px;
            position: absolute; top: 242px" Width="46px"></asp:TextBox>
        <asp:TextBox ID="DZVcationM" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 121; left: 231px;
            position: absolute; top: 273px" Width="46px"></asp:TextBox>
        <asp:TextBox ID="DMVcationM" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 122; left: 231px;
            position: absolute; top: 307px" Width="46px"></asp:TextBox>
        <asp:TextBox ID="DXVcationM" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 123; left: 231px;
            position: absolute; top: 341px" Width="46px"></asp:TextBox>
        <asp:TextBox ID="DPVcationM" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 124; left: 231px;
            position: absolute; top: 375px" Width="46px"></asp:TextBox>
        <asp:TextBox ID="DQVcationM" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 125; left: 231px;
            position: absolute; top: 406px" Width="46px"></asp:TextBox>
        <asp:TextBox ID="DZ1M" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 126; left: 231px;
            position: absolute; top: 443px" Width="46px"></asp:TextBox>
        <asp:TextBox ID="DX2" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 127; left: 231px;
            position: absolute; top: 475px" Width="46px"></asp:TextBox>
        <asp:TextBox ID="DYearDay" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 128; left: 465px;
            position: absolute; top: 178px" Width="46px"></asp:TextBox>
        <asp:TextBox ID="DVacationday" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 129; left: 465px;
            position: absolute; top: 208px" Width="46px"></asp:TextBox>
        <asp:TextBox ID="DWorkRate" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 130; left: 592px;
            position: absolute; top: 320px" Width="46px"></asp:TextBox>
        <asp:TextBox ID="DY2" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 131; left: 465px;
            position: absolute; top: 242px" Width="46px"></asp:TextBox>
        <asp:TextBox ID="DY3" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 132; left: 465px;
            position: absolute; top: 273px" Width="46px"></asp:TextBox>
        <asp:TextBox ID="DAWorkday" runat="server" BackColor="#FFC080" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 133; left: 231px;
            position: absolute; top: 146px" Width="46px"></asp:TextBox>
        <asp:TextBox ID="DBworkday" runat="server" BackColor="#FFC080" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 134; left: 465px;
            position: absolute; top: 146px" Width="46px"></asp:TextBox>
        <asp:TextBox ID="DIVcation" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="14px" ReadOnly="True" Style="z-index: 135; left: 129px;
            position: absolute; top: 178px" Width="46px"></asp:TextBox>
                </table>
              </asp:View>
        <!-- View-1 End -->
            <!-- ********************************************* 
                 ***                休假                   *** 
                 ********************************************* -->
        <!-- View-2 Start -->
              <asp:View ID="View2" runat="server">
                <table border="0" style="width: 300px; height: 100px; position: relative;" cellpadding="0" cellspacing="0">
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" BorderColor="#CC9966"
            BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt" ShowFooter="True"
            Style="z-index: 114; left: 12px; position: absolute; top: 54px" Width="669px">
            <Columns>
                <asp:HyperLinkField DataNavigateUrlFields="URL" DataNavigateUrlFormatString="{0}"
                    DataTextField="No" HeaderText="請假單No." Target="_blank">
                    <FooterStyle HorizontalAlign="Center" />
                    <HeaderStyle Width="100px" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:HyperLinkField>
                <asp:BoundField DataField="Vacation" HeaderText="假別" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                    <HeaderStyle Width="100px" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="VacationTime" HeaderText="期間" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                    <HeaderStyle Width="200px" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="ADays" HeaderText="日數" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                    <HeaderStyle Width="50px" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
            </Columns>
            <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
            <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
        </asp:GridView>
            </table>
              </asp:View>
        <!-- View-2 End -->
            <!-- ********************************************* 
                 ***                其他說明               *** 
                 ********************************************* -->
        <!-- View-3 Start -->
              <asp:View ID="View3" runat="server">
                <table border="0" style="width: 300px; height: 100px; position: relative;" cellpadding="0" cellspacing="0">

        <asp:Image ID="Image2" runat="server" ImageUrl="~/Images/other.PNG" Style="z-index: 99;
            left: 5px; position: absolute; top: 2px" />
        <asp:TextBox ID="DZZSum" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 104; left: 140px;
            position: absolute; top: 626px" Width="76px"></asp:TextBox>
        <asp:TextBox ID="DZZMin1" runat="server" BackColor="#FFFF80" BorderStyle="None" ForeColor="Black"
            Height="17px" ReadOnly="True" Style="z-index: 116; left: 140px; position: absolute;
            top: 459px" Width="76px"></asp:TextBox>
        <asp:TextBox ID="DZZMin2" runat="server" BackColor="#FFFF80" BorderStyle="None" ForeColor="Black"
            Height="17px" ReadOnly="True" Style="z-index: 117; left: 140px; position: absolute;
            top: 493px" Width="76px"></asp:TextBox>
        <asp:TextBox ID="DZZMin3" runat="server" BackColor="#FFFF80" BorderStyle="None" ForeColor="Black"
            Height="17px" ReadOnly="True" Style="z-index: 118; left: 140px; position: absolute;
            top: 523px" Width="76px"></asp:TextBox>
        <asp:TextBox ID="DZZMin4" runat="server" BackColor="#FFFF80" BorderStyle="None" ForeColor="Black"
            Height="17px" ReadOnly="True" Style="z-index: 119; left: 140px; position: absolute;
            top: 558px" Width="76px"></asp:TextBox>
        <asp:TextBox ID="DZZMin5" runat="server" BackColor="#FFFF80" BorderStyle="None" ForeColor="Black"
            Height="17px" ReadOnly="True" Style="z-index: 120; left: 140px; position: absolute;
            top: 592px" Width="76px"></asp:TextBox>
        <asp:TextBox ID="DZZDesc2" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 121; left: 233px;
            position: absolute; top: 493px" Width="485px"></asp:TextBox>
        <asp:TextBox ID="DZZDesc3" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 122; left: 233px;
            position: absolute; top: 523px" Width="485px"></asp:TextBox>
        <asp:TextBox ID="DZZDesc4" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 123; left: 233px;
            position: absolute; top: 558px" Width="485px"></asp:TextBox>
        <asp:TextBox ID="DZZDesc5" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 124; left: 233px;
            position: absolute; top: 592px" Width="485px"></asp:TextBox>
        <asp:TextBox ID="DZZDesc1" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 126; left: 233px;
            position: absolute; top: 459px" Width="485px"></asp:TextBox>
        <asp:TextBox ID="DX2Sum" runat="server" BackColor="#FFFF80" BorderStyle="None" ForeColor="Black"
            Height="17px" ReadOnly="True" Style="z-index: 104; left: 140px; position: absolute;
            top: 898px" Width="76px"></asp:TextBox>
        <asp:TextBox ID="TextBox4" runat="server" BackColor="#FFC080" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 105; left: 231px;
            position: absolute; top: 109px" Width="488px"></asp:TextBox>
        <asp:TextBox ID="DY3Desc1" runat="server" BackColor="#FFC080" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 107; left: 230px;
            position: absolute; top: 184px" Width="488px"></asp:TextBox>
        <asp:TextBox ID="DY3Min1" runat="server" BackColor="#FFC080" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 108; left: 128px;
            position: absolute; top: 184px" Width="77px"></asp:TextBox>
        <asp:TextBox ID="DY2Min" runat="server" BackColor="#FFC080" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 109; left: 132px;
            position: absolute; top: 78px" Width="77px"></asp:TextBox>
        <asp:TextBox ID="DY2Sum" runat="server" BackColor="#FFC080" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 110; left: 132px;
            position: absolute; top: 109px" Width="77px"></asp:TextBox>
        <asp:TextBox ID="DY3Min2" runat="server" BackColor="#FFC080" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 111; left: 128px;
            position: absolute; top: 218px" Width="77px"></asp:TextBox>
        <asp:TextBox ID="DY3Min3" runat="server" BackColor="#FFC080" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 112; left: 128px;
            position: absolute; top: 251px" Width="77px"></asp:TextBox>
        <asp:TextBox ID="DY3Min4" runat="server" BackColor="#FFC080" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 113; left: 128px;
            position: absolute; top: 284px" Width="77px"></asp:TextBox>
        <asp:TextBox ID="DY3Min5" runat="server" BackColor="#FFC080" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 114; left: 128px;
            position: absolute; top: 316px" Width="77px"></asp:TextBox>
        <asp:TextBox ID="DY3Sum" runat="server" BackColor="#FFC080" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 115; left: 128px;
            position: absolute; top: 350px" Width="77px"></asp:TextBox>
        <asp:TextBox ID="DX2Min1" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 116; left: 140px;
            position: absolute; top: 735px" Width="76px"></asp:TextBox>
        <asp:TextBox ID="DX2Min2" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 117; left: 140px;
            position: absolute; top: 765px" Width="76px"></asp:TextBox>
        <asp:TextBox ID="DX2Min3" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 118; left: 140px;
            position: absolute; top: 798px" Width="76px"></asp:TextBox>
        <asp:TextBox ID="DX2Min4" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 119; left: 140px;
            position: absolute; top: 830px" Width="76px"></asp:TextBox>
        <asp:TextBox ID="DX2Min5" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 120; left: 140px;
            position: absolute; top: 864px" Width="76px"></asp:TextBox>
        <asp:TextBox ID="DX2Desc2" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 121; left: 233px;
            position: absolute; top: 765px" Width="485px"></asp:TextBox>
        <asp:TextBox ID="DX2Desc3" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 122; left: 233px;
            position: absolute; top: 798px" Width="485px"></asp:TextBox>
        <asp:TextBox ID="DX2Desc4" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 123; left: 233px;
            position: absolute; top: 830px" Width="485px"></asp:TextBox>
        <asp:TextBox ID="DX2Desc5" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 124; left: 233px;
            position: absolute; top: 864px" Width="485px"></asp:TextBox>
        <asp:TextBox ID="DX2Desc1" runat="server" BackColor="#FFFF80" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 126; left: 233px;
            position: absolute; top: 735px" Width="485px"></asp:TextBox>
        <asp:TextBox ID="DY3Desc2" runat="server" BackColor="#FFC080" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 127; left: 230px;
            position: absolute; top: 218px" Width="488px"></asp:TextBox>
        <asp:TextBox ID="DY3Desc3" runat="server" BackColor="#FFC080" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 128; left: 230px;
            position: absolute; top: 251px" Width="488px"></asp:TextBox>
        <asp:TextBox ID="DY3Desc4" runat="server" BackColor="#FFC080" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 129; left: 230px;
            position: absolute; top: 284px" Width="488px"></asp:TextBox>
        <asp:TextBox ID="DY3Desc5" runat="server" BackColor="#FFC080" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 130; left: 230px;
            position: absolute; top: 316px" Width="488px"></asp:TextBox>
        <asp:TextBox ID="DY2Desc" runat="server" BackColor="#FFC080" BorderStyle="None"
            ForeColor="Black" Height="17px" ReadOnly="True" Style="z-index: 131; left: 231px;
            position: absolute; top: 77px" Width="488px"></asp:TextBox>
                </table>
              </asp:View>
        <!-- View-3 End -->
            </asp:MultiView>
        <!-- MultiView End -->
    </div>
  </form>
</body>
</html>