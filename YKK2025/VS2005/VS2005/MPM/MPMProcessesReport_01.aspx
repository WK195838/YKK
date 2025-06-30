<%@ Page Language="vb" AutoEventWireup="false" Inherits="MPMProcessesReport_01" aspCompat="True" CodeFile="MPMProcessesReport_01.aspx.vb" %>

<%@ Register Assembly="System.Web.Extensions, Version=1.0.61025.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"
    Namespace="System.Web.UI" TagPrefix="asp" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<script runat="server">

        Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)

        End Sub
</script>

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>機械加工進度表</title>
    

    	<script language="javascript" type="text/javascript">
			var wPop;
			//--Calendar------------------------------------
			function calendarPicker(field1)
			{
				window.open('DatePicker.aspx?field1=' + field1,'calendarPopup','width=250,height=190,resizable=yes');
			}
			
			    function EmpDatePicker(userid)
{
        
    window.open('DivisionEmpList.aspx?userid='+userid,'','status=0,toolbar=0,width=500,height=500,top=10,resizable=yes');
   
}
		 
		   
		</script>
		
</head>
<body>

  	<form id="Form1" runat="server">
  	 <div>
        
           <asp:Label ID="Label5" runat="server" Style="z-index: 100; left: 388px; position: absolute;
               top: 11px" Text="結案"></asp:Label>
       
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           <asp:DropDownList ID="DOrderby" runat="server" Style="z-index: 101; left: 320px;
               position: absolute; top: 40px" BackColor="Yellow">
               <asp:ListItem>降順</asp:ListItem>
               <asp:ListItem>昇順</asp:ListItem>
           </asp:DropDownList>
           &nbsp; &nbsp;&nbsp;
           <asp:DropDownList ID="DField" runat="server" Style="z-index: 102; left: 224px; position: absolute;
               top: 40px" BackColor="Yellow">
               <asp:ListItem Value="No">加工編號</asp:ListItem>
               <asp:ListItem Value="AppDate">收件日</asp:ListItem>
               <asp:ListItem Value="FinishDate">預定完成日</asp:ListItem>
               <asp:ListItem Value="delayday">遲納天數</asp:ListItem>
           </asp:DropDownList>
           <asp:DropDownList ID="DType1" runat="server" Style="z-index: 103; left: 54px; position: absolute;
               top: 40px" BackColor="Yellow">
               <asp:ListItem Value="全部">全部</asp:ListItem>
               <asp:ListItem Value="金型">金型</asp:ListItem>
               <asp:ListItem Value="部品">部品</asp:ListItem>
           </asp:DropDownList>
           <asp:DropDownList ID="DSdelay" runat="server" Style="z-index: 104; left: 220px; position: absolute;
               top: 9px" BackColor="Yellow">
               <asp:ListItem Value="0">全部</asp:ListItem>
               <asp:ListItem Value="1">正常</asp:ListItem>
               <asp:ListItem Value="遲納">遲納</asp:ListItem>
           </asp:DropDownList><asp:DropDownList ID="DFinish" runat="server" Style="z-index: 105; left: 429px; position: absolute;
               top: 8px" BackColor="Yellow">
               <asp:ListItem>全部</asp:ListItem>
               <asp:ListItem Value="0">已進行</asp:ListItem>
               <asp:ListItem Value="1">已完成</asp:ListItem>
               <asp:ListItem Value="2">取消</asp:ListItem>
           </asp:DropDownList>
           <asp:DropDownList ID="DType2" runat="server" Style="z-index: 106; left: 118px; position: absolute;
               top: 40px" BackColor="Yellow">
               <asp:ListItem Value="全部">全部</asp:ListItem>
               <asp:ListItem Value="製造">製造</asp:ListItem>
               <asp:ListItem Value="修配">修配</asp:ListItem>
           </asp:DropDownList>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp;&nbsp; 
           <asp:Panel ID="Panel2" runat="server" BorderColor="Black" BorderWidth="1px" Height="16px"
               Style="z-index: 107; left: 516px; position: absolute; top: 7px" Width="51px">
               未收件</asp:Panel>
           <asp:Panel ID="Panel1" runat="server" BackColor="LightBlue" BorderColor="Black" BorderWidth="1px"
               Height="16px" Style="z-index: 108; left: 572px; position: absolute; top: 7px"
               Width="51px">
               已收件</asp:Panel>
           <asp:Panel ID="Panel3" runat="server" BackColor="LightGreen" Height="16px" Style="z-index: 109;
               left: 688px; position: absolute; top: 7px" Width="51px" BorderColor="Black" BorderWidth="1px">
               &nbsp; 完成</asp:Panel>
           <asp:Panel ID="Panel5" runat="server" BackColor="Red" Height="16px" Style="z-index: 110;
               left: 746px; position: absolute; top: 7px" Width="51px" BorderColor="Black" BorderWidth="1px">
               &nbsp; 暫停</asp:Panel>
           <asp:Panel ID="Panel4" runat="server" BackColor="Yellow" Height="16px" Style="z-index: 111;
               left: 630px; position: absolute; top: 7px" Width="51px" BorderColor="Black" BorderWidth="1px">
               加工中</asp:Panel>
           &nbsp; &nbsp; &nbsp;&nbsp; 
           <asp:Button ID="Button1" runat="server" Style="z-index: 112; left: 707px; position: absolute;
               top: 40px" Text="Go" Width="40px" />
        
           <asp:HyperLink ID="LAutoReport" runat="server" Font-Size="12pt" Height="8px" NavigateUrl="BoardEdit.aspx"
               Style="z-index: 113; left: 900px; position: absolute; top: 45px" Target="_blank"
               Width="124px">自動更新及跳頁</asp:HyperLink>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; 
           <asp:TextBox ID="DMapNo" runat="server" BackColor="Yellow" Style="left: 424px;
               position: absolute; top: 40px; z-index: 114;" Width="88px" ></asp:TextBox>
           <asp:Label ID="Label7" runat="server" Style="z-index: 115; left: 16px; position: absolute;
               top: 101px" Text="筆數："></asp:Label>
           &nbsp;&nbsp;
           <asp:Label ID="Label9" runat="server" Style="z-index: 127; left: 16px; position: absolute;
               top: 42px" Text="類別"></asp:Label>
           <asp:TextBox ID="DClinter" runat="server" BackColor="Yellow" Style="z-index: 114;
               left: 74px; position: absolute; top: 69px" Width="92px"></asp:TextBox>
           <asp:Label ID="Label10" runat="server" Style="z-index: 127; left: 188px; position: absolute;
               top: 71px" Text="收件日期"></asp:Label>
           &nbsp;
           <asp:TextBox ID="DSdate" runat="server" BackColor="Yellow" Style="z-index: 114; left: 389px;
               position: absolute; top: 67px" Width="88px"></asp:TextBox>
           <asp:TextBox ID="BSdate" runat="server" BackColor="Yellow" Style="z-index: 114;
               left: 264px; position: absolute; top: 67px" Width="88px"></asp:TextBox>
           <asp:Label ID="Label12" runat="server" Style="z-index: 100; left: 525px; position: absolute;
               top: 42px" Text="工程延遲"></asp:Label>
           <asp:DropDownList ID="DDelay" runat="server" Style="z-index: 105; left: 604px; position: absolute;
               top: 39px" BackColor="Yellow" Width="86px">
               <asp:ListItem Value=" ">全部</asp:ListItem>
               <asp:ListItem>延遲</asp:ListItem>
               <asp:ListItem>終止</asp:ListItem>
           </asp:DropDownList>
           <asp:Panel ID="Panel7" runat="server" BackColor="#C0FFFF" ForeColor="White" Height="932px"
               ScrollBars="Both" Style="left: 0px; position: relative; top: 91px; z-index: 118;" Width="1800px">
               <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" BackColor="Silver"
                   BorderColor="Silver" BorderStyle="None" BorderWidth="1px" CellPadding="3" CellSpacing="2"
                   EmptyDataText="無資料" Font-Size="9pt" ForeColor="Black" OnPageIndexChanging="GridView1_PageIndexChanging1"
                   OnRowCreated="GridView1_RowCreated1" OnRowDataBound="GridView1_RowDataBound1"
                   PageSize="8" Style="z-index: 120; left: 11px; position: absolute; top: 30px" Width="1750px">
                   <RowStyle BackColor="White" ForeColor="Black" />
                   <Columns>
                       <asp:BoundField DataField="No" HeaderText="加工編號">
                           <ItemStyle ForeColor="Blue" Width="80px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="MapNo" HeaderText="圖號">
                           <ItemStyle ForeColor="Blue" Width="80px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="AppDate" HeaderText="收件日">
                           <ItemStyle ForeColor="Blue" Width="80px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="FinishDate" HeaderText="預定完成日">
                           <ItemStyle ForeColor="Blue" Width="180px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="delayday" HeaderText="遲納天數" HtmlEncode="False" >
                           <ItemStyle ForeColor="Blue" Width="120px" />
                           <HeaderStyle Width="120px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Clinter" HeaderText="依賴者 ">
                           <ItemStyle ForeColor="Blue" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Division1" HeaderText="部門">
                           <ItemStyle ForeColor="Blue" Width="120px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Type1" HeaderText="類別1">
                           <ItemStyle ForeColor="Blue" Width="120px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Type2" HeaderText="類別2">
                           <ItemStyle ForeColor="Blue" Width="120px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Engine1" HeaderText="工程">
                           <ItemStyle Width="200px" />
                       </asp:BoundField>
                       <asp:BoundField HeaderText="工程" DataField="Engine2" >
                           <ItemStyle Width="200px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Engine3" HeaderText="工程">
                           <ItemStyle Width="200px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Engine4" HeaderText="工程">
                           <ItemStyle Width="200px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Engine5" HeaderText="工程">
                           <ItemStyle Width="200px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Engine6" HeaderText="工程">
                           <ItemStyle Width="200px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Engine7" HeaderText="工程">
                           <ItemStyle Width="200px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Engine8" HeaderText="工程">
                           <ItemStyle Width="200px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Engine9" HeaderText="工程">
                           <ItemStyle Width="200px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Engine10" HeaderText="工程">
                           <ItemStyle Width="200px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Engine11" HeaderText="工程">
                           <ItemStyle Width="200px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Engine12" HeaderText="工程">
                           <ItemStyle Width="200px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Engine13" HeaderText="工程">
                           <ItemStyle Width="200px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="Engine14" HeaderText="工程">
                           <ItemStyle Width="200px" />
                       </asp:BoundField>
                       <asp:BoundField DataField="worktime1" HeaderText="工時" />
                       <asp:BoundField DataField="worktime2" HeaderText="工時" />
                       <asp:BoundField DataField="worktime3" HeaderText="工時" />
                       <asp:BoundField DataField="worktime4" HeaderText="工時" />
                       <asp:BoundField DataField="worktime5" HeaderText="工時" />
                       <asp:BoundField DataField="worktime6" HeaderText="工時" />
                       <asp:BoundField DataField="worktime7" HeaderText="工時" />
                       <asp:BoundField DataField="worktime8" HeaderText="工時" />
                       <asp:BoundField DataField="worktime9" HeaderText="工時" />
                       <asp:BoundField DataField="worktime10" HeaderText="工時" />
                       <asp:BoundField DataField="worktime11" HeaderText="工時" />
                       <asp:BoundField DataField="worktime12" HeaderText="工時" />
                       <asp:BoundField DataField="worktime13" HeaderText="工時" />
                       <asp:BoundField DataField="worktime14" HeaderText="工時" />
                       <asp:BoundField DataField="Delay1" HeaderText="Delay1" />
                       <asp:BoundField DataField="Delay2" HeaderText="Delay2" />
                       <asp:BoundField DataField="Delay3" HeaderText="Delay3" />
                       <asp:BoundField DataField="Delay4" HeaderText="Delay4" />
                       <asp:BoundField DataField="Delay5" HeaderText="Delay5" />
                       <asp:BoundField DataField="Delay6" HeaderText="Delay6" />
                       <asp:BoundField DataField="Delay7" HeaderText="Delay7" />
                       <asp:BoundField DataField="Delay8" HeaderText="Delay8" />
                       <asp:BoundField DataField="Delay9" HeaderText="Delay9" />
                       <asp:BoundField DataField="Delay10" HeaderText="Delay10" />
                       <asp:BoundField DataField="Delay11" HeaderText="Delay11" />
                       <asp:BoundField DataField="Delay12" HeaderText="Delay12" />
                       <asp:BoundField DataField="Delay13" HeaderText="Delay13" />
                       <asp:BoundField DataField="Delay14" HeaderText="Delay14" />
                   </Columns>
                   <FooterStyle BackColor="#F7DFB5" ForeColor="#8C4510" />
                   <PagerStyle ForeColor="Black" HorizontalAlign="Center" />
                   <SelectedRowStyle BackColor="#738A9C" Font-Bold="True" ForeColor="White" />
                   <HeaderStyle BackColor="#A55129" Font-Bold="True" ForeColor="White" HorizontalAlign="Center"
                       VerticalAlign="Middle" />
               </asp:GridView>
           </asp:Panel>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;<asp:DropDownList ID="DEngine" runat="server" Style="z-index: 119; left: 56px;
               position: absolute; top: 8px" BackColor="Yellow" Width="118px">
           </asp:DropDownList>
           &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp;&nbsp;
           <asp:Panel ID="Panel6" runat="server" BackColor="LightGray" Height="16px" Style="z-index: 120;
               left: 805px; position: absolute; top: 7px" Width="51px" BorderColor="Black" BorderWidth="1px">
               &nbsp; 終止</asp:Panel>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;&nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;<asp:ScriptManager ID="ScriptManager1"
               runat="server">
           </asp:ScriptManager>
           <asp:UpdatePanel ID="UpdatePanel1" runat="server">
               <ContentTemplate>
                   &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
               </ContentTemplate>
               <Triggers>
                   <asp:AsyncPostBackTrigger ControlID="Timer1" EventName="Tick" />
               </Triggers>
           </asp:UpdatePanel>
          
           &nbsp;
           <asp:Timer ID="Timer1" runat="server" OnTick="Timer1_Tick" Interval="600000">
                   </asp:Timer>
      
           <asp:Label ID="Label1" runat="server" Style="z-index: 121; left: 188px; position: absolute;
               top: 41px" Text="排序"></asp:Label>
           <asp:Label ID="Label8" runat="server" Style="z-index: 122; left: 189px; position: absolute;
               top: 10px" Text="狀態"></asp:Label>
           <asp:Label ID="Label6" runat="server" Style="z-index: 123; left: 384px; position: absolute;
               top: 40px" Text="圖號"></asp:Label>
           <asp:Label ID="Label3" runat="server" Style="z-index: 124; left: 278px; position: absolute;
               top: 10px" Text="進度"></asp:Label>
           <asp:DropDownList ID="DSts" runat="server" BackColor="Yellow" Style="z-index: 125;
               left: 317px; position: absolute; top: 8px">
               <asp:ListItem Value="全部">全部</asp:ListItem>
               <asp:ListItem Value="O">未收件</asp:ListItem>
               <asp:ListItem Value="R">已收件</asp:ListItem>
               <asp:ListItem Value="S">加工中</asp:ListItem>
               <asp:ListItem Value="E">完成</asp:ListItem>
               <asp:ListItem Value="X">暫停</asp:ListItem>
               <asp:ListItem Value="A">終止</asp:ListItem>
           </asp:DropDownList>
           <asp:Label ID="Label4" runat="server" Height="19px" Style="z-index: 126; left: 16px;
               position: absolute; top: 9px" Text="工程"></asp:Label>
           <asp:Label ID="Label2" runat="server" Style="z-index: 127; left: 14px; position: absolute;
               top: 70px" Text="依賴者"></asp:Label>
           <asp:Label ID="Label11" runat="server" Style="z-index: 127; left: 360px; position: absolute;
               top: 69px" Text="~"></asp:Label>
           &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp;
           &nbsp;&nbsp;&nbsp;<br />
        <br />
              <asp:ImageButton ID="BExcel" runat="server" Height="21px" ImageUrl="Images\msexcel.gif"
            Style="z-index: 128; left: 763px; position: absolute; top: 40px" Width="24px" />     &nbsp; &nbsp;
           &nbsp;
           &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
           &nbsp;&nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;<br />
        <br />
           &nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp;<br />
           &nbsp; &nbsp;<br />
        <br />
        <br />
        <br />
           &nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br />
        <br />
        <br />
        <br />
           &nbsp;
           <br />
           &nbsp; &nbsp;&nbsp;&nbsp;
           <br />
           &nbsp;
               
          

        </div>
    </form>
</body>
</html>
