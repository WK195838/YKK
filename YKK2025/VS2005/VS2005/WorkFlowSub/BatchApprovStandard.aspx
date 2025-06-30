<%@ Page Language="VB" AutoEventWireup="false" CodeFile="BatchApprovStandard.aspx.vb" Inherits="BatchApprovStandard" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>標準批簽</title>
 <script type="text/javascript">  
function SelectAllCheckboxesYES(spanChk)  
{  
    
     elm=document.forms[0];  
  
    for(i=0;i<elm.length;i++)  
        {          
            
         if(elm[i].type=="checkbox" && elm[i].id!=spanChk.id)  
            {  
         
            if(elm.elements[i].id.match('NODETAIL')!=null)
              {
              
               if(elm.elements[i].checked==spanChk.checked)  
                 elm.elements[i].click();                                       
              
              }           
              
                       
              if(elm.elements[i].id.match('YESDETAIL')!=null)
              {
              
               if(elm.elements[i].checked!=spanChk.checked)  
                 elm.elements[i].click();                                       
              
              }     
              
         
             if(elm.elements[i].id.match('CheckNOHEAD')!=null)
              {
              
               if(elm.elements[i].checked==spanChk.checked)  
                 elm.elements[i].click();                                       
              
              }  
              
       
                       
          
                
           
                      
            }  
      

    }  
}  

function SelectAllCheckboxesNO(spanChk)  
{  


    elm=document.forms[0]; 
  
    for(i=0;i<elm.length;i++)  
        {         
        
          
       
                
         if(elm[i].type=="checkbox" && elm[i].id!=spanChk.id)           
            {  
            
             
             
            if(elm.elements[i].id.match('YESDETAIL')!=null)
              {
              
               if(elm.elements[i].checked==spanChk.checked)  
                 elm.elements[i].click();                                       
              
              }                 
                
                
              if(elm.elements[i].id.match('NODETAIL')!=null)
              {
               if(elm.elements[i].checked!=spanChk.checked)               
                 elm.elements[i].click();                     
                      
              
              }       
             
                
                 if(elm.elements[i].id.match('CheckYESHEAD')!=null)
              {
              
               if(elm.elements[i].checked==spanChk.checked)  
                 elm.elements[i].click();                                       
              
              }  
               
                
           
                      
            }  
      

    }  
}  


function CheckboxesYES(spanChk)  
{
   elm=document.forms[0]; 
  
    for(i=0;i<elm.length;i++)  
        {         
                
    
                
         if(elm[i].type=="checkbox" && elm[i].id!=spanChk.id)           
            {  
            
             
              if(elm.elements[i].id.match('NODETAIL')!=null)
              {
               if(elm.elements[i].checked==spanChk.checked)               
                 elm.elements[i].click();                     
                     
              
              }       
           
                
                       
                      
            }  
      

    }  
}

 function CheckboxesYES(spanChk)  
{
  
    var row = spanChk.id.split('_')[0] + '_' +spanChk.id.split('_')[1]; // 取得名稱的前兩段為識別列  
    var YES =document.getElementById(row + '_' +'CheckYESDETAIL');
    var NO =document.getElementById(row + '_' +'CheckNODETAIL');
    var Text =document.getElementById(row + '_' +'TextBox1');
  
 
    if (this.checked)  then 
    {
       NO.checked = false;
       Text.value = 'OK.';
   
    }
    
 }

 
 function CheckboxesNO(spanChk)  
{
 
   var row = spanChk.id.split('_')[0] + '_' +spanChk.id.split('_')[1]; // 取得名稱的前兩段為識別列  
    var YES =document.getElementById(row + '_' +'CheckYESDETAIL');
    var NO =document.getElementById(row + '_' +'CheckNODETAIL');
    var Text =document.getElementById(row + '_' +'TextBox1');
    if (this.checked)  then 
    {
       YES.checked = false;
       Text.value = 'NG.'; 
  
    }
    
    
}


</script>  
		
</head>
<body>

  	<form id="Form1" runat="server">
  	 <div>
        	<FONT face="新細明體"></FONT>
           <asp:Button ID="BOK" runat="server" Text="簽核" style="z-index: 100; left: 20px; position: absolute; top: 9px" Font-Size="Medium" Width="120px" onclientclick="javascript:return confirm('確定執行？');" onclick="Button1_Click" />

<asp:Button ID="BRefresh" runat="server" Text="重新整理" style="z-index: 102; left: 168px; position: absolute; top: 9px" Font-Size="Medium" Width="120px" />
<asp:Label ID="Label1" runat="server" ForeColor="Red" Style="z-index: 104; left: 296px; position: absolute; top: 16px" Text="單獨簽核完成，請按重新整理"></asp:Label>
               
           <asp:GridView ID="GridView1" runat="server" Style="z-index: 101; left: 27px; position: absolute;
               top: 47px" AutoGenerateColumns="False" DataKeyNames="FormSno" ShowFooter="True" BackColor="White" BorderColor="#3366CC" BorderStyle="Solid" BorderWidth="1px" CellPadding="4" Width="1500px" Font-Size="Small">
               <Columns>

                   <asp:TemplateField HeaderText="OK">
                       <HeaderTemplate>
                           OK
                             <asp:CheckBox ID="CheckYESHEAD" runat="server" onclick= "SelectAllCheckboxesYES(this);"  ToolTip="按一次全選，再按一次取消全選" />
                           / NG<asp:CheckBox ID="CheckNOHEAD" runat="server" onclick= "SelectAllCheckboxesNO(this);"  ToolTip="按一次全選，再按一次取消全選" />
                       </HeaderTemplate>
                       <ItemTemplate>
                           OK
                            <asp:CheckBox ID="CheckYESDETAIL" runat="server" onclick= "CheckboxesYES(this);"    />
                           / NG<asp:CheckBox ID="CheckNODETAIL" runat="server" onclick= "CheckboxesNO(this);"   />

                       </ItemTemplate>
                       <ItemStyle Width="140px" />
                   </asp:TemplateField>

                    <asp:HyperLinkField DataNavigateUrlFields="URL" DataNavigateUrlFormatString="{0}"
                        DataTextField="A1" HeaderText="A1" Target="_blank">
                        <FooterStyle HorizontalAlign="left" />
                        <ItemStyle HorizontalAlign="left" />
                    </asp:HyperLinkField>

                   <asp:BoundField DataField="B1" HeaderText="B1" />
                   <asp:BoundField DataField="C1" HeaderText="C1" />
                   <asp:BoundField DataField="D1" HeaderText="D1" />
                   <asp:BoundField DataField="E1" HeaderText="E1" />
                   <asp:BoundField DataField="F1" HeaderText="F1" />
                   <asp:BoundField DataField="G1" HeaderText="G1" />
                   <asp:BoundField DataField="H1" HeaderText="H1" />
                   <asp:BoundField DataField="I1" HeaderText="I1" />
                   <asp:BoundField DataField="J1" HeaderText="J1" />
                   <asp:BoundField DataField="K1" HeaderText="K1" />
                   <asp:BoundField DataField="L1" HeaderText="L1" />
                   <asp:BoundField DataField="M1" HeaderText="M1" />
                   <asp:BoundField DataField="N1" HeaderText="N1" />
                   <asp:BoundField DataField="O1" HeaderText="O1" />
                   <asp:BoundField DataField="P1" HeaderText="P1" />
                   <asp:BoundField DataField="Q1" HeaderText="Q1" />
                   <asp:BoundField DataField="R1" HeaderText="R1" />
                   <asp:BoundField DataField="S1" HeaderText="S1" />
                   <asp:BoundField DataField="T1" HeaderText="T1" />
                   <asp:BoundField DataField="U1" HeaderText="U1" />
                   <asp:BoundField DataField="V1" HeaderText="V1" />
                   <asp:BoundField DataField="W1" HeaderText="W1" />
                   <asp:BoundField DataField="X1" HeaderText="X1" />
                   <asp:BoundField DataField="Y1" HeaderText="Y1" />
                   <asp:BoundField DataField="Z1" HeaderText="Z1" />
                   <asp:TemplateField HeaderText="簽核說明">
                       <ItemTemplate>
                           <asp:TextBox ID="TextBox1" runat="server" TextMode="MultiLine" BackColor="GreenYellow" Width="108px"></asp:TextBox>
                       </ItemTemplate>
                       <ItemStyle Width="100px" />
                   </asp:TemplateField>

                    <asp:HyperLinkField DataNavigateUrlFields="OPURL" DataNavigateUrlFormatString="{0}"
                        DataTextField="History" HeaderText="履歷" Target="_blank">
                        <FooterStyle HorizontalAlign="Center" />
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:HyperLinkField>

                   <asp:BoundField DataField="FormNo" HeaderText="FormNo29" />
                   <asp:BoundField DataField="FormSno" HeaderText="FormSno30" />
                   <asp:BoundField DataField="Step" HeaderText="Step31" />
                   <asp:BoundField DataField="Seqno" HeaderText="Seqno32" />
                   <asp:BoundField DataField="Applyid" HeaderText="Applyid33" />
                   <asp:BoundField DataField="Decideid" HeaderText="Decideid34" />
                   <asp:BoundField DataField="TableName" HeaderText="TableName35" />
               </Columns>
               <RowStyle BackColor="White" ForeColor="#003399" />
               <FooterStyle BackColor="#99CCCC" ForeColor="#003399" />
               <PagerStyle BackColor="#99CCCC" ForeColor="#003399" HorizontalAlign="Left" />
               <SelectedRowStyle BackColor="#009999" Font-Bold="True" ForeColor="#CCFF99" />
               <HeaderStyle BackColor="#003399" Font-Bold="True" ForeColor="White" />
           </asp:GridView>
      </div>
    </form>
</body>
</html>