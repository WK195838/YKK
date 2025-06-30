<%@ Page Language="VB" AutoEventWireup="false" CodeFile="BatchQAApprov.aspx.vb" Inherits="BatchQAApprov" %>

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
        	<FONT face="新細明體"></FONT>&nbsp;&nbsp;
      </div>
    </form>
</body>
</html>