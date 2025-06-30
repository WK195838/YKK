//-----------------------------------------------------------------------------------------------
//
//  JavaScript Program 
//
//-----------------------------------------------------------------------------------------------
//
//***********************************************************************************************
//  開啟日期選擇  
//  
//***********************************************************************************************
function OpenDatePicker(xDepo, xField)
{
    window.open('DatePicker.aspx?field1='+xField+'&pDepo='+xDepo,'','status=0,toolbar=0,width=300,height=200');
}
//***********************************************************************************************
//  閱讀核定履歷  
//  
//***********************************************************************************************
function ReadHistory(xType)
{
    var wReadHistory = document.getElementById('DReadHistory');
    
    if (!wReadHistory.checked)  {
        wReadHistory.checked = true;
    }
    else    {
        wReadHistory.checked = false;
    }     
}
//***********************************************************************************************
//  回試作工程CheckBox選擇  
//  
//***********************************************************************************************
function GoBackOP(xOP)
{
    if (document.getElementById('DOP1PER').value!="")  document.getElementById('DOP40').checked = false;
    if (document.getElementById('DOP2PER').value!="")  document.getElementById('DOP50').checked = false;
    if (document.getElementById('DOP3PER').value!="")  document.getElementById('DOP60').checked = false;
    if (document.getElementById('DOP4PER').value!="")  document.getElementById('DOP70').checked = false;
    if (document.getElementById('DOP5PER').value!="")  document.getElementById('DOP80').checked = false;
    if (document.getElementById('DOP6PER').value!="")  document.getElementById('DOP90').checked = false;
    if (document.getElementById('DOP7PER').value!="")  document.getElementById('DOP100').checked = false;
    if (document.getElementById('DOP8PER').value!="")  document.getElementById('DOP110').checked = false;
   
    if (xOP=="40")   document.getElementById('DOP40').checked = true;
    if (xOP=="50")   document.getElementById('DOP50').checked = true;
    if (xOP=="60")   document.getElementById('DOP60').checked = true;
    if (xOP=="70")   document.getElementById('DOP70').checked = true;
    if (xOP=="80")   document.getElementById('DOP80').checked = true;
    if (xOP=="90")   document.getElementById('DOP90').checked = true;
    if (xOP=="100")  document.getElementById('DOP100').checked = true;
    if (xOP=="110")  document.getElementById('DOP110').checked = true;
}
//***********************************************************************************************
//  設定鏈齒/丸扭屬行  
//  
//***********************************************************************************************
function SetInputDataAttr(xType)
{
    //鏈齒
    if (xType == "ECOL")  {
        if (document.getElementById('DECOLSEL').value!="其他")  {
            document.getElementById('DECOL').value = "";
            alert('不可輸入鏈齒顏色!');
        }
    }
    //丸扭
    if (xType == "CCOL")  {
        if (document.getElementById('DCCOLSEL').value!="MF昭安")  {
            document.getElementById('DCCOL').value = "";
            alert('不可輸入丸扭顏色!');
        }
    }
}
//***********************************************************************************************
//  已開發-需樣品 / 需登錄CheckBox選擇  
//  
//***********************************************************************************************
function NeedSample(xSAMPLE)
{
    if (document.getElementById('DModify').checked==true) {
        //----10工程可以輸入
        if (xSAMPLE=="SAMPLE")  {
            //document.getElementById('DNeedSample').checked = true;
            document.getElementById('DNeedItemRegister').checked = false;
        }

        if (xSAMPLE=="ITEM")  {
            document.getElementById('DNeedSample').checked = false;
            //document.getElementById('DNeedItemRegister').checked = true;
        }
        //----設定開發規格欄位屬性
        if (document.getElementById('DNeedSample').checked == true || document.getElementById('DNeedItemRegister').checked == true)   {
            //----ReadOnly
            //alert('enabled=false');
            document.getElementById('DSIZENO').disabled = true;
            document.getElementById('DITEM').disabled = true;
            document.getElementById('DTATYPE').disabled = true;
            document.getElementById('DTAWIDTH').disabled = true;        
            document.getElementById('DECOLSEL').disabled = true;       

            document.getElementById('DECOL').disabled = true;        
            document.getElementById('DCCOLSEL').disabled = true;        
            document.getElementById('DCCOL').disabled = true;        
            document.getElementById('DTACOL').disabled = true;        
            document.getElementById('DTACOLNO').disabled = true;        

            document.getElementById('DTAYCOLNO').disabled = true;        
            document.getElementById('DTALCOL').disabled = true;        
            document.getElementById('DTALCOLNO').disabled = true;        
            document.getElementById('DTALYCOLNO').disabled = true;        
            document.getElementById('DTARCOL').disabled = true;        

            document.getElementById('DTARCOLNO').disabled = true;        
            document.getElementById('DTARYCOLNO').disabled = true;        
            document.getElementById('DTHUPCOL').disabled = true;        
            document.getElementById('DTHUPCOLNO').disabled = true;        
            document.getElementById('DTHUPYCOLNO').disabled = true;        

            document.getElementById('DTHLUPCOL').disabled = true;        
            document.getElementById('DTHLUPCOLNO').disabled = true;        
            document.getElementById('DTHLUPYCOLNO').disabled = true;        
            document.getElementById('DTHRUPCOL').disabled = true;        
            document.getElementById('DTHRUPCOLNO').disabled = true;        

            document.getElementById('DTHRUPYCOLNO').disabled = true;        
            document.getElementById('DTHLOCOL').disabled = true;        
            document.getElementById('DTHLOCOLNO').disabled = true;        
            document.getElementById('DTHLOYCOLNO').disabled = true;        
            document.getElementById('DTHLLOCOL').disabled = true;        

            document.getElementById('DTHLLOCOLNO').disabled = true;        
            document.getElementById('DTHLLOYCOLNO').disabled = true;        
            document.getElementById('DTHRLOCOL').disabled = true;        
            document.getElementById('DTHRLOCOLNO').disabled = true;        
            document.getElementById('DTHRLOYCOLNO').disabled = true;        

            document.getElementById('DXMLEN').disabled = true;        
            document.getElementById('DXMCOL').disabled = true;        
            document.getElementById('DXMCOLNO').disabled = true;        
            document.getElementById('DXMYCOLNO').disabled = true;  
      
            document.getElementById('DAMLEN').disabled = true;        
            document.getElementById('DAMCOL').disabled = true;        
            document.getElementById('DAMCOLNO').disabled = true;        
            document.getElementById('DAMYCOLNO').disabled = true;        

            document.getElementById('DBMLEN').disabled = true;        
            document.getElementById('DBMCOL').disabled = true;        
            document.getElementById('DBMCOLNO').disabled = true;        
            document.getElementById('DBMYCOLNO').disabled = true;        

            document.getElementById('DCMLEN').disabled = true;        
            document.getElementById('DCMCOL').disabled = true;        
            document.getElementById('DCMCOLNO').disabled = true;        
            document.getElementById('DCMYCOLNO').disabled = true;        

            document.getElementById('DDMLEN').disabled = true;        
            document.getElementById('DDMCOL').disabled = true;        
            document.getElementById('DDMCOLNO').disabled = true;        
            document.getElementById('DDMYCOLNO').disabled = true;        

            document.getElementById('DEMLEN').disabled = true;        
            document.getElementById('DEMCOL').disabled = true;        
            document.getElementById('DEMCOLNO').disabled = true;        
            document.getElementById('DEMYCOLNO').disabled = true;        

            document.getElementById('DFMLEN').disabled = true;        
            document.getElementById('DFMCOL').disabled = true;        
            document.getElementById('DFMCOLNO').disabled = true;        
            document.getElementById('DFMYCOLNO').disabled = true;        

            document.getElementById('DGMLEN').disabled = true;        
            document.getElementById('DGMCOL').disabled = true;        
            document.getElementById('DGMCOLNO').disabled = true;        
            document.getElementById('DGMYCOLNO').disabled = true;        

            document.getElementById('DHMLEN').disabled = true;        
            document.getElementById('DHMCOL').disabled = true;        
            document.getElementById('DHMCOLNO').disabled = true;        
            document.getElementById('DHMYCOLNO').disabled = true;        

            document.getElementById('DLYLEN').disabled = true;        
            document.getElementById('DLYCOL').disabled = true;        
            document.getElementById('DLYCOLNO').disabled = true;        
            document.getElementById('DLYYCOLNO').disabled = true;        

            document.getElementById('DOTCON').disabled = true;       
        }
        else    {
            //----Edit
            //alert('enabled=true');
            document.getElementById('DSIZENO').disabled = false;
            document.getElementById('DITEM').disabled = false;
            document.getElementById('DTATYPE').disabled = false;
            document.getElementById('DTAWIDTH').disabled = false;        
            document.getElementById('DECOLSEL').disabled = false;       

            document.getElementById('DECOL').disabled = false;        
            document.getElementById('DCCOLSEL').disabled = false;        
            document.getElementById('DCCOL').disabled = false;        
            document.getElementById('DTACOL').disabled = false;        
            document.getElementById('DTACOLNO').disabled = false;        

            document.getElementById('DTAYCOLNO').disabled = false;        
            document.getElementById('DTALCOL').disabled = false;        
            document.getElementById('DTALCOLNO').disabled = false;        
            document.getElementById('DTALYCOLNO').disabled = false;        
            document.getElementById('DTARCOL').disabled = false;        

            document.getElementById('DTARCOLNO').disabled = false;        
            document.getElementById('DTARYCOLNO').disabled = false;        
            document.getElementById('DTHUPCOL').disabled = false;        
            document.getElementById('DTHUPCOLNO').disabled = false;        
            document.getElementById('DTHUPYCOLNO').disabled = false;        

            document.getElementById('DTHLUPCOL').disabled = false;        
            document.getElementById('DTHLUPCOLNO').disabled = false;        
            document.getElementById('DTHLUPYCOLNO').disabled = false;        
            document.getElementById('DTHRUPCOL').disabled = false;        
            document.getElementById('DTHRUPCOLNO').disabled = false;        

            document.getElementById('DTHRUPYCOLNO').disabled = false;        
            document.getElementById('DTHLOCOL').disabled = false;        
            document.getElementById('DTHLOCOLNO').disabled = false;        
            document.getElementById('DTHLOYCOLNO').disabled = false;        
            document.getElementById('DTHLLOCOL').disabled = false;        

            document.getElementById('DTHLLOCOLNO').disabled = false;        
            document.getElementById('DTHLLOYCOLNO').disabled = false;        
            document.getElementById('DTHRLOCOL').disabled = false;        
            document.getElementById('DTHRLOCOLNO').disabled = false;        
            document.getElementById('DTHRLOYCOLNO').disabled = false;        

            document.getElementById('DXMLEN').disabled = false;        
            document.getElementById('DXMCOL').disabled = false;        
            document.getElementById('DXMCOLNO').disabled = false;        
            document.getElementById('DXMYCOLNO').disabled = false;  
      
            document.getElementById('DAMLEN').disabled = false;        
            document.getElementById('DAMCOL').disabled = false;        
            document.getElementById('DAMCOLNO').disabled = false;        
            document.getElementById('DAMYCOLNO').disabled = false;        

            document.getElementById('DBMLEN').disabled = false;        
            document.getElementById('DBMCOL').disabled = false;        
            document.getElementById('DBMCOLNO').disabled = false;        
            document.getElementById('DBMYCOLNO').disabled = false;        

            document.getElementById('DCMLEN').disabled = false;        
            document.getElementById('DCMCOL').disabled = false;        
            document.getElementById('DCMCOLNO').disabled = false;        
            document.getElementById('DCMYCOLNO').disabled = false;        

            document.getElementById('DDMLEN').disabled = false;        
            document.getElementById('DDMCOL').disabled = false;        
            document.getElementById('DDMCOLNO').disabled = false;        
            document.getElementById('DDMYCOLNO').disabled = false;        

            document.getElementById('DEMLEN').disabled = false;        
            document.getElementById('DEMCOL').disabled = false;        
            document.getElementById('DEMCOLNO').disabled = false;        
            document.getElementById('DEMYCOLNO').disabled = false;        

            document.getElementById('DFMLEN').disabled = false;        
            document.getElementById('DFMCOL').disabled = false;        
            document.getElementById('DFMCOLNO').disabled = false;        
            document.getElementById('DFMYCOLNO').disabled = false;        

            document.getElementById('DGMLEN').disabled = false;        
            document.getElementById('DGMCOL').disabled = false;        
            document.getElementById('DGMCOLNO').disabled = false;        
            document.getElementById('DGMYCOLNO').disabled = false;        

            document.getElementById('DHMLEN').disabled = false;        
            document.getElementById('DHMCOL').disabled = false;        
            document.getElementById('DHMCOLNO').disabled = false;        
            document.getElementById('DHMYCOLNO').disabled = false;        

            document.getElementById('DLYLEN').disabled = false;        
            document.getElementById('DLYCOL').disabled = false;        
            document.getElementById('DLYCOLNO').disabled = false;        
            document.getElementById('DLYYCOLNO').disabled = false;        

            document.getElementById('DOTCON').disabled = false;       
        }     
    }
    else    {
        //----其他工程不可以輸入
        if (xSAMPLE=="SAMPLE")  {
            document.getElementById('DNeedSample').checked = true;
            document.getElementById('DNeedItemRegister').checked = false;
        }

        if (xSAMPLE=="ITEM")  {
            document.getElementById('DNeedSample').checked = false;
            document.getElementById('DNeedItemRegister').checked = true;
        }
    }
    
}
//
//***********************************************************************************************
//  製作開發樣品  
//  
//***********************************************************************************************
function CreateSampleSheet()
{
//alert('out');
}
//
//-----------------------------------------------------------------------------------------------
//
//  End
//
//-----------------------------------------------------------------------------------------------

 