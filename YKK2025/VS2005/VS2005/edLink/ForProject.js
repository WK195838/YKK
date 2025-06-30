//
//***********************************************************************************************
//** 功能按鈕點取
//**    設定各元件顯示內容
//** 
//***********************************************************************************************
//
function CheckAttribute(Action)
{
//alert("CheckAttribute-in");
//alert(Action);

    //
    //顧客主檔維護 
    if (Action == "CUSTOMER")
    {
        //--啟動程式        
        var xFilePath = document.getElementById('DFilePath1');
        var executableFullPath = "excel.exe " + xFilePath.value;
        try
        {
            var shellActiveXObject = new ActiveXObject("WScript.Shell");

            if ( !shellActiveXObject )
            {
                alert('Could not get reference to WScript.Shell');
            }
            shellActiveXObject.Run(executableFullPath, 1, false);
            shellActiveXObject = null;
        }
        catch (errorObject)
        {
            alert('Error:\n' + errorObject.message);
        }            
    }
    //
    //OP報表 
    if (Action == "OPREPORT")
    {
        //--啟動程式        
        var xFilePath = document.getElementById('DFilePath3');
        var executableFullPath = "excel.exe " + xFilePath.value;
        try
        {
            var shellActiveXObject = new ActiveXObject("WScript.Shell");

            if ( !shellActiveXObject )
            {
                alert('Could not get reference to WScript.Shell');
            }
            shellActiveXObject.Run(executableFullPath, 1, false);
            shellActiveXObject = null;
        }
        catch (errorObject)
        {
            alert('Error:\n' + errorObject.message);
        }            
    }
    //
    //PI報表(未PACKCP) 
    if (Action == "PIREPORT")
    {
        //--啟動程式        
        var xFilePath = document.getElementById('DFilePath2');
        var executableFullPath = "excel.exe " + xFilePath.value;
        try
        {
            var shellActiveXObject = new ActiveXObject("WScript.Shell");

            if ( !shellActiveXObject )
            {
                alert('Could not get reference to WScript.Shell');
            }
            shellActiveXObject.Run(executableFullPath, 1, false);
            shellActiveXObject = null;
        }
        catch (errorObject)
        {
            alert('Error:\n' + errorObject.message);
        }            
    }
    //
    //Run Program
    if (Action == "Run")
    {
        //--啟動程式        
        var xFilePath = document.getElementById('DFilePath1');
        var executableFullPath = xFilePath.value;

        try
        {
            var shellActiveXObject = new ActiveXObject("WScript.Shell");

            if ( !shellActiveXObject )
            {
                alert('Could not get reference to WScript.Shell');
            }
            shellActiveXObject.Run(executableFullPath, 1, false);
            shellActiveXObject = null;
        }
        catch (errorObject)
        {
            alert('Error:\n' + errorObject.message);
        }            
    }
 }
 //
//***********************************************************************************************
//** 
//**
//***********************************************************************************************
//

       
