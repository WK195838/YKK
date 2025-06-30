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
    if (Action == "NEWREPORT")
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

 }
 //
//***********************************************************************************************
//** 
//**
//***********************************************************************************************
//

       
