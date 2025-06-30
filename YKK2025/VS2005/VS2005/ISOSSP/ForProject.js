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
    //ADVREPORTExcel
    if (Action == "ADVREPORTExcel")
    {
        //--啟動Excel程式        
        var xFilePath = document.getElementById('DADVREPORTFile');
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
    //ADVSALESExcel
    if (Action == "ADVSALESExcel")
    {
        //--啟動Excel程式        
        var xFilePath = document.getElementById('DADVSALESFile');
        var executableFullPath = "excel.exe " + xFilePath.value;

alert(executableFullPath);

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
    //PULLERLISTExcel
    if (Action == "PULLERLISTExcel")
    {
        //--啟動Excel程式        
        var xFilePath = document.getElementById('DPULLERLISTFile');
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

//alert("CheckAttribute-out");
}

