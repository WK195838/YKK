//Wave's Item登錄申請用JavaScript檔

//*********************************************************************
//行事月曆
//*********************************************************************
function calendarPicker(strField)
{
//alert("calendarPicker-in");
    window.open('DatePicker.aspx?field=' + strField,'calendarPopup','width=250,height=190,resizable=yes');
//alert("calendarPicker-out");
}
//*********************************************************************
//單價版本
//*********************************************************************
function PriceVersion(Version)
{
//alert("PriceVersion-in");
    var A206 = document.getElementById('DA206');
    var K206 = document.getElementById('DK206');
    var A001 = document.getElementById('DA001');
    var A999 = document.getElementById('DA999');
    var A211 = document.getElementById('DA211');
    var K211 = document.getElementById('DK211');
     
    A211.checked = true;
    K211.checked = true;
    if (Version == "A001")
    {
        A206.checked = false;
        A999.checked = false;
        K206.checked = false;
    }
    if (Version == "A206")
    {
        A001.checked = false;
        A999.checked = false;
        K206.checked = false;
    }
    if (Version == "A999")
    {
        A001.checked = false;
        A206.checked = false;
        K206.checked = false;
    }
    if (Version == "K206")
    {
        A001.checked = false;
        A206.checked = false;
        A999.checked = false;
    }
//    
//    if (A001.checked == true && A999.checked == true) 
//    {
//        if (Version == "A001") A999.checked = false;
//        if (Version == "A999") A001.checked = false;
//    }
//alert("PriceVersion-out");
}


//*********************************************************************
//單價版本
//*********************************************************************
function ValuationPriceVersion(Version)
{
    var A206 = document.getElementById('DA206');
    var K206 = document.getElementById('DK206');
    var A001 = document.getElementById('DA001');
    var A999 = document.getElementById('DA999');
    var A211 = document.getElementById('DA211');
    var K211 = document.getElementById('DK211');
     
    if (Version == "A211")
    {
        A206.checked = false;
        A999.checked = false;
        K206.checked = false;
        A001.checked = false;
        K211.checked = false;
    }
    if (Version == "K211")
    {
        A206.checked = false;
        A999.checked = false;
        K206.checked = false;
        A001.checked = false;
        A211.checked = false;
    }
    if (Version == "A001")
    {
        A206.checked = false;
        A999.checked = false;
        K206.checked = false;
        A211.checked = false;
        K211.checked = false;
    }
    if (Version == "A206")
    {
        A001.checked = false;
        A999.checked = false;
        K206.checked = false;
        A211.checked = false;
        K211.checked = false;
    }
    if (Version == "A999")
    {
        A001.checked = false;
        A206.checked = false;
        K206.checked = false;
        A211.checked = false;
        K211.checked = false;
    }
    if (Version == "K206")
    {
        A001.checked = false;
        A206.checked = false;
        A999.checked = false;
        A211.checked = false;
        K211.checked = false;
    }
}

//*********************************************************************
//CALL EXCEL
//*********************************************************************
function CheckAttribute(Action)
{
//alert("CheckAttribute-in");
//alert(Action);
    //
    //ITEMSUITABLE
    if (Action == "ITEMSUITABLEExcel")
    {
        //--啟動Excel程式        
        var xCheck = document.getElementById('DSUITABLECHECK');
        var xFilePath = document.getElementById('DITEMSUITABLEFile');
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
            xCheck.value = "OK";
        }
        catch (errorObject)
        {
            xCheck.value = "";
            alert('Error:\n' + errorObject.message);
        }            
    }

//alert("CheckAttribute-out");
}

