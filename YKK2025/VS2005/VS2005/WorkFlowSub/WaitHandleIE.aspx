﻿<%@ Page Language="VB" AutoEventWireup="false" CodeFile="WaitHandleIE.aspx.vb" Inherits="WaitHandleIE" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>未命名頁面</title>
    <script type="text/javascript">
        function StartUpPage(url) {
            var objShell = new ActiveXObject("Wscript.Shell");
            objShell.Run("\"C:\\Program Files\\internet explorer\\iexplore.exe\"" + url);
            objShell = null;

            window.opener=null;
            window.open('','_self');
            window.close();

            return false;
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
    </div>
    </form>
</body>
</html>
