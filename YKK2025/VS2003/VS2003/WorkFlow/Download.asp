<%
Response.Buffer = True
Dim strFilePath, strFileSize, strFileName
Const adTypeBinary = 1
strFilePath = Request.QueryString("Path")
strFileName = Request.QueryString("Name")
strFileSize = Request.QueryString("Size")
Response.Clear
Set objStream = Server.CreateObject("ADODB.Stream")
objStream.Open
objStream.Type = adTypeBinary
objStream.LoadFromFile server.mappath(strFilePath & strFileName)
strFileType = lcase(Right(strFileName, 4))
Select Case strFileType
Case ".asf"
ContentType = "video/x-ms-asf"
Case ".avi"
ContentType = "video/avi"
Case ".doc"
ContentType = "application/msword"
Case ".zip"
ContentType = "application/zip"
Case ".xls"
ContentType = "application/vnd.ms-excel"
Case ".gif"
ContentType = "image/gif"
Case ".jpg", "jpeg"
ContentType = "image/jpeg"
Case ".wav"
ContentType = "audio/wav"
Case ".mp3"
ContentType = "audio/mpeg3"
Case ".mpg", "mpeg"
ContentType = "video/mpeg"
Case ".rtf"
ContentType = "application/rtf"
Case ".htm", "html"
ContentType = "text/html"
Case ".asp"
ContentType = "text/asp"
Case Else
ContentType = "application/octet-stream"
End Select
Response.AddHeader "Content-Disposition", "attachment; filename=" & request("Name")
Response.AddHeader "Content-Length", strFileSize
Response.Charset = "UTF-8"
Response.ContentType = ContentType
Response.BinaryWrite objStream.Read
Response.Flush
objStream.Close
Set objStream = Nothing
%>
