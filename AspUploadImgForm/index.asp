<% if session("APP")<>"true" then reurl("/") end if %>
<%
if isempty(Session("admin_id")) then response.Write "<script>window.close();</script>":response.End() end if

if request("act")="delimg" then
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	if  fso.fileExists(Server.mappath(request("imgpath"))) then
		fso.DeleteFile(Server.mappath(request("imgpath")))
		set fso=nothing
		response.Write 1
	else
	response.Write 0
	end if
response.End()
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>上传文件</title>
</head>

<body>
<form action="upload.asp" method="post" enctype="multipart/form-data">
选择文件 <input type="file" name="file1" /> <input type="submit" value="上传" />
</form>
</body>
</html>
