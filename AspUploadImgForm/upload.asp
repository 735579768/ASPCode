<% if session("APP")<>"true" then reurl("/") end if %>
<!--#include file="UpLoad_Class.asp"-->
<%
dim upload
set upload = new AnUpLoad
upload.Exe = "*"
upload.MaxSize = 2 * 1024 * 1024 '2M
upload.GetData()
if upload.ErrorID>0 then 
	response.Write upload.Description
else
	dim file,savpath
	'当天上传文件夹
	dim todayfolder:todayfolder=Year(Now) & Right("0"& Month(Now),2) & Right("0"& Day(Now),2)
	savepath = "/uploads/"&todayfolder&"/"
	for each frm in upload.forms("-1")
		response.Write frm & "=" & upload.forms(frm) & "<br />"
	next
	set file = upload.files("file1")
	if not(file is nothing) then
		result = file.saveToFile(savepath,0,true)
		if result then
			'response.Write "文件'" & file.LocalName & "'上传成功，保存位置'" & server.MapPath(savepath & "/" & file.filename) & "',文件大小" & file.size & "字节"
			response.Write "<script type=""text/javascript"">opener.document.getElementById('pic').value='"&savepath& file.filename&"';opener.cha();window.close();</script>"
		else
			response.Write file.Exception
		end if
	end if
end if
set upload = nothing
%>