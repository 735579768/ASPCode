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
	'�����ϴ��ļ���
	dim todayfolder:todayfolder=Year(Now) & Right("0"& Month(Now),2) & Right("0"& Day(Now),2)
	savepath = "/uploads/"&todayfolder&"/"
	for each frm in upload.forms("-1")
		response.Write frm & "=" & upload.forms(frm) & "<br />"
	next
	set file = upload.files("file1")
	if not(file is nothing) then
		result = file.saveToFile(savepath,0,true)
		if result then
			'response.Write "�ļ�'" & file.LocalName & "'�ϴ��ɹ�������λ��'" & server.MapPath(savepath & "/" & file.filename) & "',�ļ���С" & file.size & "�ֽ�"
			response.Write "<script type=""text/javascript"">opener.document.getElementById('pic').value='"&savepath& file.filename&"';opener.cha();window.close();</script>"
		else
			response.Write file.Exception
		end if
	end if
end if
set upload = nothing
%>