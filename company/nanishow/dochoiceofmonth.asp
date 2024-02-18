<%@ language=vbscript %>
<% option explicit %>

<%
dim tmp
dim uprequest
dim fs, file1, file1_size,file1_name
dim file2, file2_size,file2_name, file_dest2
dim extension,upfolder,file_dest
dim user_file

dim oldfilename,oldfilename2
dim eventlink, explain
dim referer

referer = request.ServerVariables("HTTP_REFERER")

Set uprequest	= Server.CreateObject("SiteGalaxyUpload.Form")
Set fs		= server.CreateObject("Scripting.FileSystemObject")
file1	= uprequest.item("file1")
file2	= uprequest.item("file2")
oldfilename = uprequest.item("oldfilename")
oldfilename2 = uprequest.item("oldfilename2")
eventlink = uprequest.item("eventlink")
explain  = uprequest.item("explain")

if (file1 <> "") then
        file1_size = CLng(uprequest("file1").size)
        file1_name = fs.GetFileName(uprequest("file1").FilePath)
        extension = LCase(Right(file1_name,3))
        if ((extension <> "gif") and (extension <> "jpg") and (extension <> "bmp")) then
                response.write "<script language='javascript'>alert('이미지(gif,jpg,bmp) 화일만 지원됩니다.'); history.go(-1);</script>"
                dbget.close()	:	response.End
        end if
else
        file1_size = 0
        file1_name = ""
end if

if (file2 <> "") then
        file2_size = CLng(uprequest("file2").size)
        file2_name = fs.GetFileName(uprequest("file2").FilePath)
        extension = LCase(Right(file2_name,3))
        if ((extension <> "gif") and (extension <> "jpg") and (extension <> "bmp")) then
                response.write "<script language='javascript'>alert('이미지(gif,jpg,bmp) 화일만 지원됩니다.'); history.go(-1);</script>"
                dbget.close()	:	response.End
        end if
else
        file2_size = 0
        file2_name = ""
end if

upfolder = "/down"
if (file1_size <> 0) then
		''Server.MapPath(upfolder)
        tmp = "C:\home\cube1010\www\down" + "\NaniChoicemonth" + "-" + Session.SessionID + "-" + file1_name
        uprequest("file1").saveas(tmp)
        file_dest = upfolder + "/NaniChoicemonth" + "-" + Session.SessionID + "-" + file1_name
end if

if (file2_size <> 0) then
		''Server.MapPath(upfolder)
        tmp = "C:\home\cube1010\www\down" + "\NaniEventForyou" + "-" + Session.SessionID + "-" + file2_name
        uprequest("file2").saveas(tmp)
        file_dest2 = upfolder + "/NaniEventForyou" + "-" + Session.SessionID + "-" + file2_name
end if

Set uprequest	=Nothing
Set fs		= Nothing

dim svfilename
dim fso,ofile
dim buf

if file_dest<>"" then
	buf = file_dest + vbcrlf 
else
	buf = oldfilename + vbcrlf 
end if

if file_dest2<>"" then
	buf = buf + file_dest2 + vbcrlf 
else
	buf = buf + oldfilename2 + vbcrlf 
end if

buf = buf + eventlink + vbcrlf + explain


svfilename = "C:/home/cube1010/www" & "/ext/nanishow/choiceofmonth.inc"
set fso = server.CreateObject("scripting.filesystemobject")
set ofile = fso.CreateTextFile(svfilename,True)
	
ofile.Writeline buf
ofile.close
set ofile = nothing
set fso = nothing

response.write "<script>alert('저장되었습니다.');</script>"
response.write "<script>location.replace('" + referer + "');</script>"
	
%>