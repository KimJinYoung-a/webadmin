<%@ language=vbscript %>
<% option explicit %>
<% 
Response.AddHeader "Cache-Control","no-cache" 
Response.AddHeader "Expires","0" 
Response.AddHeader "Pragma","no-cache" 
%> 
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/noreplyboardcls.asp" -->
<%
dim yyyy1,mm1,dd1
dim sitename,writer,buyname
dim orderserial
dim title,txmemo
dim yyyymmdd

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")

yyyymmdd = yyyy1 + "-" + mm1 + "-" + dd1
sitename = request("sitename")
writer = request("writer")
buyname = html2db(request("buyname"))
orderserial = request("orderserial")
title = html2db(request("title"))
txmemo = html2db(request("txmemo"))

''�ʼ� �Է� üũ.
if (yyyymmdd="") or (sitename="") or _
	(writer="") or (buyname="") then 
		dbget.close()	:	response.End
end if

dim oneboard,errmsg
set oneboard = new CNoReplyBoard
errmsg = oneboard.writeboard(yyyymmdd,sitename,writer,buyname,orderserial,title,txmemo)
set oneboard = Nothing

if errmsg<>"" then
	response.write "�ý��� ���� : " + errmsg
else
	response.write "<script>alert('����Ǿ����ϴ�.')</script>"
	response.write "<script>location.replace('/admin/board/bct_admin_deliver.asp')</script>"
	
end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->