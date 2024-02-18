<% option Explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp" -->


<%

dim mailto, title, contents
dim orderserial, userid

mailto = request.form("mailto")
title = request.form("title")
contents = html2db(request.form("contents"))
orderserial = request.form("orderserial")
userid = request.form("userid")

call sendmailCS(mailto,title,nl2br(contents))

response.write "<script>alert('메일 발송되었습니다.')</script>"

if (orderserial<>"") or (userid<>"") then
call AddCsMemo(orderserial,"1",userid,session("ssBctId"),"[Mail]" + title + VbCrlf + contents)
response.write "<script>alert('발송내용에 MEMO에 저장되었습니다.')</script>"
end if

response.write "<script>window.close();</script>"
dbget.close()	:	response.End


%>


<!-- #include virtual="/lib/db/dbclose.asp" -->
