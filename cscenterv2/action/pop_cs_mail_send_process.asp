<% option Explicit %>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/cscenterv2/lib/csAsfunction.asp" -->


<%

dim mailto, title, contents
dim orderserial, userid

mailto = request.form("mailto")
title = request.form("title")
contents = html2db(request.form("contents"))
orderserial = RequestCheckvar(request.form("orderserial"),16)
userid = RequestCheckvar(request.form("userid"),32)
if mailto <> "" then
	if checkNotValidHTML(mailto) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	response.End
	end if
end If
if title <> "" then
	if checkNotValidHTML(title) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	response.End
	end if
end If
if contents <> "" then
	if checkNotValidHTML(contents) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	response.End
	end if
end if
call sendmailFingersCS(mailto,title,nl2br(contents))

response.write "<script>alert('���� �߼۵Ǿ����ϴ�.')</script>"

if (orderserial<>"") or (userid<>"") then
call AddCsMemo(orderserial,"1",userid,session("ssBctId"),"[Mail]" + title + VbCrlf + contents)
response.write "<script>alert('�߼۳��뿡 MEMO�� ����Ǿ����ϴ�.')</script>"
end if

response.write "<script>window.close();</script>"
dbget.close()	:	response.End


%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
