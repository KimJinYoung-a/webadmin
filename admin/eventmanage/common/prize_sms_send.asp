<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/email/smslib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

'// SMS �߼� ������

dim smstext, i, title, k

smstext 		= request("smstxt")
Dim vCnt : vCnt = Request("usercell").count
title = "[�ٹ����� ��÷�ȳ�]"
dim sqlstr
'�迭�� ó��
redim arrreqhp(vCnt)
for i=1 to vCnt
	arrreqhp(i) = Request("usercell")(i)
next
'Response.write Request("usercell") & "<br>"
'Response.write Request("smstxt")
'Response.end

if (vCnt>0 and smstext<>"") then
	For k=1 To vCnt
		if LenB(smstext) > 80 then
			Call SendNormalLMS(arrreqhp(k), title, "", smstext)
		else
			Call SendNormalSMS_LINK(arrreqhp(k), "", smstext)
		end if
	Next
    response.write "<script>alert('���۵Ǿ����ϴ�.');self.close();</script>"
    dbget.close()	:	response.End
end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->