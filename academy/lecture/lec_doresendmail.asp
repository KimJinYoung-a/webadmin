<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/classes/requestlecturecls.asp"-->
<!-- #include virtual="/academy/lib/classes/fingers_lecturecls.asp"-->
<!-- #include virtual="/academy/lib/email/maillib.asp"-->
<%

dim oordermaster, oorderdetail

dim orderserial
dim mode, i, j, k, tmp
dim sqlStr

mode = RequestCheckvar(request("mode"),16)

orderserial     = RequestCheckvar((request("orderserial"),16)

call ReSendmailLectureOrder(orderserial, "customer@10x10.co.kr")

response.write "<script>alert('주문메일이 재발송 되었습니다.');</script>"
response.write "<script>opener.focus(); window.close();</script>"
dbget.close()	:	response.End

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->