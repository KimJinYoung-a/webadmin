<% option Explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_LogisticsOpen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp" -->
<%

dim mode
dim buyhp, title, message, orderserial
dim sqlStr

mode  		= requestCheckvar(request("mode"),32)
buyhp  		= requestCheckvar(request("buyhp"),32)
title  		= requestCheckvar(request("title"),100)
message  	= html2db(requestCheckvar(request("message"),1000))
orderserial	= requestCheckvar(request("orderserial"),100)


if (mode = "sendlms") then
    sqlStr = " INSERT INTO [db_kakaoSMS].[dbo].MMS_MSG ( REQDATE, STATUS, TYPE, PHONE, CALLBACK, SUBJECT, MSG, FILE_CNT ) "
    sqlStr = sqlStr & " select getdate(), '1' , '0', '" & buyhp & "', '1644-6030', '" & title & "', '" & Replace(message, "'", "") & "', '1' "
    ''response.write sqlStr
    dbget_Logistics.Execute sqlStr

    call AddCsMemo(orderserial,"1", "",session("ssBctId"),"[LMS]" + title + VbCrlf + message)
    response.write "<script>alert('발송내용에 MEMO에 저장되었습니다.')</script>"

    response.write "<script>window.close();</script>"
    dbget_Logistics.close()	:	response.End
end if

%>

잘못된 접근입니다.

<!-- #include virtual="/lib/db/db_LogisticsClose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
