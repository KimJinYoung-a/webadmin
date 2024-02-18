<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'#######################################################
'	History	:  2015.11.09 한용민 생성
'	Description : 포장 서비스
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/pack_cls.asp"-->

<%
dim mode ,title, message, sqlStr, i, orderserial, midx
	mode = requestCheckVar(request("mode"),16)
    title = request("title")
    message = request("message")
	orderserial = requestcheckvar(request("orderserial"),11)
	midx = requestCheckVar(request("midx"),10)

dim refip
	refip = request.ServerVariables("HTTP_REFERER")

if (InStr(refip,"10x10.co.kr")<1) then
	response.write "<script type='text/javascript'>alert('정상적인 유입 경로가 아닙니다.');</script>"
	dbget.close()	:	response.end
end if

'//선물포장 수정
if mode="editpojang" then
	if midx="" then
		response.write "<script type='text/javascript'>alert('일렬번호가 없습니다.'); location.replace('"& refip &"');</script>"
		dbget.close()	:	response.end
	end if
	midx = trim(midx)

	if title<>"" then
		if checkNotValidHTML(title) then
			response.write "<script type='text/javascript'>alert('선물포장명에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요.'); location.replace('"& refip &"');</script>"
			dbget.close()	:	response.end
		end if
	end if
	if message<>"" then
		if checkNotValidHTML(message) then
			response.write "<script type='text/javascript'>alert('선물메세지에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요.'); location.replace('"& refip &"');</script>"
			dbget.close()	:	response.end
		end if
	end if

	'//마스터 테이블 저장
    sqlStr = "update db_order.dbo.tbl_order_pack_master" + vbcrlf
    sqlStr = sqlStr & " set title='"& html2db(title) &"'" + vbcrlf
    sqlStr = sqlStr & " , message='"& html2db(message) &"' where" + vbcrlf
    sqlStr = sqlStr & " midx="& midx &""

	'response.write sqlStr & "<br>"
    dbget.Execute sqlStr

	response.write "<script type='text/javascript'>"
	response.write "	alert('수정 완료 되었습니다.');"
	response.write "	location.replace('/cscenter/pack/pojang_view.asp?orderserial="& orderserial &"&midx="& midx &"');"
	response.write "</script>"
	dbget.close()	:	response.end

else
	'response.write "<script type='text/javascript'>location.replace('"& SSLURL &"/inipay/pack/pack_step1.asp');</script>"
	response.write "<script type='text/javascript'>alert('정상적인 경로가 아닙니다.');</script>"
	dbget.close()	:	response.end
end if

%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->