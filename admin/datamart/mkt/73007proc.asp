<%@ language=vbscript %>
<% option Explicit %>
<% Response.CharSet = "euc-kr" %>
<%
'####################################################
' Description :  [2016 S/S 웨딩] Wedding Membership 승인처리 페이지
' History : 2016.09.12 유태욱
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim mode, evt_code, sub_idx, sqlStr, userid
dim wdcp1code, wdcp2code, wdcp3code, wdcp4code, wdcp5code
	mode = requestcheckvar(request("mode"),32)
	userid = requestcheckvar(request("userid"),32)
	evt_code = getNumeric(requestcheckvar(request("evt_code"),32))
	sub_idx = getNumeric(requestcheckvar(request("sub_idx"),10))

	If session("ssBctId") ="greenteenz" Or session("ssBctId") = "djjung" Then
	Else
		response.write "<script>alert('관계자만 볼 수 있는 페이지 입니다.');window.close();</script>"
		response.End
	End If
	
	dim refer
		refer = request.ServerVariables("HTTP_REFERER")
	if InStr(refer,"10x10.co.kr")<1 then
		Response.Write "<script type='text/javascript'>alert('잘못된 접속입니다.');</script>"
		dbget.close() : Response.End
	end If

''mkt 관리자가 어드민에서 승인
If mode = "gubunok" Then
	if evt_code="" then
		Response.Write "<script type='text/javascript'>alert('이벤트코드가 없습니다.');</script>"
		dbget.close() : Response.End
	end If
	if userid="" then
		Response.Write "<script type='text/javascript'>alert('승인할 ID가 없습니다.');</script>"
		dbget.close() : Response.End
	end If
	if sub_idx="" then
		Response.Write "<script type='text/javascript'>alert('idx가 없습니다.');</script>"
		dbget.close() : Response.End
	end If

'' 실섭 899,900,901,902,903
'' 테섭 2809,2810,2811,2812,2813

	IF application("Svr_Info") = "Dev" THEN
		wdcp1code   =  2813		''20만원 이상 2만원쿠폰
		wdcp2code   =  2812		''50만원 이상 6만원쿠폰
		wdcp3code   =  2811		''100만원 이상 15만원쿠폰
		wdcp4code   =  2810		''텐배 무료배송쿠폰1 (1만원이상 구매시)
		wdcp5code   =  2809		''텐배 무료배송쿠폰2 (1만원이상 구매시)
	Else
		wdcp1code   =  903		''20만원 이상 2만원쿠폰
		wdcp2code   =  902		''50만원 이상 6만원쿠폰
		wdcp3code   =  901		''100만원 이상 15만원쿠폰
		wdcp4code   =  900		''텐배 무료배송쿠폰1 (1만원이상 구매시)
		wdcp5code   =  899		''텐배 무료배송쿠폰2 (1만원이상 구매시)
	End If

	dim CPdate
	CPdate = Year(now) &"-"& right("0"&Month(now)+3,2) &"-"& right("0"&Day(now),2) &" "& right("0"& Hour(now),2) &":"& right("0"&Minute(now),2) &":"& right("0"&Second(now),2)

	''쿠폰 5종 발급
	sqlstr = "insert into [db_user].[dbo].tbl_user_coupon" + vbcrlf
	sqlstr = sqlstr & " (masteridx,userid,coupontype,couponvalue, couponname,minbuyprice,targetitemlist,startdate,expiredate,couponmeaipprice,validsitename)" + vbcrlf
	sqlstr = sqlstr & " 	SELECT idx, '"& userid &"',coupontype,couponvalue,couponname,minbuyprice,targetitemlist,startdate,expiredate,couponmeaipprice,validsitename" + vbcrlf
	sqlstr = sqlstr & " 	from [db_user].[dbo].tbl_user_coupon_master m" + vbcrlf
	sqlstr = sqlstr & " 	where idx in ("& wdcp1code &", "& wdcp2code &", "& wdcp3code &", "& wdcp4code &", "& wdcp5code &")"
'	response.write sqlstr
	dbget.execute sqlstr

	''sub_opt1 N를 Y로 변경
	sqlStr = "update db_event.dbo.tbl_event_subscript set" + vbcrlf
	sqlStr = sqlStr & " sub_opt1 = replace( sub_opt1, '/!/N', '/!/Y') where" + vbcrlf
	sqlStr = sqlStr & " evt_code='"& evt_code &"' and sub_idx='"& sub_idx &"'"
	'response.write sqlStr & "<br>"
	dbget.execute sqlStr

	Response.Write "<script type='text/javascript'>"
	Response.Write "	alert('OK');"
	Response.Write "	parent.top.location.replace('/admin/datamart/mkt/73007_manage.asp');"
	Response.Write "</script>"
	dbget.close() : Response.End

else
	Response.Write "<script type='text/javascript'>alert('구분자가 없습니다.');</script>"
	dbget.close() : Response.End
end if

%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
