<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"

Server.ScriptTimeOut = 60*10		' 10분
%>
<%
'####################################################
' Description : 제휴몰 수수료정산
' History : 2017.04.06 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/commissionjungsan_cls.asp"-->
<%
dim refer
	refer = request.ServerVariables("HTTP_REFERER")

if InStr(refer,"10x10.co.kr")<1 and session("ssBctId") <> "tozzinet" then
	Response.Write "잘못된 접속입니다."
	dbget.close() : response.end
end if

dim yyyy, mm, mode, sql, arrlist, bufStr, i, tendb, orderserial, itemnoptionname, cjungsan, rdsite
dim ismobile, csum, arrsum, bufsumStr, jungsandategubun, jungsanCount
	yyyy = requestcheckvar(getNumeric(request("yyyy")),4)
	mm = requestcheckvar(getNumeric(request("mm")),2)
	orderserial = requestcheckvar(getNumeric(request("orderserial")),11)
	itemnoptionname = requestcheckvar(request("itemnoptionname"),10)
	mode = requestcheckvar(request("mode"),32)
	rdsite = requestcheckvar(request("rdsite"),32)
	ismobile = requestcheckvar(getNumeric(request("ismobile")),1)

jungsanCount="N"
if yyyy="" then
	stdate = CStr(Now)
	stdate = DateSerial(Left(stdate,4), CLng(Mid(stdate,6,2)),1)
	yyyy = Left(stdate,4)
	mm = Mid(stdate,6,2)
end if

IF application("Svr_Info")="Dev" THEN
	tendb = "tendb."
end IF

'비트윈 csv 다운로드
if mode="csvbetween" then
	if trim(yyyy="") or trim(mm="") then
		response.write "<script type='text/javascript'>"
		response.write "	alert('기간이 없습니다.');"
		response.write "	parent.location.reload()"
		response.write "</script>"
		dbget.close() : response.end
	end if
	yyyy = trim(yyyy)
	mm = trim(mm)

	Set cjungsan = New Ccommission
		cjungsan.FRectyyyymm = yyyy + "-" + mm
		cjungsan.FPageSize = 100000
		cjungsan.FCurrPage = 1
		cjungsan.frectorderserial = orderserial
		cjungsan.frectitemname = itemnoptionname
		cjungsan.Getcommissionjungsan_between_notpaging()

		if cjungsan.FTotalCount < 1 then
			response.write "<script type='text/javascript'>"
			response.write "	alert('데이터가 없습니다.');"
			response.write "	parent.location.reload()"
			response.write "</script>"
			dbget.close() : response.end
		end if

		arrlist = cjungsan.farrlist

'	Response.Buffer=False
	Response.Expires=0
	response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=수수료정산between_"& yyyy &"년"& mm &"월.csv"
	Response.CacheControl = "public"

	response.write "주문일자,구매확정일자,주문번호,상품명,주문수량,주문금액(V.A.T제외),수수료율,수수료,주문상태,취소날짜" & vbcrlf

	if isarray(arrlist) then
		For i = 0 To ubound(arrlist,2)
		bufStr = ""
		bufStr = bufStr & arrlist(0,i)
		bufStr = bufStr & "," & arrlist(1,i)
		bufStr = bufStr & "," & arrlist(2,i)
		bufStr = bufStr & "," & escapedstring(arrlist(3,i))
		bufStr = bufStr & "," & arrlist(4,i)
		bufStr = bufStr & "," & arrlist(5,i)
		bufStr = bufStr & "," & arrlist(6,i)
		bufStr = bufStr & "," & arrlist(7,i)
		bufStr = bufStr & "," & arrlist(8,i)
		bufStr = bufStr & "," & arrlist(9,i)

		response.write bufStr & VbCrlf

		if i>0 and (i mod 10000)=0 then response.flush	'버퍼 초과를 막기위해 중간 플러시
		next
	end if

	set cjungsan = nothing

'비트윈 정산작성
elseif mode="regbetween" then
	if trim(yyyy="") or trim(mm="") then
		response.write "<script type='text/javascript'>"
		response.write "	alert('기간이 없습니다.');"
		response.write "	parent.location.reload()"
		response.write "</script>"
		dbget.close() : response.end
	end if
	yyyy = trim(yyyy)
	mm = trim(mm)

	jungsanCount="0"
	sql = "SELECT count(jd.orderserial) as jungsanCount"
	sql = sql & " from db_jungsan.dbo.tbl_nvshop_jungsan_detail jd with (nolock)" + vbcrlf
	sql = sql & " where jd.rdsite='betweenshop' and jd.jmonth='"& yyyy + "-" + mm &"'"

	'response.write sql & "<Br>"
	rsget.CursorLocation = adUseClient
	rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly

	If Not rsget.Eof Then
		jungsanCount = rsget("jungsancount")
	End If

	rsget.close

	if jungsanCount>0 then
		response.write "<script type='text/javascript'>"
		response.write "	alert('이미 작성된 정산 데이터가 있습니다.');"
		response.write "	parent.location.reload()"
		response.write "</script>"
		dbget.close() : response.end
	end if

	' 정산 기준 바뀜.
	if DateSerial(yyyy, mm, "01") >= "2021-03-01" then
		jungsandategubun="d.jungsanfixdate"
	else
		jungsandategubun="m.beadaldate"
	end if

	'/출고 후 취소(정산 작성 후 취소) 가끔식 업체가 송장을 잘못찍엇다던가. 가라로 찍는 경우도 있음.
	sql = "DECLARE @STDT varchar(10)" & vbcrlf
	sql = sql & " DECLARE @EDDT varchar(10)" & vbcrlf
	sql = sql & " SET @STDT='2014-07-01'" & vbcrlf		'고정값(바꾸지말것)
	sql = sql & " SET @EDDT='"& dateadd("m", +1, DateSerial(yyyy, mm, "01")) &"'" & vbcrlf 
	sql = sql & " insert into db_jungsan.dbo.tbl_nvshop_jungsan_detail" & vbcrlf
	sql = sql & " (orderserial,itemid,itemoption,cancelDT,jMonth,ismobile,orgorderserial,jumundiv,rdsite" & vbcrlf
	sql = sql & " ,rDate,fixedDate,itemNOptionName,itemno,suppPrc,commpro,commissoin,ordStatName)" & vbcrlf
	sql = sql & " 	select" & vbcrlf
	sql = sql & " 	d.orderserial,d.itemid,d.itemoption,isnull(convert(varchar(10),isNULL(M.canceldate,D.canceldate),21),'')" & vbcrlf
	sql = sql & " 	,convert(varchar(7),"& jungsandategubun &",21) as jMonth" & vbcrlf
	sql = sql & " 	,1 as ismobile,isNULL(m.linkorderserial,m.orderserial),m.jumundiv,isNULL(m.rdsite,'')" & vbcrlf
	sql = sql & " 	,convert(varchar(10),m.regdate,21) as rDate,convert(varchar(10),"& jungsandategubun &",21) as fixedDate" & vbcrlf
	sql = sql & " 	,d.itemname+(CASE WHEN d.itemoptionname<>'' THEN ' ['+d.itemoptionname+']' ELSE '' END)" & vbcrlf
	sql = sql & " 	,d.itemno*-1" & vbcrlf
	sql = sql & " 	,Round((d.reducedprice*d.itemno)/(CASE WHEN d.vatinclude='N' THEN 1 ELSE 1.1 END),0)*-1 as SuppPrc" & vbcrlf
	sql = sql & " 	,0.06" & vbcrlf
	sql = sql & " 	,Round((d.reducedprice*d.itemno)/(CASE WHEN d.vatinclude='N' THEN 1 ELSE 1.1 END),0)*(0.06)*-1" & vbcrlf
	sql = sql & " 	,CASE WHEN m.jumundiv='9' THEN '반품후취소' " & vbcrlf
	sql = sql & " 		WHEN m.jumundiv='6' THEN '교환완료'  ELSE '출고후취소' END" & vbcrlf
	sql = sql & " 	from db_order.dbo.tbl_order_master m" & vbcrlf
	sql = sql & " 	Join db_order.dbo.tbl_order_detail d" & vbcrlf
	sql = sql & " 		on m.orderserial=d.orderserial" & vbcrlf
	sql = sql & " 	join db_jungsan.dbo.tbl_nvshop_jungsan_detail J" & vbcrlf
	sql = sql & " 		on d.orderserial=J.orderserial" & vbcrlf
	sql = sql & " 		and d.itemid=j.itemid" & vbcrlf
	sql = sql & " 		and d.itemoption=j.itemoption" & vbcrlf 	'and ordStatName<>'출고후취소'
	sql = sql & " 	left join db_item.[dbo].[tbl_item_standing_item] si" & vbcrlf
	sql = sql & " 		on d.itemid = si.orgitemid" & vbcrlf
	sql = sql & " 		and d.itemoption = si.orgitemoption" & vbcrlf
	sql = sql & " 	left join db_item.[dbo].[tbl_item_standing_order] so" & vbcrlf		' 정기구독 마지막회차 완료 되면 원래 상품이 취소로 돌아가서 여기 나옴. 제낌
	sql = sql & " 		on d.itemid = so.orgitemid" & vbcrlf
	sql = sql & " 		and d.itemoption = so.orgitemoption" & vbcrlf
	sql = sql & " 		and si.endreserveidx = so.reserveidx" & vbcrlf
	sql = sql & " 	where 1=1 and m.rdsite='betweenshop'" & vbcrlf
	sql = sql & " 	and m.ipkumdiv>3" & vbcrlf
	'sql = sql & " 	and m.ipkumdiv=8" & vbcrlf
	sql = sql & " 	and (m.cancelyn<>'N' or  d.cancelyn='Y')" & vbcrlf	'취소의 CASE
	sql = sql & " 	and "& jungsandategubun &">@STDT and "& jungsandategubun &"<@EDDT" & vbcrlf
	sql = sql & " 	and d.itemid not in (0,100)" & vbcrlf
	sql = sql & " 	and so.orgitemid is null" & vbcrlf

	'response.write sql & "<br>"
	dbget.CommandTimeout = 60*5   ' 5분
	dbget.execute sql

	sql = "DECLARE @STDT varchar(10)" & vbcrlf
	sql = sql & " DECLARE @EDDT varchar(10)" & vbcrlf
	sql = sql & " SET @STDT='2014-07-01'" & vbcrlf		'고정값(바꾸지말것)
	sql = sql & " SET @EDDT='"& dateadd("m", +1, DateSerial(yyyy, mm, "01")) &"'" & vbcrlf
	sql = sql & " insert into db_jungsan.dbo.tbl_nvshop_jungsan_detail" & vbcrlf
	sql = sql & " (orderserial,itemid,itemoption,cancelDT,jMonth,ismobile,orgorderserial,jumundiv,rdsite" & vbcrlf
	sql = sql & " ,rDate,fixedDate,itemNOptionName,itemno,suppPrc,commpro,commissoin,ordStatName)" & vbcrlf
	sql = sql & " 	select" & vbcrlf
	sql = sql & " 	d.orderserial,d.itemid,d.itemoption,'',convert(varchar(7),"& jungsandategubun &",21) as jMonth" & vbcrlf
	sql = sql & " 	,1 as ismobile,isNULL(m.linkorderserial,m.orderserial),m.jumundiv,isNULL(m.rdsite,'')" & vbcrlf
	sql = sql & " 	,convert(varchar(10),m.regdate,21) as rDate" & vbcrlf
	sql = sql & " 	,convert(varchar(10),"& jungsandategubun &",21) as fixedDate" & vbcrlf
	sql = sql & " 	,d.itemname+(CASE WHEN d.itemoptionname<>'' THEN ' ['+d.itemoptionname+']' ELSE '' END)" & vbcrlf
	sql = sql & " 	,d.itemno" & vbcrlf
	sql = sql & " 	,Round((d.reducedprice*d.itemno)/(CASE WHEN d.vatinclude='N' THEN 1 ELSE 1.1 END),0) as SuppPrc" & vbcrlf
	sql = sql & " 	,(0.06)" & vbcrlf
	sql = sql & " 	,Round((d.reducedprice*d.itemno)/(CASE WHEN d.vatinclude='N' THEN 1 ELSE 1.1 END),0)*(0.06)" & vbcrlf
	sql = sql & " 	,CASE WHEN m.jumundiv='9' THEN '반품완료'" & vbcrlf 
	sql = sql & " 			WHEN	m.jumundiv='6' THEN '교환완료' ELSE '출고완료' END" & vbcrlf
	sql = sql & " 	from db_order.dbo.tbl_order_master m" & vbcrlf
	sql = sql & " 	Join db_order.dbo.tbl_order_detail d" & vbcrlf
	sql = sql & " 		on m.orderserial=d.orderserial" & vbcrlf
	sql = sql & " 	left join db_jungsan.dbo.tbl_nvshop_jungsan_detail J" & vbcrlf
	sql = sql & " 		on d.orderserial=J.orderserial" & vbcrlf
	sql = sql & " 		and d.itemid=j.itemid" & vbcrlf
	sql = sql & " 		and d.itemoption=j.itemoption" & vbcrlf
	sql = sql & " 	where 1=1 and m.rdsite='betweenshop'" & vbcrlf
	sql = sql & " 	and J.itemoption is NULL" & vbcrlf
	'sql = sql & " 	and R.ismobile=0" & vbcrlf 	'모바일인경우
	sql = sql & " 	and m.ipkumdiv>3" & vbcrlf
	sql = sql & " 	and m.ipkumdiv=8" & vbcrlf
	sql = sql & " 	and m.cancelyn='N'" & vbcrlf
	sql = sql & " 	and "& jungsandategubun &">@STDT and "& jungsandategubun &"<@EDDT" & vbcrlf
	sql = sql & " 	and d.itemid not in (0,100)" & vbcrlf
	sql = sql & " 	and d.cancelyn<>'Y'" & vbcrlf   '취소는 포함 안함.
	'sql = sql & " 	order by rDate,d.orderserial" & vbcrlf

	'response.write sql & "<br>"
	dbget.CommandTimeout = 60*5   ' 5분
	dbget.execute sql

	response.write "<script type='text/javascript'>"
	response.write "	alert('"& yyyy &"년 "& mm&"월 between 수수료정산 작성 완료.');"
	response.write "	parent.location.reload()"
	response.write "</script>"
	dbget.close() : response.end

'다음, 네이트 csv 다운로드
elseif mode="csvdaum" or mode="csvnate" then
	if rdsite="" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('제휴몰 구분이 없습니다.');"
		response.write "	parent.location.reload()"
		response.write "</script>"
		dbget.close() : response.end
	end if

	if trim(yyyy="") or trim(mm="") then
		response.write "<script type='text/javascript'>"
		response.write "	alert('기간이 없습니다.');"
		response.write "	parent.location.reload()"
		response.write "</script>"
		dbget.close() : response.end
	end if
	yyyy = trim(yyyy)
	mm = trim(mm)

	Set cjungsan = New Ccommission
		cjungsan.FRectyyyymm = yyyy + "-" + mm
		cjungsan.FPageSize = 100000
		cjungsan.FCurrPage = 1
		cjungsan.frectorderserial = orderserial
		cjungsan.frectitemname = itemnoptionname
		cjungsan.frectrdsite = rdsite
		cjungsan.Getcommissionjungsan_daum_notpaging()

		if cjungsan.FTotalCount < 1 then
			response.write "<script type='text/javascript'>"
			response.write "	alert('데이터가 없습니다.');"
			response.write "	parent.location.reload()"
			response.write "</script>"
			dbget.close() : response.end
		end if

		arrlist = cjungsan.farrlist

'	Response.Buffer=False
	Response.Expires=0
	response.ContentType = "application/vnd.ms-excel"

	if mode="csvdaum" then
		Response.AddHeader "Content-Disposition", "attachment; filename=수수료정산daum_"& yyyy &"년"& mm &"월.csv"
	ELSE
		Response.AddHeader "Content-Disposition", "attachment; filename=수수료정산nate_"& yyyy &"년"& mm &"월.csv"
	end if

	Response.CacheControl = "public"

	response.write "주문일자,출고일자/확정일자,매출코드,모바일구분,주문번호,상품명,주문수량,주문금액(V.A.T제외),수수료율,수수료,주문상태,취소날짜" & vbcrlf

	if isarray(arrlist) then
		For i = 0 To ubound(arrlist,2)
		bufStr = ""
		bufStr = bufStr & arrlist(0,i)
		bufStr = bufStr & "," & arrlist(1,i)
		bufStr = bufStr & "," & arrlist(2,i)
		bufStr = bufStr & "," & arrlist(3,i)
		bufStr = bufStr & "," & arrlist(4,i)
		bufStr = bufStr & "," & escapedstring(arrlist(5,i))
		bufStr = bufStr & "," & arrlist(6,i)
		bufStr = bufStr & "," & arrlist(7,i)
		bufStr = bufStr & "," & arrlist(8,i)
		bufStr = bufStr & "," & arrlist(9,i)
		bufStr = bufStr & "," & arrlist(10,i)
		bufStr = bufStr & "," & arrlist(11,i)

		response.write bufStr & VbCrlf

		if i>0 and (i mod 10000)=0 then response.flush	'버퍼 초과를 막기위해 중간 플러시
		next
	end if

	set cjungsan = nothing

'다음, 네이트 정산작성
elseif mode="regdaum" or mode="regnate" then
	if rdsite="" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('제휴몰 구분이 없습니다.');"
		response.write "	parent.location.reload()"
		response.write "</script>"
		dbget.close() : response.end
	end if

	if trim(yyyy="") or trim(mm="") then
		response.write "<script type='text/javascript'>"
		response.write "	alert('기간이 없습니다.');"
		response.write "	parent.location.reload()"
		response.write "</script>"
		dbget.close() : response.end
	end if
	yyyy = trim(yyyy)
	mm = trim(mm)

	jungsanCount="0"
	sql = "SELECT count(jd.orderserial) as jungsanCount"
	sql = sql & " from db_jungsan.dbo.tbl_nvshop_jungsan_detail jd with (nolock)" + vbcrlf
	sql = sql & " Join db_item.dbo.tbl_Outmall_RdsiteGubun R with (nolock)"
	sql = sql & " 	on jd.rdsite=R.rdsite"
	sql = sql & " 	and R.gubun in ('"& rdsite &"')" + vbcrlf
	sql = sql & " where jd.jmonth='"& yyyy + "-" + mm &"'"

	'response.write sql & "<Br>"
	rsget.CursorLocation = adUseClient
	rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly

	If Not rsget.Eof Then
		jungsanCount = rsget("jungsancount")
	End If

	rsget.close

	if jungsanCount>0 then
		response.write "<script type='text/javascript'>"
		response.write "	alert('이미 작성된 정산 데이터가 있습니다.');"
		response.write "	parent.location.reload()"
		response.write "</script>"
		dbget.close() : response.end
	end if

	' 정산 기준 바뀜.
	if DateSerial(yyyy, mm, "01") >= "2021-03-01" then
		jungsandategubun="d.jungsanfixdate"
	else
		jungsandategubun="m.beadaldate"
	end if

	'/출고 후 취소(정산 작성 후 취소) 가끔식 업체가 송장을 잘못찍엇다던가. 가라로 찍는 경우도 있음.
	sql = "DECLARE @STDT varchar(10)" & vbcrlf
	sql = sql & " DECLARE @EDDT varchar(10)" & vbcrlf
	sql = sql & " SET @STDT='2014-08-01'" & vbcrlf 		'고정값(바꾸지말것)
	sql = sql & " SET @EDDT='"& dateadd("m", +1, DateSerial(yyyy, mm, "01")) &"'" & vbcrlf
	sql = sql & " insert into db_jungsan.dbo.tbl_nvshop_jungsan_detail" & vbcrlf
	sql = sql & " (orderserial,itemid,itemoption,cancelDT,jMonth,ismobile,orgorderserial,jumundiv,rdsite" & vbcrlf
	sql = sql & " ,rDate,fixedDate,itemNOptionName,itemno,suppPrc,commpro,commissoin,ordStatName)" & vbcrlf
	sql = sql & " 	select" & vbcrlf
	sql = sql & " 	d.orderserial,d.itemid,d.itemoption,isnull(convert(varchar(10),isNULL(M.canceldate,D.canceldate),21),'')" & vbcrlf
	sql = sql & " 	,convert(varchar(7),"& jungsandategubun &",21) as jMonth" & vbcrlf
	sql = sql & " 	,R.ismobile,isNULL(m.linkorderserial,m.orderserial),m.jumundiv,isNULL(m.rdsite,'')" & vbcrlf
	sql = sql & " 	,convert(varchar(10),m.regdate,21) as rDate" & vbcrlf
	sql = sql & " 	,convert(varchar(10),"& jungsandategubun &",21) as fixedDate" & vbcrlf
	sql = sql & " 	,d.itemname+(CASE WHEN d.itemoptionname<>'' THEN ' ['+d.itemoptionname+']' ELSE '' END)" & vbcrlf
	sql = sql & " 	,d.itemno*-1" & vbcrlf
	sql = sql & " 	,Round((d.reducedprice*d.itemno)/(CASE WHEN d.vatinclude='N' THEN 1 ELSE 1.1 END),0)*-1 as SuppPrc" & vbcrlf
	sql = sql & " 	,(CASE WHEN R.isCharge='Y' THEN 0.03 ELSE 0 END)" & vbcrlf
	sql = sql & " 	,Round((d.reducedprice*d.itemno)/(CASE WHEN d.vatinclude='N' THEN 1 ELSE 1.1 END),0)*(CASE WHEN R.isCharge='Y' THEN 0.03 ELSE 0 END)*-1" & vbcrlf
	sql = sql & " 	,CASE WHEN m.jumundiv='9' THEN '반품후취소'" & vbcrlf
	sql = sql & " 		WHEN m.jumundiv='6' THEN '교환완료'  ELSE '출고후취소' END" & vbcrlf
	sql = sql & " 	from db_order.dbo.tbl_order_master m with (nolock)" & vbcrlf
	sql = sql & " 	Join db_order.dbo.tbl_order_detail d with (nolock)" & vbcrlf
	sql = sql & " 		on m.orderserial=d.orderserial" & vbcrlf
	sql = sql & " 	Join db_item.dbo.tbl_Outmall_RdsiteGubun R with (nolock)" & vbcrlf
	sql = sql & " 		on m.rdsite=R.rdsite" & vbcrlf
	sql = sql & " 		and R.gubun in ('"& rdsite &"')" & vbcrlf
	sql = sql & " 	join db_jungsan.dbo.tbl_nvshop_jungsan_detail J with (nolock)" & vbcrlf
	sql = sql & " 		on d.orderserial=J.orderserial" & vbcrlf
	sql = sql & " 		and d.itemid=j.itemid" & vbcrlf
	sql = sql & " 		and d.itemoption=j.itemoption" & vbcrlf 	'and ordStatName<>'출고후취소'
	sql = sql & " 	left join db_item.[dbo].[tbl_item_standing_item] si with (nolock)" & vbcrlf
	sql = sql & " 		on d.itemid = si.orgitemid" & vbcrlf
	sql = sql & " 		and d.itemoption = si.orgitemoption" & vbcrlf
	sql = sql & " 	left join db_item.[dbo].[tbl_item_standing_order] so with (nolock)" & vbcrlf		' 정기구독 마지막회차 완료 되면 원래 상품이 취소로 돌아가서 여기 나옴. 제낌
	sql = sql & " 		on d.itemid = so.orgitemid" & vbcrlf
	sql = sql & " 		and d.itemoption = so.orgitemoption" & vbcrlf
	sql = sql & " 		and si.endreserveidx = so.reserveidx" & vbcrlf
	sql = sql & " 	left join db_jungsan.dbo.tbl_nvshop_jungsan_detail JJ with (nolock)" & vbcrlf
	sql = sql & " 		on d.orderserial=JJ.orderserial" & vbcrlf
	sql = sql & " 		and d.itemid=JJ.itemid" & vbcrlf
	sql = sql & " 		and isnull(convert(varchar(10),isNULL(M.canceldate,D.canceldate),21),'')=JJ.cancelDT" & vbcrlf
	sql = sql & " 	where 1=1" & vbcrlf
	sql = sql & " 	and m.ipkumdiv>3" & vbcrlf
	sql = sql & " 	and (m.cancelyn<>'N' or  d.cancelyn='Y')" & vbcrlf	'취소의 CASE
	sql = sql & " 	and "& jungsandategubun &">@STDT and "& jungsandategubun &"<@EDDT" & vbcrlf
	sql = sql & " 	and d.itemid not in (0,100)" & vbcrlf
	sql = sql & " 	and so.orgitemid is null" & vbcrlf
	sql = sql & " 	and JJ.orderserial is null" & vbcrlf	' 이미 정산된건 제낌
	sql = sql & " 	and m.orderserial not in ("
	sql = sql & " 		'21022272847'"
	sql = sql & " 	)"

	'response.write sql & "<br>"
	dbget.CommandTimeout = 60*5   ' 5분
	dbget.execute sql

	sql = "DECLARE @STDT varchar(10)" & vbcrlf
	sql = sql & " DECLARE @EDDT varchar(10)" & vbcrlf
	sql = sql & " SET @STDT='2014-08-01'" & vbcrlf 		'고정값(바꾸지말것)
	sql = sql & " SET @EDDT='"& dateadd("m", +1, DateSerial(yyyy, mm, "01")) &"'" & vbcrlf
	sql = sql & " insert into db_jungsan.dbo.tbl_nvshop_jungsan_detail" & vbcrlf
	sql = sql & " (orderserial,itemid,itemoption,cancelDT,jMonth,ismobile,orgorderserial,jumundiv,rdsite" & vbcrlf
	sql = sql & " ,rDate,fixedDate,itemNOptionName,itemno,suppPrc,commpro,commissoin,ordStatName)" & vbcrlf
	sql = sql & " 	select" & vbcrlf
	sql = sql & " 	d.orderserial,d.itemid,d.itemoption,''" & vbcrlf
	sql = sql & " 	,convert(varchar(7),"& jungsandategubun &",21) as jMonth" & vbcrlf
	sql = sql & " 	,R.ismobile,isNULL(m.linkorderserial,m.orderserial),m.jumundiv,isNULL(m.rdsite,'')" & vbcrlf
	sql = sql & " 	,convert(varchar(10),m.regdate,21) as rDate" & vbcrlf
	sql = sql & " 	,convert(varchar(10),"& jungsandategubun &",21) as fixedDate" & vbcrlf
	sql = sql & " 	,d.itemname+(CASE WHEN d.itemoptionname<>'' THEN ' ['+d.itemoptionname+']' ELSE '' END)" & vbcrlf
	sql = sql & " 	,d.itemno" & vbcrlf
	sql = sql & " 	,Round((d.reducedprice*d.itemno)/(CASE WHEN d.vatinclude='N' THEN 1 ELSE 1.1 END),0) as SuppPrc" & vbcrlf
	sql = sql & " 	,(CASE WHEN R.isCharge='Y' THEN 0.03 ELSE 0 END)" & vbcrlf
	sql = sql & " 	,Round((d.reducedprice*d.itemno)/(CASE WHEN d.vatinclude='N' THEN 1 ELSE 1.1 END),0)*(CASE WHEN R.isCharge='Y' THEN 0.03 ELSE 0 END)" & vbcrlf
	sql = sql & " 	,CASE WHEN m.jumundiv='9' THEN '반품완료' WHEN	m.jumundiv='6' THEN '교환완료' ELSE '출고완료' END" & vbcrlf
	sql = sql & " 	from db_order.dbo.tbl_order_master m with (nolock)" & vbcrlf
	sql = sql & " 	Join db_order.dbo.tbl_order_detail d with (nolock)" & vbcrlf
	sql = sql & " 		on m.orderserial=d.orderserial" & vbcrlf
	sql = sql & " 	Join db_item.dbo.tbl_Outmall_RdsiteGubun R with (nolock)" & vbcrlf
	sql = sql & " 		on m.rdsite=R.rdsite" & vbcrlf
	sql = sql & " 		and R.gubun in ('"& rdsite &"')" & vbcrlf
	sql = sql & " 	left join db_jungsan.dbo.tbl_nvshop_jungsan_detail J with (nolock)" & vbcrlf
	sql = sql & " 		on d.orderserial=J.orderserial" & vbcrlf
	sql = sql & " 		and d.itemid=j.itemid" & vbcrlf
	sql = sql & " 		and d.itemoption=j.itemoption" & vbcrlf
	sql = sql & " 	where 1=1" & vbcrlf
	sql = sql & " 	and (J.itemoption is NULL" & vbcrlf
	sql = sql & " 	and m.ipkumdiv>3" & vbcrlf
	sql = sql & " 	and m.ipkumdiv=8" & vbcrlf
	sql = sql & " 	and m.cancelyn='N'" & vbcrlf
	sql = sql & " 	and "& jungsandategubun &">@STDT and "& jungsandategubun &"<@EDDT" & vbcrlf
	sql = sql & " 	and d.itemid not in (0,100)" & vbcrlf
	sql = sql & " 	and d.cancelyn<>'Y')" & vbcrlf   '취소는 포함 안함.
	sql = sql & " 	or (m.orderserial = '21022272847' and J.itemoption is NULL)" & vbcrlf   '임시

	'response.write sql & "<br>"
	dbget.CommandTimeout = 60*5   ' 5분
	dbget.execute sql

	response.write "<script type='text/javascript'>"
	
	if mode="regdaum" then
		response.write "	alert('"& yyyy &"년 "& mm&"월 daum 수수료정산 작성 완료.');"
	ELSE
		response.write "	alert('"& yyyy &"년 "& mm&"월 nate 수수료정산 작성 완료.');"
	end if

	response.write "	parent.location.reload()"
	response.write "</script>"
	dbget.close() : response.end

'네이버 csv 다운로드
elseif mode="csvnaver" then
	if ismobile="" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('구분이 없습니다.');"
		response.write "	parent.location.reload()"
		response.write "</script>"
		dbget.close() : response.end
	end if

	if trim(yyyy="") or trim(mm="") then
		response.write "<script type='text/javascript'>"
		response.write "	alert('기간이 없습니다.');"
		response.write "	parent.location.reload()"
		response.write "</script>"
		dbget.close() : response.end
	end if
	yyyy = trim(yyyy)
	mm = trim(mm)

	Set cjungsan = New Ccommission
		cjungsan.FRectyyyymm = yyyy + "-" + mm
		cjungsan.FPageSize = 100000
		cjungsan.FCurrPage = 1
		cjungsan.frectorderserial = orderserial
		cjungsan.frectitemname = itemnoptionname
		cjungsan.frectismobile = ismobile
		cjungsan.Getcommissionjungsan_naver_notpaging()

		if cjungsan.FTotalCount < 1 then
			response.write "<script type='text/javascript'>"
			response.write "	alert('데이터가 없습니다.');"
			response.write "	parent.location.reload()"
			response.write "</script>"
			dbget.close() : response.end
		end if

		arrlist = cjungsan.farrlist

	Set csum = New Ccommission
		csum.FRectyyyymm = yyyy + "-" + mm
		csum.FPageSize = 500
		csum.FCurrPage = 1
'		csum.frectorderserial = orderserial
'		csum.frectitemname = itemnoptionname
		csum.frectismobile = ismobile
		csum.Getcommissionjungsan_naver_sum_notpaging()

		if csum.FTotalCount < 1 then
			response.write "<script type='text/javascript'>"
			response.write "	alert('영역 상세 데이터가 없습니다.');"
			response.write "	parent.location.reload()"
			response.write "</script>"
			dbget.close() : response.end
		end if

		arrsum = csum.farrlist

'	Response.Buffer=False
	Response.Expires=0
	response.ContentType = "application/vnd.ms-excel"

	Response.AddHeader "Content-Disposition", "attachment; filename=수수료정산naver_"& yyyy &"년"& mm &"월.csv"
	Response.CacheControl = "public"
	response.write "사이트,순판매수량,순주문금액,수수료,수수료정산여부,영역상세" & vbcrlf

	if isarray(arrsum) then
		For i = 0 To ubound(arrsum,2)
		bufsumStr = ""
		bufsumStr = bufsumStr & arrsum(0,i)
		bufsumStr = bufsumStr & "," & arrsum(1,i)
		bufsumStr = bufsumStr & "," & arrsum(2,i)
		bufsumStr = bufsumStr & "," & arrsum(3,i)
		bufsumStr = bufsumStr & "," & escapedstring(arrsum(4,i))

		response.write bufsumStr & VbCrlf
		next
	end if

	response.write VbCrlf

	response.write "주문일자,구매확정일자(결제완료일자),주문번호,상품명,주문수량,주문금액(V.A.T제외),수수료율,수수료,주문상태,취소날짜" & vbcrlf

	if isarray(arrlist) then
		For i = 0 To ubound(arrlist,2)
			bufStr = ""
			bufStr = bufStr & arrlist(0,i)
			bufStr = bufStr & "," & arrlist(1,i)
			bufStr = bufStr & "," & arrlist(2,i)
			bufStr = bufStr & "," & escapedstring(arrlist(3,i))
			bufStr = bufStr & "," & arrlist(4,i)
			bufStr = bufStr & "," & arrlist(5,i)
			bufStr = bufStr & "," & arrlist(6,i)
			bufStr = bufStr & "," & arrlist(7,i)
			bufStr = bufStr & "," & arrlist(8,i)
			bufStr = bufStr & "," & arrlist(9,i)

			response.write bufStr & VbCrlf

			if i>0 and (i mod 10000)=0 then response.flush	'버퍼 초과를 막기위해 중간 플러시
		next
	end if

	set cjungsan = nothing

'네이버 정산작성
elseif mode="regnaver" then
	if trim(yyyy="") or trim(mm="") then
		response.write "<script type='text/javascript'>"
		response.write "	alert('기간이 없습니다.');"
		response.write "	parent.location.reload()"
		response.write "</script>"
		dbget.close() : response.end
	end if
	yyyy = trim(yyyy)
	mm = trim(mm)

	jungsanCount="0"
	sql = "SELECT count(jd.orderserial) as jungsanCount"
	sql = sql & " from db_jungsan.dbo.tbl_nvshop_jungsan_detail jd with (nolock)" + vbcrlf
	sql = sql & " join db_item.dbo.tbl_Outmall_RdsiteGubun R with (nolock)" + vbcrlf
	sql = sql & " 	on jd.rdsite=R.rdsite" + vbcrlf
	sql = sql & " 	and R.gubun='nvshop'" + vbcrlf
	sql = sql & " where jd.jmonth='"& yyyy + "-" + mm &"'"

	'response.write sql & "<Br>"
	rsget.CursorLocation = adUseClient
	rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly

	If Not rsget.Eof Then
		jungsanCount = rsget("jungsancount")
	End If

	rsget.close

	if jungsanCount>0 then
		response.write "<script type='text/javascript'>"
		response.write "	alert('이미 작성된 정산 데이터가 있습니다.');"
		response.write "	parent.location.reload()"
		response.write "</script>"
		dbget.close() : response.end
	end if

	' 정산 기준 바뀜.
	if DateSerial(yyyy, mm, "01") >= "2021-03-01" then
		jungsandategubun="d.jungsanfixdate"
	else
		jungsandategubun="m.beadaldate"
	end if

	'/출고 후 취소(정산 작성 후 취소) 가끔식 업체가 송장을 잘못찍엇다던가. 가라로 찍는 경우도 있음.
	sql = "DECLARE @STDT varchar(10)" & vbcrlf
	sql = sql & " DECLARE @EDDT varchar(10)" & vbcrlf
	sql = sql & " DECLARE @jMonth varchar(7)" & vbcrlf
	sql = sql & " DECLARE @STORDERSERIAL varchar(11)" & vbcrlf
	sql = sql & " SET @STDT='2014-07-26'" & vbcrlf '고정 매월26~ 25일까지
	sql = sql & " SET @EDDT='"& DateSerial(yyyy, mm, "26") &"'" & vbcrlf
	sql = sql & " SET @STORDERSERIAL=RIGHT(REPLACE(@STDT,'-',''),6)+'00000'" & vbcrlf
	sql = sql & " SET @jMonth=LEFT(@EDDT,7)" & vbcrlf
	sql = sql & " insert into db_jungsan.dbo.tbl_nvshop_jungsan_detail" & vbcrlf
	sql = sql & " (orderserial,itemid,itemoption,cancelDT,jMonth,ismobile,orgorderserial,jumundiv,rdsite" & vbcrlf
	sql = sql & " ,rDate,fixedDate,itemNOptionName,itemno,suppPrc,commpro,commissoin,ordStatName)" & vbcrlf
	sql = sql & " 	select" & vbcrlf
	sql = sql & " 	d.orderserial,d.itemid,d.itemoption,isnull(convert(varchar(10),isNULL(M.canceldate,D.canceldate),21),'')" & vbcrlf
	'sql = sql & " 	,convert(varchar(7),"& jungsandategubun &",21) as jMonth" & vbcrlf
	sql = sql & " 	,(case when day("& jungsandategubun &")>25 then convert(varchar(7),dateadd(month,+1,"& jungsandategubun &"),121) else convert(varchar(7),"& jungsandategubun &",121) end) as jMonth" & vbcrlf
	sql = sql & " 	,R.ismobile,isNULL(m.linkorderserial,m.orderserial),m.jumundiv,isNULL(m.rdsite,'')" & vbcrlf
	sql = sql & " 	,convert(varchar(10),m.regdate,21) as rDate" & vbcrlf
	sql = sql & " 	,convert(varchar(10),"& jungsandategubun &",21) as fixedDate" & vbcrlf
	sql = sql & " 	,d.itemname+(CASE WHEN d.itemoptionname<>'' THEN ' ['+d.itemoptionname+']' ELSE '' END)" & vbcrlf
	sql = sql & " 	,d.itemno*-1" & vbcrlf
	sql = sql & " 	,Round((d.reducedprice*d.itemno)/(CASE WHEN d.vatinclude='N' THEN 1 ELSE 1.1 END),0)*-1 as SuppPrc" & vbcrlf
	sql = sql & " 	,(CASE WHEN R.isCharge='Y' THEN 0.02 ELSE 0 END)" & vbcrlf
	sql = sql & " 	,Round((d.reducedprice*d.itemno)/(CASE WHEN d.vatinclude='N' THEN 1 ELSE 1.1 END),0)*(CASE WHEN R.isCharge='Y' THEN 0.02 ELSE 0 END)*-1" & vbcrlf
	sql = sql & " 	,CASE" & vbcrlf
	sql = sql & " 		WHEN m.jumundiv='9' THEN '반품후취소'" & vbcrlf
	sql = sql & " 		WHEN m.jumundiv='6' THEN '교환완료'  ELSE '출고후취소' END" & vbcrlf
	sql = sql & " 	from db_order.dbo.tbl_order_master m with (nolock)" & vbcrlf
	sql = sql & " 	Join db_order.dbo.tbl_order_detail d with (nolock)" & vbcrlf
	sql = sql & " 		on m.orderserial=d.orderserial" & vbcrlf
	sql = sql & " 	Join db_item.dbo.tbl_Outmall_RdsiteGubun R with (nolock)" & vbcrlf
	sql = sql & " 		on m.rdsite=R.rdsite" & vbcrlf
	sql = sql & " 		and R.gubun='nvshop'" & vbcrlf
	sql = sql & " 	Join db_jungsan.dbo.tbl_nvshop_jungsan_detail J with (nolock)" & vbcrlf	'이미 정산에 들어가 있고.
	sql = sql & " 		on d.orderserial=J.orderserial" & vbcrlf
	sql = sql & " 		and d.itemid=j.itemid" & vbcrlf
	sql = sql & " 		and d.itemoption=j.itemoption and J.ordStatName<>'출고후취소'" & vbcrlf
	sql = sql & " 	left join db_jungsan.dbo.tbl_nvshop_jungsan_detail J2 with (nolock)" & vbcrlf
	sql = sql & " 		on d.orderserial=J2.orderserial" & vbcrlf
	sql = sql & " 		and d.itemid=j2.itemid" & vbcrlf
	sql = sql & " 		and d.itemoption=j2.itemoption and J2.ordStatName='출고후취소'" & vbcrlf
	sql = sql & " 	left join db_item.[dbo].[tbl_item_standing_item] si with (nolock)" & vbcrlf
	sql = sql & " 		on d.itemid = si.orgitemid" & vbcrlf
	sql = sql & " 		and d.itemoption = si.orgitemoption" & vbcrlf
	sql = sql & " 	left join db_item.[dbo].[tbl_item_standing_order] so with (nolock)" & vbcrlf		' 정기구독 마지막회차 완료 되면 원래 상품이 취소로 돌아가서 여기 나옴. 제낌
	sql = sql & " 		on d.itemid = so.orgitemid" & vbcrlf
	sql = sql & " 		and d.itemoption = so.orgitemoption" & vbcrlf
	sql = sql & " 		and si.endreserveidx = so.reserveidx" & vbcrlf
	sql = sql & " 	where m.ipkumdiv>3" & vbcrlf
	sql = sql & " 	and (m.cancelyn<>'N' or d.cancelyn='Y')" & vbcrlf 	'취소의 CASE (출고후)
	sql = sql & " 	and "& jungsandategubun &">@STDT and "& jungsandategubun &"<@EDDT" & vbcrlf
	sql = sql & " 	and d.itemid not in (0,100)" & vbcrlf
	sql = sql & " 	and J2.orderserial is NULL" & vbcrlf
	sql = sql & " 	and so.orgitemid is null" & vbcrlf

	'response.write sql & "<br>"
	dbget.CommandTimeout = 60*5   ' 5분
	dbget.execute sql

	sql = "DECLARE @STDT varchar(10)" & vbcrlf
	sql = sql & " DECLARE @EDDT varchar(10)" & vbcrlf
	sql = sql & " DECLARE @jMonth varchar(7)" & vbcrlf
	sql = sql & " DECLARE @STORDERSERIAL varchar(11)" & vbcrlf
	sql = sql & " SET @STDT='2014-05-26'" & vbcrlf '고정 매월26~ 25일까지
	sql = sql & " SET @EDDT='"& DateSerial(yyyy, mm, "26") &"'" & vbcrlf
	sql = sql & " SET @STORDERSERIAL=RIGHT(REPLACE(@STDT,'-',''),6)+'00000'" & vbcrlf
	sql = sql & " SET @jMonth=LEFT(@EDDT,7)" & vbcrlf
	sql = sql & " insert into db_jungsan.dbo.tbl_nvshop_jungsan_detail" & vbcrlf
	sql = sql & " (orderserial,itemid,itemoption,cancelDT,jMonth,ismobile,orgorderserial,jumundiv,rdsite" & vbcrlf
	sql = sql & " ,rDate,fixedDate,itemNOptionName,itemno,suppPrc,commpro,commissoin,ordStatName)" & vbcrlf
	sql = sql & " 	select" & vbcrlf
	sql = sql & " 	d.orderserial,d.itemid,d.itemoption,''" & vbcrlf
	'sql = sql & " 	,convert(varchar(7),"& jungsandategubun &",21) as jMonth" & vbcrlf
	sql = sql & " 	,(case when day("& jungsandategubun &")>25 then convert(varchar(7),dateadd(month,+1,"& jungsandategubun &"),121) else convert(varchar(7),"& jungsandategubun &",121) end) as jMonth" & vbcrlf
	sql = sql & " 	,R.ismobile,isNULL(m.linkorderserial,m.orderserial),m.jumundiv,isNULL(m.rdsite,'')" & vbcrlf
	sql = sql & " 	,convert(varchar(10),m.regdate,21) as rDate" & vbcrlf
	sql = sql & " 	,convert(varchar(10),"& jungsandategubun &",21) as fixedDate" & vbcrlf
	sql = sql & " 	,d.itemname+(CASE WHEN d.itemoptionname<>'' THEN ' ['+d.itemoptionname+']' ELSE '' END)" & vbcrlf
	sql = sql & " 	,d.itemno" & vbcrlf
	sql = sql & " 	,Round((d.reducedprice*d.itemno)/(CASE WHEN d.vatinclude='N' THEN 1 ELSE 1.1 END),0) as SuppPrc" & vbcrlf
	sql = sql & " 	,(CASE WHEN R.isCharge='Y' THEN 0.02 ELSE 0 END)" & vbcrlf
	sql = sql & " 	,Round((d.reducedprice*d.itemno)/(CASE WHEN d.vatinclude='N' THEN 1 ELSE 1.1 END),0)*(CASE WHEN R.isCharge='Y' THEN 0.02 ELSE 0 END)" & vbcrlf
	sql = sql & " 	,CASE" & vbcrlf
	sql = sql & " 		WHEN m.jumundiv='9' THEN '반품완료'" & vbcrlf
	sql = sql & " 		WHEN	m.jumundiv='6' THEN '교환완료' ELSE '출고완료' END" & vbcrlf
	sql = sql & " 	from db_order.dbo.tbl_order_master m with (nolock)" & vbcrlf
	sql = sql & " 	Join db_order.dbo.tbl_order_detail d with (nolock)" & vbcrlf
	sql = sql & " 		on m.orderserial=d.orderserial" & vbcrlf
	sql = sql & " 	Join db_item.dbo.tbl_Outmall_RdsiteGubun R with (nolock)" & vbcrlf
	sql = sql & " 		on m.rdsite=R.rdsite" & vbcrlf
	sql = sql & " 		and R.gubun='nvshop'" & vbcrlf
	sql = sql & " 	left join db_jungsan.dbo.tbl_nvshop_jungsan_detail J with (nolock)" & vbcrlf
	sql = sql & " 		on d.orderserial=J.orderserial" & vbcrlf
	sql = sql & " 		and d.itemid=j.itemid" & vbcrlf
	sql = sql & " 		and d.itemoption=j.itemoption" & vbcrlf
	sql = sql & " 	where J.itemoption is NULL" & vbcrlf
	sql = sql & " 	and m.ipkumdiv>3" & vbcrlf
	sql = sql & " 	and m.ipkumdiv=8" & vbcrlf
	sql = sql & " 	and m.cancelyn='N'" & vbcrlf
	sql = sql & " 	and "& jungsandategubun &">@STDT and "& jungsandategubun &"<@EDDT" & vbcrlf
	sql = sql & " 	and d.itemid not in (0,100)" & vbcrlf
	sql = sql & " 	and d.cancelyn<>'Y'" & vbcrlf   '취소는 포함 안함.

	'response.write sql & "<br>"
	dbget.CommandTimeout = 60*5   ' 5분
	dbget.execute sql

	response.write "<script type='text/javascript'>"
	response.write "	alert('"& yyyy &"년 "& mm&"월 naver 수수료정산 작성 완료.');"
	response.write "	parent.location.reload()"
	response.write "</script>"
	dbget.close() : response.end

else
	response.write "<script type='text/javascript'>"
	response.write "	alert('잘못된 경로 입니다.');"
	response.write "	parent.location.reload()"
	response.write "</script>"
	dbget.close() : response.end
end if
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->