<%@  codepage="65001" language="VBScript" %>
<% option explicit %>
<%
session.codePage = 65001
%>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description :  고객데이터 통계(숨은 페이지 이며, scm에 매뉴로 노출 되어 있지 안음.)
' History : 2017.03.10 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->

<%
'if not(session("ssBctId")="tozzinet" or session("ssBctId")="djjung") then
'	response.write "권한이 없습니다. 관리자에게 문의 하세요."
'	session.codePage = 949
'	dbget.close() : response.end
'end if

dim refer
	refer = request.ServerVariables("HTTP_REFERER")
if InStr(refer,"10x10.co.kr")<1 then
	Response.Write "잘못된 접속입니다."
	session.codePage = 949
	dbget.close() : response.end
end if

dim fromDate, toDate, mode, sqlstr, arrlist, bufStr, i, tendb
	fromDate = requestcheckvar(request("fromDate"),10)
	toDate = requestcheckvar(request("toDate"),10)
	mode = requestcheckvar(request("mode"),32)

IF application("Svr_Info")="Dev" THEN
	tendb = "tendb."
end IF

'전체보너스쿠폰데이터
if mode="bonuscoupon" then
	if trim(fromDate="") or trim(toDate="") then
		Response.Write "기간이 없습니다."
		session.codePage = 949
		dbget.close() : response.end
	end if
	fromDate = trim(fromDate)
	toDate = trim(toDate)

	sqlstr = "select top 10000" & vbcrlf
	sqlstr = sqlstr & " SUM(tencardspend) as tencardspend, SUM(USECNT) as USECNT,userlevel,isUser,masteridx,coupontype,couponvalue" & vbcrlf
	sqlstr = sqlstr & " , replace(couponname,char(13)+char(10),'') as couponname, validsitename, evtprize_code" & vbcrlf
	sqlstr = sqlstr & " from (" & vbcrlf
	sqlstr = sqlstr & " 	select T.*" & vbcrlf
	sqlstr = sqlstr & " 	,isNULL(c.masteridx,c2.masteridx) as masteridx" & vbcrlf
	sqlstr = sqlstr & " 	,isNULL(c.coupontype,c2.coupontype) as coupontype" & vbcrlf
	sqlstr = sqlstr & " 	,isNULL(c.couponvalue,c2.couponvalue) as couponvalue" & vbcrlf
	sqlstr = sqlstr & " 	,isNULL(c.couponname,c2.couponname) as couponname" & vbcrlf
	sqlstr = sqlstr & " 	,isNULL(isNULL(c.validsitename,c2.validsitename),'') as validsitename" & vbcrlf
	sqlstr = sqlstr & " 	,isNULL(c.evtprize_code,0) as evtprize_code" & vbcrlf
	sqlstr = sqlstr & " 	from (" & vbcrlf
	sqlstr = sqlstr & " 		select m.bCpnIdx,sum(m.tencardspend) as tencardspend, m.userlevel" & vbcrlf
	sqlstr = sqlstr & " 		, (CASE WHEN isNULL(m.userid,'')='' THEN 0 ELSE 1 END) as isUser" & vbcrlf
	sqlstr = sqlstr & " 		, SUM(CASE WHEN m.jumundiv in (6,9) THEN 0 ELSE 1 END) as USECNT" & vbcrlf
	sqlstr = sqlstr & " 		from "& tendb &"db_order.dbo.tbl_order_master m" & vbcrlf
	sqlstr = sqlstr & " 		where  m.cancelyn='N'" & vbcrlf
	sqlstr = sqlstr & " 		and m.ipkumdiv>3 " & vbcrlf
	sqlstr = sqlstr & " 		and m.regdate>='"& fromDate &"'" & vbcrlf
	sqlstr = sqlstr & " 		and m.regdate<'"& toDate &"'" & vbcrlf
	sqlstr = sqlstr & " 		and m.beadaldiv not in (90)" & vbcrlf
	sqlstr = sqlstr & " 		and m.bCpnIdx is Not NULL" & vbcrlf
	sqlstr = sqlstr & " 		group by m.bCpnIdx, m.userlevel" & vbcrlf
	sqlstr = sqlstr & " 			, (CASE WHEN isNULL(m.userid,'')='' THEN 0 ELSE 1 END)" & vbcrlf
	sqlstr = sqlstr & " 	) T" & vbcrlf
	sqlstr = sqlstr & " 	left join "& tendb &"db_user.dbo.tbl_user_coupon c" & vbcrlf
	sqlstr = sqlstr & " 		on T.bCpnIdx=C.idx" & vbcrlf
	sqlstr = sqlstr & " 	left join "& tendb &"db_log.dbo.tbl_old_user_coupon c2" & vbcrlf
	sqlstr = sqlstr & " 		on T.bCpnIdx=C2.idx" & vbcrlf
	sqlstr = sqlstr & " )TT" & vbcrlf
	sqlstr = sqlstr & " group by userlevel,isUser,masteridx,coupontype,couponvalue,couponname,validsitename, evtprize_code" & vbcrlf
	sqlstr = sqlstr & " order by masteridx, evtprize_code,userlevel" & vbcrlf

	'response.write sqlstr & "<br>"
	db3_rsget.open sqlstr,db3_dbget,1
	If Not db3_rsget.Eof Then
		arrlist = db3_rsget.getrows()
	End If
	db3_rsget.close

'	Response.Buffer=False
	Response.Expires=0
	response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=전체보너스쿠폰데이터.csv"
	Response.CacheControl = "public"

	response.write "쿠폰사용액,건수,userlevel,isUser,masteridx,coupontype,couponvalue,couponname,validsitename,이벤트발행KEY,,등급쿠폰,등급쿠폰-배송비쿠폰포함" & vbcrlf

	if isarray(arrlist) then
		For i = 0 To ubound(arrlist,2)
		bufStr = ""
		bufStr = bufStr & arrlist(0,i)
		bufStr = bufStr & "," & arrlist(1,i)
		bufStr = bufStr & "," & arrlist(2,i)
		bufStr = bufStr & "," & arrlist(3,i)
		bufStr = bufStr & "," & arrlist(4,i)
		bufStr = bufStr & "," & arrlist(5,i)
		bufStr = bufStr & "," & arrlist(6,i)
		bufStr = bufStr & "," & arrlist(7,i)
		bufStr = bufStr & "," & arrlist(8,i)
		bufStr = bufStr & "," & arrlist(9,i)

		response.write bufStr & VbCrlf
		next
	end if

'전시카테고리 전체 리스트
elseif mode="categorylist" then
	sqlstr = "select top 10000" & vbcrlf
	sqlstr = sqlstr & " l.code_large+m.code_mid+s.code_small as catecd, l.code_large AS cdlarge, m.code_mid AS cdmid, s.code_small AS cdsmall" & vbcrlf
	sqlstr = sqlstr & " , replace(replace(replace(l.code_nm,char(9),''),char(10),''),char(13),'') as nmlarge" & vbcrlf
	sqlstr = sqlstr & " , replace(replace(replace(m.code_nm,char(9),''),char(10),''),char(13),'') as nmmid" & vbcrlf
	sqlstr = sqlstr & " , replace(replace(replace(s.code_nm,char(9),''),char(10),''),char(13),'') as nmsmall" & vbcrlf
	sqlstr = sqlstr & " , replace(replace(replace(m.code_nm_eng,char(9),''),char(10),''),char(13),'') as nmmid_eng" & vbcrlf
	sqlstr = sqlstr & " , replace(replace(replace(s.code_nm_eng,char(9),''),char(10),''),char(13),'') as nmsmall_eng" & vbcrlf
	sqlstr = sqlstr & " , replace(replace(replace(m.code_nm_cn_gan,char(9),''),char(10),''),char(13),'') as mid_nm_cn_gan" & vbcrlf
	sqlstr = sqlstr & " , replace(replace(replace(m.code_nm_cn_bun,char(9),''),char(10),''),char(13),'') as mid_nm_cn_bun" & vbcrlf
	sqlstr = sqlstr & " , replace(replace(replace(s.code_nm_cn_gan,char(9),''),char(10),''),char(13),'') as small_nm_cn_gan" & vbcrlf
	sqlstr = sqlstr & " , replace(replace(replace(s.code_nm_cn_bun,char(9),''),char(10),''),char(13),'') as small_nm_cn_bun" & vbcrlf
	sqlstr = sqlstr & " FROM "& tendb &"db_item.[dbo].tbl_Cate_large l" & vbcrlf
	sqlstr = sqlstr & " INNER JOIN "& tendb &"db_item.[dbo].tbl_Cate_mid m" & vbcrlf
	sqlstr = sqlstr & " 	ON l.code_large = m.code_large" & vbcrlf
	sqlstr = sqlstr & " INNER JOIN "& tendb &"db_item.[dbo].tbl_Cate_small s" & vbcrlf
	sqlstr = sqlstr & " 	ON l.code_large = s.code_large" & vbcrlf
	sqlstr = sqlstr & " 	AND m.code_mid = s.code_mid" & vbcrlf
	sqlstr = sqlstr & " order by cdlarge asc, cdsmall asc, nmlarge asc" & vbcrlf

	'response.write sqlstr & "<br>"
	db3_rsget.open sqlstr,db3_dbget,1
	If Not db3_rsget.Eof Then
		arrlist = db3_rsget.getrows()
	End If
	db3_rsget.close
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html;charset=utf-8" />
<style type='text/css'>
	.txt {mso-number-format:'\@'}
</style>
</head>
<td class='txt'>

	<table>
	<tr>
	<td>카테고리코드</td>
	<td>대카테코드</td>
	<td>중카테코드</td>
	<td>소카테코드</td>
	<td>대카테명</td>
	<td>중카테명</td>
	<td>소카테명</td>
	<td>중카테영문명</td>
	<td>소카테영문명</td>
	<td>중카테간자체</td>
	<td>중카테번자체</td>
	<td>소카테간자체</td>
	<td>소카테번자체</td>
	</tr>

	<%
	if isarray(arrlist) then
		For i = 0 To ubound(arrlist,2)
	%>	
		<tr>
			<td class='txt'><%= arrlist(0,i) %></td>
			<td class='txt'><%= arrlist(1,i) %></td>
			<td class='txt'><%= arrlist(2,i) %></td>
			<td class='txt'><%= arrlist(3,i) %></td>
			<td><%= arrlist(4,i) %></td>
			<td><%= arrlist(5,i) %></td>
			<td><%= arrlist(6,i) %></td>
			<td><%= arrlist(7,i) %></td>
			<td><%= arrlist(8,i) %></td>
			<td><%= arrlist(9,i) %></td>
			<td><%= arrlist(10,i) %></td>
			<td><%= arrlist(11,i) %></td>
			<td><%= arrlist(12,i) %></td>
		</tr>
	<%
		next
	end if
	%>
	</table>
</html>
<%
'전시카테고리 전체 다운로드
elseif mode="category" then
	sqlstr = "select top 100" & vbcrlf
	sqlstr = sqlstr & " l.code_large+m.code_mid+s.code_small as catecd, l.code_large AS cdlarge, m.code_mid AS cdmid, s.code_small AS cdsmall" & vbcrlf
	sqlstr = sqlstr & " , l.code_nm AS nmlarge, m.code_nm AS nmmid, s.code_nm AS nmsmall, s.code_nm_eng AS nmsmall_eng" & vbcrlf
	sqlstr = sqlstr & " , s.code_nm_cn_gan, s.code_nm_cn_bun" & vbcrlf
	sqlstr = sqlstr & " FROM "& tendb &"db_item.[dbo].tbl_Cate_large l" & vbcrlf
	sqlstr = sqlstr & " INNER JOIN "& tendb &"db_item.[dbo].tbl_Cate_mid m" & vbcrlf
	sqlstr = sqlstr & " 	ON l.code_large = m.code_large" & vbcrlf
	sqlstr = sqlstr & " INNER JOIN "& tendb &"db_item.[dbo].tbl_Cate_small s" & vbcrlf
	sqlstr = sqlstr & " 	ON l.code_large = s.code_large" & vbcrlf
	sqlstr = sqlstr & " 	AND m.code_mid = s.code_mid" & vbcrlf
	sqlstr = sqlstr & " order by cdlarge asc, cdsmall asc, nmlarge asc" & vbcrlf

	'response.write sqlstr & "<br>"
	db3_rsget.open sqlstr,db3_dbget,1
	If Not db3_rsget.Eof Then
		arrlist = db3_rsget.getrows()
	End If
	db3_rsget.close

'	Response.Buffer=False
	Response.Expires=0
	response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=전체관리카테고리.xls"
	Response.CacheControl = "public"

	response.write "카테고리코드,대카테코드,중카테코드,소카테코드,소카테명,소카테영문명,간자체,번자체" & vbcrlf

	if isarray(arrlist) then
		For i = 0 To ubound(arrlist,2)
		bufStr = ""
		bufStr = bufStr & arrlist(0,i)
		bufStr = bufStr & "," & arrlist(1,i)
		bufStr = bufStr & "," & arrlist(2,i)
		bufStr = bufStr & "," & arrlist(3,i)
		bufStr = bufStr & "," & arrlist(4,i)
		bufStr = bufStr & "," & arrlist(5,i)
		bufStr = bufStr & "," & arrlist(6,i)
		bufStr = bufStr & "," & arrlist(7,i)
		bufStr = bufStr & "," & arrlist(8,i)
		bufStr = bufStr & "," & arrlist(9,i)

		response.write bufStr & VbCrlf
		next
	end if

else
	response.write "잘못된 경로 입니다."
	session.codePage = 949
	dbget.close() : response.end
end if
%>
<%
session.codePage = 949
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->