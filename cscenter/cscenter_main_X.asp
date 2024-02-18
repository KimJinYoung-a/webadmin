<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/upchebeasongcls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/classes/cooperate/cooperateCls.asp"-->
<%

'==============================================================================
dim lastPageTime, pageElapsedTime
lastPageTime = Timer

function checkAndWriteElapsedTime(memo)
	pageElapsedTime = Timer - lastPageTime
	lastPageTime = Timer
	response.write "<!-- Page Execute Time Check : " & FormatNumber(pageElapsedTime, 4) & " : " & memo & " -->" & vbCrLf
end function

dim IsUpdateUserListNeed, IsUpdateOneToOneBoardNeed, IsUpdateMichulgoListUpcheNeed, IsUpdateMichulgoListTenTenNeed, IsUpdateIpjumMichulgoNeed, IsUpdateCSListNeed, IsUpdateUpcheReturnNeed, IsUpdateIpjumStockOutNeed, IsUpdateTenTenStockOutNeed
dim IsUpdateMaxCSMasterIdxNeed

IsUpdateUserListNeed			= False
IsUpdateOneToOneBoardNeed		= False
IsUpdateMichulgoListUpcheNeed	= False
IsUpdateMichulgoListTenTenNeed	= False
IsUpdateIpjumMichulgoNeed		= False
IsUpdateCSListNeed				= False
IsUpdateUpcheReturnNeed			= False
IsUpdateIpjumStockOutNeed		= False
IsUpdateTenTenStockOutNeed		= False
IsUpdateMaxCSMasterIdxNeed		= False

'' IsUpdateUserListNeed			= True
'' IsUpdateOneToOneBoardNeed		= True
'' IsUpdateMichulgoListUpcheNeed	= True
'' IsUpdateMichulgoListTenTenNeed	= True
'' IsUpdateIpjumMichulgoNeed		= True
'' IsUpdateCSListNeed				= True
'' IsUpdateUpcheReturnNeed			= True
'' IsUpdateIpjumStockOutNeed		= True
'' IsUpdateTenTenStockOutNeed		= True
'' IsUpdateMaxCSMasterIdxNeed		= True


'==============================================================================
'// 담당자 목록
If Trim(application("csTimeUserList")) = "" Or Trim(application("csBoardUserArr")) = "" Or DateDiff("s", application("csTimeUserList"), Now() ) > 1800 Then						'// 30분(1800초) 초과시 새로 쿼리함
	application("csTimeUserList") = Now()
	IsUpdateUserListNeed = True
end if

'// 1:1 상담게시판
If Trim(application("csTimeOneToOneBoard")) = "" Or DateDiff("s", application("csTimeOneToOneBoard"), Now() ) > 600 Then			'// 10분(600초) 초과시 새로 쿼리함
	application("csTimeOneToOneBoard") = Now()
	IsUpdateOneToOneBoardNeed = True
end if

'// 미출고리스트[업체배송]
If Trim(application("csTimeMichulgoListUpche")) = "" Or Trim(application("csMiSend")) = "" Or DateDiff("s", application("csTimeMichulgoListUpche"), Now() ) > 900 Then	'// 15분(900초) 초과시 새로 쿼리함
	application("csTimeMichulgoListUpche") = Now()
	IsUpdateMichulgoListUpcheNeed = True
end if

'// 미출고리스트[텐바이텐]
If Trim(application("csTimeMichulgoListTenTen")) = "" Or Trim(application("csTenMiSend")) = "" Or DateDiff("s", application("csTimeMichulgoListTenTen"), Now() ) > 900 Then	'// 15분(900초) 초과시 새로 쿼리함
	application("csTimeMichulgoListTenTen") = Now()
	IsUpdateMichulgoListTenTenNeed = True
end if

'// 미출고리스트[입점몰]
If Trim(application("csTimeIpjumMichulgo")) = "" Or Trim(application("csMiSend")) = "" Or DateDiff("s", application("csTimeIpjumMichulgo"), Now() ) > 900 Then	'// 15분(900초) 초과시 새로 쿼리함
	application("csTimeIpjumMichulgo") = Now()
	IsUpdateIpjumMichulgoNeed = True
end if

'// CS처리리스트
If Trim(application("csTimeCSList")) = "" Or DateDiff("s", application("csTimeCSList"), Now() ) > 900 Then	'// 15분(900초) 초과시 새로 쿼리함
	application("csTimeCSList") = Now()
	IsUpdateCSListNeed = True

	application("csTimeUpcheReturn") = Now()
	IsUpdateUpcheReturnNeed = True
end if

'// 반품접수(업체배송)
If Trim(application("csTimeUpcheReturn")) = "" Or DateDiff("s", application("csTimeUpcheReturn"), Now() ) > 900 Then	'// 15분(900초) 초과시 새로 쿼리함
	application("csTimeCSList") = Now()
	IsUpdateCSListNeed = True

	application("csTimeUpcheReturn") = Now()
	IsUpdateUpcheReturnNeed = True
end if

'// 입점몰 품절취소요청건
If Trim(application("csTimeIpjumStockOut")) = "" Or DateDiff("s", application("csTimeIpjumStockOut"), Now() ) > 900 Then	'// 15분(900초) 초과시 새로 쿼리함
	application("csTimeIpjumStockOut") = Now()
	IsUpdateIpjumStockOutNeed = True
end if

'// 품절취소요청건[텐배+업배]
If Trim(application("csTimeTenTenStockOut")) = "" Or DateDiff("s", application("csTimeTenTenStockOut"), Now() ) > 900 Then	'// 15분(900초) 초과시 새로 쿼리함
	application("csTimeTenTenStockOut") = Now()
	IsUpdateTenTenStockOutNeed = True
end if

'// CS Master Idx
If Trim(application("csTimeMaxCSMasterIdx")) = "" Or DateDiff("s", application("csTimeMaxCSMasterIdx"), Now() ) > 3600 Then			'// 60분(3600초) 초과시 새로 쿼리함
	application("csTimeMaxCSMasterIdx") = Now()
	IsUpdateMaxCSMasterIdxNeed = True
end if

if IsUpdateUserListNeed then
	application("csTimeOneToOneBoard") = Now()
	IsUpdateOneToOneBoardNeed = True

	application("csTimeMichulgoListUpche") = Now()
	IsUpdateMichulgoListUpcheNeed = True

	application("csTimeIpjumMichulgo") = Now()
	IsUpdateIpjumMichulgoNeed = True

	application("csTimeUpcheReturn") = Now()
	IsUpdateUpcheReturnNeed = True

	application("csTimeIpjumStockOut") = Now()
	IsUpdateIpjumStockOutNeed = True

	application("csTimeTenTenStockOut") = Now()
	IsUpdateTenTenStockOutNeed = True
end if


'==============================================================================
dim sqlStr, i, j
dim resultCount
dim sqlchargeidlist

'==============================================================================
Dim paramInfo3, strSql

'==============================================================================
dim onemonthbefore, nowdate1800
dim twomonthbefore, treemonthbefore
nowdate1800 = Left(now, 10) + " 18:00"
onemonthbefore = Left(DateAdd("m", -1, now), 10)
twomonthbefore = Left(DateAdd("m", -2, now), 10)
treemonthbefore= Left(DateAdd("m", -3, now), 10)


'==============================================================================
dim maxCSMasterIdx : maxCSMasterIdx = 0
dim minCSMasterIdxThreeMonth : minCSMasterIdxThreeMonth = 0

if (IsUpdateMaxCSMasterIdxNeed = True) then

	application("csMaxCSMasterIdx") = 0

	sqlStr = " select IsNull(max(m.id), 0) as id "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_cs.dbo.tbl_new_as_list m "

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
	if  not rsget.EOF  then
	    application("csMaxCSMasterIdx") = rsget("id")
	end if
	rsget.close

	Call checkAndWriteElapsedTime("001")

end if

maxCSMasterIdx = application("csMaxCSMasterIdx")

if (IsUpdateMaxCSMasterIdxNeed = True) then

	application("csMinCSMasterIdxThreeMonth") = 0

	sqlStr = " select top 1 m.id "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_cs.dbo.tbl_new_as_list m "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and datediff(m, m.regdate, getdate()) < 3 "
	sqlStr = sqlStr + " 	and m.id > " & (maxCSMasterIdx - 200000) & " "			'// 최근 20만개만 검색한다.(속도문제)
	sqlStr = sqlStr + " order by m.id "

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
	if  not rsget.EOF  then
	    application("csMinCSMasterIdxThreeMonth") = rsget("id")
	end if
	rsget.close

	Call checkAndWriteElapsedTime("002")

end if

minCSMasterIdxThreeMonth = application("csMinCSMasterIdxThreeMonth")




'==============================================================================
'// 근무중인 담당자 목록
dim csBoardUserArr, csMichulgoBoardUserArr, csMichulgoStockoutBoardUserArr, csReturnBoardUserArr

if (IsUpdateUserListNeed = True) then

	'// YN 이 N 이 아닌것이어야 한다.
	'// 분배할 때는 YN 이 Y 인것, 분배받은것 표시할때는 N 이 아닌것!!
	'// YN = T 인 경우 : 분배정지(분배받은것은 표시되고, 더이상 분배받지 않는다.)
	sqlStr = " select "
	sqlStr = sqlStr + " 	u.userid "
	sqlStr = sqlStr + " 	, (case when IsNull(u.one2oneyn, 'Y') <> 'N' then 'Y' else 'N' end) as one2oneyn "
	sqlStr = sqlStr + " 	, (case when IsNull(u.michulgoyn, 'Y') <> 'N' then 'Y' else 'N' end) as michulgoyn "
	sqlStr = sqlStr + " 	, (case when IsNull(u.stockoutyn, 'Y') <> 'N' then 'Y' else 'N' end) as stockoutyn "
	sqlStr = sqlStr + " 	, (case when IsNull(u.returnyn, 'Y') <> 'N' then 'Y' else 'N' end) as returnyn "
	sqlStr = sqlStr + " from db_cs.dbo.tbl_cs_board_user u "
	sqlStr = sqlStr + " left join ( "
	sqlStr = sqlStr + " 	select "
	sqlStr = sqlStr + " 		distinct T.userid, 'Y' as vacationyn "
	sqlStr = sqlStr + " 	from "
	sqlStr = sqlStr + " 		( "
	sqlStr = sqlStr + " 			select "
	sqlStr = sqlStr + " 				m.userid "
	sqlStr = sqlStr + " 				, (case when d.halfgubun = 'pm' then DateAdd(hh, 12, d.startday) else DateAdd(hh, -6, d.startday) end) as startday "
	sqlStr = sqlStr + " 				, (case when d.halfgubun = 'am' then DateAdd(hh, -12, d.endday) else DateAdd(hh, -6, d.endday) end) as endday "
	sqlStr = sqlStr + " 			from "
	sqlStr = sqlStr + " 			db_partner.dbo.tbl_vacation_master m "
	sqlStr = sqlStr + " 			join db_partner.dbo.tbl_vacation_detail d "
	sqlStr = sqlStr + " 			on "
	sqlStr = sqlStr + " 				m.idx = d.masteridx "
	sqlStr = sqlStr + " 			where "
	sqlStr = sqlStr + " 				1 = 1 "
	sqlStr = sqlStr + " 				and m.deleteyn = 'N' "
	sqlStr = sqlStr + " 				and d.deleteyn = 'N' "
	sqlStr = sqlStr + " 				and d.statedivcd <> 'D' "
	sqlStr = sqlStr + " 				and d.endday >= getdate() "
	sqlStr = sqlStr + " 		) T "
	sqlStr = sqlStr + " 	where "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and T.startday >= getdate() "
	sqlStr = sqlStr + " 		and T.endday < getdate() "
	sqlStr = sqlStr + " ) V "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	u.userid = V.userid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and IsNull(V.vacationyn, 'N') = 'N' "
	sqlStr = sqlStr + " 	and u.useyn = 'Y' "
	sqlStr = sqlStr + " order by indexno "

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

	'// 에러방지(담당자가 한명도 없으면 에러가 발생한다.)
	csBoardUserArr = "xxxxxxxx"
	csMichulgoBoardUserArr = "xxxxxxxx"
	csMichulgoStockoutBoardUserArr = "xxxxxxxx"
	csReturnBoardUserArr = "xxxxxxxx"

	if  not rsget.EOF  then
		do until rsget.eof

			if (rsget("one2oneyn") = "Y") then
				csBoardUserArr = csBoardUserArr + "," + rsget("userid")
			end if

			if (rsget("michulgoyn") = "Y") then
				csMichulgoBoardUserArr = csMichulgoBoardUserArr + "," + rsget("userid")
			end if

			if (rsget("stockoutyn") = "Y") then
				csMichulgoStockoutBoardUserArr = csMichulgoStockoutBoardUserArr + "," + rsget("userid")
			end if

			if (rsget("returnyn") = "Y") then
				csReturnBoardUserArr = csReturnBoardUserArr + "," + rsget("userid")
			end if

			rsget.MoveNext
	    loop
	end if
	rsget.close

	application("csBoardUserArr") 					= csBoardUserArr
	application("csMichulgoBoardUserArr") 			= csMichulgoBoardUserArr
	application("csMichulgoStockoutBoardUserArr") 	= csMichulgoStockoutBoardUserArr
	application("csReturnBoardUserArr") 			= csReturnBoardUserArr

	Call checkAndWriteElapsedTime("003")

end if

csBoardUserArr = Split(application("csBoardUserArr"), ",")
csMichulgoBoardUserArr = Split(application("csMichulgoBoardUserArr"), ",")
csMichulgoStockoutBoardUserArr = Split(application("csMichulgoStockoutBoardUserArr"), ",")
csReturnBoardUserArr = Split(application("csReturnBoardUserArr"), ",")


'==============================================================================
'// 1:1 상담게시판 관리(담당자별)
dim csBoardChargecntArr, csBoardNochargecnt

if (IsUpdateOneToOneBoardNeed = True) then

	csBoardChargecntArr = ""
	csBoardNochargecnt = 0

	sqlchargeidlist = ""
	for i = 0 to UBound(csBoardUserArr)
		if (sqlchargeidlist = "") then
			sqlchargeidlist = "'" + CStr(csBoardUserArr(i)) + "'"
		else
			sqlchargeidlist = sqlchargeidlist + ",'" + CStr(csBoardUserArr(i)) + "'"
		end if
	next

	sqlStr = " select "
	for i = 0 to UBound(csBoardUserArr)
		sqlStr = sqlStr + "     IsNull(sum(case when (IsNull(chargeid,'') = '" + CStr(csBoardUserArr(i)) + "') then 1 else 0 end), 0) as chargeid" + CStr(i) + ", "
	next
	sqlStr = sqlStr + "     IsNull(sum(case when (IsNull(chargeid,'') not in (" + sqlchargeidlist + ")) then 1 else 0 end), 0) as nochargeid "
	sqlStr = sqlStr + " from [db_cs].[dbo].tbl_myqna "
	sqlStr = sqlStr + " where 1 = 1 "
	sqlStr = sqlStr + " and isusing = 'Y' "
	sqlStr = sqlStr + " and IsNull(replyuser,'') = '' "
	sqlStr = sqlStr + " and regdate > '" + onemonthbefore + "' "
	'response.write sqlStr

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
	if  not rsget.EOF  then
		for i = 0 to UBound(csBoardUserArr)
			if (csBoardChargecntArr = "") then
				csBoardChargecntArr = rsget("chargeid" + CStr(i))
			else
				csBoardChargecntArr = csBoardChargecntArr & "|" & rsget("chargeid" + CStr(i))
			end if
		next
		csBoardNochargecnt = rsget("nochargeid")
	end if
	rsget.close

	application("csBoardChargecntArr") = csBoardChargecntArr
	application("csBoardNochargecnt") = csBoardNochargecnt

	Call checkAndWriteElapsedTime("004")

end if

csBoardChargecntArr = Split(application("csBoardChargecntArr"), "|")
csBoardNochargecnt = application("csBoardNochargecnt")


'==============================================================================
' 미출고리스트[업체배송] - 통계
Dim strMiSend, arrMiSend

if (IsUpdateMichulgoListUpcheNeed = True) then

	strMisend = "0|0|0"

	paramInfo3 = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
		,Array("@cdl"			, adVarchar	, adParamInput	, 3		, "")	_
	)
	strSql = "[db_order].[dbo].sp_Ten_MiSend_Count"
	Call fnExecSPReturnRSOutput(strSql, paramInfo3)

	If Not rsget.EOF Then
		strMisend	= rsget(0) & "|" & rsget(1) & "|" & rsget(2)
	End If
	rsget.close()

	application("csMiSend") = strMisend

end if

arrMiSend = Split(application("csMiSend"), "|")


'==============================================================================
'미출고리스트[업체배송] - 근무일수 기준 D+3 일
dim tmpSql, michulgoBaseDate

tmpSql = " exec [db_cs].[dbo].[usp_getDayMinusWorkday] '" & Left(now(), 10) & "', 3 " & VbCRLF
rsget.CursorLocation = adUseClient
rsget.Open tmpSql, dbget, adOpenForwardOnly
if Not rsget.Eof then
    '// 근무일수 기준 D+3 일
    michulgoBaseDate = rsget("minusworkday")
end if
rsget.close


'==============================================================================
'미출고리스트[업체배송] - 담당자별 서머리
dim csMichulgoBrandCountByUserid, csMichulgoOrderCountByUserid, nochargeMichulgoBrandcnt

if (IsUpdateMichulgoListUpcheNeed = True) then

	csMichulgoBrandCountByUserid = ""
	csMichulgoOrderCountByUserid = ""
	nochargeMichulgoBrandcnt = "0"

	''sqlchargeidlist = "'xxxxxxxxx'"
	sqlchargeidlist = ""
	for i = 0 to UBound(csMichulgoBoardUserArr)
		if (sqlchargeidlist = "") then
			sqlchargeidlist = "'" + CStr(csMichulgoBoardUserArr(i)) + "'"
		else
			sqlchargeidlist = sqlchargeidlist + ",'" + CStr(csMichulgoBoardUserArr(i)) + "'"
		end if
	next

	sqlStr = " select "
	for i = 0 to UBound(csMichulgoBoardUserArr)
		sqlStr = sqlStr + "     IsNull(sum(case when (IsNull(T.userid,'') = '" + CStr(csMichulgoBoardUserArr(i)) + "') then T.makeridcnt else 0 end), 0) as chargeid" + CStr(i) + ", "
		sqlStr = sqlStr + "     IsNull(sum(case when (IsNull(T.userid,'') = '" + CStr(csMichulgoBoardUserArr(i)) + "') then T.ordercnt else 0 end), 0) as chargeidorder" + CStr(i) + ", "
	next
	sqlStr = sqlStr + "     IsNull(sum(case when (IsNull(T.userid,'') not in (" + sqlchargeidlist + ")) then T.makeridcnt else 0 end), 0) as nochargeid "
	sqlStr = sqlStr + " FROM "
	sqlStr = sqlStr + " 	( "
	sqlStr = sqlStr + " 		SELECT "
	sqlStr = sqlStr + " 			u.userid, count(distinct d.makerid) as makeridcnt, count(d.itemid) as ordercnt "
	sqlStr = sqlStr + " 		FROM "
	sqlStr = sqlStr + " 			[db_order].[dbo].tbl_order_master m "
	sqlStr = sqlStr + " 			JOIN [db_order].[dbo].tbl_order_detail d "
	sqlStr = sqlStr + " 			on "
	sqlStr = sqlStr + " 				m.orderserial=d.orderserial "
	sqlStr = sqlStr + " 			LEFT JOIN [db_temp].dbo.tbl_mibeasong_list T "
	sqlStr = sqlStr + " 			on "
	sqlStr = sqlStr + " 				1 = 1 "
	sqlStr = sqlStr + " 				and d.orderserial=T.orderserial "
	sqlStr = sqlStr + " 				and d.idx=T.detailidx "
	sqlStr = sqlStr + " 			LEFT JOIN db_cs.dbo.tbl_cs_michulgo_upche_brand u "
	sqlStr = sqlStr + " 			on "
	sqlStr = sqlStr + " 				d.makerid = u.brandid "
	sqlStr = sqlStr + " 		WHERE "
	sqlStr = sqlStr + " 			m.regdate >= DateAdd(m, -2, getdate()) "
	sqlStr = sqlStr + " 			and m.ipkumdiv < '8' "
	sqlStr = sqlStr + " 			and m.ipkumdiv > '3' "
	sqlStr = sqlStr + " 			and m.cancelyn = 'N' "
	sqlStr = sqlStr + " 			and m.jumundiv <> '9' "
	sqlStr = sqlStr + " 			and d.itemid<>0 "
	sqlStr = sqlStr + " 			and IsNull(d.currstate,'0')<'7' "
	'sqlStr = sqlStr + " 			and d.currstate='3' "
	sqlStr = sqlStr + " 			and d.isupchebeasong='Y' "
	sqlStr = sqlStr + " 			and m.sitename = '10x10' "																'// 입점몰주문 제외
	sqlStr = sqlStr + " 			and d.cancelyn <> 'Y' "
	''sqlStr = sqlStr + " 			and datediff(d,m.baljudate,getdate())>=4 "
	sqlStr = sqlStr + " 			and datediff(d,m.baljudate,'" + CStr(michulgoBaseDate) + "')>=0 "						'// 근무일수 기준 D+3 일
	sqlStr = sqlStr + " 			and not ((T.ipgodate is not null) and (datediff(d, T.ipgodate, getdate()) <= 0)) "		'// 출고예정일 이전 주문 제외
	sqlStr = sqlStr + " 			and IsNULL(T.code,'00')<>'05' "															'// 품절출고불가 제외
	sqlStr = sqlStr + " 		GROUP BY "
	sqlStr = sqlStr + " 			u.userid "
	sqlStr = sqlStr + " 	) T "
	''rw sqlStr

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
	if  not rsget.EOF  then
		for i = 0 to UBound(csMichulgoBoardUserArr)
			if (csMichulgoBrandCountByUserid = "") then
				csMichulgoBrandCountByUserid = rsget("chargeid" + CStr(i))
			else
				csMichulgoBrandCountByUserid = csMichulgoBrandCountByUserid & "|" & rsget("chargeid" + CStr(i))
			end if

			if (csMichulgoOrderCountByUserid = "") then
				csMichulgoOrderCountByUserid = rsget("chargeidorder" + CStr(i))
			else
				csMichulgoOrderCountByUserid = csMichulgoOrderCountByUserid & "|" & rsget("chargeidorder" + CStr(i))
			end if
		next
		nochargeMichulgoBrandcnt = rsget("nochargeid")
	end if
	rsget.close

	application("csMichulgoBrandCountByUserid") = csMichulgoBrandCountByUserid
	application("csMichulgoOrderCountByUserid") = csMichulgoOrderCountByUserid
	application("nochargeMichulgoBrandcnt") = nochargeMichulgoBrandcnt

end if

csMichulgoBrandCountByUserid = Split(application("csMichulgoBrandCountByUserid"), "|")
csMichulgoOrderCountByUserid = Split(application("csMichulgoOrderCountByUserid"), "|")
nochargeMichulgoBrandcnt = application("nochargeMichulgoBrandcnt")


'==============================================================================
'미출고리스트[업체배송] - 브랜드별 서머리
dim csMichulgoBrandUseridArr, csMichulgoBrandNameArr, csMichulgoBrandcntArr

if (IsUpdateMichulgoListUpcheNeed = True) then

	csMichulgoBrandUseridArr = ""
	csMichulgoBrandNameArr = ""
	csMichulgoBrandcntArr = ""

	resultCount = 0

	sqlStr = " SELECT "
	sqlStr = sqlStr + " 	u.userid, d.makerid, count(d.itemid) as cnt "
	sqlStr = sqlStr + " FROM "
	sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_master m "
	sqlStr = sqlStr + " 	JOIN [db_order].[dbo].tbl_order_detail d "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		m.orderserial=d.orderserial "
	sqlStr = sqlStr + " 	LEFT JOIN [db_temp].dbo.tbl_mibeasong_list T "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and d.orderserial=T.orderserial "
	sqlStr = sqlStr + " 		and d.idx=T.detailidx "
	sqlStr = sqlStr + " 	LEFT JOIN db_cs.dbo.tbl_cs_michulgo_upche_brand u "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		d.makerid = u.brandid "
	sqlStr = sqlStr + " WHERE "
	sqlStr = sqlStr + " 	m.regdate >= DateAdd(m, -2, getdate()) "
	sqlStr = sqlStr + " 	and u.userid is not NULL "
	sqlStr = sqlStr + " 	and m.ipkumdiv < '8' "
	sqlStr = sqlStr + " 	and m.ipkumdiv > '3' "
	sqlStr = sqlStr + " 	and m.cancelyn = 'N' "
	sqlStr = sqlStr + " 	and m.jumundiv <> '9' "
	sqlStr = sqlStr + " 	and d.itemid<>0 "
	sqlStr = sqlStr + " 	and IsNull(d.currstate,'0')<'7' "
	'sqlStr = sqlStr + " 	and d.currstate='3' "
	sqlStr = sqlStr + " 	and d.isupchebeasong='Y' "
	sqlStr = sqlStr + " 	and m.sitename = '10x10' "																		'// 입점몰주문 제외
	sqlStr = sqlStr + " 	and d.cancelyn <> 'Y' "
	''sqlStr = sqlStr + " 	and datediff(d,m.baljudate,getdate())>=4 "
	sqlStr = sqlStr + " 	and datediff(d,m.baljudate,'" + CStr(michulgoBaseDate) + "')>=0 "								'// 근무일수 기준 D+3 일
	sqlStr = sqlStr + " 	and not ((T.ipgodate is not null) and (datediff(d, T.ipgodate, getdate()) <= 0)) "				'// 출고예정일 이전 주문 제외
	sqlStr = sqlStr + " 	and IsNULL(T.code,'00')<>'05' "																	'// 품절출고불가 제외
	sqlStr = sqlStr + " GROUP BY "
	sqlStr = sqlStr + " 	u.userid, d.makerid "
	sqlStr = sqlStr + " ORDER BY "
	sqlStr = sqlStr + " 	u.userid, count(d.itemid) desc, d.makerid "
	''response.write sqlStr

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

	if not rsget.EOF  then
		resultCount = rsget.RecordCount

		for i = 0 to resultCount - 1
			if (csMichulgoBrandUseridArr = "") then
				csMichulgoBrandUseridArr = rsget("userid")
			else
				csMichulgoBrandUseridArr = csMichulgoBrandUseridArr & "|" & rsget("userid")
			end if

			if (csMichulgoBrandNameArr = "") then
				csMichulgoBrandNameArr = rsget("makerid")
			else
				csMichulgoBrandNameArr = csMichulgoBrandNameArr & "|" & rsget("makerid")
			end if

			if (csMichulgoBrandcntArr = "") then
				csMichulgoBrandcntArr = rsget("cnt")
			else
				csMichulgoBrandcntArr = csMichulgoBrandcntArr & "|" & rsget("cnt")
			end if

			rsget.movenext
		next
	end if
	rsget.close

	application("csMichulgoBrandUseridArr") = csMichulgoBrandUseridArr
	application("csMichulgoBrandNameArr") = csMichulgoBrandNameArr
	application("csMichulgoBrandcntArr") = csMichulgoBrandcntArr

	Call checkAndWriteElapsedTime("005")

end if

csMichulgoBrandUseridArr = Split(application("csMichulgoBrandUseridArr"), "|")
csMichulgoBrandNameArr = Split(application("csMichulgoBrandNameArr"), "|")
csMichulgoBrandcntArr = Split(application("csMichulgoBrandcntArr"), "|")


'==============================================================================
'미출고리스트[입점몰,텐배+업배] - 입점몰 리스트
dim csMichulgoOtherSitenameArr, csMichulgoOtherSiteOrderCountArr, csMichulgoOtherSiteOrderDetailCountArr

if (IsUpdateIpjumMichulgoNeed = True) then

	csMichulgoOtherSitenameArr = ""
	csMichulgoOtherSiteOrderCountArr = ""
	csMichulgoOtherSiteOrderDetailCountArr = ""

	resultCount = 0

	sqlStr = " SELECT "
	sqlStr = sqlStr + " 	m.sitename, count(distinct m.orderserial) as ordercnt, count(d.idx) as cnt "
	sqlStr = sqlStr + " FROM "
	sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_master m "
	sqlStr = sqlStr + " 	JOIN [db_order].[dbo].tbl_order_detail d "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		m.orderserial=d.orderserial "
	sqlStr = sqlStr + " 	JOIN [db_temp].dbo.tbl_mibeasong_list T "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and d.orderserial=T.orderserial "
	sqlStr = sqlStr + " 		and d.idx=T.detailidx "
	sqlStr = sqlStr + " WHERE "
	sqlStr = sqlStr + " 	m.regdate >= DateAdd(m, -2, getdate()) "
	sqlStr = sqlStr + " 	and m.ipkumdiv < '8' "
	sqlStr = sqlStr + " 	and m.ipkumdiv > '3' "
	sqlStr = sqlStr + " 	and m.cancelyn = 'N' "
	sqlStr = sqlStr + " 	and m.jumundiv <> '9' "
	sqlStr = sqlStr + " 	and d.itemid<>0 "
	sqlStr = sqlStr + " 	and IsNull(d.currstate,'0')<'7' "
	'sqlStr = sqlStr + " 	and d.currstate='3' "
	sqlStr = sqlStr + " 	and m.sitename <> '10x10' "																		'// 입점몰주문
	'sqlStr = sqlStr + " 	and d.isupchebeasong='Y' "
	sqlStr = sqlStr + " 	and d.cancelyn <> 'Y' "
	''sqlStr = sqlStr + " 	and datediff(d,m.baljudate,getdate())>=4 "
	sqlStr = sqlStr + " 	and datediff(d,m.baljudate,'" + CStr(michulgoBaseDate) + "')>=0 "								'// 근무일수 기준 D+3 일
	sqlStr = sqlStr + " 	and not ((T.ipgodate is not null) and (datediff(d, T.ipgodate, getdate()) <= 0)) "				'// 출고예정일 이전 주문 제외
	sqlStr = sqlStr + " 	and IsNULL(T.code,'00')<>'05' "																	'// 품절출고불가 제외
	sqlStr = sqlStr + " GROUP BY "
	sqlStr = sqlStr + " 	m.sitename "
	sqlStr = sqlStr + " ORDER BY "
	sqlStr = sqlStr + " 	m.sitename "
	''response.write sqlStr

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

	if not rsget.EOF  then

		resultCount = rsget.RecordCount

		for i = 0 to resultCount - 1
			if (csMichulgoOtherSitenameArr = "") then
				csMichulgoOtherSitenameArr = rsget("sitename")
			else
				csMichulgoOtherSitenameArr = csMichulgoOtherSitenameArr & "|" & rsget("sitename")
			end if

			if (csMichulgoOtherSiteOrderCountArr = "") then
				csMichulgoOtherSiteOrderCountArr = rsget("ordercnt")
			else
				csMichulgoOtherSiteOrderCountArr = csMichulgoOtherSiteOrderCountArr & "|" & rsget("ordercnt")
			end if

			if (csMichulgoOtherSiteOrderDetailCountArr = "") then
				csMichulgoOtherSiteOrderDetailCountArr = rsget("cnt")
			else
				csMichulgoOtherSiteOrderDetailCountArr = csMichulgoOtherSiteOrderDetailCountArr & "|" & rsget("cnt")
			end if

			rsget.movenext
		next
	end if
	rsget.close

	application("csMichulgoOtherSitenameArr") = csMichulgoOtherSitenameArr
	application("csMichulgoOtherSiteOrderCountArr") = csMichulgoOtherSiteOrderCountArr
	application("csMichulgoOtherSiteOrderDetailCountArr") = csMichulgoOtherSiteOrderDetailCountArr

	Call checkAndWriteElapsedTime("013")

end if

csMichulgoOtherSitenameArr = Split(application("csMichulgoOtherSitenameArr"), "|")
csMichulgoOtherSiteOrderCountArr = Split(application("csMichulgoOtherSiteOrderCountArr"), "|")
csMichulgoOtherSiteOrderDetailCountArr = Split(application("csMichulgoOtherSiteOrderDetailCountArr"), "|")


'==============================================================================
'미출고리스트[입점몰,텐배+업배] - 입점몰 리스트 : 주문별 서머리
dim csMichulgoOtherSiteOrderSitenameArr, csMichulgoOtherSiteOrderserialArr, csMichulgoOtherSiteOrderdetailCntcntArr, csMichulgoOtherSiteOrderdetaiidx

if (IsUpdateIpjumMichulgoNeed = True) then

	csMichulgoOtherSiteOrderSitenameArr = ""
	csMichulgoOtherSiteOrderserialArr = ""
	csMichulgoOtherSiteOrderdetailCntcntArr = ""
	csMichulgoOtherSiteOrderdetaiidx = ""

	resultCount = 0

	sqlStr = " SELECT "
	sqlStr = sqlStr + " 	m.sitename, m.orderserial, max(d.idx) as orderdetailidx, count(d.itemid) as cnt "
	sqlStr = sqlStr + " FROM "
	sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_master m "
	sqlStr = sqlStr + " 	JOIN [db_order].[dbo].tbl_order_detail d "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		m.orderserial=d.orderserial "
	sqlStr = sqlStr + " 	JOIN [db_temp].dbo.tbl_mibeasong_list T "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and d.orderserial=T.orderserial "
	sqlStr = sqlStr + " 		and d.idx=T.detailidx "
	sqlStr = sqlStr + " WHERE "
	sqlStr = sqlStr + " 	m.regdate >= DateAdd(m, -2, getdate()) "
	sqlStr = sqlStr + " 	and m.ipkumdiv < '8' "
	sqlStr = sqlStr + " 	and m.ipkumdiv > '3' "
	sqlStr = sqlStr + " 	and m.cancelyn = 'N' "
	sqlStr = sqlStr + " 	and m.jumundiv <> '9' "
	sqlStr = sqlStr + " 	and d.itemid<>0 "
	sqlStr = sqlStr + " 	and IsNull(d.currstate,'0')<'7' "
	'sqlStr = sqlStr + " 	and d.currstate='3' "
	sqlStr = sqlStr + " 	and m.sitename <> '10x10' "																		'// 입점몰주문
	'sqlStr = sqlStr + " 	and d.isupchebeasong='Y' "
	sqlStr = sqlStr + " 	and d.cancelyn <> 'Y' "
	''sqlStr = sqlStr + " 	and datediff(d,m.baljudate,getdate())>=4 "
	sqlStr = sqlStr + " 	and datediff(d,m.baljudate,'" + CStr(michulgoBaseDate) + "')>=0 "								'// 근무일수 기준 D+3 일
	sqlStr = sqlStr + " 	and not ((T.ipgodate is not null) and (datediff(d, T.ipgodate, getdate()) <= 0)) "				'// 출고예정일 이전 주문 제외
	sqlStr = sqlStr + " 	and IsNULL(T.code,'00')<>'05' "																	'// 품절출고불가 제외
	sqlStr = sqlStr + " GROUP BY "
	sqlStr = sqlStr + " 	m.sitename, m.orderserial "
	sqlStr = sqlStr + " ORDER BY "
	sqlStr = sqlStr + " 	m.sitename, m.orderserial "
	''response.write sqlStr

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

	if not rsget.EOF  then

		resultCount = rsget.RecordCount

		for i = 0 to resultCount - 1
			if (csMichulgoOtherSiteOrderSitenameArr = "") then
				csMichulgoOtherSiteOrderSitenameArr = rsget("sitename")
			else
				csMichulgoOtherSiteOrderSitenameArr = csMichulgoOtherSiteOrderSitenameArr & "|" & rsget("sitename")
			end if

			if (csMichulgoOtherSiteOrderserialArr = "") then
				csMichulgoOtherSiteOrderserialArr = rsget("orderserial")
			else
				csMichulgoOtherSiteOrderserialArr = csMichulgoOtherSiteOrderserialArr & "|" & rsget("orderserial")
			end if

			if (csMichulgoOtherSiteOrderdetailCntcntArr = "") then
				csMichulgoOtherSiteOrderdetailCntcntArr = rsget("cnt")
			else
				csMichulgoOtherSiteOrderdetailCntcntArr = csMichulgoOtherSiteOrderdetailCntcntArr & "|" & rsget("cnt")
			end if

			if (csMichulgoOtherSiteOrderdetaiidx = "") then
				csMichulgoOtherSiteOrderdetaiidx = rsget("orderdetailidx")
			else
				csMichulgoOtherSiteOrderdetaiidx = csMichulgoOtherSiteOrderdetaiidx & "|" & rsget("orderdetailidx")
			end if

			rsget.movenext
		next
	end if
	rsget.close

	application("csMichulgoOtherSiteOrderSitenameArr") 		= csMichulgoOtherSiteOrderSitenameArr
	application("csMichulgoOtherSiteOrderserialArr") 		= csMichulgoOtherSiteOrderserialArr
	application("csMichulgoOtherSiteOrderdetailCntcntArr") 	= csMichulgoOtherSiteOrderdetailCntcntArr
	application("csMichulgoOtherSiteOrderdetaiidx") 		= csMichulgoOtherSiteOrderdetaiidx

	Call checkAndWriteElapsedTime("014")

end if

csMichulgoOtherSiteOrderSitenameArr 	= Split(application("csMichulgoOtherSiteOrderSitenameArr"), "|")
csMichulgoOtherSiteOrderserialArr 		= Split(application("csMichulgoOtherSiteOrderserialArr"), "|")
csMichulgoOtherSiteOrderdetailCntcntArr = Split(application("csMichulgoOtherSiteOrderdetailCntcntArr"), "|")
csMichulgoOtherSiteOrderdetaiidx 		= Split(application("csMichulgoOtherSiteOrderdetaiidx"), "|")


'==============================================================================
' 미출고리스트 텐바이텐
Dim strTenMiSend, arrTenMiSend

if (IsUpdateMichulgoListTenTenNeed = True) then


	strSql = "[db_order].[dbo].sp_Ten_TenMiSend_Count"
	Call fnExecSPReturnRSOutput(strSql, "")

	If Not rsget.EOF Then
		strTenMiSend	= rsget(0) & "|" & rsget(1) & "|" & rsget(2) & "|" & rsget(3)
	End If
	rsget.close()

	application("csTenMiSend") = strTenMiSend

	Call checkAndWriteElapsedTime("006")

end if

arrTenMiSend = Split(application("csTenMiSend"), "|")


'==============================================================================
'CS처리리스트 관리
''환불 미처리(A003), 마일리지 환불 미처리, 카드취소미처리(A007), 주문취소미처리(A008),
''출고시유의사항, 업체미처리, 업체처리완료, 회수요청미처리, 확인요청, 외부몰환불미처리
dim csRefundRequestRegCount, csRefundRequestConfirmCount, csRequestCardCancelCount, csReturnNotFinish
dim csNotFinA008, csNotFinA005, csUpcheNotFin

if (IsUpdateCSListNeed = True) or (IsUpdateUpcheReturnNeed = True) then

	csRefundRequestRegCount			= 0
	csRefundRequestConfirmCount		= 0
	csRequestCardCancelCount		= 0
	csReturnNotFinish				= 0
	csNotFinA008					= 0
	csNotFinA005					= 0
	csUpcheNotFin					= 0

	sqlStr = " select "
	sqlStr = sqlStr + "     sum(case when (divcd = 'A003') and (currstate='B001') then 1 else 0 end) as csRefundRequestRegCount, "
	sqlStr = sqlStr + "     sum(case when (divcd = 'A003') and (currstate='B005') then 1 else 0 end) as csRefundRequestConfirmCount, "
	sqlStr = sqlStr + "     sum(case when (divcd = 'A007') then 1 else 0 end) as csRequestCardCancelCount, "
	sqlStr = sqlStr + "     sum(case when (divcd = 'A005') then 1 else 0 end) as csNotFinA005, "
	sqlStr = sqlStr + "     sum(case when (divcd = 'A008') and (regdate>'2008-04-23') then 1 else 0 end) as csNotFinA008, "
	sqlStr = sqlStr + "     sum(case when (requireupche='Y') and (currstate<'B006') then 1 else 0 end) as csUpcheNotFin, "
	sqlStr = sqlStr + "     sum(case when (divcd in ('A010', 'A011', 'A111')) and (currstate < 'B002') then 1 else 0 end) as csReturnNotFinish "
	sqlStr = sqlStr + " from [db_cs].[dbo].tbl_new_as_list "
	sqlStr = sqlStr + " where 1=1 "
	sqlStr = sqlStr + " and deleteyn = 'N' "
	sqlStr = sqlStr + " and currstate < 'B007'"
	sqlStr = sqlStr + " and regdate>'"&treemonthbefore&"'"        ''한달지난 접수건은 잡에서 삭제X =>반품접수건만.

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

	if  not rsget.EOF  then
	    csRefundRequestRegCount     = rsget("csRefundRequestRegCount")
	    csRefundRequestConfirmCount = rsget("csRefundRequestConfirmCount")
	    csRequestCardCancelCount 	= rsget("csRequestCardCancelCount")
	    csNotFinA008        		= rsget("csNotFinA008")

	    csUpcheNotFin       		= rsget("csUpcheNotFin")
	    csNotFinA005        		= rsget("csNotFinA005")
	    csReturnNotFinish  			= rsget("csReturnNotFinish")
	end if
	rsget.close

	application("csRefundRequestRegCount") 		= csRefundRequestRegCount
	application("csRefundRequestConfirmCount") 	= csRefundRequestConfirmCount
	application("csRequestCardCancelCount") 	= csRequestCardCancelCount
	application("csNotFinA008") 				= csNotFinA008

	application("csUpcheNotFin") 				= csUpcheNotFin
	application("csNotFinA005") 				= csNotFinA005
	application("csReturnNotFinish") 			= csReturnNotFinish

	Call checkAndWriteElapsedTime("007")

end if

csRefundRequestRegCount = application("csRefundRequestRegCount")
csRefundRequestConfirmCount = application("csRefundRequestConfirmCount")
csRequestCardCancelCount = application("csRequestCardCancelCount")
csNotFinA008 = application("csNotFinA008")

csUpcheNotFin = application("csUpcheNotFin")
csNotFinA005 = application("csNotFinA005")
csReturnNotFinish = application("csReturnNotFinish")


'==============================================================================
dim upReturnMiFinishBaseDate7, upReturnMiFinishBaseDate3

tmpSql = " exec [db_cs].[dbo].[usp_getDayMinusWorkday] '" & Left(now(), 10) & "', 7 " & VbCRLF
rsget.CursorLocation = adUseClient
rsget.Open tmpSql, dbget, adOpenForwardOnly
if Not rsget.Eof then
    '// 근무일수 기준 D+7 일
    upReturnMiFinishBaseDate7 = rsget("minusworkday")
end if
rsget.close

tmpSql = " exec [db_cs].[dbo].[usp_getDayMinusWorkday] '" & Left(now(), 10) & "', 3 " & VbCRLF
rsget.CursorLocation = adUseClient
rsget.Open tmpSql, dbget, adOpenForwardOnly
if Not rsget.Eof then
    '// 근무일수 기준 D+3 일
    upReturnMiFinishBaseDate3 = rsget("minusworkday")
end if
rsget.close


'==============================================================================
'반품접수(업체배송) D+7 미처리
dim CSUpcheReturnNotFinish7, CSUpcheReturnNotFinish3

if (IsUpdateCSListNeed = True) or (IsUpdateUpcheReturnNeed = True) then

	sqlStr = " select "
	sqlStr = sqlStr + " IsNull(sum(case when datediff(d, m.regdate, '" + CStr(upReturnMiFinishBaseDate7) + "') >= 0 then 1 else 0 end), 0) as cnt7 "
	sqlStr = sqlStr + " , IsNull(sum(case when (datediff(d, m.regdate, '" + CStr(upReturnMiFinishBaseDate3) + "') >= 0) then 1 else 0 end), 0) as cnt3 "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_cs.dbo.tbl_new_as_list m "
	sqlStr = sqlStr + " 	join db_cs.dbo.tbl_new_as_detail d "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		m.id = d.masterid "
	sqlStr = sqlStr + " 	left join [db_temp].dbo.tbl_csmifinish_list T "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		d.id=T.csdetailidx "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and m.deleteyn = 'N' "
	sqlStr = sqlStr + " 	and m.currstate < 'B006' "
	sqlStr = sqlStr + " 	and d.itemid <> 0 "
	sqlStr = sqlStr + " 	and d.isupchebeasong='Y' "
	sqlStr = sqlStr + " 	and m.divcd = 'A004' "
	sqlStr = sqlStr + " 	and datediff(d, m.regdate, '" + CStr(upReturnMiFinishBaseDate3) + "') >= 0 "
	sqlStr = sqlStr + " 	and not ((T.ipgodate is not null) and (datediff(d, T.ipgodate, getdate()) <= 0)) "
	sqlStr = sqlStr + " 	and (datediff(m, m.regdate, getdate()) < 3) "
	sqlStr = sqlStr + " 	and m.id > " & minCSMasterIdxThreeMonth & " "

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

	if Not rsget.Eof then
	    application("CSUpcheReturnNotFinish7") = rsget("cnt7")
		application("CSUpcheReturnNotFinish3") = rsget("cnt3")
	end if
	rsget.close

	Call checkAndWriteElapsedTime("008")

end if

CSUpcheReturnNotFinish7 = application("CSUpcheReturnNotFinish7")
CSUpcheReturnNotFinish3 = application("CSUpcheReturnNotFinish3")


'==============================================================================
'반품접수(업체배송) D+7 미처리(담당자별 서머리)
dim csReturn7BrandCountByUserid, csReturn7OrderCountByUserid, nochargeReturn7Brandcnt

if (IsUpdateCSListNeed = True) or (IsUpdateUpcheReturnNeed = True) then

	csReturn7BrandCountByUserid = ""
	csReturn7OrderCountByUserid = ""
	nochargeReturn7Brandcnt = 0

	sqlchargeidlist = ""
	for i = 0 to UBound(csReturnBoardUserArr)
		if (sqlchargeidlist = "") then
			sqlchargeidlist = "'" + CStr(csReturnBoardUserArr(i)) + "'"
		else
			sqlchargeidlist = sqlchargeidlist + ",'" + CStr(csReturnBoardUserArr(i)) + "'"
		end if
	next

	sqlStr = " select "
	for i = 0 to UBound(csReturnBoardUserArr)
		sqlStr = sqlStr + "     IsNull(sum(case when (IsNull(T.userid,'') = '" + CStr(csReturnBoardUserArr(i)) + "') then T.makeridcnt else 0 end), 0) as chargeid" + CStr(i) + ", "
		sqlStr = sqlStr + "     IsNull(sum(case when (IsNull(T.userid,'') = '" + CStr(csReturnBoardUserArr(i)) + "') then T.ordercnt else 0 end), 0) as chargeidorder" + CStr(i) + ", "
	next
	sqlStr = sqlStr + "     IsNull(sum(case when (IsNull(T.userid,'') not in (" + sqlchargeidlist + ")) then T.makeridcnt else 0 end), 0) as nochargeid "
	sqlStr = sqlStr + " FROM "
	sqlStr = sqlStr + " 	( "
	sqlStr = sqlStr + " 		SELECT "
	sqlStr = sqlStr + " 			u.userid, count(distinct d.makerid) as makeridcnt, count(d.itemid) as ordercnt "
	sqlStr = sqlStr + " 		from "
	sqlStr = sqlStr + " 			db_cs.dbo.tbl_new_as_list m "
	sqlStr = sqlStr + " 			join db_cs.dbo.tbl_new_as_detail d "
	sqlStr = sqlStr + " 			on "
	sqlStr = sqlStr + " 				m.id = d.masterid "
	sqlStr = sqlStr + " 			left join [db_temp].dbo.tbl_csmifinish_list T "
	sqlStr = sqlStr + " 			on "
	sqlStr = sqlStr + " 				d.id=T.csdetailidx "
	sqlStr = sqlStr + " 			LEFT JOIN db_cs.dbo.tbl_cs_return_upche_brand u "
	sqlStr = sqlStr + " 			on "
	sqlStr = sqlStr + " 				d.makerid = u.brandid "
	sqlStr = sqlStr + " 		where "
	sqlStr = sqlStr + " 			1 = 1 "
	if application("Svr_Info") <> "Dev" then
		sqlStr = sqlStr + " 	and m.id >= 1200000 "		'// 속도개선
	else
		sqlStr = sqlStr + " 	and m.id >= 600000 "		'// 속도개선
	end if
	sqlStr = sqlStr + " 			and m.deleteyn = 'N' "
	sqlStr = sqlStr + " 			and m.currstate < 'B006' "
	sqlStr = sqlStr + " 			and d.itemid <> 0 "
	sqlStr = sqlStr + " 			and d.isupchebeasong='Y' "
	sqlStr = sqlStr + " 			and m.divcd = 'A004' "
	sqlStr = sqlStr + " 			and datediff(d, m.regdate, '" + CStr(upReturnMiFinishBaseDate7) + "') >= 0 "
	sqlStr = sqlStr + " 			and not ((T.ipgodate is not null) and (datediff(d, T.ipgodate, getdate()) <= 0)) "
	sqlStr = sqlStr + " 			and (datediff(m, m.regdate, getdate()) < 3) "
	sqlStr = sqlStr + " 			and m.id > " & minCSMasterIdxThreeMonth & " "
	sqlStr = sqlStr + " 		GROUP BY "
	sqlStr = sqlStr + " 			u.userid "
	sqlStr = sqlStr + " 	) T "
	''rw sqlStr
	'response.end

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
	if  not rsget.EOF  then
		for i = 0 to UBound(csReturnBoardUserArr)
			if (csReturn7BrandCountByUserid = "") then
				csReturn7BrandCountByUserid = rsget("chargeid" + CStr(i))
			else
				csReturn7BrandCountByUserid = csReturn7BrandCountByUserid & "|" & rsget("chargeid" + CStr(i))
			end if

			if (csReturn7OrderCountByUserid = "") then
				csReturn7OrderCountByUserid = rsget("chargeidorder" + CStr(i))
			else
				csReturn7OrderCountByUserid = csReturn7OrderCountByUserid & "|" & rsget("chargeidorder" + CStr(i))
			end if
		next
		nochargeReturn7Brandcnt = rsget("nochargeid")
	end if
	rsget.close

	application("csReturn7BrandCountByUserid") = csReturn7BrandCountByUserid
	application("csReturn7OrderCountByUserid") = csReturn7OrderCountByUserid
	application("nochargeReturn7Brandcnt") = nochargeReturn7Brandcnt

	Call checkAndWriteElapsedTime("009")

end if

csReturn7BrandCountByUserid = Split(application("csReturn7BrandCountByUserid"), "|")
csReturn7OrderCountByUserid = Split(application("csReturn7OrderCountByUserid"), "|")
nochargeReturn7Brandcnt = application("nochargeReturn7Brandcnt")


'==============================================================================
'반품접수(업체배송) D+7 미처리(브랜드별 서머리)
dim csReturn7BrandUseridArr, csReturn7BrandNameArr, csReturn7BrandcntArr

if (IsUpdateCSListNeed = True) or (IsUpdateUpcheReturnNeed = True) then

	csReturn7BrandUseridArr = ""
	csReturn7BrandNameArr = ""
	csReturn7BrandcntArr = ""

	resultCount = 0

	sqlStr = " SELECT "
	sqlStr = sqlStr + " 	u.userid, d.makerid, count(d.itemid) as cnt "
	sqlStr = sqlStr + " FROM "
	sqlStr = sqlStr + " 	db_cs.dbo.tbl_new_as_list m "
	sqlStr = sqlStr + " 	join db_cs.dbo.tbl_new_as_detail d "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		m.id = d.masterid "
	sqlStr = sqlStr + " 	left join [db_temp].dbo.tbl_csmifinish_list T "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		d.id=T.csdetailidx "
	sqlStr = sqlStr + " 	JOIN db_cs.dbo.tbl_cs_return_upche_brand u "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		d.makerid = u.brandid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	if application("Svr_Info") <> "Dev" then
		sqlStr = sqlStr + " 	and m.id >= 1200000 "		'// 속도개선
	else
		sqlStr = sqlStr + " 	and m.id >= 600000 "		'// 속도개선
	end if
	sqlStr = sqlStr + " 	and m.deleteyn = 'N' "
	sqlStr = sqlStr + " 	and m.currstate < 'B006' "
	sqlStr = sqlStr + " 	and d.itemid <> 0 "
	sqlStr = sqlStr + " 	and d.isupchebeasong='Y' "
	sqlStr = sqlStr + " 	and m.divcd = 'A004' "
	sqlStr = sqlStr + " 	and datediff(d, m.regdate, '" + CStr(upReturnMiFinishBaseDate7) + "') >= 0 "
	sqlStr = sqlStr + " 	and not ((T.ipgodate is not null) and (datediff(d, T.ipgodate, getdate()) <= 0)) "
	sqlStr = sqlStr + " 	and (datediff(m, m.regdate, getdate()) < 3) "
	sqlStr = sqlStr + " 	and m.id > " & minCSMasterIdxThreeMonth & " "
	sqlStr = sqlStr + " GROUP BY "
	sqlStr = sqlStr + " 	u.userid, d.makerid "
	sqlStr = sqlStr + " ORDER BY "
	sqlStr = sqlStr + " 	u.userid, count(d.itemid) desc, d.makerid "
	''rw sqlStr

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

	if not rsget.EOF  then

		resultCount = rsget.RecordCount

		for i = 0 to resultCount - 1
			if (csReturn7BrandUseridArr = "") then
				csReturn7BrandUseridArr = rsget("userid")
			else
				csReturn7BrandUseridArr = csReturn7BrandUseridArr & "|" & rsget("userid")
			end if

			if (csReturn7BrandNameArr = "") then
				csReturn7BrandNameArr = rsget("makerid")
			else
				csReturn7BrandNameArr = csReturn7BrandNameArr & "|" & rsget("makerid")
			end if

			if (csReturn7BrandcntArr = "") then
				csReturn7BrandcntArr = rsget("cnt")
			else
				csReturn7BrandcntArr = csReturn7BrandcntArr & "|" & rsget("cnt")
			end if

			rsget.movenext
		next
	end if
	rsget.close

	application("csReturn7BrandUseridArr") = csReturn7BrandUseridArr
	application("csReturn7BrandNameArr") = csReturn7BrandNameArr
	application("csReturn7BrandcntArr") = csReturn7BrandcntArr

	Call checkAndWriteElapsedTime("010")

end if

csReturn7BrandUseridArr = Split(application("csReturn7BrandUseridArr"), "|")
csReturn7BrandNameArr = Split(application("csReturn7BrandNameArr"), "|")
csReturn7BrandcntArr = Split(application("csReturn7BrandcntArr"), "|")


'==============================================================================
'반품접수(업체배송) D+3 미처리(담당자별 서머리)
dim csReturn3BrandCountByUserid, csReturn3OrderCountByUserid, nochargeReturn3Brandcnt

if (IsUpdateCSListNeed = True) or (IsUpdateUpcheReturnNeed = True) then

	csReturn3BrandCountByUserid = ""
	csReturn3OrderCountByUserid = ""
	nochargeReturn3Brandcnt = 0

	sqlchargeidlist = ""
	for i = 0 to UBound(csReturnBoardUserArr)
		if (sqlchargeidlist = "") then
			sqlchargeidlist = "'" + CStr(csReturnBoardUserArr(i)) + "'"
		else
			sqlchargeidlist = sqlchargeidlist + ",'" + CStr(csReturnBoardUserArr(i)) + "'"
		end if
	next

	sqlStr = " select "
	for i = 0 to UBound(csReturnBoardUserArr)
		sqlStr = sqlStr + "     IsNull(sum(case when (IsNull(T.userid,'') = '" + CStr(csReturnBoardUserArr(i)) + "') then T.makeridcnt else 0 end), 0) as chargeid" + CStr(i) + ", "
		sqlStr = sqlStr + "     IsNull(sum(case when (IsNull(T.userid,'') = '" + CStr(csReturnBoardUserArr(i)) + "') then T.ordercnt else 0 end), 0) as chargeidorder" + CStr(i) + ", "
	next
	sqlStr = sqlStr + "     IsNull(sum(case when (IsNull(T.userid,'') not in (" + sqlchargeidlist + ")) then T.makeridcnt else 0 end), 0) as nochargeid "
	sqlStr = sqlStr + " FROM "
	sqlStr = sqlStr + " 	( "
	sqlStr = sqlStr + " 		SELECT "
	sqlStr = sqlStr + " 			u.userid, count(distinct d.makerid) as makeridcnt, count(d.itemid) as ordercnt "
	sqlStr = sqlStr + " 		from "
	sqlStr = sqlStr + " 			db_cs.dbo.tbl_new_as_list m "
	sqlStr = sqlStr + " 			join db_cs.dbo.tbl_new_as_detail d "
	sqlStr = sqlStr + " 			on "
	sqlStr = sqlStr + " 				m.id = d.masterid "
	sqlStr = sqlStr + " 			left join [db_temp].dbo.tbl_csmifinish_list T "
	sqlStr = sqlStr + " 			on "
	sqlStr = sqlStr + " 				d.id=T.csdetailidx "
	sqlStr = sqlStr + " 			LEFT JOIN db_cs.dbo.tbl_cs_return_upche_brand u "
	sqlStr = sqlStr + " 			on "
	sqlStr = sqlStr + " 				d.makerid = u.brandid "
	sqlStr = sqlStr + " 		where "
	sqlStr = sqlStr + " 			1 = 1 "
	if application("Svr_Info") <> "Dev" then
		sqlStr = sqlStr + " 	and m.id >= 1200000 "		'// 속도개선
	else
		sqlStr = sqlStr + " 	and m.id >= 600000 "		'// 속도개선
	end if
	sqlStr = sqlStr + " 			and m.deleteyn = 'N' "
	sqlStr = sqlStr + " 			and m.currstate < 'B006' "
	sqlStr = sqlStr + " 			and d.itemid <> 0 "
	sqlStr = sqlStr + " 			and d.isupchebeasong='Y' "
	sqlStr = sqlStr + " 			and IsNull(T.code, '00') = '00' "			'// 사유미입력만
	sqlStr = sqlStr + " 			and m.divcd = 'A004' "
	sqlStr = sqlStr + " 			and datediff(d, m.regdate, '" + CStr(upReturnMiFinishBaseDate3) + "') >= 0 "
	sqlStr = sqlStr + " 			and not ((T.ipgodate is not null) and (datediff(d, T.ipgodate, getdate()) <= 0)) "
	sqlStr = sqlStr + " 			and (datediff(m, m.regdate, getdate()) < 3) "
	sqlStr = sqlStr + " 			and m.id > " & minCSMasterIdxThreeMonth & " "
	sqlStr = sqlStr + " 		GROUP BY "
	sqlStr = sqlStr + " 			u.userid "
	sqlStr = sqlStr + " 	) T "
	''rw sqlStr
	'response.end

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
	if  not rsget.EOF  then
		for i = 0 to UBound(csReturnBoardUserArr)
			if (csReturn3BrandCountByUserid = "") then
				csReturn3BrandCountByUserid = rsget("chargeid" + CStr(i))
			else
				csReturn3BrandCountByUserid = csReturn3BrandCountByUserid & "|" & rsget("chargeid" + CStr(i))
			end if

			if (csReturn3OrderCountByUserid = "") then
				csReturn3OrderCountByUserid = rsget("chargeidorder" + CStr(i))
			else
				csReturn3OrderCountByUserid = csReturn3OrderCountByUserid & "|" & rsget("chargeidorder" + CStr(i))
			end if
		next
		nochargeReturn3Brandcnt = rsget("nochargeid")
	end if
	rsget.close

	application("csReturn3BrandCountByUserid") = csReturn3BrandCountByUserid
	application("csReturn3OrderCountByUserid") = csReturn3OrderCountByUserid
	application("nochargeReturn3Brandcnt") = nochargeReturn3Brandcnt

	Call checkAndWriteElapsedTime("009")

end if

csReturn3BrandCountByUserid = Split(application("csReturn3BrandCountByUserid"), "|")
csReturn3OrderCountByUserid = Split(application("csReturn3OrderCountByUserid"), "|")
nochargeReturn3Brandcnt = application("nochargeReturn3Brandcnt")


'==============================================================================
'반품접수(업체배송) D+3 미처리(브랜드별 서머리)
dim csReturn3BrandUseridArr, csReturn3BrandNameArr, csReturn3BrandcntArr

if (IsUpdateCSListNeed = True) or (IsUpdateUpcheReturnNeed = True) then

	csReturn3BrandUseridArr = ""
	csReturn3BrandNameArr = ""
	csReturn3BrandcntArr = ""

	resultCount = 0

	sqlStr = " SELECT "
	sqlStr = sqlStr + " 	u.userid, d.makerid, count(d.itemid) as cnt "
	sqlStr = sqlStr + " FROM "
	sqlStr = sqlStr + " 	db_cs.dbo.tbl_new_as_list m "
	sqlStr = sqlStr + " 	join db_cs.dbo.tbl_new_as_detail d "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		m.id = d.masterid "
	sqlStr = sqlStr + " 	left join [db_temp].dbo.tbl_csmifinish_list T "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		d.id=T.csdetailidx "
	sqlStr = sqlStr + " 	JOIN db_cs.dbo.tbl_cs_return_upche_brand u "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		d.makerid = u.brandid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	if application("Svr_Info") <> "Dev" then
		sqlStr = sqlStr + " 	and m.id >= 1200000 "		'// 속도개선
	else
		sqlStr = sqlStr + " 	and m.id >= 600000 "		'// 속도개선
	end if
	sqlStr = sqlStr + " 	and m.deleteyn = 'N' "
	sqlStr = sqlStr + " 	and m.currstate < 'B006' "
	sqlStr = sqlStr + " 	and d.itemid <> 0 "
	sqlStr = sqlStr + " 	and d.isupchebeasong='Y' "
	sqlStr = sqlStr + " 	and IsNull(T.code, '00') = '00' "			'// 사유미입력만
	sqlStr = sqlStr + " 	and m.divcd = 'A004' "
	sqlStr = sqlStr + " 	and datediff(d, m.regdate, '" + CStr(upReturnMiFinishBaseDate3) + "') >= 0 "
	sqlStr = sqlStr + " 	and not ((T.ipgodate is not null) and (datediff(d, T.ipgodate, getdate()) <= 0)) "
	sqlStr = sqlStr + " 	and (datediff(m, m.regdate, getdate()) < 3) "
	sqlStr = sqlStr + " 	and m.id > " & minCSMasterIdxThreeMonth & " "
	sqlStr = sqlStr + " GROUP BY "
	sqlStr = sqlStr + " 	u.userid, d.makerid "
	sqlStr = sqlStr + " ORDER BY "
	sqlStr = sqlStr + " 	u.userid, count(d.itemid) desc, d.makerid "
	''rw sqlStr

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

	if not rsget.EOF  then

		resultCount = rsget.RecordCount

		for i = 0 to resultCount - 1
			if (csReturn3BrandUseridArr = "") then
				csReturn3BrandUseridArr = rsget("userid")
			else
				csReturn3BrandUseridArr = csReturn3BrandUseridArr & "|" & rsget("userid")
			end if

			if (csReturn3BrandNameArr = "") then
				csReturn3BrandNameArr = rsget("makerid")
			else
				csReturn3BrandNameArr = csReturn3BrandNameArr & "|" & rsget("makerid")
			end if

			if (csReturn3BrandcntArr = "") then
				csReturn3BrandcntArr = rsget("cnt")
			else
				csReturn3BrandcntArr = csReturn3BrandcntArr & "|" & rsget("cnt")
			end if

			rsget.movenext
		next
	end if
	rsget.close

	application("csReturn3BrandUseridArr") = csReturn3BrandUseridArr
	application("csReturn3BrandNameArr") = csReturn3BrandNameArr
	application("csReturn3BrandcntArr") = csReturn3BrandcntArr

	Call checkAndWriteElapsedTime("010")

end if

csReturn3BrandUseridArr = Split(application("csReturn3BrandUseridArr"), "|")
csReturn3BrandNameArr = Split(application("csReturn3BrandNameArr"), "|")
csReturn3BrandcntArr = Split(application("csReturn3BrandcntArr"), "|")


'==============================================================================
''업체처리완료건.
dim csUpcheFinished

if (IsUpdateCSListNeed = True) or (IsUpdateUpcheReturnNeed = True) then

	sqlStr = " select count(*) as cnt "
	sqlStr = sqlStr + " from  [db_cs].[dbo].tbl_new_as_list A"
	sqlStr = sqlStr + " left join [db_cs].[dbo].tbl_new_as_list B "
	sqlStr = sqlStr + " on A.id = B.refasid "
	sqlStr = sqlStr + " where (A.requireupche='Y') and (A.currstate='B006') and A.deleteyn = 'N' and (A.divcd not in ('A000', 'A100') or IsNull(B.currstate, 'B007') = 'B007') "

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
	if  not rsget.EOF  then
	    csUpcheFinished = rsget("cnt")
	end if
	rsget.close

	application("csUpcheFinished") = csUpcheFinished

	Call checkAndWriteElapsedTime("011")

end if

csUpcheFinished = application("csUpcheFinished")


'==============================================================================
''마일리지환불미처리
dim csNotFinMileRefund

if (IsUpdateCSListNeed = True) or (IsUpdateUpcheReturnNeed = True) then

	sqlStr = " select count(A.id) as csNotFinMileRefund from [db_cs].[dbo].tbl_new_as_list A "
	sqlStr = sqlStr + "     Left Join [db_cs].[dbo].tbl_as_refund_info r on A.id=r.asid "
	sqlStr = sqlStr + " where 1 = 1 "
	sqlStr = sqlStr + " and A.currstate<'B007' "
	sqlStr = sqlStr + " and A.divcd='A003' "
	sqlStr = sqlStr + " and A.deleteyn='N' "
	sqlStr = sqlStr + " and R.returnmethod='R900'"
	sqlStr = sqlStr + " and A.regdate>'"&treemonthbefore&"'"

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
	if  not rsget.EOF  then
	    csNotFinMileRefund = rsget("csNotFinMileRefund")
	end if
	rsget.close

	application("csNotFinMileRefund") = csNotFinMileRefund

	Call checkAndWriteElapsedTime("012")

end if

csNotFinMileRefund = application("csNotFinMileRefund")


'==============================================================================
'입점몰 품절취소요청[텐배+업배] : 입점몰 리스트
dim csStockoutOtherSitenameArr, csStockoutOtherSiteOrderCountArr, csStockoutOtherSiteOrderDetailCountArr

if (IsUpdateIpjumStockOutNeed = True) then

	csStockoutOtherSitenameArr = ""
	csStockoutOtherSiteOrderCountArr = ""
	csStockoutOtherSiteOrderDetailCountArr = ""

	resultCount = 0

	sqlStr = " SELECT "
	sqlStr = sqlStr + " 	m.sitename, count(distinct m.orderserial) as ordercnt, count(d.idx) as cnt "
	sqlStr = sqlStr + " FROM "
	sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_master m "
	sqlStr = sqlStr + " 	JOIN [db_order].[dbo].tbl_order_detail d "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		m.orderserial=d.orderserial "
	sqlStr = sqlStr + " 	JOIN [db_temp].dbo.tbl_mibeasong_list T "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and d.orderserial=T.orderserial "
	sqlStr = sqlStr + " 		and d.idx=T.detailidx "
	sqlStr = sqlStr + " WHERE "
	sqlStr = sqlStr + " 	m.regdate >= DateAdd(m, -2, getdate()) "
	sqlStr = sqlStr + " 	and m.ipkumdiv < '8' "
	sqlStr = sqlStr + " 	and m.ipkumdiv > '3' "
	sqlStr = sqlStr + " 	and m.cancelyn = 'N' "
	sqlStr = sqlStr + " 	and m.jumundiv <> '9' "
	sqlStr = sqlStr + " 	and d.itemid<>0 "
	''sqlStr = sqlStr + " 	and d.currstate='3' "
	sqlStr = sqlStr + " 	and m.sitename <> '10x10' "																'// 입점몰주문
	''sqlStr = sqlStr + " 	and d.isupchebeasong='Y' "																'// 텐배+업배
	sqlStr = sqlStr + " 	and d.cancelyn <> 'Y' "
	sqlStr = sqlStr + " 	and IsNull(T.code,'00')='05' "															'// 품절출고불가
	sqlStr = sqlStr + " 	and IsNull(T.state, '0')='0' "															'// 고객안내 이전
	sqlStr = sqlStr + " GROUP BY "
	sqlStr = sqlStr + " 	m.sitename "
	sqlStr = sqlStr + " ORDER BY "
	sqlStr = sqlStr + " 	m.sitename "
	''response.write sqlStr

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

	if not rsget.EOF  then

		resultCount = rsget.RecordCount

		for i = 0 to resultCount - 1
			if (csStockoutOtherSitenameArr = "") then
				csStockoutOtherSitenameArr = rsget("sitename")
			else
				csStockoutOtherSitenameArr = csStockoutOtherSitenameArr & "|" & rsget("sitename")
			end if

			if (csStockoutOtherSiteOrderCountArr = "") then
				csStockoutOtherSiteOrderCountArr = rsget("ordercnt")
			else
				csStockoutOtherSiteOrderCountArr = csStockoutOtherSiteOrderCountArr & "|" & rsget("ordercnt")
			end if

			if (csStockoutOtherSiteOrderDetailCountArr = "") then
				csStockoutOtherSiteOrderDetailCountArr = rsget("cnt")
			else
				csStockoutOtherSiteOrderDetailCountArr = csStockoutOtherSiteOrderDetailCountArr & "|" & rsget("cnt")
			end if

			rsget.movenext
		next
	end if
	rsget.close

	application("csStockoutOtherSitenameArr") = csStockoutOtherSitenameArr
	application("csStockoutOtherSiteOrderCountArr") = csStockoutOtherSiteOrderCountArr
	application("csStockoutOtherSiteOrderDetailCountArr") = csStockoutOtherSiteOrderDetailCountArr

	Call checkAndWriteElapsedTime("013")

end if

csStockoutOtherSitenameArr = Split(application("csStockoutOtherSitenameArr"), "|")
csStockoutOtherSiteOrderCountArr = Split(application("csStockoutOtherSiteOrderCountArr"), "|")
csStockoutOtherSiteOrderDetailCountArr = Split(application("csStockoutOtherSiteOrderDetailCountArr"), "|")


'==============================================================================
'입점몰 품절취소요청[텐배+업배](주문별 서머리)
dim csStockoutOtherSiteOrderSitenameArr, csStockoutOtherSiteOrderserialArr, csStockoutOtherSiteOrderdetailCntcntArr, csStockoutOtherSiteOrderdetaiidx

if (IsUpdateIpjumStockOutNeed = True) then

	csStockoutOtherSiteOrderSitenameArr = ""
	csStockoutOtherSiteOrderserialArr = ""
	csStockoutOtherSiteOrderdetailCntcntArr = ""
	csStockoutOtherSiteOrderdetaiidx = ""

	resultCount = 0

	sqlStr = " SELECT "
	sqlStr = sqlStr + " 	m.sitename, m.orderserial, max(d.idx) as orderdetailidx, count(d.itemid) as cnt "
	sqlStr = sqlStr + " FROM "
	sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_master m "
	sqlStr = sqlStr + " 	JOIN [db_order].[dbo].tbl_order_detail d "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		m.orderserial=d.orderserial "
	sqlStr = sqlStr + " 	JOIN [db_temp].dbo.tbl_mibeasong_list T "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and d.orderserial=T.orderserial "
	sqlStr = sqlStr + " 		and d.idx=T.detailidx "
	sqlStr = sqlStr + " WHERE "
	sqlStr = sqlStr + " 	m.regdate >= DateAdd(m, -2, getdate()) "
	sqlStr = sqlStr + " 	and m.ipkumdiv < '8' "
	sqlStr = sqlStr + " 	and m.ipkumdiv > '3' "
	sqlStr = sqlStr + " 	and m.cancelyn = 'N' "
	sqlStr = sqlStr + " 	and m.jumundiv <> '9' "
	sqlStr = sqlStr + " 	and d.itemid<>0 "
	sqlStr = sqlStr + " 	and m.sitename <> '10x10' "
	''sqlStr = sqlStr + " 	and d.isupchebeasong='Y' "
	sqlStr = sqlStr + " 	and d.cancelyn <> 'Y' "
	sqlStr = sqlStr + " 	and IsNull(T.code,'00')='05' "
	sqlStr = sqlStr + " 	and IsNull(T.state, '0')='0' "
	sqlStr = sqlStr + " GROUP BY "
	sqlStr = sqlStr + " 	m.sitename, m.orderserial "
	sqlStr = sqlStr + " ORDER BY "
	sqlStr = sqlStr + " 	m.sitename, m.orderserial "
	''response.write sqlStr

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

	if not rsget.EOF  then

		resultCount = rsget.RecordCount

		for i = 0 to resultCount - 1
			if (csStockoutOtherSiteOrderSitenameArr = "") then
				csStockoutOtherSiteOrderSitenameArr = rsget("sitename")
			else
				csStockoutOtherSiteOrderSitenameArr = csStockoutOtherSiteOrderSitenameArr & "|" & rsget("sitename")
			end if

			if (csStockoutOtherSiteOrderserialArr = "") then
				csStockoutOtherSiteOrderserialArr = rsget("orderserial")
			else
				csStockoutOtherSiteOrderserialArr = csStockoutOtherSiteOrderserialArr & "|" & rsget("orderserial")
			end if

			if (csStockoutOtherSiteOrderdetailCntcntArr = "") then
				csStockoutOtherSiteOrderdetailCntcntArr = rsget("cnt")
			else
				csStockoutOtherSiteOrderdetailCntcntArr = csStockoutOtherSiteOrderdetailCntcntArr & "|" & rsget("cnt")
			end if

			if (csStockoutOtherSiteOrderdetaiidx = "") then
				csStockoutOtherSiteOrderdetaiidx = rsget("orderdetailidx")
			else
				csStockoutOtherSiteOrderdetaiidx = csStockoutOtherSiteOrderdetaiidx & "|" & rsget("orderdetailidx")
			end if

			rsget.movenext
		next
	end if
	rsget.close

	application("csStockoutOtherSiteOrderSitenameArr") 		= csStockoutOtherSiteOrderSitenameArr
	application("csStockoutOtherSiteOrderserialArr") 		= csStockoutOtherSiteOrderserialArr
	application("csStockoutOtherSiteOrderdetailCntcntArr") 	= csStockoutOtherSiteOrderdetailCntcntArr
	application("csStockoutOtherSiteOrderdetaiidx") 		= csStockoutOtherSiteOrderdetaiidx

	Call checkAndWriteElapsedTime("014")

end if

csStockoutOtherSiteOrderSitenameArr 	= Split(application("csStockoutOtherSiteOrderSitenameArr"), "|")
csStockoutOtherSiteOrderserialArr 		= Split(application("csStockoutOtherSiteOrderserialArr"), "|")
csStockoutOtherSiteOrderdetailCntcntArr = Split(application("csStockoutOtherSiteOrderdetailCntcntArr"), "|")
csStockoutOtherSiteOrderdetaiidx 		= Split(application("csStockoutOtherSiteOrderdetaiidx"), "|")


'==============================================================================
'품절취소요청[텐배+업배](담당자별 서머리)
dim csStockoutOrderCountByUserid			'// 담당자당 주문 수
dim csStockoutOrderDetailCountByUserid		'// 담당자당 주문 디테일 수
dim nochargeStockoutOrdercnt				'// 미지정 주문 수

if (IsUpdateTenTenStockOutNeed = True) then

	csStockoutOrderCountByUserid = ""
	csStockoutOrderDetailCountByUserid = ""
	nochargeStockoutOrdercnt = ""

	sqlchargeidlist = ""
	for i = 0 to UBound(csMichulgoStockoutBoardUserArr)
		if (sqlchargeidlist = "") then
			sqlchargeidlist = "'" + CStr(csMichulgoStockoutBoardUserArr(i)) + "'"
		else
			sqlchargeidlist = sqlchargeidlist + ",'" + CStr(csMichulgoStockoutBoardUserArr(i)) + "'"
		end if
	next

	sqlStr = " select "
	for i = 0 to UBound(csMichulgoStockoutBoardUserArr)
		sqlStr = sqlStr + "     IsNull(sum(case when (IsNull(T.userid,'') = '" + CStr(csMichulgoStockoutBoardUserArr(i)) + "') then T.ordercnt else 0 end), 0) as order" + CStr(i) + ", "
		sqlStr = sqlStr + "     IsNull(sum(case when (IsNull(T.userid,'') = '" + CStr(csMichulgoStockoutBoardUserArr(i)) + "') then T.detailcnt else 0 end), 0) as orderdetail" + CStr(i) + ", "
	next
	sqlStr = sqlStr + "     IsNull(sum(case when (IsNull(T.userid,'') not in (" + sqlchargeidlist + ")) then T.ordercnt else 0 end), 0) as nochargeid "
	sqlStr = sqlStr + " FROM "
	sqlStr = sqlStr + " 	( "
	sqlStr = sqlStr + " 		SELECT "
	sqlStr = sqlStr + " 			u.userid, count(distinct m.orderserial) as ordercnt, count(d.itemid) as detailcnt "
	sqlStr = sqlStr + " 		FROM "
	sqlStr = sqlStr + " 			[db_order].[dbo].tbl_order_master m "
	sqlStr = sqlStr + " 			JOIN [db_order].[dbo].tbl_order_detail d "
	sqlStr = sqlStr + " 			on "
	sqlStr = sqlStr + " 				m.orderserial=d.orderserial "
	sqlStr = sqlStr + " 			JOIN [db_temp].dbo.tbl_mibeasong_list T "
	sqlStr = sqlStr + " 			on "
	sqlStr = sqlStr + " 				1 = 1 "
	sqlStr = sqlStr + " 				and d.orderserial=T.orderserial "
	sqlStr = sqlStr + " 				and d.idx=T.detailidx "
	sqlStr = sqlStr + " 			LEFT JOIN db_cs.dbo.tbl_cs_michulgo_stockout_order u "
	sqlStr = sqlStr + " 			on "
	sqlStr = sqlStr + " 				1 = 1 "
	sqlStr = sqlStr + " 				and m.orderserial = u.orderserial "
	sqlStr = sqlStr + " 				and u.deleteyn = 'N' "
	sqlStr = sqlStr + " 		WHERE "
	sqlStr = sqlStr + " 			m.regdate >= DateAdd(m, -2, getdate()) "
	sqlStr = sqlStr + " 			and m.ipkumdiv < '8' "
	sqlStr = sqlStr + " 			and m.ipkumdiv > '3' "
	sqlStr = sqlStr + " 			and m.cancelyn = 'N' "
	sqlStr = sqlStr + " 			and m.jumundiv <> '9' "
	sqlStr = sqlStr + " 			and d.itemid<>0 "
	''sqlStr = sqlStr + " 			and d.currstate='3' "
	sqlStr = sqlStr + " 			and m.sitename = '10x10' "																'// 입점몰주문 제외
	''sqlStr = sqlStr + " 			and d.isupchebeasong='Y' "																'// 텐배+업배
	sqlStr = sqlStr + " 			and d.cancelyn <> 'Y' "
	sqlStr = sqlStr + " 			and IsNull(T.code,'00')='05' "															'// 품절출고불가
	sqlStr = sqlStr + " 			and IsNull(T.state, '0')='0' "															'// 고객안내 이전
	sqlStr = sqlStr + " 		GROUP BY "
	sqlStr = sqlStr + " 			u.userid "
	sqlStr = sqlStr + " 	) T "
	''response.write sqlStr

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
	if  not rsget.EOF  then
		for i = 0 to UBound(csMichulgoStockoutBoardUserArr)
			if (csStockoutOrderCountByUserid = "") then
				csStockoutOrderCountByUserid = rsget("order" + CStr(i))
			else
				csStockoutOrderCountByUserid = csStockoutOrderCountByUserid & "|" & rsget("order" + CStr(i))
			end if

			if (csStockoutOrderDetailCountByUserid = "") then
				csStockoutOrderDetailCountByUserid = rsget("orderdetail" + CStr(i))
			else
				csStockoutOrderDetailCountByUserid = csStockoutOrderDetailCountByUserid & "|" & rsget("orderdetail" + CStr(i))
			end if
		next
		nochargeStockoutOrdercnt = rsget("nochargeid")
	end if
	rsget.close

	application("csStockoutOrderCountByUserid") = csStockoutOrderCountByUserid
	application("csStockoutOrderDetailCountByUserid") = csStockoutOrderDetailCountByUserid
	application("nochargeStockoutOrdercnt") = nochargeStockoutOrdercnt

	Call checkAndWriteElapsedTime("015")

end if

csStockoutOrderCountByUserid = Split(application("csStockoutOrderCountByUserid"), "|")
csStockoutOrderDetailCountByUserid = Split(application("csStockoutOrderDetailCountByUserid"), "|")
nochargeStockoutOrdercnt = application("nochargeStockoutOrdercnt")




'==============================================================================
'품절취소요청[텐배+업배](주문별 서머리)
dim csStockoutOrderUseridArr, csStockoutOrderserialArr, csStockoutOrderdetailCntcntArr, csStockoutHasChargeUser, csStockoutOrderdetaiidx

if (IsUpdateTenTenStockOutNeed = True) then

	csStockoutOrderUseridArr = ""
	csStockoutOrderserialArr = ""
	csStockoutOrderdetailCntcntArr = ""
	csStockoutHasChargeUser = ""
	csStockoutOrderdetaiidx = ""

	resultCount = 0

	sqlStr = " SELECT "
	sqlStr = sqlStr + " 	u.userid, m.orderserial, max(d.idx) as orderdetailidx, count(d.itemid) as cnt "
	sqlStr = sqlStr + " FROM "
	sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_master m "
	sqlStr = sqlStr + " 	JOIN [db_order].[dbo].tbl_order_detail d "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		m.orderserial=d.orderserial "
	sqlStr = sqlStr + " 	JOIN [db_temp].dbo.tbl_mibeasong_list T "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and d.orderserial=T.orderserial "
	sqlStr = sqlStr + " 		and d.idx=T.detailidx "
	sqlStr = sqlStr + " 	LEFT JOIN db_cs.dbo.tbl_cs_michulgo_stockout_order u "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and m.orderserial = u.orderserial "
	sqlStr = sqlStr + " 		and u.deleteyn = 'N' "
	sqlStr = sqlStr + " WHERE "
	sqlStr = sqlStr + " 	m.regdate >= DateAdd(m, -2, getdate()) "
	sqlStr = sqlStr + " 	and m.ipkumdiv < '8' "
	sqlStr = sqlStr + " 	and m.ipkumdiv > '3' "
	sqlStr = sqlStr + " 	and m.cancelyn = 'N' "
	sqlStr = sqlStr + " 	and m.jumundiv <> '9' "
	sqlStr = sqlStr + " 	and d.itemid<>0 "
	sqlStr = sqlStr + " 	and m.sitename = '10x10' "
	''sqlStr = sqlStr + " 	and d.isupchebeasong='Y' "
	sqlStr = sqlStr + " 	and d.cancelyn <> 'Y' "
	sqlStr = sqlStr + " 	and IsNull(T.code,'00')='05' "
	sqlStr = sqlStr + " 	and IsNull(T.state, '0')='0' "
	sqlStr = sqlStr + " GROUP BY "
	sqlStr = sqlStr + " 	u.userid, m.orderserial "
	sqlStr = sqlStr + " ORDER BY "
	sqlStr = sqlStr + " 	u.userid, m.orderserial "
	''response.write sqlStr

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

	if not rsget.EOF  then

		resultCount = rsget.RecordCount

		for i = 0 to resultCount - 1
			if (csStockoutOrderUseridArr = "") then
				csStockoutOrderUseridArr = rsget("userid")
			else
				csStockoutOrderUseridArr = csStockoutOrderUseridArr & "|" & rsget("userid")
			end if

			if (csStockoutOrderserialArr = "") then
				csStockoutOrderserialArr = rsget("orderserial")
			else
				csStockoutOrderserialArr = csStockoutOrderserialArr & "|" & rsget("orderserial")
			end if

			if (csStockoutOrderdetailCntcntArr = "") then
				csStockoutOrderdetailCntcntArr = rsget("cnt")
			else
				csStockoutOrderdetailCntcntArr = csStockoutOrderdetailCntcntArr & "|" & rsget("cnt")
			end if

			if (csStockoutHasChargeUser = "") then
				csStockoutHasChargeUser = "N"
			else
				csStockoutHasChargeUser = csStockoutHasChargeUser & "|" & "N"
			end if

			if (csStockoutOrderdetaiidx = "") then
				csStockoutOrderdetaiidx = rsget("orderdetailidx")
			else
				csStockoutOrderdetaiidx = csStockoutOrderdetaiidx & "|" & rsget("orderdetailidx")
			end if

			rsget.movenext
		next
	end if
	rsget.close

	application("csStockoutOrderUseridArr") 		= csStockoutOrderUseridArr
	application("csStockoutOrderserialArr") 		= csStockoutOrderserialArr
	application("csStockoutOrderdetailCntcntArr") 	= csStockoutOrderdetailCntcntArr
	application("csStockoutHasChargeUser") 			= csStockoutHasChargeUser
	application("csStockoutOrderdetaiidx") 			= csStockoutOrderdetaiidx

	Call checkAndWriteElapsedTime("016")

end if

csStockoutOrderUseridArr 		= Split(application("csStockoutOrderUseridArr"), "|")
csStockoutOrderserialArr 		= Split(application("csStockoutOrderserialArr"), "|")
csStockoutOrderdetailCntcntArr 	= Split(application("csStockoutOrderdetailCntcntArr"), "|")
csStockoutHasChargeUser 		= Split(application("csStockoutHasChargeUser"), "|")
csStockoutOrderdetaiidx 		= Split(application("csStockoutOrderdetaiidx"), "|")


'==============================================================================
' 아이디별 확인사항
Dim csMainUserID
csMainUserID	= req("csMainUserID", session("ssBctId") )

'미처리메모
dim CSMemoNotFinish


sqlStr = " select count(*) as CSMemoNotFinish from db_cs.dbo.tbl_cs_memo"
sqlStr = sqlStr + " where finishyn='N'"
sqlStr = sqlStr + " and regdate>'"&treemonthbefore&"'"
sqlStr = sqlStr & " AND writeUser = '" & csMainUserID & "'	" & vbCrLf

rsget.CursorLocation = adUseClient
rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
if  not rsget.EOF  then
    CSMemoNotFinish = rsget("CSMemoNotFinish")
end if
rsget.close

%>
<script language="JavaScript" src="/cscenter/js/cscenter.js"></script>
<script language="JavaScript" src="/cscenter/js/convert.date.js"></script>
<script language='javascript'>
function popUpcheMisend(misendCode, dplusOver, currState)
{
	var MisendState = "";
	if (misendCode)
		MisendState = "0";

	var window_width = 1024;
    var window_height = 800;
	var popwin = window.open("/admin/upchebeasong/upchemibeasonglist.asp?menupos=246&popupFlag=Y&MisendState="+MisendState+"&MisendReason=" + misendCode + "&dplusOver=" + dplusOver + "&currState=" + currState , "cs_action","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
	popwin.focus();
}

function popUpcheMisendByBrand(makerid, misendCode, dplusOver, currState)
{
	var MisendState = "";
	if (misendCode)
		MisendState = "0";

	var window_width = 1024;
    var window_height = 800;
	var popwin = window.open("/admin/upchebeasong/upchemibeasonglist.asp?menupos=246&popupFlag=Y&exinmaychulgoday=Y&makerid="+makerid+"&MisendState="+MisendState+"&MisendReason=" + misendCode + "&dplusOver=" + dplusOver + "&currState=" + currState , "popUpcheMisendByBrand","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
	popwin.focus();

}

function popUpcheStockout10x10ByOrderserial(orderserial)
{
	var window_width = 1024;
    var window_height = 800;

	var popwin = window.open("/admin/ordermaster/misendmaster_main.asp?orderserial=" + orderserial, "popUpcheStockout10x10ByOrderserial","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
	popwin.focus();

}

function popUpcheStockoutOtherSiteByOrderserial(orderserial, sitename)
{
	var window_width = 1024;
    var window_height = 800;

	var popwin = window.open("/admin/ordermaster/misendmaster_main.asp?orderserial=" + orderserial, "popUpcheStockoutOtherSiteByOrderserial","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
	popwin.focus();

}

function setUpcheStockout10x10ByOrderserial(frm)
{
	// alert("작업중");
	// return;

	if (confirm("미지정건을 분배 하시겠습니까?") == true) {
		frm.submit();
	}
}

function Cscenter_Action_List2(searchtype) {
    var window_width = 1280;
    var window_height = 960;

    var popwin = window.open("/cscenter/action/cs_action.asp?searchtype=" + searchtype ,"cs_action","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
    /*
    if (searchtype=="upreturnmifinish"){
        var popwin = window.open("/cscenter/action/cs_action.asp?divcd=A004&currstate=B004&delYN=N","cs_action","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
    }else{
	    var popwin = window.open("/cscenter/action/cs_action.asp?searchtype=" + searchtype ,"cs_action","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
	}
	*/

	popwin.focus();
}

function Cscenter_Action_MiFinishReturnList() {
    var window_width = 1280;
    var window_height = 960;
	var url = "/cscenter/mifinish/cs_mifinishlist.asp?divcd=returncs&Dtype=dday&dplusOver=&exinmaychulgoday=Y&exoldcs=Y";

    var popwin = window.open(url,"Cscenter_Action_MiFinishReturnList","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");

	popwin.focus();
}

function Cscenter_Action_MiFinishReturnListDPlus(dday) {
    var window_width = 1280;
    var window_height = 960;
	var url = "/cscenter/mifinish/cs_mifinishlist.asp?divcd=returncs&Dtype=dday&dplusOver=" + dday + "&exinmaychulgoday=Y&exoldcs=Y";

    var popwin = window.open(url,"Cscenter_Action_MiFinishReturnListDPlus","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");

	popwin.focus();
}

function Cscenter_Action_MiFinishReturnListByBrand(makerid, dday) {
    var window_width = 1280;
    var window_height = 960;
	var url = "/cscenter/mifinish/cs_mifinishlist.asp?divcd=returncs&Dtype=dday&dplusOver=" + dday + "&exinmaychulgoday=Y&exoldcs=Y&makerid=" + makerid;

	if (dday*1 == 3) {
		//MifinishReason
		url = url + "&MifinishReason=00";
	}

    var popwin = window.open(url,"Cscenter_Action_MiFinishReturnListByBrand","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");

	popwin.focus();
}

// 접수자아이디별
function Cscenter_Action_List3(searchfield,searchstring,divcd,currstate) {
    var window_width = 1280;
    var window_height = 960;
	var popwin = window.open("/cscenter/action/cs_action.asp?delYN=N&searchfield=" + searchfield + "&searchstring=" + searchstring + "&divcd=" + divcd + "&currstate=" + currstate ,"cs_action","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
	popwin.focus();
}

function Cscenter_Action_List4(searchfield,searchstring,searchtype) {
    var window_width = 1280;
    var window_height = 960;
	var popwin = window.open("/cscenter/action/cs_action.asp?delYN=N&searchfield=" + searchfield + "&searchstring=" + searchstring + "&searchtype=" + searchtype ,"cs_action","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
	popwin.focus();
}

//

function upchebeasongmaster(){
	var popwin = window.open("/admin/upchebeasong/upchebeasongmaster.asp","upchebeasongmaster","width=1000 height=700 scrollbars=yes resizable=yes");
	popwin.focus();
}

function misendmaster(v){
	var popwin = window.open("/admin/ordermaster/misendmaster_top.asp?inputyn=" + v ,"misendmaster","width=1000 height=700 scrollbars=yes resizable=yes");
	popwin.focus();
}


function cscenter_memo_list(orderserial, userid, finishyn, writeUser) {
	var popwin = window.open("/cscenter/memo/cs_memo.asp?orderserial=" + orderserial + "&userid=" + userid + "&finishyn=" + finishyn + "&writeUser=" + writeUser,"cs_memo","width=1000 height=700 scrollbars=yes resizable=yes");
	popwin.focus();
}



function cscenter_mileage_list(userid) {
	var popwin = window.open("/cscenter/mileage/cs_mileage.asp?userid=" + userid ,"cs_mileage","width=1000 height=700 scrollbars=yes resizable=yes");
	popwin.focus();
}

function cscenter_coupon_list(userid) {
	var popwin = window.open("/cscenter/coupon/cs_coupon.asp?userid=" + userid ,"cs_coupon","width=1000 height=700 scrollbars=yes resizable=yes");
	popwin.focus();
}

function cscenter_payment_list(v) {
	var popwin = window.open("/cscenter/payment/cs_paymentlist.asp?ipkumstate=" + v ,"cs_paymentlist","width=1000 height=700 scrollbars=yes resizable=yes");
	popwin.focus();
}

function cscenter_cashreceipt_list() {
	var popwin = window.open("/cscenter/taxsheet/cashreceiptlist.asp","cscenter_cashreceipt","width=1000 height=700 scrollbars=yes resizable=yes");
	popwin.focus();
}

function cscenter_tax_list() {
	var popwin = window.open("/cscenter/taxsheet/tax_list.asp","cscenter_tax","width=1000 height=700 scrollbars=yes resizable=yes");
	popwin.focus();
}

function cscenter_refund_list() {
	var popwin = window.open("/cscenter/refund/refundlist.asp","cscenter_refund","width=1000 height=700 scrollbars=yes resizable=yes");
	popwin.focus();
}

function cscenter_eventprize_list()
{
	window.open("/admin/eventmanage/event/eventprize_list.asp?menupos=1056","cscenter_eventprize","width=1000 height=700 scrollbars=yes resizable=yes");
}

function cscenter_member_list()
{
	window.open("/cscenter/member/customerinfo.asp?menupos=1166","cscenter_member","width=1000 height=700 scrollbars=yes resizable=yes");
}

function cscenter_deposit_list()
{
	var popwin = window.open("/cscenter/deposit/cs_deposit.asp?menupos=1324","cscenter_deposit","width=1000 height=700 scrollbars=yes resizable=yes");
	popwin.focus();
}

function cscenter_eventjoin_list()
{
	var popwin = window.open("/admin/eventmanage/event/eventjoin_list.asp?menupos=1529","eventjoin","width=1100 height=700 scrollbars=yes resizable=yes");
	popwin.focus();
}

function showHideBrand(id) {
	tr = document.getElementsByTagName('tr');

	for (i = 0; i < tr.length; i++) {
		// if (tr[i].getAttribute(thisname) == true) {
		if (tr[i].id == id) {
			if ( tr[i].style.display=='none' ) {
				tr[i].style.display = '';
			} else {
				tr[i].style.display = 'none';
			}
		}
	}
}
//업무협조
function popCooperate(){
	 var winCooperate = window.open("/admin/cooperate/popIndex.asp","popCooperate","width="+screen.availWidth+", height="+ screen.availHeight +",resizable=yes, scrollbars=yes");
	 winCooperate.focus();
}

var csTimeOneToOneBoard = new Date(getDateFromFormat("<%= application("csTimeOneToOneBoard") %>", "yyyy-MM-dd a h:mm:ss"));
var csTimeMichulgoListUpche = new Date(getDateFromFormat("<%= application("csTimeMichulgoListUpche") %>", "yyyy-MM-dd a h:mm:ss"));
var csTimeIpjumMichulgo = new Date(getDateFromFormat("<%= application("csTimeIpjumMichulgo") %>", "yyyy-MM-dd a h:mm:ss"));
var csTimeMichulgoListTenTen = new Date(getDateFromFormat("<%= application("csTimeMichulgoListTenTen") %>", "yyyy-MM-dd a h:mm:ss"));
var csTimeCSList = new Date(getDateFromFormat("<%= application("csTimeCSList") %>", "yyyy-MM-dd a h:mm:ss"));
var csTimeUpcheReturn = new Date(getDateFromFormat("<%= application("csTimeUpcheReturn") %>", "yyyy-MM-dd a h:mm:ss"));
var csTimeIpjumStockOut = new Date(getDateFromFormat("<%= application("csTimeIpjumStockOut") %>", "yyyy-MM-dd a h:mm:ss"));
var csTimeTenTenStockOut = new Date(getDateFromFormat("<%= application("csTimeTenTenStockOut") %>", "yyyy-MM-dd a h:mm:ss"));

function DisplayClock() {
	var v = new Date();

	var objOneToOneBoard = document.getElementById("objOneToOneBoard");
	var objMichulgoListUpche = document.getElementById("objMichulgoListUpche");
	var objIpjumMichulgo = document.getElementById("objIpjumMichulgo");
	var objMichulgoListTenTen = document.getElementById("objMichulgoListTenTen");
	var objCSList = document.getElementById("objCSList");

	var objUpcheReturn7 = document.getElementById("objUpcheReturn7");
	var objUpcheReturn3 = document.getElementById("objUpcheReturn3");

	var objIpjumStockOut = document.getElementById("objIpjumStockOut");
	var objTenTenStockOut = document.getElementById("objTenTenStockOut");

	objOneToOneBoard.innerHTML = GetDateDiffString(v.getTime() - csTimeOneToOneBoard.getTime());
	objMichulgoListUpche.innerHTML = GetDateDiffString(v.getTime() - csTimeMichulgoListUpche.getTime());
	objIpjumMichulgo.innerHTML = GetDateDiffString(v.getTime() - csTimeIpjumMichulgo.getTime());
	objMichulgoListTenTen.innerHTML = GetDateDiffString(v.getTime() - csTimeMichulgoListTenTen.getTime());
	objCSList.innerHTML = GetDateDiffString(v.getTime() - csTimeCSList.getTime());

	objUpcheReturn7.innerHTML = GetDateDiffString(v.getTime() - csTimeUpcheReturn.getTime());
	objUpcheReturn3.innerHTML = GetDateDiffString(v.getTime() - csTimeUpcheReturn.getTime());

	objIpjumStockOut.innerHTML = GetDateDiffString(v.getTime() - csTimeIpjumStockOut.getTime());
	objTenTenStockOut.innerHTML = GetDateDiffString(v.getTime() - csTimeTenTenStockOut.getTime());

	setTimeout('DisplayClock();','1000');
}

function GetDateDiffString(v) {
	var result = "";

	if (v < (60 * 1000)) {
		v = v / 1000;
		result = parseInt(v) + "초 전";
	} else if (v < (60 * 60 * 1000)) {
		v = v / (60 * 1000);
		result = parseInt(v) + "분 전";
	} else {
		result =  "1시간 전";
	}

	return result;
}

function RefreshData(v) {
	var frm = document.frm;

	frm.mode.value = "RefreshData";
	frm.csTime.value = v;
	frm.submit();
}

window.onload = function() {
	DisplayClock();
}

</script>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
<form name="frm" method="post" action="cscenter_main_process.asp">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mode" value="">
<input type="hidden" name="csTime" value="">
</form>
<tr>
    <!-- 왼쪽메뉴 시작 -->
	<td width="33%" valign="top">
	    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
        <tr valign="top">
            <td>
        	    <!-- 주문내역검색 -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("menubar") %>">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
                        <tr height="25">
            			    <td>
            			    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>주문내역 검색</b>
            			    </td>
            			    <td align="right">
            			        <a href="javascript:PopOrderMasterWithCallRing();"> 주문내역보기 <img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a>
            			    </td>
            			</tr>
            	        </table>
            	    </td>
            	</tr>
            	</table>
        	    <!-- 주문내역검색 -->
        	</td>
        </tr>
        <tr valign="top">
            <td height="10"></td>
        </tr>
        <tr valign="top">
            <td>
        	    <!-- 게시판 관리 시작-->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("menubar") %>">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
                        <tr height="25">
            			    <td style="border-bottom:1px solid #BABABA">
            			    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>1:1 상담게시판 관리</b>
            			    	(<span id="objOneToOneBoard"></span>) <a href="javascript:RefreshData('csTimeOneToOneBoard')"><img src="/images/icon_reload.gif" border="0"></a>
            			    </td>
            			    <td align="right" style="border-bottom:1px solid #BABABA">
            			        <a href="javascript:PopMyQnaList('', '', '');">
        				        바로가기
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td>담당자별 미처리 상담건</td>
            			    <td align="right"></td>
            			</tr>
            			<%
            			'// 첫번째것 스킵한다.(에러방지)
            			for i = 1 to UBound(csBoardUserArr)
            			%>
            			<tr height="25">
            			    <td>&nbsp;&nbsp;- <%= csBoardUserArr(i) %></td>
            			    <td align="right">
            			        <a href="javascript:PopMyQnaListByChargeId('<%= csBoardUserArr(i) %>', 'N');">
        				        <b><%= csBoardChargecntArr(i) %></b> 건
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
            			</tr>
            			<% next %>
            			<tr height="25">
            			    <td>&nbsp;&nbsp;- 미분배 상담건</td>
            			    <td align="right">
            			        <a href="javascript:PopMyQnaListByChargeId('all', 'N');">
        				        <b><%= csBoardNochargecnt %></b> 건
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
            			</tr>
            	        </table>
            	    </td>
            	</tr>
            	</table>
        	    <!-- 게시판 관리 끝-->
        	</td>
        </tr>
        <tr valign="top">
            <td height="10"></td>
        </tr>
        <tr valign="top">
            <td width="50%">
        	    <!-- 업체배송 관리 시작 -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("menubar") %>">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
                        <tr height="25">
            			    <td style="border-bottom:1px solid #BABABA">
            			    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>미출고리스트[업체배송]</b>
								(<span id="objMichulgoListUpche"></span>) <a href="javascript:RefreshData('csTimeMichulgoListUpche')"><img src="/images/icon_reload.gif" border="0"></a>
            			    </td>
            			    <td align="right" style="border-bottom:1px solid #BABABA">
            			        <a href="javascript:upchebeasongmaster();">
        				        바로가기
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td>2일 이상 미확인건</td>
            			    <td align="right"><b><%=FormatNumber(arrMiSend(0),0)%></b> 건<a href="javascript:popUpcheMisend('','2','2');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<tr height="25">
            			    <td>3일 이상 미발송건</td>
            			    <td align="right"><b><%=FormatNumber(arrMiSend(1),0)%></b> 건<a href="javascript:popUpcheMisend('','3','3');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<tr height="25">
            			    <td>담당자별 D+3 미발송건(입점몰,출고예정 제외)</td>
            			    <td align="right"></td>
            			</tr>
            			<%
            			'// 첫번째것 스킵한다.(에러방지)
            			for i = 1 to UBound(csMichulgoBoardUserArr)
            			%>
            			<tr height="25">
            			    <td>&nbsp;&nbsp;* <%= csMichulgoBoardUserArr(i) %></td>
            			    <td align="right">
            			        <a href="javascript:showHideBrand('<%= csMichulgoBoardUserArr(i) %>');">
        				        <b><%= csMichulgoBrandCountByUserid(i) %></b> (<%= csMichulgoOrderCountByUserid(i) %>)
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
            			</tr>
	            			<% for j = 0 to UBound(csMichulgoBrandNameArr) %>
	            				<% if (csMichulgoBoardUserArr(i) = csMichulgoBrandUseridArr(j)) then %>
		            			<tr height="25" id="<%= csMichulgoBoardUserArr(i) %>" style="display:none">
		            			    <td>&nbsp;&nbsp;&nbsp;&nbsp;- <%= csMichulgoBrandNameArr(j) %></td>
		            			    <td align="right">
		            			        <a href="javascript:popUpcheMisendByBrand('<%= csMichulgoBrandNameArr(j) %>', '','3','');">
		        				        <%= csMichulgoBrandcntArr(j) %> 건
		        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
		        				        </a>
		            			    </td>
		            			</tr>
	            				<% end if %>
	            			<% next %>
            			<% next %>
            			<tr height="25">
            			    <td>&nbsp;&nbsp;- 미분배 미발송건</td>
            			    <td align="right">
        				        <b><%= nochargeMichulgoBrandcnt %></b>
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td>품절취소요청건</td>
            			    <td align="right"><b><%=FormatNumber(arrMiSend(2),0)%></b> 건<a href="javascript:popUpcheMisend('05','','');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<!--
            			<tr height="25">
            			    <td colspan="2"></td>
            			 	too Many
            			    <td>미출고사유 미입력건 (D+2이상)</td>
            			    <td align="right"><b>??</b> 건<img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></td>
            			</tr>
            			-->

            	        </table>
            	    </td>
            	</tr>
            	</table>
        	    <!-- 업체배송 관리 끝-->
        	</td>
		</tr>


        <tr valign="top">
            <td height="10"></td>
        </tr>

        <tr valign="top">
            <td width="50%">
        	    <!-- 업체배송 관리 시작 -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("menubar") %>">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
                        <tr height="25">
            			    <td style="border-bottom:1px solid #BABABA">
            			    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>입점몰 미출고리스트[텐배+업배]</b>
								(<span id="objIpjumMichulgo"></span>) <a href="javascript:RefreshData('csTimeIpjumMichulgo')"><img src="/images/icon_reload.gif" border="0"></a>
            			    </td>
            			    <td align="right" style="border-bottom:1px solid #BABABA">
            			    	&nbsp;
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td>입점몰별(출고예정 제외)</td>
            			    <td align="right">
            			    </td>
            			</tr>
            			<%
            			for i = 0 to UBound(csMichulgoOtherSitenameArr)
            			%>
            			<tr height="25">
            			    <td>&nbsp;&nbsp;* <%= csMichulgoOtherSitenameArr(i) %></td>
            			    <td align="right">
            			        <a href="javascript:showHideBrand('Michulgoothersite<%= csMichulgoOtherSitenameArr(i) %>');">
        				        <b><%= csMichulgoOtherSiteOrderCountArr(i) %></b> (<%= csMichulgoOtherSiteOrderDetailCountArr(i) %>)
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
            			</tr>
	            			<% for j = 0 to UBound(csMichulgoOtherSiteOrderserialArr) %>
	            				<% if (csMichulgoOtherSitenameArr(i) = csMichulgoOtherSiteOrderSitenameArr(j)) then %>
		            			<tr height="25" id="Michulgoothersite<%= csMichulgoOtherSitenameArr(i) %>" style="display:none">
		            			    <td>
		            			    	&nbsp;&nbsp;&nbsp;&nbsp;-
		            			    	<a href="javascript:PopOrderMasterWithCallRingOrderserial('<%= csMichulgoOtherSiteOrderserialArr(j) %>')">
			            			    	<%= csMichulgoOtherSiteOrderserialArr(j) %>
			            			    	<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
		            			    	</a>
		            			    </td>
		            			    <td align="right">
		            			        <a href="javascript:popUpcheMichulgoOtherSiteByOrderserial('<%= csMichulgoOtherSiteOrderserialArr(j) %>', '<%= csMichulgoOtherSitenameArr(i) %>');">
		        				        <%= csMichulgoOtherSiteOrderdetailCntcntArr(j) %> 건
		        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
		        				        </a>
		            			    </td>
		            			</tr>
	            				<% end if %>
	            			<% next %>
            			<% next %>
            	        </table>
            	    </td>
            	</tr>
            	</table>
        	    <!-- 업체배송 관리 끝-->
        	</td>
        </tr>

        <tr valign="top">
            <td height="10"></td>
        </tr>
        <tr valign="top">
            <td width="50%">
        	    <!-- 텐바이텐 미출고건 시작-->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("menubar") %>">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
                        <tr height="25">
            			    <td style="border-bottom:1px solid #BABABA">
            			    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>미출고리스트[텐바이텐]</b>
								(<span id="objMichulgoListTenTen"></span>) <a href="javascript:RefreshData('csTimeMichulgoListTenTen')"><img src="/images/icon_reload.gif" border="0"></a>
            			    </td>
            			    <td align="right" style="border-bottom:1px solid #BABABA">
                    			<a href="javascript:misendmaster('Y');">
        				        바로가기
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td>사유 미입력건</td>
            			    <td align="right"><a href="javascript:misendmaster('N');"><b><%=FormatNumber(arrTenMiSend(0),0)%></b> 건<img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<tr height="25">
            			    <td>고객 미안내건</td>
            			    <td align="right"><a href="javascript:misendmaster('N');"><b><%=FormatNumber(arrTenMiSend(1),0)%></b> 건<img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<tr height="25">
            			    <td>SMS/Mail/통화완료건</td>
            			    <td align="right"><a href="javascript:misendmaster('4');"><b><%=FormatNumber(arrTenMiSend(2),0)%></b> 건<img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<tr height="25">
            			    <td>CS처리완료(물류팀처리요청)건</td>
            			    <td align="right"><a href="javascript:misendmaster('6');"><b><%=FormatNumber(arrTenMiSend(3),0)%></b> 건<img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            	        </table>
            	    </td>
            	</tr>
            	</table>
        	    <!--  cs처리요청 끝-->
        	</td>
        </tr>

        </table>
    </td>
    <!-- 왼쪽메뉴 끝 -->

    <td width="10"></td>

    <!-- 가운데메뉴 시작 -->
    <td width="33%" valign="top">
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">

        <tr valign="top">
        	<td>
        	    <!-- CS처리리스트 시작-->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("menubar") %>">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
                        <tr height="25">
            			    <td style="border-bottom:1px solid #BABABA">
            			    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>CS처리리스트 관리</b>
            			    	(<span id="objCSList"></span>) <a href="javascript:RefreshData('csTimeCSList')"><img src="/images/icon_reload.gif" border="0"></a>
            			    </td>
            			    <td align="right" style="border-bottom:1px solid #BABABA">
            			        <a href="javascript:Cscenter_Action_List2('');">
        				        바로가기
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td>환불미처리(접수)</td>
            			    <td align="right"><b><%= csRefundRequestRegCount %></b> 건<a href="javascript:Cscenter_Action_List2('norefund');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<tr height="25">
            			    <td>환불미처리(확인요청)</td>
            			    <td align="right"><b><%= csRefundRequestConfirmCount %></b> 건<a href="javascript:Cscenter_Action_List2('confirm');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<tr height="25">
            			    <td>마일리지/예치금 환불미처리</td>
            			    <td align="right"><b><%= csNotFinMileRefund %></b> 건<a href="javascript:Cscenter_Action_List2('norefundmile');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<tr height="25">
            			    <td>카드취소 미처리</td>
            			    <td align="right"><b><%= csRequestCardCancelCount %></b> 건<a href="javascript:Cscenter_Action_List2('cardnocheck');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<tr height="25">
            			    <td>주문취소 미처리</td>
            			    <td align="right"><b><%= csNotFinA008 %></b> 건<a href="javascript:Cscenter_Action_List2('cancelnofinish');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<tr height="25">
            			    <td>회수요청 미처리</td>
            			    <td align="right"><b><%= csReturnNotFinish %></b> 건<a href="javascript:Cscenter_Action_List2('returnmifinish');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<!--
            			<tr height="25">
            			    <td>업체 미처리</td>
            			    <td align="right"><b><%= csUpcheNotFin %></b> 건<a href="javascript:Cscenter_Action_List2('upchemifinish');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			-->
            			<tr height="25">
            			    <td>반품접수(업체배송) D+7 미처리</td>
            			    <td align="right"><b><%= CSUpcheReturnNotFinish7 %></b> 건<a href="javascript:Cscenter_Action_MiFinishReturnListDPlus(7);"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>

            			<tr height="25">
            			    <td>업체처리완료</td>
            			    <td align="right"><b><%= csUpcheFinished %></b> 건<a href="javascript:Cscenter_Action_List2('upchefinish');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<!--
            			<tr height="25">
            			    <td>외부몰환불 미처리</td>
            			    <td align="right"><b><%= csNotFinA005 %></b> 건<a href="javascript:Cscenter_Action_List2('norefundetc');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			-->
            	        </table>
            	    </td>
            	</tr>
            	</table>
        	    <!--  CS처리리스트 끝-->
        	</td>
        </tr>

        <tr valign="top">
            <td height="10"></td>
        </tr>

        <tr valign="top">
            <td width="50%">
        	    <!-- 반품접수(업체배송) D+3 미처리 시작 -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("menubar") %>">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
                        <tr height="25">
            			    <td style="border-bottom:1px solid #BABABA">
            			    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>반품접수(업배) D+3 미처리</b>
            			    	(<span id="objUpcheReturn3"></span>) <a href="javascript:RefreshData('csTimeUpcheReturn')"><img src="/images/icon_reload.gif" border="0"></a>
            			    </td>
            			    <td align="right" style="border-bottom:1px solid #BABABA">
            			        <a href="javascript:Cscenter_Action_MiFinishReturnListDPlus(3);">
        				        바로가기
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td>반품접수(업체배송) D+3 미처리</td>
            			    <td align="right"><b><%= CSUpcheReturnNotFinish3 %></b> 건<a href="javascript:Cscenter_Action_MiFinishReturnListDPlus(3);"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<tr height="25">
            			    <td>담당자별 D+3 미처리건(사유입력 제외)</td>
            			    <td align="right"></td>
            			</tr>
            			<%
            			'// 첫번째것 스킵한다.(에러방지)
            			for i = 1 to UBound(csReturnBoardUserArr)
            			%>
            			<tr height="25">
            			    <td>&nbsp;&nbsp;* <%= csReturnBoardUserArr(i) %></td>
            			    <td align="right">
            			        <a href="javascript:showHideBrand('return3<%= csReturnBoardUserArr(i) %>');">
        				        <b><%= csReturn3BrandCountByUserid(i) %></b> (<%= csReturn3OrderCountByUserid(i) %>)
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
            			</tr>
	            			<% for j = 0 to UBound(csReturn3BrandNameArr) %>
	            				<% if (csReturnBoardUserArr(i) = csReturn3BrandUseridArr(j)) then %>
		            			<tr height="25" id="return3<%= csReturnBoardUserArr(i) %>" style="display:none">
		            			    <td>&nbsp;&nbsp;&nbsp;&nbsp;- <%= csReturn3BrandNameArr(j) %></td>
		            			    <td align="right">
		            			        <a href="javascript:Cscenter_Action_MiFinishReturnListByBrand('<%= csReturn3BrandNameArr(j) %>', 3);">
		        				        <%= csReturn3BrandcntArr(j) %> 건
		        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
		        				        </a>
		            			    </td>
		            			</tr>
	            				<% end if %>
	            			<% next %>
            			<% next %>
            			<tr height="25">
            			    <td>&nbsp;&nbsp;- 미분배 미처리건</td>
            			    <td align="right">
        				        <b><%= nochargeReturn3Brandcnt %></b>
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
            			    </td>
            			</tr>
            	        </table>
            	    </td>
            	</tr>
            	</table>
        	    <!-- 반품접수(업체배송) D+3 미처리 끝-->
        	</td>
        </tr>

        <tr valign="top">
            <td height="10"></td>
        </tr>

        <tr valign="top">
            <td width="50%">
        	    <!-- 반품접수(업체배송) D+7 미처리 시작 -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("menubar") %>">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
                        <tr height="25">
            			    <td style="border-bottom:1px solid #BABABA">
            			    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>반품접수(업배) D+7 미처리</b>
            			    	(<span id="objUpcheReturn7"></span>) <a href="javascript:RefreshData('csTimeUpcheReturn')"><img src="/images/icon_reload.gif" border="0"></a>
            			    </td>
            			    <td align="right" style="border-bottom:1px solid #BABABA">
            			        <a href="javascript:Cscenter_Action_MiFinishReturnListDPlus(7);">
        				        바로가기
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td>반품접수(업체배송) D+7 미처리</td>
            			    <td align="right"><b><%= CSUpcheReturnNotFinish7 %></b> 건<a href="javascript:Cscenter_Action_MiFinishReturnListDPlus(7);"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<tr height="25">
            			    <td>담당자별 D+7 미처리건(반품예정 제외)</td>
            			    <td align="right"></td>
            			</tr>
            			<%
            			'// 첫번째것 스킵한다.(에러방지)
            			for i = 1 to UBound(csReturnBoardUserArr)
            			%>
            			<tr height="25">
            			    <td>&nbsp;&nbsp;* <%= csReturnBoardUserArr(i) %></td>
            			    <td align="right">
            			        <a href="javascript:showHideBrand('return7<%= csReturnBoardUserArr(i) %>');">
        				        <b><%= csReturn7BrandCountByUserid(i) %></b> (<%= csReturn7OrderCountByUserid(i) %>)
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
            			</tr>
	            			<% for j = 0 to UBound(csReturn7BrandNameArr) %>
	            				<% if (csReturnBoardUserArr(i) = csReturn7BrandUseridArr(j)) then %>
		            			<tr height="25" id="return7<%= csReturnBoardUserArr(i) %>" style="display:none">
		            			    <td>&nbsp;&nbsp;&nbsp;&nbsp;- <%= csReturn7BrandNameArr(j) %></td>
		            			    <td align="right">
		            			        <a href="javascript:Cscenter_Action_MiFinishReturnListByBrand('<%= csReturn7BrandNameArr(j) %>', 7);">
		        				        <%= csReturn7BrandcntArr(j) %> 건
		        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
		        				        </a>
		            			    </td>
		            			</tr>
	            				<% end if %>
	            			<% next %>
            			<% next %>
            			<tr height="25">
            			    <td>&nbsp;&nbsp;- 미분배 미처리건</td>
            			    <td align="right">
        				        <b><%= nochargeReturn7Brandcnt %></b>
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
            			    </td>
            			</tr>
            	        </table>
            	    </td>
            	</tr>
            	</table>
        	    <!-- 반품접수(업체배송) D+7 미처리 끝-->
        	</td>
        </tr>

        <tr valign="top">
            <td height="10"></td>
        </tr>

        <tr valign="top">
            <td width="50%">
        	    <!-- 업체배송 관리 시작 -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("menubar") %>">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
                        <tr height="25">
            			    <td style="border-bottom:1px solid #BABABA">
            			    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>입점몰 품절취소요청건</b>
								(<span id="objIpjumStockOut"></span>) <a href="javascript:RefreshData('csTimeIpjumStockOut')"><img src="/images/icon_reload.gif" border="0"></a>
            			    </td>
            			    <td align="right" style="border-bottom:1px solid #BABABA">
            			    	&nbsp;
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td>입점몰별(고객 안내이전)</td>
            			    <td align="right">
            			    </td>
            			</tr>
            			<%
            			for i = 0 to UBound(csStockoutOtherSitenameArr)
            			%>
            			<tr height="25">
            			    <td>&nbsp;&nbsp;* <%= csStockoutOtherSitenameArr(i) %></td>
            			    <td align="right">
            			        <a href="javascript:showHideBrand('stockoutothersite<%= csStockoutOtherSitenameArr(i) %>');">
        				        <b><%= csStockoutOtherSiteOrderCountArr(i) %></b> (<%= csStockoutOtherSiteOrderDetailCountArr(i) %>)
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
            			</tr>
	            			<% for j = 0 to UBound(csStockoutOtherSiteOrderserialArr) %>
	            				<% if (csStockoutOtherSitenameArr(i) = csStockoutOtherSiteOrderSitenameArr(j)) then %>
		            			<tr height="25" id="stockoutothersite<%= csStockoutOtherSitenameArr(i) %>" style="display:none">
		            			    <td>
		            			    	&nbsp;&nbsp;&nbsp;&nbsp;-
		            			    	<a href="javascript:PopOrderMasterWithCallRingOrderserial('<%= csStockoutOtherSiteOrderserialArr(j) %>')">
			            			    	<%= csStockoutOtherSiteOrderserialArr(j) %>
			            			    	<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
		            			    	</a>
		            			    </td>
		            			    <td align="right">
		            			        <a href="javascript:popUpcheStockoutOtherSiteByOrderserial('<%= csStockoutOtherSiteOrderserialArr(j) %>', '<%= csStockoutOtherSitenameArr(i) %>');">
		        				        <%= csStockoutOtherSiteOrderdetailCntcntArr(j) %> 건
		        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
		        				        </a>
		            			    </td>
		            			</tr>
	            				<% end if %>
	            			<% next %>
            			<% next %>
            	        </table>
            	    </td>
            	</tr>
            	</table>
        	    <!-- 업체배송 관리 끝-->
        	</td>
        </tr>
        <tr valign="top">
            <td height="10"></td>
        </tr>
        <tr valign="top">
        	<td>
        	    <!-- 업체배송 관리 시작 -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("menubar") %>">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
                        <tr height="25">
            			    <td style="border-bottom:1px solid #BABABA">
            			    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>품절취소요청건[텐배+업배]</b>
								(<span id="objTenTenStockOut"></span>) <a href="javascript:RefreshData('csTimeTenTenStockOut')"><img src="/images/icon_reload.gif" border="0"></a>
            			    </td>
            			    <td align="right" style="border-bottom:1px solid #BABABA">
            			    	&nbsp;
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td>담당자별(입점몰제외, 고객 안내이전)</td>
            			    <td align="right">
            			    </td>
            			</tr>
            			<%
            			'// 첫번째것 스킵한다.(에러방지)
            			for i = 1 to UBound(csMichulgoStockoutBoardUserArr)
            			%>
            			<tr height="25">
            			    <td>&nbsp;&nbsp;* <%= csMichulgoStockoutBoardUserArr(i) %></td>
            			    <td align="right">
            			        <a href="javascript:showHideBrand('stockout<%= csMichulgoStockoutBoardUserArr(i) %>');">
        				        <b><%= csStockoutOrderCountByUserid(i) %></b> (<%= csStockoutOrderDetailCountByUserid(i) %>)
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
            			</tr>
	            			<% for j = 0 to UBound(csStockoutOrderserialArr) %>
	            				<% if (csMichulgoStockoutBoardUserArr(i) = csStockoutOrderUseridArr(j)) then %>
									<% csStockoutHasChargeUser(j) = "Y" %>
		            			<tr height="25" id="stockout<%= csMichulgoStockoutBoardUserArr(i) %>" style="display:none">
		            			    <td>
		            			    	&nbsp;&nbsp;&nbsp;&nbsp;-
		            			    	<a href="javascript:PopOrderMasterWithCallRingOrderserial('<%= csStockoutOrderserialArr(j) %>')">
			            			    	<%= csStockoutOrderserialArr(j) %>
			            			    	<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
		            			    	</a>
		            			    </td>
		            			    <td align="right">
		            			        <a href="javascript:popUpcheStockout10x10ByOrderserial('<%= csStockoutOrderserialArr(j) %>');">
		        				        <%= csStockoutOrderdetailCntcntArr(j) %> 건
		        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
		        				        </a>
		            			    </td>
		            			</tr>
	            				<% end if %>
	            			<% next %>
            			<% next %>
            			<tr height="25">
            			    <td>&nbsp;&nbsp;- 미분배</td>
            			    <td align="right">
        				        <b><%= nochargeStockoutOrdercnt %></b>
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
            			    </td>
            			</tr>
            			<% for j = 0 to UBound(csStockoutOrderserialArr) %>
            				<% if (csStockoutHasChargeUser(j) = "N") then %>
							<form name="frm<%= j %>" method="post" action="cscenter_main_process.asp">
							<input type="hidden" name="menupos" value="<%= menupos %>">
							<input type="hidden" name="mode" value="setstockoutchargeuser">
							<input type="hidden" name="orderdetailidx" value="<%= csStockoutOrderdetaiidx(j) %>">
	            			<tr height="25">
	            			    <td>
	            			    	&nbsp;&nbsp;&nbsp;&nbsp;- <%= csStockoutOrderserialArr(j) %>
	            			    	<!--
	            			    	<% if Not IsNull(csStockoutOrderUseridArr(i)) and csStockoutOrderUseridArr(i) <> "" then %>
	            			    		(<%= csStockoutOrderUseridArr(i) %>)
	            			    	<% end if %>
	            			    	-->
	            			    </td>
	            			    <td align="right">
	            			        <input type="button" class="button" value="분배" onclick="setUpcheStockout10x10ByOrderserial(frm<%= j %>)">
	            			    </td>
	            			</tr>
	            			</form>
            				<% end if %>
            			<% next %>
            	        </table>
            	    </td>
            	</tr>
            	</table>
        	    <!-- 업체배송 관리 끝-->
        	</td>
        </tr>

      	</table>
    </td>
    <!-- 가운데메뉴 끝 -->

    <td width="10"></td>

    <!-- 오픈쪽메뉴 시작 -->
    <td valign="top">
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
        <tr valign="top">
            <td>
                <!-- 새로고침 시작 -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("tabletop") %>">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
                        <tr height="25">
                        	<td>
            			    	<img src="/images/icon_star.gif" align="absbottom">
								<b>ID : </b>
								<input type="text" class="text" id="csMainUserID" value="<%=csMainUserID%>" size="10">
								<input type="button" class="button" value="검색" onclick="location.href = 'cscenter_main.asp?menupos=757&csMainUserID=' + document.getElementById('csMainUserID').value;">
								<!-- 초기로그인시 로그인 아이디로 설정 / 다른아이디로도 검색가능하도록 -->
            			    </td>
            			    <td align="right">
            			    	<a href="javascript:document.location.reload();">
        				        새로고침
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
            			</tr>
            	        </table>
            	    </td>
            	</tr>
            	</table>
            	<!-- 새로고침 끝 -->
            </td>
        </tr>

        <tr valign="top">
            <td height="10"></td>
        </tr>

        <tr valign="top">
            <td>
                <!-- 아이디별 확인사항 -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("tabletop") %>">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
            			<tr height="25">
            			    <td style="border-bottom:1px solid #BABABA">
            			        <img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>아이디별 확인사항</b>
            			    </td>
            			    <td align="right" style="border-bottom:1px solid #BABABA">
            			        &nbsp;
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td>미처리메모</td>
            			    <td align="right">
            			        <b><%= CSMemoNotFinish %></b> 건
        				    	<a href="javascript:cscenter_memo_list('','','N', '<%=csMainUserID%>');">
                    		    <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
                    		    </a>
            			    </td>
            			</tr>
            	        </table>
            	    </td>
            	</tr>
            	</table>
        	    <!--  아이디별 확인사항 끝-->
            </td>
        </tr>

        <tr valign="top">
            <td height="10"></td>
        </tr>
        <%
        	Dim NewCoop
        	Set NewCoop = new CCooperate
        	NewCoop.FDoc_Id = session("ssBctId")
        	NewCoop.fnGetCooperateCount			' 리얼에 올리기전 확인 후 주석 제거


        %>
        <tr valign="top">
            <td>
                <!-- 협조문 시작 -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("tabletop") %>">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
            			<tr height="25">
            			    <td>받은 협조문 (미처리)</td>
            			    <td align="right">
            			        <a href="javascript:popCooperate();">
            			        <%
            			        	If NewCoop.FComeCnt = 0 Then
            			        		Response.Write "[" & NewCoop.FComeCnt & "] 건"
            			        	Else
            			        		Response.Write "[<b>" & NewCoop.FComeCnt & "</b>] 건...<img src='http://fiximage.10x10.co.kr/web2009/main/news_icon_new.gif' border='0'>"
            			        	End If
            			        %>
            			        </a>
        				    	<a href="javascript:popCooperate();">
                    		    <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
                    		    </a>
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td>보낸 협조문 (미처리)</td>
            			    <td align="right">
            			        <a href="javascript:popCooperate();">
            			        <%
            			        	If NewCoop.FSendCnt = 0 Then
            			        		Response.Write "[" & NewCoop.FSendCnt & "] 건"
            			        	Else
            			        		Response.Write "[<b>" & NewCoop.FSendCnt & "</b>] 건...<img src='http://fiximage.10x10.co.kr/web2009/main/news_icon_new.gif' border='0'>"
            			        	End If
            			        %>
            			        </a>
        				    	<a href="javascript:popCooperate();">
                    		    <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
                    		    </a>
            			    </td>
            			</tr>
            	        </table>
            	    </td>
            	</tr>
            	</table>
        	    <!--  협조문 끝-->
            </td>
        </tr>


        <tr valign="top">
            <td height="10"></td>
        </tr>

        <tr valign="top">
            <td>
        		<!-- 각종 서류 발행 시작 -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("tabletop") %>">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
            			<tr height="25">
            			    <td style="border-bottom:1px solid #BABABA">
            			        <img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>고객환불처리</b>
            			    </td>
            			    <td align="right" style="border-bottom:1px solid #BABABA">
            			        <a href="javascript:cscenter_refund_list();">
        				        바로가기
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td style="border-bottom:1px solid #BABABA">
            			        <img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>현금영수증 발행</b>
            			    </td>
            			    <td align="right" style="border-bottom:1px solid #BABABA">
            			        <a href="javascript:cscenter_cashreceipt_list();">
        				        바로가기
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
				                </a>
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td style="border-bottom:1px solid #BABABA">
            			        <img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>세금계산서 발행</b>
            			    </td>
            			    <td align="right" style="border-bottom:1px solid #BABABA">
            			        <a href="javascript:cscenter_tax_list();">
        				        바로가기
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td style="border-bottom:1px solid #BABABA">
            			        <img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>무통장입금 관리</b>
            			    </td>
            			    <td align="right" style="border-bottom:1px solid #BABABA">
            			        <a href="javascript:cscenter_payment_list(1);">
        				        바로가기
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
            			</tr>
            	        </table>
            	    </td>
            	</tr>
            	</table>
        	    <!--  각종 서류 발행 끝-->
        	</td>
        </tr>

        <tr valign="top">
            <td height="10"></td>
        </tr>

		<tr valign="top">
            <td>
            	<!-- 각종조회 -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("tabletop") %>">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
                        <tr height="25">
            			    <td>
            			    	<table width="100%" style="border:1px solid #BABABA" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#FFFFFF">
		                        <tr height="20">
		            			    <td align="center">
		            			    	<a href="javascript:cscenter_mileage_list('');">마일리지조회</a>
		            			    </td>
		            			</tr>
		            	        </table>
            			    </td>
            			    <td width="5"></td>
            			    <td>
            			    	<table width="100%" style="border:1px solid #BABABA" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#FFFFFF">
		                        <tr height="20">
		            			    <td align="center">
		            			    	<a href="javascript:cscenter_deposit_list('');">예치금조회</a>
		            			    </td>
		            			</tr>
		            	        </table>
            			    </td>
            			    <td width="5"></td>
            			    <td>
            			    	<table width="100%" style="border:1px solid #BABABA" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#FFFFFF">
		                        <tr height="20">
		            			    <td align="center">
		            			    	<a href="javascript:cscenter_coupon_list('');">쿠폰조회</a>
		            			    </td>
		            			</tr>
		            	        </table>
            			    </td>
            			    <td width="5"></td>
            			 </tr>
            			 <tr>
            			    <td>
            			    	<table width="100%" style="border:1px solid #BABABA" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#FFFFFF">
		                        <tr height="20">
		            			    <td align="center">
		            			    	<a href="javascript:cscenter_eventjoin_list('');">참여이벤트조회</a>
		            			    </td>
		            			</tr>
		            	        </table>
            			    </td>
            			    <td width="5"></td>
            			    <td>
            			    	<table width="100%" style="border:1px solid #BABABA" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#FFFFFF">
		                        <tr height="20">
		            			    <td align="center">
		            			    	<a href="javascript:cscenter_member_list('');">회원조회</a>
		            			    </td>
		            			</tr>
		            	        </table>
            			    </td>
            			    <td width="5"></td>
            			    <td>
            			    	<table width="100%" style="border:1px solid #BABABA" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#FFFFFF">
		                        <tr height="20">
		            			    <td align="center">
		            			    	<a href="javascript:cscenter_eventprize_list('');">당첨조회</a>
		            			    </td>
		            			</tr>
		            	        </table>
            			    </td>
            			    <td width="5"></td>
            			</tr>
            	        </table>
            	    </td>
            	</tr>
            	</table>
        	    <!-- 각종조회 -->
            </td>
        </tr>

        <tr valign="top">
            <td height="10"></td>
        </tr>

        <tr valign="top">
            <td>
            	<!-- SMS MAIL -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("tabletop") %>">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
                        <tr height="25">
            			    <td>
            			    	<table width="100%" style="border:1px solid #BABABA" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#FFFFFF">
		                        <tr height="20">
		            			    <td align="center">
		            			    	<a href="javascript:PopCSSMSSend('','','','');">SMS발송</a>
		            			    </td>
		            			</tr>
		            	        </table>
            			    </td>
            			    <td width="5"></td>
            			    <td>
            			    	<table width="100%" style="border:1px solid #BABABA" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#FFFFFF">
		                        <tr height="20">
		            			    <td align="center">
		            			    	<a href="javascript:PopCSMailSend('','');">메일발송</a>
		            			    </td>
		            			</tr>
		            	        </table>
            			    </td>
            			</tr>
            	        </table>
            	    </td>
            	</tr>
            	</table>
        	    <!-- SMS MAIL -->
            </td>
        </tr>

        </table>
    </td>
    <!-- 오픈쪽메뉴 끝 -->

</tr>
</table>

<% Call checkAndWriteElapsedTime("017") %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
