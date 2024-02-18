<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : cs센터
' History : 2009.04.17 이상구 생성
'			2016.03.25 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/upchebeasongcls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/classes/cooperate/cooperateCls.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_taxsheetReqCls.asp"-->
<%
dim CSMainIsDev

if application("Svr_Info") <> "Dev" then
	CSMainIsDev = "N"
else
	CSMainIsDev = "N"
end if

dim lastPageTime, pageElapsedTime
lastPageTime = Timer

function checkAndWriteElapsedTime(memo)
	pageElapsedTime = Timer - lastPageTime
	lastPageTime = Timer
	response.write "<!-- Page Execute Time Check : " & FormatNumber(pageElapsedTime, 4) & " : " & memo & " -->" & vbCrLf
end function

dim IsUpdateUserListNeed, IsUpdateOneToOneBoardNeed, IsUpdateMichulgoListUpcheNeed, IsUpdateMiAnswerUpcheNeed, IsUpdateMichulgoListTenTenNeed, IsUpdateIpjumMichulgoNeed, IsUpdateCSListNeed, IsUpdateUpcheReturnListNeed, IsUpdateIpjumStockOutNeed, IsUpdateIpjumRefundNeed, IsUpdateTenTenStockOutNeed
dim IsUpdateMaxCSMasterIdxNeed

IsUpdateUserListNeed			= False
IsUpdateOneToOneBoardNeed		= False
IsUpdateMichulgoListUpcheNeed	= False
IsUpdateMiAnswerUpcheNeed		= False
IsUpdateMichulgoListTenTenNeed	= False
IsUpdateIpjumMichulgoNeed		= False
IsUpdateCSListNeed				= False
IsUpdateUpcheReturnListNeed		= False
IsUpdateIpjumStockOutNeed		= False
IsUpdateIpjumRefundNeed			= False
IsUpdateTenTenStockOutNeed		= False
IsUpdateMaxCSMasterIdxNeed		= False

'' IsUpdateUserListNeed			= True
'' IsUpdateOneToOneBoardNeed		= True
'' IsUpdateMichulgoListUpcheNeed	= True
'' IsUpdateMiAnswerUpcheNeed		= True
'' IsUpdateMichulgoListTenTenNeed	= True
'' IsUpdateIpjumMichulgoNeed		= True
'' IsUpdateCSListNeed				= True
'' IsUpdateUpcheReturnListNeed		= True
'' IsUpdateIpjumStockOutNeed		= True
'' IsUpdateIpjumRefundNeed			= True
'' IsUpdateTenTenStockOutNeed		= True
'' IsUpdateMaxCSMasterIdxNeed		= True

'==============================================================================
'// 담당자 목록
If Trim(application("csTimeUserList")) = "" Or Trim(application("csVipBoardUserArr")) = "" Or Trim(application("csBoardUserArr")) = "" Or DateDiff("s", application("csTimeUserList"), Now() ) > 1800 Then						'// 30분(1800초) 초과시 새로 쿼리함
	application("csTimeUserList") = Now()
	IsUpdateUserListNeed = True
end if

'// 1:1 상담게시판
If Trim(application("csTimeOneToOneBoard")) = "" Or DateDiff("s", application("csTimeOneToOneBoard"), Now() ) > 600 Then			'// 10분(600초) 초과시 새로 쿼리함
	application("csTimeOneToOneBoard") = Now()
	IsUpdateOneToOneBoardNeed = True
end if

'// 미출고리스트[업체배송]
If Trim(application("csTimeMichulgoListUpcheNew")) = "" Or Trim(application("csMiSend")) = "" Or DateDiff("s", application("csTimeMichulgoListUpcheNew"), Now() ) > 1800 Then	'// 30분(1800초) 초과시 새로 쿼리함
	application("csTimeMichulgoListUpcheNew") = Now()
	IsUpdateMichulgoListUpcheNeed = True
end if

'// 상품문의 답변이전
If Trim(application("csTimeMiAnswerList")) = "" Or DateDiff("s", application("csTimeMiAnswerList"), Now() ) > 7200 Then	'// 120분(7200초) 초과시 새로 쿼리함
	application("csTimeMiAnswerList") = Now()
	IsUpdateMiAnswerUpcheNeed = True
end if

'// 미출고리스트[텐바이텐]
If Trim(application("csTimeMichulgoListTenTen")) = "" Or Trim(application("csTenMiSend")) = "" Or DateDiff("s", application("csTimeMichulgoListTenTen"), Now() ) > 900 Then	'// 15분(900초) 초과시 새로 쿼리함
	application("csTimeMichulgoListTenTen") = Now()
	IsUpdateMichulgoListTenTenNeed = True
end if

'// 미출고리스트[입점몰]
If Trim(application("csTimeIpjumMichulgo")) = "" Or Trim(application("csMiSend")) = "" Or DateDiff("s", application("csTimeIpjumMichulgo"), Now() ) > 1800 Then	'// 30분(1800초) 초과시 새로 쿼리함
	application("csTimeIpjumMichulgo") = Now()
	IsUpdateIpjumMichulgoNeed = True
end if

'// CS처리리스트
If Trim(application("csTimeCSList")) = "" Or DateDiff("s", application("csTimeCSList"), Now() ) > 900 Then	'// 15분(900초) 초과시 새로 쿼리함
	application("csTimeCSList") = Now()
	IsUpdateCSListNeed = True

	application("csTimeUpcheReturn") = Now()
	IsUpdateUpcheReturnListNeed = True
end if

'// 반품접수(업체배송)
If Trim(application("csTimeUpcheReturn")) = "" Or DateDiff("s", application("csTimeUpcheReturn"), Now() ) > 1800 Then	'// 30분(1800초) 초과시 새로 쿼리함
	application("csTimeCSList") = Now()
	IsUpdateCSListNeed = True

	application("csTimeUpcheReturn") = Now()
	IsUpdateUpcheReturnListNeed = True
end if

'// 입점몰 품절취소요청건
If Trim(application("csTimeIpjumStockOut")) = "" Or DateDiff("s", application("csTimeIpjumStockOut"), Now() ) > 900 Then	'// 15분(900초) 초과시 새로 쿼리함
	application("csTimeIpjumStockOut") = Now()
	IsUpdateIpjumStockOutNeed = True
end if

'// 입점몰 환불미처리건
If Trim(application("csTimeIpjumRefund")) = "" Or DateDiff("n", application("csTimeIpjumRefund"), Now() ) > 120 Then	'// 120분 초과시 새로 쿼리함
	application("csTimeIpjumRefund") = Now()
	IsUpdateIpjumRefundNeed = True
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

if IsUpdateUserListNeed or request("force") = "Y" then
	application("csTimeOneToOneBoard") = Now()
	IsUpdateOneToOneBoardNeed = True

	application("csTimeMichulgoListUpcheNew") = Now()
	IsUpdateMichulgoListUpcheNeed = True

	application("csTimeIpjumMichulgo") = Now()
	IsUpdateIpjumMichulgoNeed = True

	application("csTimeUpcheReturn") = Now()
	IsUpdateUpcheReturnListNeed = True

	application("csTimeIpjumStockOut") = Now()
	IsUpdateIpjumStockOutNeed = True

	application("csTimeIpjumRefund") = Now()
	IsUpdateIpjumRefundNeed = True

	application("csTimeTenTenStockOut") = Now()
	IsUpdateTenTenStockOutNeed = True
end if
'==============================================================================

dim sqlStr, i, j, k, resultCount, Found, sqlchargeidlist, tmpResultArr(), paramInfo3, strSql
dim onemonthbefore, nowdate1800, twomonthbefore, treemonthbefore, sixmonthbefore, nowdate, tomorrowdate
nowdate1800 = Left(now, 10) + " 18:00"
nowdate        = Left(now, 10)
tomorrowdate   = Left(DateAdd("d",  1, now), 10)
onemonthbefore = Left(DateAdd("m", -1, now), 10)
twomonthbefore = Left(DateAdd("m", -2, now), 10)
treemonthbefore= Left(DateAdd("m", -3, now), 10)
sixmonthbefore = Left(DateAdd("m", -6, now), 10)

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
dim csVipBoardUserArr, csVVipBoardUserArr, csBoardUserArr, csMichulgoStockoutBoardUserArr, csReturnBoardUserArr

if (IsUpdateUserListNeed = True) then
	'// YN 이 N 이 아닌것이어야 한다.
	'// 분배할 때는 YN 이 Y 인것, 분배받은것 표시할때는 N 이 아닌것!!
	'// YN = T 인 경우 : 분배정지(분배받은것은 표시되고, 더이상 분배받지 않는다.)
	sqlStr = " exec [db_cs].[dbo].[usp_Ten_Cs_ChargeUserList] "
	''response.write sqlStr
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

	'// 에러방지(담당자가 한명도 없으면 에러가 발생한다.)
	csVipBoardUserArr = "xxxxxxxx"
	csVVipBoardUserArr = "xxxxxxxx"
	csBoardUserArr = "xxxxxxxx"
	csMichulgoStockoutBoardUserArr = "xxxxxxxx"
	csReturnBoardUserArr = "xxxxxxxx"

	if  not rsget.EOF  then
		do until rsget.eof

			if (rsget("vipone2oneyn") = "Y") then
				csVipBoardUserArr = csVipBoardUserArr + "," + rsget("userid")
			end if

			if (rsget("vvipone2oneyn") = "Y") then
				csVVipBoardUserArr = csVVipBoardUserArr + "," + rsget("userid")
			end if

			if (rsget("one2oneyn") = "Y") then
				csBoardUserArr = csBoardUserArr + "," + rsget("userid")
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
	rsget.Close

	''response.write csVipBoardUserArr & "aa"
	''response.write csBoardUserArr & "aa"

	application("csVipBoardUserArr") 				= csVipBoardUserArr
	application("csVVipBoardUserArr") 				= csVVipBoardUserArr
	application("csBoardUserArr") 					= csBoardUserArr
	application("csMichulgoStockoutBoardUserArr") 	= csMichulgoStockoutBoardUserArr
	application("csReturnBoardUserArr") 			= csReturnBoardUserArr

	Call checkAndWriteElapsedTime("003")
end if

csVipBoardUserArr = Split(application("csVipBoardUserArr"), ",")
csVVipBoardUserArr = Split(application("csVVipBoardUserArr"), ",")
csBoardUserArr = Split(application("csBoardUserArr"), ",")
csMichulgoStockoutBoardUserArr = Split(application("csMichulgoStockoutBoardUserArr"), ",")
csReturnBoardUserArr = Split(application("csReturnBoardUserArr"), ",")

'==============================================================================
'// 1:1 상담게시판 관리(담당자별)
dim csVipBoardChargecntArr, csVipBoardNochargecnt, csVVipBoardChargecntArr, csVVipBoardNochargecnt
dim csBoardChargecntArr, csBoardNochargecnt, csrecommendBoardcnt, csstaffBoardcnt, csStaffStockoutOrderCount, csStaffStockoutOrderArr
dim csExtSiteCntArr

if (IsUpdateOneToOneBoardNeed = True) then

	'// 등급/담당자아이디/카운트
	sqlStr = " select"
	sqlStr = sqlStr & "	(case"

	if date() >= "2018-08-01" then
		sqlStr = sqlStr & " 	when userlevel in (2, 3) then 'V'"
		sqlStr = sqlStr & " 	when userlevel in (4) then 'VV'"
	else
		sqlStr = sqlStr & " 	when userlevel in (3, 4) then 'V'"
		sqlStr = sqlStr & " 	when userlevel in (6) then 'VV'"
	end if

	sqlStr = sqlStr & " 	else 'R' end) as userlevel"
	sqlStr = sqlStr & "	, IsNull(chargeid,'') as chargeid, count(*) as cnt"
	sqlStr = sqlStr & "	from [db_cs].[dbo].tbl_myqna "
	sqlStr = sqlStr & "	where isusing = 'Y' "
	sqlStr = sqlStr & "	and replydate is NULL "
	sqlStr = sqlStr & "	and isnull(qadiv,'')<>'26' "	'고객건의사항 문의는 제외
	sqlStr = sqlStr & "	and userlevel not in (7)"	'직원제외
	sqlStr = sqlStr & "	and IsNull(sitename, '10x10') = '10x10'"	'제휴몰CS 제외
	sqlStr = sqlStr & "	and regdate > '" + CStr(onemonthbefore) + "' "
	sqlStr = sqlStr & "	group by"
	sqlStr = sqlStr & "		(case"

	if date() >= "2018-08-01" then
		sqlStr = sqlStr & " 		when userlevel in (2, 3) then 'V'"
		sqlStr = sqlStr & " 		when userlevel in (4) then 'VV'"
	else
		sqlStr = sqlStr & " 		when userlevel in (3, 4) then 'V'"
		sqlStr = sqlStr & " 		when userlevel in (6) then 'VV'"
	end if

	sqlStr = sqlStr & " 		else 'R' end)"
	sqlStr = sqlStr & "		, IsNull(chargeid,'')"

	'response.write sqlStr & "<br>"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

	csVVipBoardChargecntArr = ""
	csVVipBoardNochargecnt = 0
	csVipBoardChargecntArr = ""
	csVipBoardNochargecnt = 0
	csBoardChargecntArr = ""
	csBoardNochargecnt = 0

	if not rsget.EOF then
		resultCount = rsget.RecordCount

		for i = 0 to resultCount - 1
			Found = False

			'// VIP
			If (rsget("userlevel") = "V") Then
				For j = 0 to UBound(csVipBoardUserArr)
					if (rsget("chargeid") = csVipBoardUserArr(j)) then
						Found = True
						If (csVipBoardChargecntArr = "") Then
							csVipBoardChargecntArr = rsget("chargeid") & "," & rsget("cnt")
						Else
							csVipBoardChargecntArr = csVipBoardChargecntArr & "|" & rsget("chargeid") & "," & rsget("cnt")
						End If

						Exit For
					End If
				Next

				If Found = False Then
					csVipBoardNochargecnt = csVipBoardNochargecnt + rsget("cnt")
				End If

			'// VVIP
			ElseIf (rsget("userlevel") = "VV") Then
				For j = 0 to UBound(csVVipBoardUserArr)
					if (rsget("chargeid") = csVVipBoardUserArr(j)) then
						Found = True
						If (csVVipBoardChargecntArr = "") Then
							csVVipBoardChargecntArr = rsget("chargeid") & "," & rsget("cnt")
						Else
							csVVipBoardChargecntArr = csVVipBoardChargecntArr & "|" & rsget("chargeid") & "," & rsget("cnt")
						End If

						Exit For
					End If
				Next

				If Found = False Then
					csVVipBoardNochargecnt = csVVipBoardNochargecnt + rsget("cnt")
				End If

			'// 일반고객
			Else
				For j = 0 to UBound(csBoardUserArr)
					if (rsget("chargeid") = csBoardUserArr(j)) then
						Found = True
						If (csBoardChargecntArr = "") Then
							csBoardChargecntArr = rsget("chargeid") & "," & rsget("cnt")
						Else
							csBoardChargecntArr = csBoardChargecntArr & "|" & rsget("chargeid") & "," & rsget("cnt")
						End If

						Exit For
					End If
				Next

				If Found = False Then
					csBoardNochargecnt = csBoardNochargecnt + rsget("cnt")
				End If
			End If

			rsget.movenext
		next
	end if
	rsget.Close

	application("csVVipBoardChargecntArr") = csVVipBoardChargecntArr
	application("csVVipBoardNochargecnt") = csVVipBoardNochargecnt
	application("csVipBoardChargecntArr") = csVipBoardChargecntArr
	application("csVipBoardNochargecnt") = csVipBoardNochargecnt
	application("csBoardChargecntArr") = csBoardChargecntArr
	application("csBoardNochargecnt") = csBoardNochargecnt

	'//고객건의사항문의 카운트		'/2016.03.28 한용민 추가
	sqlStr = " select count(*) as cnt "
	sqlStr = sqlStr + "	from [db_cs].[dbo].tbl_myqna "
	sqlStr = sqlStr + "	where isusing = 'Y' "
	sqlStr = sqlStr + "	and replydate is NULL "
	sqlStr = sqlStr + "	and isnull(qadiv,'')='26' "	'고객건의사항 문의
	sqlStr = sqlStr + "	and userlevel not in (7)"	'직원제외
	sqlStr = sqlStr + "	and regdate > '" + CStr(onemonthbefore) + "' "

	'response.write sqlStr & "<br>"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

	csrecommendBoardcnt = 0

	if not rsget.EOF  then
		resultCount = rsget.RecordCount
		csrecommendBoardcnt = rsget("cnt")
	end if
	rsget.Close

	application("csrecommendBoardcnt") = csrecommendBoardcnt

	'//직원문의사항  카운트		'/2016.04.15 한용민 추가
	sqlStr = " select count(*) as cnt "
	sqlStr = sqlStr + "	from [db_cs].[dbo].tbl_myqna "
	sqlStr = sqlStr + "	where isusing = 'Y' "
	sqlStr = sqlStr + "	and replydate is NULL "
	sqlStr = sqlStr + "	and userlevel in (7)"	'직원제외
	sqlStr = sqlStr + "	and regdate > '" + CStr(onemonthbefore) + "' "

	'response.write sqlStr & "<br>"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

	csstaffBoardcnt = 0

	if not rsget.EOF  then
		resultCount = rsget.RecordCount
		csstaffBoardcnt = rsget("cnt")
	end if
	rsget.Close

	application("csstaffBoardcnt") = csstaffBoardcnt

	'// 제휴몰별 답변이전 건수
	sqlStr = " select sitename, count(*) as cnt "
	sqlStr = sqlStr + "	from [db_cs].[dbo].tbl_myqna "
	sqlStr = sqlStr + "	where 1 = 1 "
	sqlStr = sqlStr + "	and replydate is NULL "
	sqlStr = sqlStr + "	and isusing = 'Y' "
	sqlStr = sqlStr + "	and dispyn = 'Y' "
	sqlStr = sqlStr + "	and sitename <> '10x10' "
	sqlStr = sqlStr + "	and DateDiff(day, regdate, getdate()) < 14 "
	sqlStr = sqlStr + "	group by sitename "
	sqlStr = sqlStr + "	order by sitename "
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

	csExtSiteCntArr = ""

	if not rsget.EOF  then
		do until rsget.eof
			csExtSiteCntArr = csExtSiteCntArr & rsget("sitename") & "," & rsget("cnt") & "|"
			rsget.MoveNext
		loop
	end if
	rsget.Close

	application("csExtSiteCntArr") = csExtSiteCntArr

	'### STAFF 품절취소요청건
	csStaffStockoutOrderCount = 0
	csStaffStockoutOrderArr = ""

	sqlStr = "SELECT "
	sqlStr = sqlStr & "	m.userid, m.orderserial, isNull(u.userid,'') as csname "
	sqlStr = sqlStr & " FROM [db_order].[dbo].tbl_order_master m with (nolock)"
	sqlStr = sqlStr & " JOIN [db_order].[dbo].tbl_order_detail d with (nolock) on m.orderserial=d.orderserial "
	sqlStr = sqlStr & " JOIN [db_temp].dbo.tbl_mibeasong_list T with (nolock) on 1 = 1 and d.orderserial=T.orderserial and d.idx=T.detailidx "
	sqlStr = sqlStr & " LEFT JOIN db_cs.dbo.tbl_cs_michulgo_stockout_order u with (nolock) on 1 = 1 and m.orderserial = u.orderserial and u.deleteyn = 'N' "
	sqlStr = sqlStr & " WHERE m.regdate >= DateAdd(m, -2, getdate()) and m.ipkumdiv < '8' and m.ipkumdiv > '3' and m.cancelyn = 'N' and m.jumundiv <> '9' "
	'sqlStr = sqlStr & " WHERE m.regdate >= DateAdd(m, -20, getdate()) and m.jumundiv <> '9'  "
	sqlStr = sqlStr & "	and d.itemid<>0 and d.cancelyn <> 'Y' and IsNull(T.code,'00') in ('05','06') and IsNull(T.state, '0')='0' "
	sqlStr = sqlStr & "	and m.userlevel = '7' "
	sqlStr = sqlStr & " GROUP BY m.userid, m.orderserial, u.userid "

	'response.write sqlStr & "<br>"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

	if not rsget.EOF then
		resultCount = rsget.RecordCount

		for i = 0 to resultCount - 1
			If (csStaffStockoutOrderArr = "") Then
				csStaffStockoutOrderArr = rsget("userid") & "," & rsget("orderserial") & "," & rsget("csname")
			Else
				csStaffStockoutOrderArr = csStaffStockoutOrderArr & "|" & rsget("userid") & "," & rsget("orderserial") & "," & rsget("csname")
			End If
			rsget.movenext
		Next
		csStaffStockoutOrderCount = i
	end if
	rsget.Close
	application("csStaffStockoutOrderCount") = csStaffStockoutOrderCount
	application("csStaffStockoutOrderArr") = csStaffStockoutOrderArr

	Call checkAndWriteElapsedTime("004")
end if

csVVipBoardChargecntArr = Split(application("csVVipBoardChargecntArr"), "|")
csVVipBoardNochargecnt = application("csVVipBoardNochargecnt")
csVipBoardChargecntArr = Split(application("csVipBoardChargecntArr"), "|")
csVipBoardNochargecnt = application("csVipBoardNochargecnt")
csBoardChargecntArr = Split(application("csBoardChargecntArr"), "|")
csBoardNochargecnt = application("csBoardNochargecnt")

csrecommendBoardcnt = application("csrecommendBoardcnt")
csstaffBoardcnt = application("csstaffBoardcnt")

csStaffStockoutOrderArr = Split(application("csStaffStockoutOrderArr"), "|")
csStaffStockoutOrderCount = application("csStaffStockoutOrderCount")

csExtSiteCntArr = Split(application("csExtSiteCntArr"), "|")

'==============================================================================
' 미출고리스트[업체배송] - 통계
Dim strMiSend, arrMiSend

if (IsUpdateMichulgoListUpcheNeed = True) then

	strMisend = "0|0|0|0|0|0|0|0|0|0"

	strSql = " exec db_temp.dbo.usp_TEN_GetMichulgoList_CS_SUM_Cnt "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly
	if Not rsget.Eof then
		strMisend = rsget("upcheNocheckCnt") & "|" & rsget("upcheNoSendCnt") & "|" & rsget("upcheStockOutCnt")
		strMisend = strMisend & "|" & rsget("upcheNoSendNoReasonCnt")
		strMisend = strMisend & "|" & rsget("upcheNoSendOverDayCnt")
		strMisend = strMisend & "|" & rsget("upcheNoSendStockOutCnt")
		strMisend = strMisend & "|" & rsget("upcheNoSendHandCnt")
		strMisend = strMisend & "|" & rsget("upcheNoSendFurniCnt")
		strMisend = strMisend & "|" & rsget("upcheNoSendPrepareStockCnt")
		strMisend = strMisend & "|" & rsget("upcheNoSendEtcCnt")
	end if
	rsget.close

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
'반품미처리[업체배송,입점몰제외] - 근무일수 기준 D+4 일
dim upcheReturnBaseDate

tmpSql = " exec [db_cs].[dbo].[usp_getDayMinusWorkday] '" & Left(now(), 10) & "', 4 " & VbCRLF
rsget.CursorLocation = adUseClient
rsget.Open tmpSql, dbget, adOpenForwardOnly
if Not rsget.Eof then
    '// 근무일수 기준 D+4 일
    upcheReturnBaseDate = rsget("minusworkday")
end if
rsget.close

'==============================================================================
'// 상품문의 답변이전
'// -- 텐배 : 1개월 이내 전체
'// -- 업배 : D+3 일 초과, 1개월 이내 전체

dim csTotalCountMiAnswer, csTotalCountMiAnswerTen, csTotalCountMiAnswerUpche
dim csMiAnswerListByBrandTen, csMiAnswerListByBrandUpche
dim tmpVal, tmpVal2

if (IsUpdateMiAnswerUpcheNeed = True) then
	tmpSql = " exec [db_cs].[dbo].[usp_Ten_Cs_ItemQNA_Count] '" + CStr(CSMainIsDev) + "' " & VbCRLF
	rsget.CursorLocation = adUseClient
	rsget.Open tmpSql, dbget, adOpenForwardOnly
	if Not rsget.Eof then
		csTotalCountMiAnswer		= rsget("totcnt")
		csTotalCountMiAnswerTen		= rsget("tencnt")
		csTotalCountMiAnswerUpche	= rsget("upchecnt")
	end if
	rsget.close

	application("csTotalCountMiAnswer") = csTotalCountMiAnswer & "|" & csTotalCountMiAnswerTen & "|" & csTotalCountMiAnswerUpche
end if

csTotalCountMiAnswer		= 0
csTotalCountMiAnswerTen		= 0
csTotalCountMiAnswerUpche	= 0

if (application("csTotalCountMiAnswer") <> "") then
	tmpVal = Split(application("csTotalCountMiAnswer"), "|")

	if (UBound(tmpVal) = 2) and tmpVal(0) <> "" then
		csTotalCountMiAnswer		= tmpVal(0)
		csTotalCountMiAnswerTen		= tmpVal(1)
		csTotalCountMiAnswerUpche	= tmpVal(2)
	end if
end if

'==============================================================================
'// 상품문의 답변이전
'// -- 텐배 : 1개월 이내 전체
'// -- 업배 : D+3 일 초과, 1개월 이내 전체

if (IsUpdateMiAnswerUpcheNeed = True) then
	'// 상품문의 답변이전[텐배]
	tmpSql = " exec [db_cs].[dbo].[usp_Ten_Cs_ItemQNA_ListByBrand] '" + CStr(CSMainIsDev) + "', 'T' " & VbCRLF
	rsget.CursorLocation = adUseClient
	rsget.Open tmpSql, dbget, adOpenForwardOnly
	if Not rsget.Eof then
		resultCount = rsget.RecordCount

		for i = 0 to resultCount - 1
			csMiAnswerListByBrandTen = csMiAnswerListByBrandTen & "|" & rsget("makerid") & "," & rsget("totcnt")
			rsget.movenext
		next
	end if
	rsget.close

	application("csMiAnswerListByBrandTen") = csMiAnswerListByBrandTen

	'// 상품문의 답변이전[업배]
	tmpSql = " exec [db_cs].[dbo].[usp_Ten_Cs_ItemQNA_ListByBrand] '" + CStr(CSMainIsDev) + "', 'U' " & VbCRLF
	rsget.CursorLocation = adUseClient
	rsget.Open tmpSql, dbget, adOpenForwardOnly
	if Not rsget.Eof then
		resultCount = rsget.RecordCount

		for i = 0 to resultCount - 1
			csMiAnswerListByBrandUpche = csMiAnswerListByBrandUpche & "|" & rsget("makerid") & "," & rsget("totcnt")
			rsget.movenext
		next
	end if
	rsget.close

	application("csMiAnswerListByBrandUpche") = csMiAnswerListByBrandUpche
end if

csMiAnswerListByBrandTen = Split(application("csMiAnswerListByBrandTen"), "|")
csMiAnswerListByBrandUpche = Split(application("csMiAnswerListByBrandUpche"), "|")

'==============================================================================
'미출고리스트[업체배송]
dim tmpcsMichulgoList, tmpcsMichulgoListByBrand
dim csMichulgoList, csMichulgoListByBrand
dim csMichulgoItem, csMichulgoItemByBrand
dim nochargeMichulgoBrandcnt

if (IsUpdateMichulgoListUpcheNeed = True) then
	tmpcsMichulgoList = ""
	tmpcsMichulgoListByBrand = ""

	'// 담당자별 서머리
	sqlStr = " exec [db_cs].[dbo].[usp_Ten_Cs_Upche_MichulgoList] '" + CStr(michulgoBaseDate) + "' "
	''response.write sqlStr
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly
	if Not rsget.Eof then
		resultCount = rsget.RecordCount

		for i = 0 to resultCount - 1
			tmpcsMichulgoList = tmpcsMichulgoList + "|" + CStr(rsget("userid")) + "," + CStr(rsget("brandCnt")) + "," + CStr(rsget("orderCnt")) + ""
			rsget.movenext
		next
	end if
	rsget.close

	'// 브랜드별 서머리
	sqlStr = " exec [db_cs].[dbo].[usp_Ten_Cs_Upche_MichulgoListByBrand] '" + CStr(michulgoBaseDate) + "' "
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly
	if Not rsget.Eof then
		resultCount = rsget.RecordCount

		for i = 0 to resultCount - 1
			if (rsget("makerid") <> "10x10Jinair") then
				tmpcsMichulgoListByBrand = tmpcsMichulgoListByBrand + "|" + CStr(rsget("userid")) + "," + CStr(rsget("makerid")) + "," + CStr(rsget("orderCnt")) + ""
			end if

			rsget.movenext
		next
	end if
	rsget.close

	application("csMichulgoList") = tmpcsMichulgoList
	application("csMichulgoListByBrand") = tmpcsMichulgoListByBrand

 	Call checkAndWriteElapsedTime("005")
end if

csMichulgoList = Split(application("csMichulgoList"), "|")
csMichulgoListByBrand = Split(application("csMichulgoListByBrand"), "|")

'// 미분배 브랜드
nochargeMichulgoBrandcnt = 0
for i = 0 to UBound(csMichulgoList)
	if (Trim(csMichulgoList(i)) <> "") then
		csMichulgoItem = Split(csMichulgoList(i), ",")
		if (csMichulgoItem(0) = "") then
			nochargeMichulgoBrandcnt = csMichulgoItem(1)
		end if
	end if
next

'==============================================================================
'반품 미처리[업체배송,입점몰제외]
dim tmpcsUpcheReturnList, tmpcsUpcheReturnListByBrand
dim csUpcheReturnList, csUpcheReturnListByBrand
dim csUpcheReturnItem, csUpcheReturnItemByBrand
dim nochargeUpcheReturnBrandcnt

if (IsUpdateUpcheReturnListNeed = True) then

	tmpcsUpcheReturnList = ""
	tmpcsUpcheReturnListByBrand = ""

	'// 담당자별 서머리
	sqlStr = " exec [db_cs].[dbo].[usp_Ten_Cs_Upche_ReturnList] '" + CStr(upcheReturnBaseDate) + "' "
	''response.write sqlStr
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly
	if Not rsget.Eof then
		resultCount = rsget.RecordCount

		for i = 0 to resultCount - 1
			tmpcsUpcheReturnList = tmpcsUpcheReturnList + "|" + CStr(rsget("userid")) + "," + CStr(rsget("brandCnt")) + "," + CStr(rsget("orderCnt")) + ""
			rsget.movenext
		next
	end if
	rsget.close

	'// 브랜드별 서머리
	sqlStr = " exec [db_cs].[dbo].[usp_Ten_Cs_Upche_ReturnListByBrand] '" + CStr(upcheReturnBaseDate) + "' "
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly
	if Not rsget.Eof then
		resultCount = rsget.RecordCount

		for i = 0 to resultCount - 1
			tmpcsUpcheReturnListByBrand = tmpcsUpcheReturnListByBrand + "|" + CStr(rsget("userid")) + "," + CStr(rsget("makerid")) + "," + CStr(rsget("orderCnt")) + ""
			rsget.movenext
		next
	end if
	rsget.close

	application("csUpcheReturnList") = tmpcsUpcheReturnList
	application("csUpcheReturnListByBrand") = tmpcsUpcheReturnListByBrand

 	Call checkAndWriteElapsedTime("005")
end if

csUpcheReturnList = Split(application("csUpcheReturnList"), "|")
csUpcheReturnListByBrand = Split(application("csUpcheReturnListByBrand"), "|")

'// 미분배 브랜드
nochargeUpcheReturnBrandcnt = 0
for i = 0 to UBound(csUpcheReturnList)
	if (Trim(csUpcheReturnList(i)) <> "") then
		csUpcheReturnItem = Split(csUpcheReturnList(i), ",")
		if (csUpcheReturnItem(0) = "") then
			nochargeUpcheReturnBrandcnt = csUpcheReturnItem(1)
		end if
	end if
next

'==============================================================================
'미출고리스트[입점몰,업배] - 브랜드별 서머리
dim csMichulgoOtherSiteBrandSiteNameArr, csMichulgoOtherSiteBrandNameArr, csMichulgoOtherSiteBrandcntArr
dim prevSiteName, totBrandCount, totItemCount

if (IsUpdateIpjumMichulgoNeed = True) then

	csMichulgoOtherSiteBrandSiteNameArr = ""
	csMichulgoOtherSiteBrandNameArr = ""
	csMichulgoOtherSiteBrandcntArr = ""

	resultCount = 0

	sqlStr = " SELECT "
	sqlStr = sqlStr + " 	m.sitename, d.makerid, count(d.itemid) as cnt "
	sqlStr = sqlStr + " FROM "
	sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_master m with (nolock)"
	sqlStr = sqlStr + " 	JOIN [db_order].[dbo].tbl_order_detail d with (nolock)"
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		m.orderserial=d.orderserial "
	sqlStr = sqlStr + " 	LEFT JOIN [db_temp].dbo.tbl_mibeasong_list T with (nolock)"
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
	sqlStr = sqlStr + " 	and d.isupchebeasong='Y' "
	sqlStr = sqlStr + " 	and m.sitename <> '10x10' "																		'// 입점몰주문
	sqlStr = sqlStr + " 	and d.cancelyn <> 'Y' "
	''sqlStr = sqlStr + " 	and datediff(d,m.baljudate,getdate())>=4 "
	sqlStr = sqlStr + " 	and datediff(d,m.baljudate,'" + CStr(michulgoBaseDate) + "')>=0 "								'// 근무일수 기준 D+3 일
	sqlStr = sqlStr + " 	and not ((T.ipgodate is not null) and (datediff(d, T.ipgodate, getdate()) <= 0)) "				'// 출고예정일 이전 주문 제외
	sqlStr = sqlStr + " 	and IsNULL(T.code,'00') not in ('05','06')"																	'// 품절출고불가/택배파업 제외
	sqlStr = sqlStr + " GROUP BY "
	sqlStr = sqlStr + " 	m.sitename, d.makerid "
	sqlStr = sqlStr + " ORDER BY "
	sqlStr = sqlStr + " 	m.sitename, count(d.itemid) desc, d.makerid "

	'response.write sqlStr & "<Br>"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

	if not rsget.EOF  then
		resultCount = rsget.RecordCount

		for i = 0 to resultCount - 1
			if (csMichulgoOtherSiteBrandSiteNameArr = "") then
				csMichulgoOtherSiteBrandSiteNameArr = rsget("sitename")
			else
				csMichulgoOtherSiteBrandSiteNameArr = csMichulgoOtherSiteBrandSiteNameArr & "|" & rsget("sitename")
			end if

			if (csMichulgoOtherSiteBrandNameArr = "") then
				csMichulgoOtherSiteBrandNameArr = rsget("makerid")
			else
				csMichulgoOtherSiteBrandNameArr = csMichulgoOtherSiteBrandNameArr & "|" & rsget("makerid")
			end if

			if (csMichulgoOtherSiteBrandcntArr = "") then
				csMichulgoOtherSiteBrandcntArr = rsget("cnt")
			else
				csMichulgoOtherSiteBrandcntArr = csMichulgoOtherSiteBrandcntArr & "|" & rsget("cnt")
			end if

			rsget.movenext
		next
	end if
	rsget.close

	application("csMichulgoOtherSiteBrandSiteNameArr") = csMichulgoOtherSiteBrandSiteNameArr
	application("csMichulgoOtherSiteBrandNameArr") = csMichulgoOtherSiteBrandNameArr
	application("csMichulgoOtherSiteBrandcntArr") = csMichulgoOtherSiteBrandcntArr

	Call checkAndWriteElapsedTime("005")
end if

csMichulgoOtherSiteBrandSiteNameArr = Split(application("csMichulgoOtherSiteBrandSiteNameArr"), "|")
csMichulgoOtherSiteBrandNameArr = Split(application("csMichulgoOtherSiteBrandNameArr"), "|")
csMichulgoOtherSiteBrandcntArr = Split(application("csMichulgoOtherSiteBrandcntArr"), "|")

'==============================================================================
'미출고리스트[입점몰,업배] - 입점몰 리스트 : 주문별 서머리
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
	sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_master m with (nolock)"
	sqlStr = sqlStr + " 	JOIN [db_order].[dbo].tbl_order_detail d with (nolock)"
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		m.orderserial=d.orderserial "
	sqlStr = sqlStr + " 	LEFT JOIN [db_temp].dbo.tbl_mibeasong_list T with (nolock)"
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
	sqlStr = sqlStr + " 	and d.isupchebeasong='Y' "
	sqlStr = sqlStr + " 	and d.cancelyn <> 'Y' "
	''sqlStr = sqlStr + " 	and datediff(d,m.baljudate,getdate())>=4 "
	sqlStr = sqlStr + " 	and datediff(d,m.baljudate,'" + CStr(michulgoBaseDate) + "')>=0 "								'// 근무일수 기준 D+3 일
	sqlStr = sqlStr + " 	and not ((T.ipgodate is not null) and (datediff(d, T.ipgodate, getdate()) <= 0)) "				'// 출고예정일 이전 주문 제외
	sqlStr = sqlStr + " 	and IsNULL(T.code,'00') not in ('05','06') "																	'// 품절출고불가/택배파업 제외
	sqlStr = sqlStr + " GROUP BY "
	sqlStr = sqlStr + " 	m.sitename, m.orderserial "
	sqlStr = sqlStr + " ORDER BY "
	sqlStr = sqlStr + " 	m.sitename, m.orderserial "

	'response.write sqlStr & "<br>"
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

	strSql = " exec db_temp.dbo.usp_TEN_GetMichulgoList_CS_SUM_Cnt "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly
	if Not rsget.Eof then
		strTenMiSend = rsget("tenNoReasonCnt") & "|" & rsget("tenNoCallCnt") & "|" & "0" & "|" & "0"
	end if
	rsget.close

	application("csTenMiSend") = strTenMiSend

	Call checkAndWriteElapsedTime("006")
end if

arrTenMiSend = Split(application("csTenMiSend"), "|")

'==============================================================================
'CS처리리스트 관리
''환불 미처리(A003), 마일리지 환불 미처리, 카드취소미처리(A007), 주문취소미처리(A008),
''출고시유의사항, 업체미처리, 업체처리완료, 회수요청미처리, 확인요청, 외부몰환불미처리
dim csRefundRequestRegCount, csRefundRequestConfirmCount, csRequestCardCancelCount, csReturnNotFinish
dim csNotFinA008, csNotFinA005, csUpcheNotFin, csCustomerAddPayment

if (IsUpdateCSListNeed = True) or (IsUpdateUpcheReturnListNeed = True) then

	csRefundRequestRegCount			= 0
	csRefundRequestConfirmCount		= 0
	csRequestCardCancelCount		= 0
	csReturnNotFinish				= 0
	csNotFinA008					= 0
	csNotFinA005					= 0
	csUpcheNotFin					= 0
	csCustomerAddPayment			= 0

	sqlStr = " update a "
	sqlStr = sqlStr + " set a.currstate = 'B006' "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	[db_cs].[dbo].[tbl_new_as_list] a "
	sqlStr = sqlStr + " 	join [db_cs].[dbo].[tbl_as_customer_addbeasongpay_info] i on a.id = i.asid "
	sqlStr = sqlStr + " 	join [db_order].[dbo].[tbl_order_master] m on i.payorderserial = m.orderserial "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and a.divcd = 'A999' "
	sqlStr = sqlStr + " 	and a.deleteyn = 'N' "
	sqlStr = sqlStr + " 	and a.currstate = 'B001' "
	sqlStr = sqlStr + " 	and m.ipkumdiv >= '4' "
	sqlStr = sqlStr + " 	and m.cancelyn = 'N' "
	dbget.Execute sqlStr

	sqlStr = " select "
	sqlStr = sqlStr + "     sum(case when (divcd = 'A003') and (currstate='B001') and (requireupche = 'N') then 1 else 0 end) as csRefundRequestRegCount, "
	sqlStr = sqlStr + "     sum(case when (divcd = 'A003') and (currstate='B005') and (requireupche = 'N') then 1 else 0 end) as csRefundRequestConfirmCount, "
	sqlStr = sqlStr + "     sum(case when (divcd = 'A007') and regdate < convert(varchar(10), getdate(), 121) then 1 else 0 end) as csRequestCardCancelCount, "
	sqlStr = sqlStr + "     sum(case when (divcd = 'A005') then 1 else 0 end) as csNotFinA005, "
	sqlStr = sqlStr + "     sum(case when (divcd = 'A008') and (regdate>'2008-04-23') then 1 else 0 end) as csNotFinA008, "
	sqlStr = sqlStr + "     sum(case when (requireupche='Y') and (currstate<'B006') then 1 else 0 end) as csUpcheNotFin, "
	sqlStr = sqlStr + "     sum(case when (divcd in ('A010', 'A011', 'A111')) and (currstate < 'B002') then 1 else 0 end) as csReturnNotFinish, "
	sqlStr = sqlStr + "     sum(case when (divcd = 'A999') and (currstate='B006') then 1 else 0 end) as csCustomerAddPayment "
	sqlStr = sqlStr + " from [db_cs].[dbo].tbl_new_as_list "
	sqlStr = sqlStr + " where 1=1 "
	sqlStr = sqlStr + " and deleteyn = 'N' "
	sqlStr = sqlStr + " and currstate < 'B007'"
	sqlStr = sqlStr + " and regdate>'"&treemonthbefore&"'"        				''한달지난 접수건은 잡에서 삭제X =>반품접수건만.
	sqlStr = sqlStr + " and id > " & (maxCSMasterIdx - 200000) & " "			'// 최근 20만개만 검색한다.(속도문제)

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
		csCustomerAddPayment  		= rsget("csCustomerAddPayment")
	end if
	rsget.close

	application("csRefundRequestRegCount") 		= csRefundRequestRegCount
	application("csRefundRequestConfirmCount") 	= csRefundRequestConfirmCount
	application("csRequestCardCancelCount") 	= csRequestCardCancelCount
	application("csNotFinA008") 				= csNotFinA008

	application("csUpcheNotFin") 				= csUpcheNotFin
	application("csNotFinA005") 				= csNotFinA005
	application("csReturnNotFinish") 			= csReturnNotFinish
	application("csCustomerAddPayment") 		= csCustomerAddPayment

	Call checkAndWriteElapsedTime("007")
end if

csRefundRequestRegCount = application("csRefundRequestRegCount")
csRefundRequestConfirmCount = application("csRefundRequestConfirmCount")
csRequestCardCancelCount = application("csRequestCardCancelCount")
csNotFinA008 = application("csNotFinA008")

csUpcheNotFin = application("csUpcheNotFin")
csNotFinA005 = application("csNotFinA005")
csReturnNotFinish = application("csReturnNotFinish")
csCustomerAddPayment = application("csCustomerAddPayment")

'' '==============================================================================
dim upReturnMiFinishBaseDate7

tmpSql = " exec [db_cs].[dbo].[usp_getDayMinusWorkday] '" & Left(now(), 10) & "', 7 " & VbCRLF
rsget.CursorLocation = adUseClient
rsget.Open tmpSql, dbget, adOpenForwardOnly
if Not rsget.Eof then
    '// 근무일수 기준 D+7 일
    upReturnMiFinishBaseDate7 = rsget("minusworkday")
end if
rsget.close

'' tmpSql = " exec [db_cs].[dbo].[usp_getDayMinusWorkday] '" & Left(now(), 10) & "', 3 " & VbCRLF
'' rsget.CursorLocation = adUseClient
'' rsget.Open tmpSql, dbget, adOpenForwardOnly
'' if Not rsget.Eof then
''     '// 근무일수 기준 D+3 일
''     upReturnMiFinishBaseDate3 = rsget("minusworkday")
'' end if
'' rsget.close

'==============================================================================
'반품접수(업체배송) D+4 미처리
dim upcheReturnNotFinish

if (IsUpdateCSListNeed = True) or (IsUpdateUpcheReturnListNeed = True) then

	sqlStr = " select "
	sqlStr = sqlStr + " IsNull(sum(case when datediff(d, m.regdate, '" + CStr(upcheReturnBaseDate) + "') >= 0 then 1 else 0 end), 0) as cnt "
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
	sqlStr = sqlStr + " 	and datediff(d, m.regdate, '" + CStr(upcheReturnBaseDate) + "') >= 0 "
	sqlStr = sqlStr + " 	and not ((T.ipgodate is not null) and (datediff(d, T.ipgodate, getdate()) <= 0)) "
	sqlStr = sqlStr + " 	and (datediff(m, m.regdate, getdate()) < 3) "
	sqlStr = sqlStr + " 	and m.id > " & minCSMasterIdxThreeMonth & " "

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

	if Not rsget.Eof then
	    application("upcheReturnNotFinish") = rsget("cnt")
	end if
	rsget.close

	Call checkAndWriteElapsedTime("008")
end if

upcheReturnNotFinish = application("upcheReturnNotFinish")

'==============================================================================
'반품접수(업체배송, 입점몰주문) D+7 미처리(브랜드별 서머리)
dim csReturn7OtherSiteSiteNameArr, csReturn7OtherSiteBrandNameArr, csReturn7OtherSiteBrandcntArr

if (IsUpdateCSListNeed = True) or (IsUpdateUpcheReturnListNeed = True) then

	csReturn7OtherSiteSiteNameArr = ""
	csReturn7OtherSiteBrandNameArr = ""
	csReturn7OtherSiteBrandcntArr = ""

	resultCount = 0

	sqlStr = " SELECT "
	sqlStr = sqlStr + " 	m.extsitename, d.makerid, count(d.itemid) as cnt "
	sqlStr = sqlStr + " FROM "
	sqlStr = sqlStr + " 	db_cs.dbo.tbl_new_as_list m "
	sqlStr = sqlStr + " 	join db_cs.dbo.tbl_new_as_detail d "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		m.id = d.masterid "
	sqlStr = sqlStr + " 	left join [db_temp].dbo.tbl_csmifinish_list T "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		d.id=T.csdetailidx "
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
	sqlStr = sqlStr + " 	and m.extsitename <> '10x10' "																		'// 입점몰주문
	sqlStr = sqlStr + " 	and m.divcd = 'A004' "
	sqlStr = sqlStr + " 	and datediff(d, m.regdate, '" + CStr(upReturnMiFinishBaseDate7) + "') >= 0 "
	sqlStr = sqlStr + " 	and not ((T.ipgodate is not null) and (datediff(d, T.ipgodate, getdate()) <= 0)) "
	sqlStr = sqlStr + " 	and (datediff(m, m.regdate, getdate()) < 3) "
	sqlStr = sqlStr + " 	and m.id > " & minCSMasterIdxThreeMonth & " "
	sqlStr = sqlStr + " GROUP BY "
	sqlStr = sqlStr + " 	m.extsitename, d.makerid "
	sqlStr = sqlStr + " ORDER BY "
	sqlStr = sqlStr + " 	m.extsitename, count(d.itemid) desc, d.makerid "
	''rw sqlStr

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

	if not rsget.EOF  then

		resultCount = rsget.RecordCount

		for i = 0 to resultCount - 1
			if (csReturn7OtherSiteSiteNameArr = "") then
				csReturn7OtherSiteSiteNameArr = rsget("extsitename")
			else
				csReturn7OtherSiteSiteNameArr = csReturn7OtherSiteSiteNameArr & "|" & rsget("extsitename")
			end if

			if (csReturn7OtherSiteBrandNameArr = "") then
				csReturn7OtherSiteBrandNameArr = rsget("makerid")
			else
				csReturn7OtherSiteBrandNameArr = csReturn7OtherSiteBrandNameArr & "|" & rsget("makerid")
			end if

			if (csReturn7OtherSiteBrandcntArr = "") then
				csReturn7OtherSiteBrandcntArr = rsget("cnt")
			else
				csReturn7OtherSiteBrandcntArr = csReturn7OtherSiteBrandcntArr & "|" & rsget("cnt")
			end if

			rsget.movenext
		next
	end if
	rsget.close

	application("csReturn7OtherSiteSiteNameArr") = csReturn7OtherSiteSiteNameArr
	application("csReturn7OtherSiteBrandNameArr") = csReturn7OtherSiteBrandNameArr
	application("csReturn7OtherSiteBrandcntArr") = csReturn7OtherSiteBrandcntArr

	Call checkAndWriteElapsedTime("010")
end if

csReturn7OtherSiteSiteNameArr = Split(application("csReturn7OtherSiteSiteNameArr"), "|")
csReturn7OtherSiteBrandNameArr = Split(application("csReturn7OtherSiteBrandNameArr"), "|")
csReturn7OtherSiteBrandcntArr = Split(application("csReturn7OtherSiteBrandcntArr"), "|")

'==============================================================================
''업체처리완료건, 물류처리완료건
dim csUpcheFinished, csLogicsFinished
csUpcheFinished = 0
csLogicsFinished = 0

if (IsUpdateCSListNeed = True) or (IsUpdateUpcheReturnListNeed = True) then
	sqlStr = " select "
	sqlStr = sqlStr + " sum(case when A.requireupche='Y' then 1 else 0 end) as upchecnt "
	sqlStr = sqlStr + " , sum(case when A.requireupche<>'Y' then 1 else 0 end) as logicscnt "
	sqlStr = sqlStr + " from  [db_cs].[dbo].tbl_new_as_list A"
	sqlStr = sqlStr + " left join [db_cs].[dbo].tbl_new_as_list B "
	sqlStr = sqlStr + " on A.id = B.refasid "
	sqlStr = sqlStr + " where (A.currstate='B006') and A.deleteyn = 'N' and (A.divcd not in ('A000', 'A100') or IsNull(B.currstate, 'B007') = 'B007') "
	sqlStr = sqlStr + " and (A.requireupche='Y' or A.divcd in ('A010','A200','A011','A111','A004','A012','A112')) "

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
	if  not rsget.EOF  then
	    csUpcheFinished = rsget("upchecnt")
		csLogicsFinished = rsget("logicscnt")
	end if
	rsget.close

	application("csUpcheFinished") = csUpcheFinished
	application("csLogicsFinished") = csLogicsFinished

	Call checkAndWriteElapsedTime("011")
end if

csUpcheFinished = application("csUpcheFinished")
csLogicsFinished = application("csLogicsFinished")

'==============================================================================
''마일리지/예치금 환불미처리
dim csNotFinMileRefund
if (IsUpdateCSListNeed = True) or (IsUpdateUpcheReturnListNeed = True) then

	sqlStr = " select count(A.id) as csNotFinMileRefund from [db_cs].[dbo].tbl_new_as_list A "
	sqlStr = sqlStr + "     Left Join [db_cs].[dbo].tbl_as_refund_info r on A.id=r.asid "
	sqlStr = sqlStr + " where 1 = 1 "
	sqlStr = sqlStr + " and A.currstate<'B007' "
	sqlStr = sqlStr + " and A.divcd='A003' "
	sqlStr = sqlStr + " and A.deleteyn='N' "
	sqlStr = sqlStr + " and R.returnmethod in ('R900', 'R910')"
	sqlStr = sqlStr + " and A.regdate>'"&treemonthbefore&"'"
	''response.write sqlStr

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
	sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_master m with (nolock)"
	sqlStr = sqlStr + " 	JOIN [db_order].[dbo].tbl_order_detail d with (nolock)"
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		m.orderserial=d.orderserial "
	sqlStr = sqlStr + " 	JOIN [db_temp].dbo.tbl_mibeasong_list T with (nolock)"
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and d.orderserial=T.orderserial "
	sqlStr = sqlStr + " 		and d.idx=T.detailidx "
	sqlStr = sqlStr + " WHERE "
	''sqlStr = sqlStr + " 	m.regdate >= DateAdd(m, -2, getdate()) "
    sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and m.ipkumdiv < '8' "
	sqlStr = sqlStr + " 	and m.ipkumdiv > '3' "
	sqlStr = sqlStr + " 	and m.cancelyn = 'N' "
	sqlStr = sqlStr + " 	and m.jumundiv <> '9' "
	sqlStr = sqlStr + " 	and d.itemid<>0 "
	''sqlStr = sqlStr + " 	and d.currstate='3' "
	sqlStr = sqlStr + " 	and m.sitename <> '10x10' "																'// 입점몰주문
	''sqlStr = sqlStr + " 	and d.isupchebeasong='Y' "																'// 텐배+업배
	sqlStr = sqlStr + " 	and d.cancelyn <> 'Y' "
	sqlStr = sqlStr + " 	and IsNull(T.code,'00') in ('05','06') "															'// 품절출고불가/택배파업
	sqlStr = sqlStr + " 	and IsNull(T.state, '0') in ('0', '4') "												'// 고객안내 이전+고객안내
	sqlStr = sqlStr + " GROUP BY "
	sqlStr = sqlStr + " 	m.sitename "
	sqlStr = sqlStr + " ORDER BY "
	sqlStr = sqlStr + " 	m.sitename "

	'response.write sqlStr & "<Br>"
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
	sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_master m with (nolock)"
	sqlStr = sqlStr + " 	JOIN [db_order].[dbo].tbl_order_detail d with (nolock)"
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		m.orderserial=d.orderserial "
	sqlStr = sqlStr + " 	JOIN [db_temp].dbo.tbl_mibeasong_list T with (nolock)"
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and d.orderserial=T.orderserial "
	sqlStr = sqlStr + " 		and d.idx=T.detailidx "
	sqlStr = sqlStr + " WHERE "
	''sqlStr = sqlStr + " 	m.regdate >= DateAdd(m, -2, getdate()) "
    sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and m.ipkumdiv < '8' "
	sqlStr = sqlStr + " 	and m.ipkumdiv > '3' "
	sqlStr = sqlStr + " 	and m.cancelyn = 'N' "
	sqlStr = sqlStr + " 	and m.jumundiv <> '9' "
	sqlStr = sqlStr + " 	and d.itemid<>0 "
	sqlStr = sqlStr + " 	and m.sitename <> '10x10' "
	''sqlStr = sqlStr + " 	and d.isupchebeasong='Y' "
	sqlStr = sqlStr + " 	and d.cancelyn <> 'Y' "
	sqlStr = sqlStr + " 	and IsNull(T.code,'00') in ('05','06') "
	sqlStr = sqlStr + " 	and IsNull(T.state, '0') in ('0', '4') "
	sqlStr = sqlStr + " GROUP BY "
	sqlStr = sqlStr + " 	m.sitename, m.orderserial "
	sqlStr = sqlStr + " ORDER BY "
	sqlStr = sqlStr + " 	m.sitename, m.orderserial "

	'response.write sqlStr & "<Br>"
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
'입점몰 환불미처리 : 입점몰 리스트
dim csRefundOtherSitenameArr, csRefundOtherSiteOrderCountArr, csRefundOtherSiteCSCountArr

if (IsUpdateIpjumRefundNeed = True) then

	csRefundOtherSitenameArr = ""
	csRefundOtherSiteOrderCountArr = ""
	csRefundOtherSiteCSCountArr = ""

	resultCount = 0

	sqlStr = " SELECT m.sitename, count(distinct m.orderserial) as ordercnt, count(a.id) as cnt "
	sqlStr = sqlStr + " FROM "
	sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_master m "
	sqlStr = sqlStr + " 	join db_cs.dbo.tbl_new_as_list a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		m.orderserial = a.orderserial "
	sqlStr = sqlStr + " WHERE "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and a.regdate >= DateAdd(m, -12, getdate()) "
	sqlStr = sqlStr + " 	and m.sitename <> '10x10' "							'// 입점몰주문
	sqlStr = sqlStr + " 	and a.divcd = 'A005' "								'// 입점몰환불
	sqlStr = sqlStr + " 	and a.currstate <> 'B007' "
	sqlStr = sqlStr + " 	and a.deleteyn <> 'Y' "
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
			if (csRefundOtherSitenameArr = "") then
				csRefundOtherSitenameArr = rsget("sitename")
			else
				csRefundOtherSitenameArr = csRefundOtherSitenameArr & "|" & rsget("sitename")
			end if

			if (csRefundOtherSiteOrderCountArr = "") then
				csRefundOtherSiteOrderCountArr = rsget("ordercnt")
			else
				csRefundOtherSiteOrderCountArr = csRefundOtherSiteOrderCountArr & "|" & rsget("ordercnt")
			end if

			if (csRefundOtherSiteCSCountArr = "") then
				csRefundOtherSiteCSCountArr = rsget("cnt")
			else
				csRefundOtherSiteCSCountArr = csRefundOtherSiteCSCountArr & "|" & rsget("cnt")
			end if

			rsget.movenext
		next
	end if
	rsget.close

	application("csRefundOtherSitenameArr") = csRefundOtherSitenameArr
	application("csRefundOtherSiteOrderCountArr") = csRefundOtherSiteOrderCountArr
	application("csRefundOtherSiteCSCountArr") = csRefundOtherSiteCSCountArr

	Call checkAndWriteElapsedTime("013-1")
end if

csRefundOtherSitenameArr = Split(application("csRefundOtherSitenameArr"), "|")
csRefundOtherSiteOrderCountArr = Split(application("csRefundOtherSiteOrderCountArr"), "|")
csRefundOtherSiteCSCountArr = Split(application("csRefundOtherSiteCSCountArr"), "|")

'==============================================================================
'입점몰 환불미처리(주문별 서머리)
dim csRefundOtherSiteOrderSitenameArr, csRefundOtherSiteOrderserialArr, csRefundOtherSiteCSCntcntArr, csRefundOtherSiteCSidx

if (IsUpdateIpjumRefundNeed = True) then

	csRefundOtherSiteOrderSitenameArr = ""
	csRefundOtherSiteOrderserialArr = ""
	csRefundOtherSiteCSCntcntArr = ""
	csRefundOtherSiteCSidx = ""

	resultCount = 0

	sqlStr = " SELECT m.sitename, m.orderserial, max(a.id) as asidx, count(a.id) as cnt "
	sqlStr = sqlStr + " FROM "
	sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_master m "
	sqlStr = sqlStr + " 	join db_cs.dbo.tbl_new_as_list a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		m.orderserial = a.orderserial "
	sqlStr = sqlStr + " WHERE "
	sqlStr = sqlStr + " 	1 = 1 "
	'sqlStr = sqlStr + " 	and a.regdate >= DateAdd(m, -2, getdate()) "	' cs하소라요청(기간상관없이 2차 완료 안된건 보여져야함)
	sqlStr = sqlStr + " 	and m.sitename <> '10x10' "							'// 입점몰주문
	sqlStr = sqlStr + " 	and a.divcd = 'A005' "								'// 입점몰환불
	sqlStr = sqlStr + " 	and a.currstate <> 'B007' "
	sqlStr = sqlStr + " 	and a.deleteyn <> 'Y' "
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
			if (csRefundOtherSiteOrderSitenameArr = "") then
				csRefundOtherSiteOrderSitenameArr = rsget("sitename")
			else
				csRefundOtherSiteOrderSitenameArr = csRefundOtherSiteOrderSitenameArr & "|" & rsget("sitename")
			end if

			if (csRefundOtherSiteOrderserialArr = "") then
				csRefundOtherSiteOrderserialArr = rsget("orderserial")
			else
				csRefundOtherSiteOrderserialArr = csRefundOtherSiteOrderserialArr & "|" & rsget("orderserial")
			end if

			if (csRefundOtherSiteCSCntcntArr = "") then
				csRefundOtherSiteCSCntcntArr = rsget("cnt")
			else
				csRefundOtherSiteCSCntcntArr = csRefundOtherSiteCSCntcntArr & "|" & rsget("cnt")
			end if

			if (csRefundOtherSiteCSidx = "") then
				csRefundOtherSiteCSidx = rsget("asidx")
			else
				csRefundOtherSiteCSidx = csRefundOtherSiteCSidx & "|" & rsget("asidx")
			end if

			rsget.movenext
		next
	end if
	rsget.close

	application("csRefundOtherSiteOrderSitenameArr") 		= csRefundOtherSiteOrderSitenameArr
	application("csRefundOtherSiteOrderserialArr") 		= csRefundOtherSiteOrderserialArr
	application("csRefundOtherSiteCSCntcntArr") 	= csRefundOtherSiteCSCntcntArr
	application("csRefundOtherSiteCSidx") 		= csRefundOtherSiteCSidx

	Call checkAndWriteElapsedTime("014")
end if

csRefundOtherSiteOrderSitenameArr 	= Split(application("csRefundOtherSiteOrderSitenameArr"), "|")
csRefundOtherSiteOrderserialArr 		= Split(application("csRefundOtherSiteOrderserialArr"), "|")
csRefundOtherSiteCSCntcntArr = Split(application("csRefundOtherSiteCSCntcntArr"), "|")
csRefundOtherSiteCSidx 		= Split(application("csRefundOtherSiteCSidx"), "|")

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

	sqlStr = " UPDATE "
	sqlStr = sqlStr + " 	db_cs.dbo.tbl_cs_michulgo_stockout_order "
	sqlStr = sqlStr + " SET "
	sqlStr = sqlStr + " 	deleteyn = 'Y' "
	sqlStr = sqlStr + " WHERE "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and orderserial not in ( "
	sqlStr = sqlStr + " 		SELECT "
	sqlStr = sqlStr + " 			distinct m.orderserial "
	sqlStr = sqlStr + " 		FROM "
	sqlStr = sqlStr + " 			db_cs.dbo.tbl_cs_michulgo_stockout_order s "
	sqlStr = sqlStr + " 			join [db_order].[dbo].tbl_order_master m "
	sqlStr = sqlStr + " 			on "
	sqlStr = sqlStr + " 				s.orderserial = m.orderserial "
	sqlStr = sqlStr + " 			JOIN [db_order].[dbo].tbl_order_detail d "
	sqlStr = sqlStr + " 			on "
	sqlStr = sqlStr + " 				m.orderserial=d.orderserial "
	sqlStr = sqlStr + " 			JOIN [db_temp].dbo.tbl_mibeasong_list T "
	sqlStr = sqlStr + " 			on "
	sqlStr = sqlStr + " 				1 = 1 "
	sqlStr = sqlStr + " 				and d.orderserial=T.orderserial "
	sqlStr = sqlStr + " 				and d.idx=T.detailidx "
	sqlStr = sqlStr + " 		WHERE "
	''sqlStr = sqlStr + " 			m.regdate >= DateAdd(m, -2, getdate()) "
	sqlStr = sqlStr + " 			1 = 1 "
	sqlStr = sqlStr + " 			and m.ipkumdiv < '8' "
	sqlStr = sqlStr + " 			and m.ipkumdiv > '3' "
    sqlStr = sqlStr + " 			and IsNull(d.currstate,'0') <> '7' "
	sqlStr = sqlStr & "				and m.userlevel <> '7' "
	sqlStr = sqlStr + " 			and m.cancelyn = 'N' "
	sqlStr = sqlStr + " 			and m.jumundiv <> '9' "
	sqlStr = sqlStr + " 			and d.itemid<>0 "
	sqlStr = sqlStr + " 			and m.sitename = '10x10' "
	sqlStr = sqlStr + " 			and d.cancelyn <> 'Y' "
	sqlStr = sqlStr + " 			and IsNull(T.code,'00') in ('05','06') "
	sqlStr = sqlStr + " 			and IsNull(T.state, '0')='0' "
	sqlStr = sqlStr + " 	) "
	sqlStr = sqlStr + " 	and deleteyn = 'N' "
	rsget.Open sqlStr, dbget, 1

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
	''sqlStr = sqlStr + " 			m.regdate >= DateAdd(m, -2, getdate()) "
	sqlStr = sqlStr + " 			1 = 1 "
	sqlStr = sqlStr + " 			and m.ipkumdiv < '8' "
	sqlStr = sqlStr + " 			and m.ipkumdiv > '3' "
	sqlStr = sqlStr & "				and m.userlevel <> '7' "
	sqlStr = sqlStr + " 			and m.cancelyn = 'N' "
	sqlStr = sqlStr + " 			and m.jumundiv <> '9' "
	sqlStr = sqlStr + " 			and d.itemid<>0 "
	sqlStr = sqlStr + " 			and IsNull(d.currstate,'0') <> '7' "
	sqlStr = sqlStr + " 			and m.sitename = '10x10' "																'// 입점몰주문 제외
	''sqlStr = sqlStr + " 			and d.isupchebeasong='Y' "																'// 텐배+업배
	sqlStr = sqlStr + " 			and d.cancelyn <> 'Y' "
	sqlStr = sqlStr + " 			and IsNull(T.code,'00') in ('05','06') "															'// 품절출고불가/택배파업
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
dim csStockoutOrderserialUseridArr, csStockoutRegdateArr

if (IsUpdateTenTenStockOutNeed = True) then
	csStockoutOrderUseridArr = ""
	csStockoutOrderserialArr = ""
	csStockoutOrderdetailCntcntArr = ""
	csStockoutHasChargeUser = ""
	csStockoutOrderdetaiidx = ""
	csStockoutOrderserialUseridArr = ""
	csStockoutRegdateArr = ""

	resultCount = 0

	sqlStr = " SELECT "
	sqlStr = sqlStr + " 	IsNull(u.userid, '') as userid, m.orderserial, max(d.idx) as orderdetailidx, count(d.itemid) as cnt, m.userid as orderserialuserid, IsNull(u.regdate, '') as regdate "
	sqlStr = sqlStr + " FROM "
	sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_master m with (nolock)"
	sqlStr = sqlStr + " 	JOIN [db_order].[dbo].tbl_order_detail d with (nolock)"
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		m.orderserial=d.orderserial "
	sqlStr = sqlStr + " 	JOIN [db_temp].dbo.tbl_mibeasong_list T with (nolock)"
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and d.orderserial=T.orderserial "
	sqlStr = sqlStr + " 		and d.idx=T.detailidx "
	sqlStr = sqlStr + " 	LEFT JOIN db_cs.dbo.tbl_cs_michulgo_stockout_order u with (nolock)"
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and m.orderserial = u.orderserial "
	sqlStr = sqlStr + " 		and u.deleteyn = 'N' "
	sqlStr = sqlStr + " WHERE "
	''sqlStr = sqlStr + " 	m.regdate >= DateAdd(m, -2, getdate()) "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and m.ipkumdiv < '8' "
	sqlStr = sqlStr + " 	and m.ipkumdiv > '3' "
	sqlStr = sqlStr + " 	and m.cancelyn = 'N' "
	sqlStr = sqlStr + " 	and m.jumundiv <> '9' "
	sqlStr = sqlStr + " 	and d.itemid<>0 "
	sqlStr = sqlStr + " 	and m.sitename = '10x10' "
	''sqlStr = sqlStr + " 	and d.isupchebeasong='Y' "
	sqlStr = sqlStr + " 	and d.cancelyn <> 'Y' "
	sqlStr = sqlStr + " 	and IsNull(T.code,'00') in ('05','06') "
	sqlStr = sqlStr + " 	and IsNull(T.state, '0')='0' "
	sqlStr = sqlStr + " GROUP BY "
	sqlStr = sqlStr + " 	IsNull(u.userid, ''), m.orderserial, m.userid, IsNull(u.regdate, '') "
	sqlStr = sqlStr + " ORDER BY "
	sqlStr = sqlStr + " 	IsNull(u.userid, ''), IsNull(u.regdate, '') desc, m.orderserial "
	''response.write sqlStr

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

	if not rsget.EOF  then

		resultCount = rsget.RecordCount

		for i = 0 to resultCount - 1
			if (csStockoutOrderserialArr = "") then
				csStockoutOrderUseridArr = rsget("userid")
				csStockoutOrderserialArr = rsget("orderserial")
				csStockoutOrderdetailCntcntArr = rsget("cnt")
				csStockoutHasChargeUser = "N"
				csStockoutOrderdetaiidx = rsget("orderdetailidx")

				csStockoutOrderserialUseridArr = rsget("orderserialuserid")
				csStockoutRegdateArr = rsget("regdate")
			else
				csStockoutOrderUseridArr = csStockoutOrderUseridArr & "|" & rsget("userid")
				csStockoutOrderserialArr = csStockoutOrderserialArr & "|" & rsget("orderserial")
				csStockoutOrderdetailCntcntArr = csStockoutOrderdetailCntcntArr & "|" & rsget("cnt")
				csStockoutHasChargeUser = csStockoutHasChargeUser & "|" & "N"
				csStockoutOrderdetaiidx = csStockoutOrderdetaiidx & "|" & rsget("orderdetailidx")

				csStockoutOrderserialUseridArr = csStockoutOrderserialUseridArr & "|" & rsget("orderserialuserid")
				csStockoutRegdateArr = csStockoutRegdateArr & "|" & rsget("regdate")
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

	application("csStockoutOrderserialUseridArr") 	= csStockoutOrderserialUseridArr
	application("csStockoutRegdateArr") 			= csStockoutRegdateArr

	Call checkAndWriteElapsedTime("016")
end if

csStockoutOrderUseridArr 		= Split(application("csStockoutOrderUseridArr"), "|")
csStockoutOrderserialArr 		= Split(application("csStockoutOrderserialArr"), "|")
csStockoutOrderdetailCntcntArr 	= Split(application("csStockoutOrderdetailCntcntArr"), "|")
csStockoutHasChargeUser 		= Split(application("csStockoutHasChargeUser"), "|")
csStockoutOrderdetaiidx 		= Split(application("csStockoutOrderdetaiidx"), "|")
csStockoutOrderserialUseridArr 	= Split(application("csStockoutOrderserialUseridArr"), "|")
csStockoutRegdateArr 			= Split(application("csStockoutRegdateArr"), "|")

'==============================================================================
' 아이디별 확인사항
Dim csMainUserID, csMainUserName
csMainUserID	= req("csMainUserID", session("ssBctId") )
csMainUserName	= session("ssBctCname")

'// 동명이인
''if (csMainUserID = "raider7942") then
''	csMainUserName = "김은주B"
''end if

'미처리메모
dim CSMemoNotFinish, CSA060NotFinish

sqlStr = " [db_cs].[dbo].[usp_Ten_GetMiFinishMemo] '" + CStr(csMainUserID) + "' "

rsget.CursorLocation = adUseClient
rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
if  not rsget.EOF  then
    CSMemoNotFinish = rsget("CSMemoNotFinish")
end if
rsget.close

sqlStr = " exec [db_cs].[dbo].[usp_Ten_CsAsCountNew] 'A060', 'B001', 'N', '' , '', '', '', '" + CStr(csMainUserID) + "', '', '" & sixmonthbefore & "', '" & tomorrowdate & "', '', '', '', '', 'regdate', '' "

rsget.CursorLocation = adUseClient
rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
if  not rsget.EOF  then
    CSA060NotFinish = rsget("cnt")
end if
rsget.close


'==============================================================================
' 무통장 입금 미확인
Dim csTotalTodayIpkum, csNotFinishTodayIpkum

sqlStr = " select Count(idx) as csTotalTodayIpkum, IsNull(sum(case when i.ipkumstate in ('0', '1') then 1 else 0 end), 0) as csNotFinishTodayIpkum "
sqlStr = sqlStr & " from [db_order].[dbo].tbl_ipkum_list i "
sqlStr = sqlStr & " where i.bankdate >= '" + Left(Now(), 10) + "' "

rsget.CursorLocation = adUseClient
rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
if  not rsget.EOF  then
    csTotalTodayIpkum = rsget("csTotalTodayIpkum")
	csNotFinishTodayIpkum = rsget("csNotFinishTodayIpkum")
end if
rsget.close


'==============================================================================
' 부서 업무협조
dim myPartWorkCnt : myPartWorkCnt = 0

sqlStr = " SELECT COUNT(r.idx) as cnt "
sqlStr = sqlStr & " FROM [db_temp].[dbo].[tbl_breakdown_request] AS r "
sqlStr = sqlStr & " INNER JOIN [db_temp].[dbo].[tbl_breakdown_request_detail] AS d ON r.idx = d.req_idx "
sqlStr = sqlStr & " WHERE 1 = 1 "
sqlStr = sqlStr & " AND d.work_part_sn = 10 "
sqlStr = sqlStr & " AND d.work_state < '5' "
sqlStr = sqlStr & " AND d.isusing = 'Y' "

rsget.CursorLocation = adUseClient
rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
if  not rsget.EOF  then
    myPartWorkCnt = rsget("cnt")
end if
rsget.close


'==============================================================================
'// 고객 직접품절취소 후 출고가능 주문목록
dim chulgoAbleOrderArr(), chulgoAbleOrderItem
dim chulgoAbleOrderRegdate

sqlStr = " select * from [db_cs].[dbo].[tbl_chulgoAbleOrderList] "
sqlStr = sqlStr + " order by canceldate, makerid, orderserial "
rsget.CursorLocation = adUseClient
rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
if  not rsget.EOF  then
	redim chulgoAbleOrderArr(rsget.RecordCount - 1, 3)
	for i = 0 to UBound(chulgoAbleOrderArr)
		if (i = 0) then
			chulgoAbleOrderRegdate = rsget("regdate")
		end if

		chulgoAbleOrderArr(i, 0) = rsget("canceldate")
		chulgoAbleOrderArr(i, 1) = rsget("makerid")
		chulgoAbleOrderArr(i, 2) = rsget("orderserial")
		chulgoAbleOrderArr(i, 3) = rsget("buyname")
		rsget.movenext
	next
end if
rsget.close


'==============================================================================
Dim oCTaxRequest
set oCTaxRequest = new CTaxRequest
	oCTaxRequest.FRectUseYN = "Y"
	oCTaxRequest.FRectFinishYN = "N"
	oCTaxRequest.FPageSize = 20
	oCTaxRequest.GetTaxRequestList

if request("force") = "Y" then
	response.write	"<script language='javascript'>" &_
					"	alert('저장되었습니다.'); location.href = 'cscenter_main.asp?menupos=" & menupos & "'; " &_
					"</script>"
	dbget.close() : response.end
end if

dim ErrExists : ErrExists = False

%>
<script language="JavaScript" src="/cscenter/js/cscenter.js?v=1.1"></script>
<script language="JavaScript" src="/cscenter/js/convert.date.js"></script>
<script language='javascript'>

function popUpcheMisendNEW(isupchedeliver, makerid, sitename, dplusOver, MisendReason, MisendState, upcheNoCheck, Dtype, exMayChulgo) {
	if (isupchedeliver == "") {
		isupchedeliver = "Y";
	}

	var window_width = 1400;
    var window_height = 800;
	var url = "/admin/upchebeasong/upchemibeasonglistNEW.asp?research=on&menupos=1750&isupchedeliver=" + isupchedeliver + "&makerid=" + makerid + "&sitename=" + sitename + "&Dtype=" + Dtype + "&dplusOver=" + dplusOver + "&MisendReason=" + MisendReason + "&MisendState=" + MisendState + "&upcheNoCheck=" + upcheNoCheck + "&exinmaychulgoday=" + exMayChulgo;

	var popwin = window.open(url, "cs_action","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
	popwin.focus();
}

<%
' function popUpcheMisend(misendCode, dplusOver, currState){
' 	var MisendState = "";
' 	if (misendCode)
' 		MisendState = "0";

' 	var window_width = 1024;
'     var window_height = 800;
' 	var popwin = window.open("/admin/upchebeasong/upchemibeasonglist.asp?menupos=246&popupFlag=Y&MisendState="+MisendState+"&MisendReason=" + misendCode + "&dplusOver=" + dplusOver + "&currState=" + currState , "cs_action","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
' 	popwin.focus();
' }

' function popUpcheMisendByBrand(makerid, misendCode, dplusOver, currState){
' 	var MisendState = "";
' 	if (misendCode)
' 		MisendState = "0";

' 	var window_width = 1024;
'     var window_height = 800;
' 	var popwin = window.open("/admin/upchebeasong/upchemibeasonglist.asp?menupos=246&popupFlag=Y&exinmaychulgoday=Y&makerid="+makerid+"&MisendState="+MisendState+"&MisendReason=" + misendCode + "&dplusOver=" + dplusOver + "&currState=" + currState , "popUpcheMisendByBrand","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
' 	popwin.focus();

' }
%>

function popUpcheStockout10x10ByOrderserial(orderserial){
	var window_width = 1024;
    var window_height = 800;

	var popwin = window.open("/admin/ordermaster/misendmaster_main.asp?orderserial=" + orderserial, "popUpcheStockout10x10ByOrderserial","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
	popwin.focus();

}

function popUpcheStockoutOtherSiteByOrderserial(orderserial, sitename){
	var window_width = 1024;
    var window_height = 800;

	var popwin = window.open("/admin/ordermaster/misendmaster_main.asp?orderserial=" + orderserial, "popUpcheStockoutOtherSiteByOrderserial","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
	popwin.focus();

}

<%
' function popUpcheMichulgoOtherSiteByOrderserial(orderserial, sitename){
' 	var window_width = 1024;
'     var window_height = 800;

' 	var popwin = window.open("/admin/ordermaster/misendmaster_main.asp?orderserial=" + orderserial, "popUpcheMichulgoOtherSiteByOrderserial","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
' 	popwin.focus();

' }
%>

function popUpcheMisendOtherSiteByBrand(makerid, sitename) {
	var window_width = 1024;
    var window_height = 800;

	var popwin = window.open("/admin/upchebeasong/upchemibeasonglist.asp?popupFlag=Y&dplusOver=3&exinmaychulgoday=Y&makerid=" + makerid + "&sitename=" + sitename, "popUpcheMisendOtherSiteByBrand","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
	popwin.focus();
}

function PopItemQnaNotAnswer(makerid, istenbea){
	var window_width = 1024;
    var window_height = 800;

	var v = "/admin/board/newitemqna_list.asp?makerid=" + makerid + "&notupbea=" + istenbea + "&mifinish=on";
	if (istenbea == "N") {
		v = v + "&dplusday=3";
	}

	var popwin = window.open(v, "PopItemQnaNotAnswer","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
	popwin.focus();
}

function setUpcheStockout10x10ByOrderserial(frm){
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

function Cscenter_Action_List2_A060(writeUser) {
    var window_width = 1280;
    var window_height = 960;
    var params;

    params = '?divcd=A060';
    params = params + '&searchfield=writeUser&searchstring=' + writeUser;
    params = params + '&startdt=<%= sixmonthbefore %>&enddt=<%= nowdate %>';

    var popwin = window.open("/cscenter/action/cs_action.asp" + params,"Cscenter_Action_List2_A060","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");

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
    var window_width = 1400;
    var window_height = 960;
	var url = "/cscenter/mifinish/cs_mifinishlist.asp?divcd=returncs&Dtype=dday&dplusOver=" + dday + "&exinmaychulgoday=Y&exoldcs=Y";

    var popwin = window.open(url,"Cscenter_Action_MiFinishReturnListDPlus","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");

	popwin.focus();
}

function Cscenter_Action_MiFinishReturnListByBrand(makerid, dday) {
    var window_width = 1280;
    var window_height = 960;
	var url = "/cscenter/mifinish/cs_mifinishlist.asp?divcd=returncs&vSiteName=10x10&Dtype=dday&dplusOver=" + dday + "&exinmaychulgoday=Y&exoldcs=Y&makerid=" + makerid;

	if (dday*1 == 3) {
		//MifinishReason
		url = url + "&MifinishReason=00";
	}

    var popwin = window.open(url,"Cscenter_Action_MiFinishReturnListByBrand","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");

	popwin.focus();
}

function Cscenter_Action_MiFinishReturnListOtherSiteByBrand(makerid, extsitename) {
    var window_width = 1280;
    var window_height = 960;
	var url = "/cscenter/mifinish/cs_mifinishlist.asp?divcd=returncs&vSiteName=" + extsitename + "&Dtype=dday&dplusOver=7&exinmaychulgoday=Y&exoldcs=Y&makerid=" + makerid;

    var popwin = window.open(url,"Cscenter_Action_MiFinishReturnListOtherSiteByBrand","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");

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

function cscenter_memo_list_FIN(orderserial, userid, finishyn, writeUser) {
	var popwin = window.open("/cscenterv2/history/history_memo_list.asp?orderserial=" + orderserial + "&userid=" + userid + "&finishyn=" + finishyn + "&writeUser=" + writeUser,"cscenter_memo_list_FIN","width=1000 height=700 scrollbars=yes resizable=yes");
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

function cscenter_eventprize_list(){
	window.open("/admin/eventmanage/event/eventprize_list.asp?menupos=1056","cscenter_eventprize","width=1000 height=700 scrollbars=yes resizable=yes");
}

function cscenter_member_list(){
	window.open("/cscenter/member/customerinfo.asp?menupos=1166","cscenter_member","width=1000 height=700 scrollbars=yes resizable=yes");
}

function cscenter_deposit_list(){
	var popwin = window.open("/cscenter/deposit/cs_deposit.asp?menupos=1324","cscenter_deposit","width=1000 height=700 scrollbars=yes resizable=yes");
	popwin.focus();
}

function cscenter_eventjoin_list(){
	var popwin = window.open("/admin/eventmanage/event/eventjoin_list.asp?menupos=1529","eventjoin","width=1100 height=700 scrollbars=yes resizable=yes");
	popwin.focus();
}

function cashreceiptInfo(iorderserial){
	var receiptUrl = "/cscenter/taxsheet/popCashReceipt.asp?orderserial=" + iorderserial;
	var popwin = window.open(receiptUrl,"Cashreceipt","width=1100,height=400,scrollbars=yes,resizable=yes");
	popwin.focus();
}

function PopCSSMSCertLog(userid, usercell, usermail) {
	var url = "/cscenter/action/pop_cs_smscert_log.asp?userid=" + userid + "&usercell=" + usercell + "&usermail=" + usermail;
	var popwin = window.open(url,"PopCSSMSCertLog","width=800,height=400,scrollbars=yes,resizable=yes");
	popwin.focus();
}

function PopCSKAKAOLog() {
	var url = "/cscenter/action/pop_cs_kakao_log.asp";
	var popwin = window.open(url,"PopCSkakaoLog","width=800,height=400,scrollbars=yes,resizable=yes");
	popwin.focus();
}

function PopCSReqWork() {
	var url = "/admin/breakdown/index.asp?menupos=1378&work_part_sn=10&search_state=N";
	var popwin = window.open(url,"PopCSReqWork","width=1500,height=800,scrollbars=yes,resizable=yes");
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

function jsSearchEvent(frm) {
	if(frm.selEvt.value == "evt_code" && frm.sEtxt.value != "") {
		frm.sEtxt.value = frm.sEtxt.value.replace(/\s/g, "");
		if(!IsDigit(frm.sEtxt.value)) {
			alert("이벤트코드는 숫자만 가능합니다.");
			frm.sEtxt.focus();
			return;
		}
	}

	var param = "&selEvt=" + frm.selEvt.value + "&sEtxt=" + frm.sEtxt.value + "&selDate=O&iSD=<%= Left(DateAdd("m",-3, Now()),10) %>&iED=<%= Left(Now(),10) %>"
	var win = window.open("/admin/eventmanage/event/v2/index.asp?menupos=1739" + param,"jsSearchEvent","width=1500, height=800, resizable=yes, scrollbars=yes");
	win.focus();
}

function PopMyQnaListBySiteName(sitename) {
	var param = "sitename=" + sitename;
	var win = window.open("/cscenter/board/cscenter_qna_board_list.asp?research=on&finishyn=N&isusing=Y&" + param,"PopMyQnaListBySiteName","width=1500, height=800, resizable=yes, scrollbars=yes");
}

var csTimeOneToOneBoard = new Date(getDateFromFormat("<%= application("csTimeOneToOneBoard") %>", "yyyy-MM-dd a h:mm:ss"));
var csTimeMichulgoListUpcheNew = new Date(getDateFromFormat("<%= application("csTimeMichulgoListUpcheNew") %>", "yyyy-MM-dd a h:mm:ss"));
var csTimeMiAnswerList = new Date(getDateFromFormat("<%= application("csTimeMiAnswerList") %>", "yyyy-MM-dd a h:mm:ss"));
var csTimeIpjumMichulgo = new Date(getDateFromFormat("<%= application("csTimeIpjumMichulgo") %>", "yyyy-MM-dd a h:mm:ss"));
var csTimeMichulgoListTenTen = new Date(getDateFromFormat("<%= application("csTimeMichulgoListTenTen") %>", "yyyy-MM-dd a h:mm:ss"));
var csTimeCSList = new Date(getDateFromFormat("<%= application("csTimeCSList") %>", "yyyy-MM-dd a h:mm:ss"));
var csTimeUpcheReturn = new Date(getDateFromFormat("<%= application("csTimeUpcheReturn") %>", "yyyy-MM-dd a h:mm:ss"));
var csTimeIpjumStockOut = new Date(getDateFromFormat("<%= application("csTimeIpjumStockOut") %>", "yyyy-MM-dd a h:mm:ss"));
var csTimeIpjumRefund = new Date(getDateFromFormat("<%= application("csTimeIpjumRefund") %>", "yyyy-MM-dd a h:mm:ss"));
var csTimeTenTenStockOut = new Date(getDateFromFormat("<%= application("csTimeTenTenStockOut") %>", "yyyy-MM-dd a h:mm:ss"));

function DisplayClock() {
	var v = new Date();

	var objOneToOneBoard = document.getElementById("objOneToOneBoard");
	var objMichulgoListUpche = document.getElementById("objMichulgoListUpche");
	var objMiAnswerList = document.getElementById("objMiAnswerList");
	var objIpjumMichulgo = document.getElementById("objIpjumMichulgo");
	var objMichulgoListTenTen = document.getElementById("objMichulgoListTenTen");
	var objCSList = document.getElementById("objCSList");

	var objUpcheReturn = document.getElementById("objUpcheReturn");

	var objIpjumStockOut = document.getElementById("objIpjumStockOut");
	var objTenTenStockOut = document.getElementById("objTenTenStockOut");

	objOneToOneBoard.innerHTML = GetDateDiffString(v.getTime() - csTimeOneToOneBoard.getTime());
	objMichulgoListUpche.innerHTML = GetDateDiffString(v.getTime() - csTimeMichulgoListUpcheNew.getTime());
	// objMiAnswerList.innerHTML = GetDateDiffString(v.getTime() - csTimeMiAnswerList.getTime());
	objIpjumMichulgo.innerHTML = GetDateDiffString(v.getTime() - csTimeIpjumMichulgo.getTime());
	objMichulgoListTenTen.innerHTML = GetDateDiffString(v.getTime() - csTimeMichulgoListTenTen.getTime());
	objCSList.innerHTML = GetDateDiffString(v.getTime() - csTimeCSList.getTime());

	objUpcheReturn.innerHTML = GetDateDiffString(v.getTime() - csTimeUpcheReturn.getTime());
	objUpcheReturn7_2.innerHTML = GetDateDiffString(v.getTime() - csTimeUpcheReturn.getTime());		// 타이머 같이 쓰자

	objIpjumStockOut.innerHTML = GetDateDiffString(v.getTime() - csTimeIpjumStockOut.getTime());
	objIpjumRefund.innerHTML = GetDateDiffString(v.getTime() - csTimeIpjumRefund.getTime());
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

function jsUpdateChulgoAbleOrder() {
	var frm = document.frm;

	frm.mode.value = "updChulgoAbleOrdr";
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
            			    <td> <font color="blue">고객건의사항 미처리 상담건</font></td>
            			    <td align="right">
            			        <a href="javascript:PopMyQna('', '', 'N', '26', '', '', '', '', '');">
        				        <b><%= csrecommendBoardcnt %></b> 건
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
						</tr>
            			<tr height="25">
            			    <td> VVIP고객 담당자별 미처리 상담건</td>
            			    <td align="right"></td>
            			</tr>
            			<%
						ReDim tmpVal(1)
						i = 0
            			for i = 0 to UBound(csVVipBoardUserArr)
							If (csVVipBoardUserArr(i) <> "xxxxxxxx") Then
								tmpVal(0) = csVVipBoardUserArr(i)
								tmpVal(1) = 0

								For j = 0 To UBound(csVVipBoardChargecntArr)
									tmpVal2 = Split(csVVipBoardChargecntArr(j), ",")
									If (UBound(tmpVal2) = 1) Then
										If tmpVal(0) = tmpVal2(0) Then
											tmpVal(1) = tmpVal2(1)
											Exit For
										End If
									End If
								Next
            			%>
            			<tr height="25">
            			    <td>
								&nbsp;&nbsp;-
								<% if (csMainUserName = tmpVal(0)) then %><font color="#DB3A00"><b><% end if %>
								<%= tmpVal(0) %>
							</td>
            			    <td align="Right">
            			        <a href="javascript:PopMyQnaListByChargeId('<%= tmpVal(0) %>', 'N');">
        				        	<b><%= tmpVal(1) %></b> 건
									<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
						</tr>
						<%
							End If
						Next
						%>
            			<tr height="25">
            			    <td>&nbsp;&nbsp;- VVIP고객 미분배 상담건</td>
            			    <td align="right">
            			        <a href="javascript:PopMyQna('', '', 'VV', '', '', '', '', '', '');">
        							<b><%= csVVipBoardNochargecnt %></b> 건
        							<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
						</tr>
            			<tr height="25">
            			    <td> VIP고객 담당자별 미처리 상담건</td>
            			    <td align="right"></td>
            			</tr>
            			<%
						ReDim tmpVal(1)
            			for i = 0 to UBound(csVipBoardUserArr)
							If (csVipBoardUserArr(i) <> "xxxxxxxx") Then
								tmpVal(0) = csVipBoardUserArr(i)
								tmpVal(1) = 0

								For j = 0 To UBound(csVipBoardChargecntArr)
									tmpVal2 = Split(csVipBoardChargecntArr(j), ",")
									If (UBound(tmpVal2) = 1) Then
										If tmpVal(0) = tmpVal2(0) Then
											tmpVal(1) = tmpVal2(1)
											Exit For
										End If
									End If
								Next
            			%>
            			<tr height="25">
            			    <td>
								&nbsp;&nbsp;-
								<% if (csMainUserName = tmpVal(0)) then %><font color="#DB3A00"><b><% end if %>
								<%= tmpVal(0) %>
							</td>
            			    <td align="Right">
            			        <a href="javascript:PopMyQnaListByChargeId('<%= tmpVal(0) %>', 'N');">
        				        	<b><%= tmpVal(1) %></b> 건
									<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
						</tr>
						<%
							End If
						Next
						%>
            			<tr height="25">
            			    <td>&nbsp;&nbsp;- VIP고객 미분배 상담건</td>
            			    <td align="right">
            			        <a href="javascript:PopMyQna('', '', 'V', '', '', '', '', '', '');">
        				        <b><%= csVipBoardNochargecnt %></b> 건
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
						</tr>
            			<tr height="25">
            			    <td> 일반고객 담당자별 미처리 상담건</td>
            			    <td align="right"></td>
            			</tr>
            			<%
						ReDim tmpVal(1)
            			for i = 0 to UBound(csBoardUserArr)
							If (csBoardUserArr(i) <> "xxxxxxxx") Then
								tmpVal(0) = csBoardUserArr(i)
								tmpVal(1) = 0

								For j = 0 To UBound(csBoardChargecntArr)
									tmpVal2 = Split(csBoardChargecntArr(j), ",")
									If (UBound(tmpVal2) = 1) Then
										If tmpVal(0) = tmpVal2(0) Then
											tmpVal(1) = tmpVal2(1)
											Exit For
										End If
									End If
								Next

            			%>
            			<tr height="25">
            			    <td>
								&nbsp;&nbsp;-
								<% if (csMainUserName = tmpVal(0)) then %><font color="#DB3A00"><b><% end if %>
								<%= tmpVal(0) %>
							</td>
            			    <td align="Right">
            			        <a href="javascript:PopMyQnaListByChargeId('<%= tmpVal(0) %>', 'N');">
        				        	<b><%= tmpVal(1) %></b> 건
									<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
						</tr>
						<%
							End If
						Next
						%>
            			<tr height="25">
            			    <td>&nbsp;&nbsp;- 일반고객 미분배 상담건</td>
            			    <td align="right">
            			        <a href="javascript:PopMyQnaListByChargeId('', 'N');">
        				        <b><%= csBoardNochargecnt %></b> 건
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td> <font color="blue">STAFF 미처리 상담건</font></td>
            			    <td align="right">
            			        <a href="javascript:PopMyQna('', '', 'N', '', '', '', '', '7', '');">
        				        <b><%= csstaffBoardcnt %></b> 건
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
						</tr>
            			<tr height="25">
            			    <td> <font color="blue">STAFF 품절취소요청건</font></td>
            			    <td align="right">
            			        <a href="javascript:showHideBrand('csstaffstockoutordertr');">
        				        <b><%= csStaffStockoutOrderCount %></b> 건
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
						</tr>
	            			<% for i = 0 to UBound(csStaffStockoutOrderArr) %>
		            			<tr height="25" id="csstaffstockoutordertr" style="display:none">
		            			    <td>
		            			    	&nbsp;&nbsp;&nbsp;&nbsp;-
		            			    	<a href="javascript:PopOrderMasterWithCallRingOrderserial('<%= Split(csStaffStockoutOrderArr(i),",")(1) %>')">
			            			    	<%= Split(csStaffStockoutOrderArr(i),",")(1) %>
			            			    	<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
		            			    	</a>
		            			    	&nbsp;<%= Split(csStaffStockoutOrderArr(i),",")(0) %>
		            			    </td>
		            			    <td align="right">
		            			        <a href="javascript:popUpcheStockout10x10ByOrderserial('<%= Split(csStaffStockoutOrderArr(i),",")(1) %>');">
		        				        확인
		        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
		        				        </a>
		            			    </td>
		            			</tr>
	            			<% next %>
            			<tr height="25">
            			    <td>제휴몰 미처리 상담건</td>
            			    <td align="right">
            			    </td>
						</tr>
	            		<%
						for i = 0 to UBound(csExtSiteCntArr)
							if (csExtSiteCntArr(i) <> "") then
								tmpVal = csExtSiteCntArr(i)
								tmpVal = Split(tmpVal, ",")
						%>
            			<tr height="25">
            			    <td>
								&nbsp;&nbsp;-
								<%= tmpVal(0) %>
							</td>
            			    <td align="Right">
            			        <a href="javascript:PopMyQnaListBySiteName('<%= tmpVal(0) %>', 'N');">
        				        	<b><%= tmpVal(1) %></b> 건
									<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
						</tr>
	            		<%
							end if
						next
						%>
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
        <!-- CS 팀장님 요청으로 숨김처리(2013-07-24 skyer9)<tr valign="top">
            <td>
        	    <!-aa- 상품문의 답변이전 관리 시작-aa->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("menubar") %>">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
                        <tr height="25">
            			    <td style="border-bottom:1px solid #BABABA">
            			    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>상품문의 답변이전 관리</b>
            			    	(<span id="objMiAnswerList"></span>) <a href="javascript:RefreshData('csTimeMiAnswerList')"><img src="/images/icon_reload.gif" border="0"></a>
            			    </td>
            			    <td align="right" style="border-bottom:1px solid #BABABA">
								&nbsp;
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td>상품문의 답변이전[1개월, 업배, D+3일이내 제외]</td>
            			    <td align="right"><a href="javascript:showHideBrand('csMiAnswerListByBrandUpche');"><b><%= FormatNumber(csTotalCountMiAnswerUpche, 0) %></b> 건</a></td>
            			</tr>
            			<%
            			for i = 0 to UBound(csMiAnswerListByBrandUpche)
							if (csMiAnswerListByBrandUpche(i) <> "") then
								tmpVal = csMiAnswerListByBrandUpche(i)
								tmpVal = Split(tmpVal, ",")
            			%>
            			<tr height="25" id="csMiAnswerListByBrandUpche" style="display:none">
            			    <td>&nbsp;&nbsp;- <%= tmpVal(0) %></td>
            			    <td align="right">
            			        <a href="javascript:PopItemQnaNotAnswer('<%= tmpVal(0) %>', 'N');">
        				        <%= tmpVal(1) %> 건
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
            			</tr>
							<% end if %>
            			<% next %>
            			<tr height="25">
            			    <td>상품문의 답변이전[1개월, 텐배]</td>
            			    <td align="right"><a href="javascript:showHideBrand('csMiAnswerListByBrandTen');"><b><%= FormatNumber(csTotalCountMiAnswerTen, 0) %></b> 건</a></td>
            			</tr>
            			<%
            			for i = 0 to UBound(csMiAnswerListByBrandTen)
							if (csMiAnswerListByBrandTen(i) <> "") then
								tmpVal = csMiAnswerListByBrandTen(i)
								tmpVal = Split(tmpVal, ",")
            			%>
            			<tr height="25" id="csMiAnswerListByBrandTen" style="display:none">
            			    <td>&nbsp;&nbsp;- <%= tmpVal(0) %></td>
            			    <td align="right">
            			        <a href="javascript:PopItemQnaNotAnswer('<%= tmpVal(0) %>', 'Y');">
        				        <%= tmpVal(1) %> 건
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
            			</tr>
							<% end if %>
            			<% next %>
            	        </table>
            	    </td>
            	</tr>
            	</table>
        	    <!-aa- 상품문의 답변이전 관리 끝-aa->
        	</td>
        </tr>
        <tr valign="top">
            <td height="10"></td>
		</tr>-->
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
								(<span id="objMichulgoListUpche"></span>) <a href="javascript:RefreshData('csTimeMichulgoListUpcheNew')"><img src="/images/icon_reload.gif" border="0"></a>
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
            			    <td align="right"><b><%=FormatNumber(arrMiSend(0),0)%></b> 건<a href="javascript:popUpcheMisendNEW('Y', '', '', '2', '', '', 'Y', 'topN', '');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<tr height="25">
            			    <td>3일 이상 미발송건</td>
            			    <td align="right"><b><%=FormatNumber(arrMiSend(1),0)%></b> 건<a href="javascript:popUpcheMisendNEW('Y', '', '', '3', '', '', '', 'topN', '');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
						<% if (UBound(arrMiSend) > 2) then %>
            			<tr height="25">
            			    <td>&nbsp;&nbsp;- 사유 미입력</td>
            			    <td align="right">
								<b><%=FormatNumber(arrMiSend(3),0)%></b> 건<a href="javascript:popUpcheMisendNEW('Y', '', '', '3', '00', '', '', 'all', '');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a>
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td>&nbsp;&nbsp;- 출고예정일 도과</td>
            			    <td align="right">
								<b><%=FormatNumber(arrMiSend(4),0)%></b> 건<a href="javascript:popUpcheMisendNEW('Y', '', '', '3', '', '', '', 'all', 'Y');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a>
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td>&nbsp;&nbsp;- 품절</td>
            			    <td align="right">
								<b><%=FormatNumber(arrMiSend(5),0)%></b> 건<a href="javascript:popUpcheMisendNEW('Y', '', '', '3', '05', '', '', 'all', '');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a>
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td>&nbsp;&nbsp;- 주문제작</td>
            			    <td align="right">
								<b><%=FormatNumber(arrMiSend(6),0)%></b> 건<a href="javascript:popUpcheMisendNEW('Y', '', '', '3', '02', '', '', 'all', '');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a>
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td>&nbsp;&nbsp;- 가구배송</td>
            			    <td align="right">
								<b><%=FormatNumber(arrMiSend(7),0)%></b> 건<a href="javascript:popUpcheMisendNEW('Y', '', '', '3', '09', '', '', 'all', '');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a>
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td>&nbsp;&nbsp;- 출고지연</td>
            			    <td align="right">
								<b><%=FormatNumber(arrMiSend(8),0)%></b> 건<a href="javascript:popUpcheMisendNEW('Y', '', '', '3', '03', '', '', 'all', '');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a>
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td>&nbsp;&nbsp;- 기타사유</td>
            			    <td align="right">
								<b><%=FormatNumber(arrMiSend(9),0)%></b> 건<a href="javascript:popUpcheMisendNEW('Y', '', '', '3', '', '', '', 'all', '');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a>
            			    </td>
            			</tr>
						<% end if %>
            			<tr height="25">
            			    <td>담당자별 D+3 미발송건(입점몰,현장수령,출고예정 제외)</td>
            			    <td align="right"></td>
            			</tr>
            			<%
						for i = 0 to UBound(csMichulgoList)
							if (Trim(csMichulgoList(i)) <> "") then
								csMichulgoItem = Split(Trim(csMichulgoList(i)), ",")

								if (csMichulgoItem(0) <> "") then
									%>
									<tr height="25">
										<td>
											&nbsp;&nbsp;*
											<% if (csMainUserName = csMichulgoItem(0)) then %><font color="#DB3A00"><b><% end if %>
											<%= csMichulgoItem(0) %>
										</td>
										<td align="right">
											<a href="javascript:showHideBrand('<%= csMichulgoItem(0) %>');">
											<b><%= csMichulgoItem(1) %></b> (<%= csMichulgoItem(2) %>)
											<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
											</a>
										</td>
									</tr>
									<% for j = 0 to UBound(csMichulgoListByBrand) %>
										<% if (Trim(csMichulgoListByBrand(j)) <> "") then %>
											<% csMichulgoItemByBrand = Split(Trim(csMichulgoListByBrand(j)), ",") %>

											<% if CStr(csMichulgoItem(0)) = CStr(csMichulgoItemByBrand(0)) then %>
												<tr height="25" id="<%= csMichulgoItem(0) %>" style="display:none">
													<td>&nbsp;&nbsp;&nbsp;&nbsp;- <%= csMichulgoItemByBrand(1) %></td>
													<td align="right">
														<!--
														<a href="javascript:popUpcheMisendByBrand('<%= csMichulgoItemByBrand(1) %>', '','3','');">
														-->
														<a href="javascript:popUpcheMisendNEW('Y', '<%= csMichulgoItemByBrand(1) %>', '', '3', '', '', '', 'topN', 'Y');">
														<%= csMichulgoItemByBrand(2) %> 건
														<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
														</a>
													</td>
												</tr>
											<% end if %>
										<% end if %>
									<% next %>
								<% end if %>
							<% end if %>
            			<% next %>
            			<tr height="25">
            			    <td>&nbsp;&nbsp;- 미분배 미발송건</td>
            			    <td align="right">
								<a href="javascript:showHideBrand('XXX');">
        				        <b><%= nochargeMichulgoBrandcnt %></b>
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
								</a>
            			    </td>
            			</tr>
						<% for j = 0 to UBound(csMichulgoListByBrand) %>
							<% if (Trim(csMichulgoListByBrand(j)) <> "") then %>
								<% csMichulgoItemByBrand = Split(Trim(csMichulgoListByBrand(j)), ",") %>

								<% if "" = CStr(csMichulgoItemByBrand(0)) then %>
								<tr height="25" id="XXX" style="display:none">
									<td>&nbsp;&nbsp;&nbsp;&nbsp;- <%= csMichulgoItemByBrand(1) %></td>
									<td align="right">
										<!--
											 <a href="javascript:popUpcheMisendByBrand('<%= csMichulgoItemByBrand(1) %>', '','3','');">
										   -->
										<a href="javascript:popUpcheMisendNEW('Y', '<%= csMichulgoItemByBrand(1) %>', '', '3', '', '', '', 'topN', 'Y');">
											<%= csMichulgoItemByBrand(2) %> 건
											<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
										</a>
									</td>
								</tr>
								<% end if %>
							<% end if %>
						<% next %>
            			<tr height="25">
            			    <td>품절취소요청건</td>
            			    <td align="right"><b><%=FormatNumber(arrMiSend(2),0)%></b> 건<a href="javascript:popUpcheMisendNEW('Y', '', '', '', '05', '0', '', 'topN', '');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<!--<tr height="25">
            			    <td colspan="2"></td>
            			 	too Many
            			    <td>미출고사유 미입력건 (D+2이상)</td>
            			    <td align="right"><b>??</b> 건<img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></td>
            			</tr>-->
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
        	    <!-- 입점몰별 미출고 관리 시작 -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="#B2CCFF">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
                        <tr height="25">
            			    <td style="border-bottom:1px solid #BABABA">
            			    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>입점몰 미출고리스트[업배]</b>
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
						prevSiteName = ""
            			for i = 0 to UBound(csMichulgoOtherSiteBrandSiteNameArr)
            			%>
							<% if (prevSiteName <> csMichulgoOtherSiteBrandSiteNameArr(i)) then %>
								<%
								prevSiteName = csMichulgoOtherSiteBrandSiteNameArr(i)

								'// 사이트별 합계 구하기
								totBrandCount = 0
								totItemCount = 0
								for j = 0 to UBound(csMichulgoOtherSiteBrandSiteNameArr)
									if (prevSiteName = csMichulgoOtherSiteBrandSiteNameArr(j)) then
										totBrandCount = totBrandCount + 1
										totItemCount = totItemCount + csMichulgoOtherSiteBrandcntArr(j)
									end if
								next
								%>
								<tr height="25">
									<td>&nbsp;&nbsp;* <%= csMichulgoOtherSiteBrandSiteNameArr(i) %></td>
									<td align="right">
										<a href="javascript:showHideBrand('MichulgoOtherSite<%= csMichulgoOtherSiteBrandSiteNameArr(i) %>');">
										<b><%= totBrandCount %></b> (<%= totItemCount %>)
										<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
										</a>
									</td>
								</tr>
								<% for j = 0 to UBound(csMichulgoOtherSiteBrandNameArr) %>
									<% if prevSiteName = csMichulgoOtherSiteBrandSiteNameArr(j) then %>
									<tr height="25" id="MichulgoOtherSite<%= csMichulgoOtherSiteBrandSiteNameArr(i) %>" style="display:none">
										<td>&nbsp;&nbsp;&nbsp;&nbsp;- <%= csMichulgoOtherSiteBrandNameArr(j) %></td>
										<td align="right">
											<a href="javascript:popUpcheMisendOtherSiteByBrand('<%= csMichulgoOtherSiteBrandNameArr(j) %>', '<%= prevSiteName %>');">
											<%= csMichulgoOtherSiteBrandcntArr(j) %> 건
											<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
											</a>
										</td>
									</tr>
									<% end if %>
								<% next %>
							<% end if %>
            			<% next %>
            	        </table>
            	    </td>
            	</tr>
            	</table>
        	    <!-- 입점몰별 미출고 관리 끝-->
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
            			    <td align="right"><a href="javascript:popUpcheMisendNEW('N', '', '', '', '00', '', '', 'topN', '');"><b><%=FormatNumber(arrTenMiSend(0),0)%></b> 건<img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<tr height="25">
            			    <td>고객 미안내건</td>
            			    <td align="right"><a href="javascript:popUpcheMisendNEW('N', '', '', '', '', '0', '', 'topN', '');"><b><%=FormatNumber(arrTenMiSend(1),0)%></b> 건<img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<tr height="25">
            			    <td>SMS/Mail/통화완료건</td>
            			    <td align="right"><a href="javascript:misendmaster('4');"><b>-</b> 건<img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<tr height="25">
            			    <td>CS처리완료(물류팀처리요청)건</td>
            			    <td align="right"><a href="javascript:misendmaster('6');"><b>-</b> 건<img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
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
            <td width="50%">
        	    <!-- 반품접수(업체배송, 입점몰제외) D+4 미처리 시작 -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("menubar") %>">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
                        <tr height="25">
            			    <td style="border-bottom:1px solid #BABABA">
            			    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>반품접수(업배) D+4 미처리</b>
            			    	(<span id="objUpcheReturn"></span>) <a href="javascript:RefreshData('csTimeUpcheReturn')"><img src="/images/icon_reload.gif" border="0"></a>
            			    </td>
            			    <td align="right" style="border-bottom:1px solid #BABABA">
            			        <a href="javascript:Cscenter_Action_MiFinishReturnListDPlus(7);">
        				        바로가기
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td>반품접수(업체배송) D+4 미처리</td>
            			    <td align="right"><b><%= upcheReturnNotFinish %></b> 건<a href="javascript:Cscenter_Action_MiFinishReturnListDPlus(4);"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<tr height="25">
            			    <td>담당자별 D+4 미처리건(반품예정 제외, 입점몰제외)</td>
            			    <td align="right"></td>
            			</tr>
            			<%
						for i = 0 to UBound(csUpcheReturnList)
							if (Trim(csUpcheReturnList(i)) <> "") then
								csUpcheReturnItem = Split(Trim(csUpcheReturnList(i)), ",")

								if (csUpcheReturnItem(0) <> "") then
									%>
									<tr height="25">
										<td>
											&nbsp;&nbsp;*
											<% if (csMainUserName = csUpcheReturnItem(0)) then %><font color="#DB3A00"><b><% end if %>
											<%= csUpcheReturnItem(0) %>
										</td>
										<td align="right">
											<a href="javascript:showHideBrand('return<%= csUpcheReturnItem(0) %>');">
											<b><%= csUpcheReturnItem(1) %></b> (<%= csUpcheReturnItem(2) %>)
											<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
											</a>
										</td>
									</tr>
									<% for j = 0 to UBound(csUpcheReturnListByBrand) %>
										<% if (Trim(csUpcheReturnListByBrand(j)) <> "") then %>
											<% csUpcheReturnItemByBrand = Split(Trim(csUpcheReturnListByBrand(j)), ",") %>

											<% if CStr(csUpcheReturnItem(0)) = CStr(csUpcheReturnItemByBrand(0)) then %>
												<tr height="25" id="return<%= csUpcheReturnItem(0) %>" style="display:none">
													<td>&nbsp;&nbsp;&nbsp;&nbsp;- <%= csUpcheReturnItemByBrand(1) %></td>
													<td align="right">
														<a href="javascript:Cscenter_Action_MiFinishReturnListByBrand('<%= csUpcheReturnItemByBrand(1) %>', 4);">
														<%= csUpcheReturnItemByBrand(2) %> 건
														<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
														</a>
													</td>
												</tr>
											<% end if %>
										<% end if %>
									<% next %>
								<% end if %>
							<% end if %>
            			<% next %>
            			<tr height="25">
            			    <td>&nbsp;&nbsp;- 미분배 미처리건</td>
            			    <td align="right">
        				        <b><%= nochargeUpcheReturnBrandcnt %></b>
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
        	    <!-- 입점몰별 반품접수 관리 시작 -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="#B2CCFF">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
                        <tr height="25">
            			    <td style="border-bottom:1px solid #BABABA">
            			    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>입점몰 반품접수 D+7 미처리</b>
								(<span id="objUpcheReturn7_2"></span>) <a href="javascript:RefreshData('csTimeUpcheReturn')"><img src="/images/icon_reload.gif" border="0"></a>
            			    </td>
            			    <td align="right" style="border-bottom:1px solid #BABABA">
            			    	&nbsp;
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td>입점몰별(반품예정 제외)</td>
            			    <td align="right">
            			    </td>
            			</tr>
            			<%
						prevSiteName = ""
            			for i = 0 to UBound(csReturn7OtherSiteSiteNameArr)
            			%>
							<% if (prevSiteName <> csReturn7OtherSiteSiteNameArr(i)) then %>
								<%
								prevSiteName = csReturn7OtherSiteSiteNameArr(i)

								'// 사이트별 합계 구하기
								totBrandCount = 0
								totItemCount = 0
								for j = 0 to UBound(csReturn7OtherSiteSiteNameArr)
									if (prevSiteName = csReturn7OtherSiteSiteNameArr(j)) then
										totBrandCount = totBrandCount + 1
										totItemCount = totItemCount + csReturn7OtherSiteBrandcntArr(j)
									end if
								next
								%>
								<tr height="25">
									<td>&nbsp;&nbsp;* <%= csReturn7OtherSiteSiteNameArr(i) %></td>
									<td align="right">
										<a href="javascript:showHideBrand('Return7OtherSite<%= csReturn7OtherSiteSiteNameArr(i) %>');">
										<b><%= totBrandCount %></b> (<%= totItemCount %>)
										<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
										</a>
									</td>
								</tr>
								<% for j = 0 to UBound(csReturn7OtherSiteBrandNameArr) %>
									<% if prevSiteName = csReturn7OtherSiteSiteNameArr(j) then %>
									<tr height="25" id="Return7OtherSite<%= csReturn7OtherSiteSiteNameArr(i) %>" style="display:none">
										<td>&nbsp;&nbsp;&nbsp;&nbsp;- <%= csReturn7OtherSiteBrandNameArr(j) %></td>
										<td align="right">
											<a href="javascript:Cscenter_Action_MiFinishReturnListOtherSiteByBrand('<%= csReturn7OtherSiteBrandNameArr(j) %>', '<%= prevSiteName %>');">
											<%= csReturn7OtherSiteBrandcntArr(j) %> 건
											<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
											</a>
										</td>
									</tr>
									<% end if %>
								<% next %>
							<% end if %>
            			<% next %>
            	        </table>
            	    </td>
            	</tr>
            	</table>
        	    <!-- 입점몰별 미출고 관리 끝-->
        	</td>
        </tr>
        <tr valign="top">
            <td height="10"></td>
        </tr>
        <tr valign="top">
            <td width="50%">
        	    <!-- 업체배송 관리 시작 -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="#B2CCFF">
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
            			    <td>입점몰별(고객 안내이전+고객안내)</td>
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
                                <%
                                ErrExists = False
                                if (Not IsArray(csStockoutOrderserialUseridArr)) or (Not IsArray(csMichulgoStockoutBoardUserArr)) or (Not IsArray(csStockoutOrderUseridArr)) then
                                    ErrExists = True
                                elseif UBound(csMichulgoStockoutBoardUserArr) <> UBound(csMichulgoStockoutBoardUserArr) then
                                    ErrExists = True
                                elseif UBound(csStockoutOrderUseridArr) < 0 then
                                    ErrExists = True
                                end if

                                %>
            			    </td>
            			</tr>
            			<%
if ErrExists = True then
                        %>
                        <tr height="25">
            			    <td>
								데이타에 오류가 있습니다. 새로고침하세요.(<%= application("csStockoutOrderUseridArr") %>)
							</td>
            			    <td align="right">
            			    </td>
            			</tr>
                        <%
else
                        '// 첫번째것 스킵한다.(에러방지)
            			for i = 1 to UBound(csMichulgoStockoutBoardUserArr)
            			%>
            			<tr height="25">
            			    <td>
								&nbsp;&nbsp;*
								<% if (csMainUserName = csMichulgoStockoutBoardUserArr(i)) then %><font color="#DB3A00"><b><% end if %>
								<%= csMichulgoStockoutBoardUserArr(i) %>
							</td>
            			    <td align="right">
            			        <a href="javascript:showHideBrand('stockout<%= csMichulgoStockoutBoardUserArr(i) %>');">
        				        <b><%= csStockoutOrderCountByUserid(i) %></b> (<%= csStockoutOrderDetailCountByUserid(i) %>)
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
            			</tr>
	            			<% for j = 0 to UBound(csStockoutOrderserialArr) %>
	            				<% if UBound(csStockoutOrderserialUseridArr) >= 0 and (csMichulgoStockoutBoardUserArr(i) = csStockoutOrderUseridArr(j)) then %>
									<%
									'// 여기서 값설정
									csStockoutHasChargeUser(j) = "Y"
									%>
		            			<tr height="25" id="stockout<%= csMichulgoStockoutBoardUserArr(i) %>" style="display:none">
		            			    <td>
		            			    	&nbsp;&nbsp;&nbsp;&nbsp;-
		            			    	<a href="javascript:PopOrderMasterWithCallRingOrderserial('<%= csStockoutOrderserialArr(j) %>')"><acronym title="<%= csStockoutRegdateArr(j) %>">
			            			    	<%= csStockoutOrderserialArr(j) %></acronym>
			            			    	<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
		            			    	</a>
										&nbsp;
										<%= csStockoutOrderserialUseridArr(j) %>
										<% if (DateDiff("h",csStockoutRegdateArr(j), Now()) <= 1) then %>
										&nbsp;
										<font color="red">new</font>
										<% end if %>
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
<% end if %>
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
	            			    	<% 'if Not IsNull(csStockoutOrderUseridArr(i)) and csStockoutOrderUseridArr(i) <> "" then %>
	            			    		(<%'= csStockoutOrderUseridArr(i) %>)
	            			    	<% 'end if %>
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

				<p />

        	    <!-- 품절상품 고객취소 후 출고가능 주문목록 -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("menubar") %>">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
                        <tr height="25">
            			    <td style="border-bottom:1px solid #BABABA" colspan="3">
            			    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>품절상품 고객취소 후 출고가능 주문</b>
								(최종업데이트 : <%= chulgoAbleOrderRegdate %>)
								<a href="javascript:jsUpdateChulgoAbleOrder()"><img src="/images/icon_reload.gif" border="0"></a>
            			    </td>
            			    <td align="right" style="border-bottom:1px solid #BABABA">
            			    	&nbsp;
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td align="center" width="170">고객 취소일자</td>
							<td></td>
							<td></td>
							<td></td>
            			</tr>
						<%
						for i = 0 to UBound(chulgoAbleOrderArr)
						%>
            			<tr height="25">
            			    <td><font color="<%= CHKIIF(DateDiff("d", chulgoAbleOrderArr(i, 0), Now) > 0, "red", "black")%>"><%= chulgoAbleOrderArr(i, 0) %></font></td>
            			    <td><%= chulgoAbleOrderArr(i, 1) %></td>
							<td><a href="javascript:PopOrderMasterWithCallRingOrderserial('<%= chulgoAbleOrderArr(i, 2) %>')"><%= chulgoAbleOrderArr(i, 2) %></a></td>
            			    <td><%= chulgoAbleOrderArr(i, 3) %></td>
            			</tr>
						<% next %>
						</table>
					</td>
				</tr>
				</table>
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
						<tr height="25">
            			    <td>업체긴급문의 미처리</td>
            			    <td align="right">
            			        <b><%= CSA060NotFinish %></b> 건
                                <a href="javascript:Cscenter_Action_List2_A060('<%= csMainUserID %>');">
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
        <tr valign="top">
            <td>
                <!-- 이벤트검색 -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("tabletop") %>">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
            			<tr height="25">
            			    <td style="border-bottom:1px solid #BABABA">
            			        <img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>이벤트검색</b>
            			    </td>
            			    <td align="right" style="border-bottom:1px solid #BABABA">
            			        &nbsp;
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td colspan="2">
								<form name="frmEvent" method="post" action="" onSubmit="return false;">
									<select class="select" name="selEvt">
			    						<option value="evt_code" >이벤트 코드</option>
			    						<option value="evt_name">이벤트명</option>
			    					</select>
									<input type="text" class="text" name="sEtxt" value="" size="20" maxlength="60" onKeyPress="if (event.keyCode == 13) jsSearchEvent(frmEvent);">
									<input type="button" class="button" value="검색" onClick="jsSearchEvent(frmEvent)">
								</form>
							</td>
            			</tr>
            	        </table>
            	    </td>
            	</tr>
            	</table>
        	    <!--  이벤트검색 끝-->
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
            			    <td style="border-bottom:1px solid #BABABA" colspan="2">
        						<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
            						<%
									if (oCTaxRequest.FResultCount < 1) then
										%>
									<tr height="25">
            							<td>&nbsp;&nbsp; 요청없음</td>
            						</tr>
										<%
									else
										for i = 0 to oCTaxRequest.FResultCount - 1
										%>
									<tr height="25">
            							<td>
											&nbsp;&nbsp; - <a href="javascript:cashreceiptInfo('<%= oCTaxRequest.FTaxList(i).Forderserial %>');"><%= oCTaxRequest.FTaxList(i).Forderserial %></a>
											<a href="javascript:cashreceiptInfo('<%= oCTaxRequest.FTaxList(i).Forderserial %>');">
			            			    		<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
		            			    		</a>
										</td>
            						</tr>
										<%
										next
									end if
									%>
            					</table>
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td style="border-bottom:1px solid #BABABA">
            			        <img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>무통장입금 관리</b>
            			    </td>
            			    <td align="right" style="border-bottom:1px solid #BABABA">
								미확인 [<b><%= csNotFinishTodayIpkum %></b>] / 오늘 전체 : [<%= csTotalTodayIpkum %>]
								&nbsp;...&nbsp;
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
            			    <td width="5"></td>
            			    <td>
            			    	<table width="100%" style="border:1px solid #BABABA" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#FFFFFF">
		                        <tr height="20">
		            			    <td align="center">
		            			    	<a href="javascript:PopCSfileSend('','','','');">고객파일전송관리</a>
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
        <tr valign="top">
            <td height="10"></td>
        </tr>
        <tr valign="top">
            <td>
            	<!-- 인증번호 -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("tabletop") %>">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
                        <tr height="25">
            			    <td width="33%">
            			    	<table width="100%" style="border:1px solid #BABABA" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#FFFFFF">
		                        <tr height="20">
		            			    <td align="center">
		            			    	<a href="javascript:PopCSSMSCertLog('','','','');">인증번호전송이력</a>
		            			    </td>
		            			</tr>
		            	        </table>
            			    </td>
            			    <td width="5"></td>
            			    <td width="33%">
            			    	<table width="100%" style="border:1px solid #BABABA" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#FFFFFF">
		                        <tr height="20">
		            			    <td align="center">
		            			    	<a href="javascript:PopCSKAKAOLog();">카카오톡전송이력</a>
		            			    </td>
		            			</tr>
		            	        </table>
            			    </td>
            			    <td width="5"></td>
            			    <td>
								<table width="100%" style="border:1px solid #BABABA" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#FFFFFF">
									<tr height="20">
		            					<td align="center">
		            			    		<a href="javascript:PopCSReqWork();">부서업무협조<%= CHKIIF(myPartWorkCnt>0, "<font color='#FF1493'><b>((( " & myPartWorkCnt & " )))</b></font>", "(" & myPartWorkCnt & ")") %></a>
		            					</td>
		            				</tr>
		            	        </table>
            			    </td>
            			</tr>
            	        </table>
            	    </td>
            	</tr>
            	</table>
        	    <!-- 인증번호 -->
            </td>
        </tr>
        <tr valign="top">
            <td height="10"></td>
        </tr>
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
            			    <td>카드취소 D+1 미처리</td>
            			    <td align="right"><b><%= csRequestCardCancelCount %></b> 건<a href="javascript:Cscenter_Action_List2('cardnocheckdp1');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<tr height="25">
            			    <td>주문취소 미처리</td>
            			    <td align="right"><b><%= csNotFinA008 %></b> 건<a href="javascript:Cscenter_Action_List2('cancelnofinish');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<tr height="25">
            			    <td>회수요청 미처리</td>
            			    <td align="right"><b><%= csReturnNotFinish %></b> 건<a href="javascript:Cscenter_Action_List2('returnmifinish');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<!--<tr height="25">
            			    <td>업체 미처리</td>
            			    <td align="right"><b><%'= csUpcheNotFin %></b> 건<a href="javascript:Cscenter_Action_List2('upchemifinish');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>-->
            			<tr height="25">
            			    <td>반품접수(업체배송) D+4 미처리</td>
            			    <td align="right"><b><%= upcheReturnNotFinish %></b> 건<a href="javascript:Cscenter_Action_MiFinishReturnListDPlus(7);"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>

            			<tr height="25">
            			    <td>업체처리완료</td>
            			    <td align="right"><b><%= csUpcheFinished %></b> 건<a href="javascript:Cscenter_Action_List2('upchefinish');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<tr height="25">
            			    <td>물류처리완료</td>
            			    <td align="right"><b><%= csLogicsFinished %></b> 건<a href="javascript:Cscenter_Action_List2('logicsfinish');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<tr height="25">
            			    <td>고객추가결제</td>
            			    <td align="right"><b><%= csCustomerAddPayment %></b> 건<a href="javascript:Cscenter_Action_List2('customeraddpay');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
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
        	    <!-- 입점몰 환불미처리 시작 -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="#B2CCFF">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
                        <tr height="25">
            			    <td style="border-bottom:1px solid #BABABA">
            			    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>입점몰 환불미처리</b>
								(<span id="objIpjumRefund"></span>) <a href="javascript:RefreshData('csTimeIpjumRefund')"><img src="/images/icon_reload.gif" border="0"></a>
            			    </td>
            			    <td align="right" style="border-bottom:1px solid #BABABA">
            			    	&nbsp;
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td>입점몰 환불미처리</td>
            			    <td align="right">
            			    </td>
            			</tr>
            			<%
            			for i = 0 to UBound(csRefundOtherSitenameArr)
            			%>
            			<tr height="25">
            			    <td>&nbsp;&nbsp;* <%= csRefundOtherSitenameArr(i) %></td>
            			    <td align="right">
            			        <a href="javascript:showHideBrand('refundothersite<%= csRefundOtherSitenameArr(i) %>');">
        				        <b><%= csRefundOtherSiteOrderCountArr(i) %></b> (<%= csRefundOtherSiteCSCountArr(i) %>)
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
            			</tr>
	            			<% for j = 0 to UBound(csRefundOtherSiteOrderserialArr) %>
	            				<% if (csRefundOtherSitenameArr(i) = csRefundOtherSiteOrderSitenameArr(j)) then %>
		            			<tr height="25" id="refundothersite<%= csRefundOtherSitenameArr(i) %>" style="display:none">
		            			    <td>
		            			    	&nbsp;&nbsp;&nbsp;&nbsp;-
		            			    	<a href="javascript:PopOrderMasterWithCallRingOrderserial('<%= csRefundOtherSiteOrderserialArr(j) %>')">
			            			    	<%= csRefundOtherSiteOrderserialArr(j) %>
			            			    	<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
		            			    	</a>
		            			    </td>
		            			    <td align="right">
		        				        <%= csRefundOtherSiteCSCntcntArr(j) %> 건
		            			    </td>
		            			</tr>
	            				<% end if %>
	            			<% next %>
            			<% next %>
            	        </table>
            	    </td>
            	</tr>
            	</table>
        	    <!-- 입점몰 환불미처리 끝-->
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
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
