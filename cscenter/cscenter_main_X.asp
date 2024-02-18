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
'// ����� ���
If Trim(application("csTimeUserList")) = "" Or Trim(application("csBoardUserArr")) = "" Or DateDiff("s", application("csTimeUserList"), Now() ) > 1800 Then						'// 30��(1800��) �ʰ��� ���� ������
	application("csTimeUserList") = Now()
	IsUpdateUserListNeed = True
end if

'// 1:1 ���Խ���
If Trim(application("csTimeOneToOneBoard")) = "" Or DateDiff("s", application("csTimeOneToOneBoard"), Now() ) > 600 Then			'// 10��(600��) �ʰ��� ���� ������
	application("csTimeOneToOneBoard") = Now()
	IsUpdateOneToOneBoardNeed = True
end if

'// �������Ʈ[��ü���]
If Trim(application("csTimeMichulgoListUpche")) = "" Or Trim(application("csMiSend")) = "" Or DateDiff("s", application("csTimeMichulgoListUpche"), Now() ) > 900 Then	'// 15��(900��) �ʰ��� ���� ������
	application("csTimeMichulgoListUpche") = Now()
	IsUpdateMichulgoListUpcheNeed = True
end if

'// �������Ʈ[�ٹ�����]
If Trim(application("csTimeMichulgoListTenTen")) = "" Or Trim(application("csTenMiSend")) = "" Or DateDiff("s", application("csTimeMichulgoListTenTen"), Now() ) > 900 Then	'// 15��(900��) �ʰ��� ���� ������
	application("csTimeMichulgoListTenTen") = Now()
	IsUpdateMichulgoListTenTenNeed = True
end if

'// �������Ʈ[������]
If Trim(application("csTimeIpjumMichulgo")) = "" Or Trim(application("csMiSend")) = "" Or DateDiff("s", application("csTimeIpjumMichulgo"), Now() ) > 900 Then	'// 15��(900��) �ʰ��� ���� ������
	application("csTimeIpjumMichulgo") = Now()
	IsUpdateIpjumMichulgoNeed = True
end if

'// CSó������Ʈ
If Trim(application("csTimeCSList")) = "" Or DateDiff("s", application("csTimeCSList"), Now() ) > 900 Then	'// 15��(900��) �ʰ��� ���� ������
	application("csTimeCSList") = Now()
	IsUpdateCSListNeed = True

	application("csTimeUpcheReturn") = Now()
	IsUpdateUpcheReturnNeed = True
end if

'// ��ǰ����(��ü���)
If Trim(application("csTimeUpcheReturn")) = "" Or DateDiff("s", application("csTimeUpcheReturn"), Now() ) > 900 Then	'// 15��(900��) �ʰ��� ���� ������
	application("csTimeCSList") = Now()
	IsUpdateCSListNeed = True

	application("csTimeUpcheReturn") = Now()
	IsUpdateUpcheReturnNeed = True
end if

'// ������ ǰ����ҿ�û��
If Trim(application("csTimeIpjumStockOut")) = "" Or DateDiff("s", application("csTimeIpjumStockOut"), Now() ) > 900 Then	'// 15��(900��) �ʰ��� ���� ������
	application("csTimeIpjumStockOut") = Now()
	IsUpdateIpjumStockOutNeed = True
end if

'// ǰ����ҿ�û��[�ٹ�+����]
If Trim(application("csTimeTenTenStockOut")) = "" Or DateDiff("s", application("csTimeTenTenStockOut"), Now() ) > 900 Then	'// 15��(900��) �ʰ��� ���� ������
	application("csTimeTenTenStockOut") = Now()
	IsUpdateTenTenStockOutNeed = True
end if

'// CS Master Idx
If Trim(application("csTimeMaxCSMasterIdx")) = "" Or DateDiff("s", application("csTimeMaxCSMasterIdx"), Now() ) > 3600 Then			'// 60��(3600��) �ʰ��� ���� ������
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
	sqlStr = sqlStr + " 	and m.id > " & (maxCSMasterIdx - 200000) & " "			'// �ֱ� 20������ �˻��Ѵ�.(�ӵ�����)
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
'// �ٹ����� ����� ���
dim csBoardUserArr, csMichulgoBoardUserArr, csMichulgoStockoutBoardUserArr, csReturnBoardUserArr

if (IsUpdateUserListNeed = True) then

	'// YN �� N �� �ƴѰ��̾�� �Ѵ�.
	'// �й��� ���� YN �� Y �ΰ�, �й������ ǥ���Ҷ��� N �� �ƴѰ�!!
	'// YN = T �� ��� : �й�����(�й�������� ǥ�õǰ�, ���̻� �й���� �ʴ´�.)
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

	'// ��������(����ڰ� �Ѹ� ������ ������ �߻��Ѵ�.)
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
'// 1:1 ���Խ��� ����(����ں�)
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
' �������Ʈ[��ü���] - ���
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
'�������Ʈ[��ü���] - �ٹ��ϼ� ���� D+3 ��
dim tmpSql, michulgoBaseDate

tmpSql = " exec [db_cs].[dbo].[usp_getDayMinusWorkday] '" & Left(now(), 10) & "', 3 " & VbCRLF
rsget.CursorLocation = adUseClient
rsget.Open tmpSql, dbget, adOpenForwardOnly
if Not rsget.Eof then
    '// �ٹ��ϼ� ���� D+3 ��
    michulgoBaseDate = rsget("minusworkday")
end if
rsget.close


'==============================================================================
'�������Ʈ[��ü���] - ����ں� ���Ӹ�
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
	sqlStr = sqlStr + " 			and m.sitename = '10x10' "																'// �������ֹ� ����
	sqlStr = sqlStr + " 			and d.cancelyn <> 'Y' "
	''sqlStr = sqlStr + " 			and datediff(d,m.baljudate,getdate())>=4 "
	sqlStr = sqlStr + " 			and datediff(d,m.baljudate,'" + CStr(michulgoBaseDate) + "')>=0 "						'// �ٹ��ϼ� ���� D+3 ��
	sqlStr = sqlStr + " 			and not ((T.ipgodate is not null) and (datediff(d, T.ipgodate, getdate()) <= 0)) "		'// ������� ���� �ֹ� ����
	sqlStr = sqlStr + " 			and IsNULL(T.code,'00')<>'05' "															'// ǰ�����Ұ� ����
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
'�������Ʈ[��ü���] - �귣�庰 ���Ӹ�
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
	sqlStr = sqlStr + " 	and m.sitename = '10x10' "																		'// �������ֹ� ����
	sqlStr = sqlStr + " 	and d.cancelyn <> 'Y' "
	''sqlStr = sqlStr + " 	and datediff(d,m.baljudate,getdate())>=4 "
	sqlStr = sqlStr + " 	and datediff(d,m.baljudate,'" + CStr(michulgoBaseDate) + "')>=0 "								'// �ٹ��ϼ� ���� D+3 ��
	sqlStr = sqlStr + " 	and not ((T.ipgodate is not null) and (datediff(d, T.ipgodate, getdate()) <= 0)) "				'// ������� ���� �ֹ� ����
	sqlStr = sqlStr + " 	and IsNULL(T.code,'00')<>'05' "																	'// ǰ�����Ұ� ����
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
'�������Ʈ[������,�ٹ�+����] - ������ ����Ʈ
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
	sqlStr = sqlStr + " 	and m.sitename <> '10x10' "																		'// �������ֹ�
	'sqlStr = sqlStr + " 	and d.isupchebeasong='Y' "
	sqlStr = sqlStr + " 	and d.cancelyn <> 'Y' "
	''sqlStr = sqlStr + " 	and datediff(d,m.baljudate,getdate())>=4 "
	sqlStr = sqlStr + " 	and datediff(d,m.baljudate,'" + CStr(michulgoBaseDate) + "')>=0 "								'// �ٹ��ϼ� ���� D+3 ��
	sqlStr = sqlStr + " 	and not ((T.ipgodate is not null) and (datediff(d, T.ipgodate, getdate()) <= 0)) "				'// ������� ���� �ֹ� ����
	sqlStr = sqlStr + " 	and IsNULL(T.code,'00')<>'05' "																	'// ǰ�����Ұ� ����
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
'�������Ʈ[������,�ٹ�+����] - ������ ����Ʈ : �ֹ��� ���Ӹ�
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
	sqlStr = sqlStr + " 	and m.sitename <> '10x10' "																		'// �������ֹ�
	'sqlStr = sqlStr + " 	and d.isupchebeasong='Y' "
	sqlStr = sqlStr + " 	and d.cancelyn <> 'Y' "
	''sqlStr = sqlStr + " 	and datediff(d,m.baljudate,getdate())>=4 "
	sqlStr = sqlStr + " 	and datediff(d,m.baljudate,'" + CStr(michulgoBaseDate) + "')>=0 "								'// �ٹ��ϼ� ���� D+3 ��
	sqlStr = sqlStr + " 	and not ((T.ipgodate is not null) and (datediff(d, T.ipgodate, getdate()) <= 0)) "				'// ������� ���� �ֹ� ����
	sqlStr = sqlStr + " 	and IsNULL(T.code,'00')<>'05' "																	'// ǰ�����Ұ� ����
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
' �������Ʈ �ٹ�����
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
'CSó������Ʈ ����
''ȯ�� ��ó��(A003), ���ϸ��� ȯ�� ��ó��, ī����ҹ�ó��(A007), �ֹ���ҹ�ó��(A008),
''�������ǻ���, ��ü��ó��, ��üó���Ϸ�, ȸ����û��ó��, Ȯ�ο�û, �ܺθ�ȯ�ҹ�ó��
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
	sqlStr = sqlStr + " and regdate>'"&treemonthbefore&"'"        ''�Ѵ����� �������� �⿡�� ����X =>��ǰ�����Ǹ�.

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
    '// �ٹ��ϼ� ���� D+7 ��
    upReturnMiFinishBaseDate7 = rsget("minusworkday")
end if
rsget.close

tmpSql = " exec [db_cs].[dbo].[usp_getDayMinusWorkday] '" & Left(now(), 10) & "', 3 " & VbCRLF
rsget.CursorLocation = adUseClient
rsget.Open tmpSql, dbget, adOpenForwardOnly
if Not rsget.Eof then
    '// �ٹ��ϼ� ���� D+3 ��
    upReturnMiFinishBaseDate3 = rsget("minusworkday")
end if
rsget.close


'==============================================================================
'��ǰ����(��ü���) D+7 ��ó��
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
'��ǰ����(��ü���) D+7 ��ó��(����ں� ���Ӹ�)
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
		sqlStr = sqlStr + " 	and m.id >= 1200000 "		'// �ӵ�����
	else
		sqlStr = sqlStr + " 	and m.id >= 600000 "		'// �ӵ�����
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
'��ǰ����(��ü���) D+7 ��ó��(�귣�庰 ���Ӹ�)
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
		sqlStr = sqlStr + " 	and m.id >= 1200000 "		'// �ӵ�����
	else
		sqlStr = sqlStr + " 	and m.id >= 600000 "		'// �ӵ�����
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
'��ǰ����(��ü���) D+3 ��ó��(����ں� ���Ӹ�)
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
		sqlStr = sqlStr + " 	and m.id >= 1200000 "		'// �ӵ�����
	else
		sqlStr = sqlStr + " 	and m.id >= 600000 "		'// �ӵ�����
	end if
	sqlStr = sqlStr + " 			and m.deleteyn = 'N' "
	sqlStr = sqlStr + " 			and m.currstate < 'B006' "
	sqlStr = sqlStr + " 			and d.itemid <> 0 "
	sqlStr = sqlStr + " 			and d.isupchebeasong='Y' "
	sqlStr = sqlStr + " 			and IsNull(T.code, '00') = '00' "			'// �������Է¸�
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
'��ǰ����(��ü���) D+3 ��ó��(�귣�庰 ���Ӹ�)
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
		sqlStr = sqlStr + " 	and m.id >= 1200000 "		'// �ӵ�����
	else
		sqlStr = sqlStr + " 	and m.id >= 600000 "		'// �ӵ�����
	end if
	sqlStr = sqlStr + " 	and m.deleteyn = 'N' "
	sqlStr = sqlStr + " 	and m.currstate < 'B006' "
	sqlStr = sqlStr + " 	and d.itemid <> 0 "
	sqlStr = sqlStr + " 	and d.isupchebeasong='Y' "
	sqlStr = sqlStr + " 	and IsNull(T.code, '00') = '00' "			'// �������Է¸�
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
''��üó���Ϸ��.
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
''���ϸ���ȯ�ҹ�ó��
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
'������ ǰ����ҿ�û[�ٹ�+����] : ������ ����Ʈ
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
	sqlStr = sqlStr + " 	and m.sitename <> '10x10' "																'// �������ֹ�
	''sqlStr = sqlStr + " 	and d.isupchebeasong='Y' "																'// �ٹ�+����
	sqlStr = sqlStr + " 	and d.cancelyn <> 'Y' "
	sqlStr = sqlStr + " 	and IsNull(T.code,'00')='05' "															'// ǰ�����Ұ�
	sqlStr = sqlStr + " 	and IsNull(T.state, '0')='0' "															'// ���ȳ� ����
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
'������ ǰ����ҿ�û[�ٹ�+����](�ֹ��� ���Ӹ�)
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
'ǰ����ҿ�û[�ٹ�+����](����ں� ���Ӹ�)
dim csStockoutOrderCountByUserid			'// ����ڴ� �ֹ� ��
dim csStockoutOrderDetailCountByUserid		'// ����ڴ� �ֹ� ������ ��
dim nochargeStockoutOrdercnt				'// ������ �ֹ� ��

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
	sqlStr = sqlStr + " 			and m.sitename = '10x10' "																'// �������ֹ� ����
	''sqlStr = sqlStr + " 			and d.isupchebeasong='Y' "																'// �ٹ�+����
	sqlStr = sqlStr + " 			and d.cancelyn <> 'Y' "
	sqlStr = sqlStr + " 			and IsNull(T.code,'00')='05' "															'// ǰ�����Ұ�
	sqlStr = sqlStr + " 			and IsNull(T.state, '0')='0' "															'// ���ȳ� ����
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
'ǰ����ҿ�û[�ٹ�+����](�ֹ��� ���Ӹ�)
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
' ���̵� Ȯ�λ���
Dim csMainUserID
csMainUserID	= req("csMainUserID", session("ssBctId") )

'��ó���޸�
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
	// alert("�۾���");
	// return;

	if (confirm("���������� �й� �Ͻðڽ��ϱ�?") == true) {
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

// �����ھ��̵�
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
//��������
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
		result = parseInt(v) + "�� ��";
	} else if (v < (60 * 60 * 1000)) {
		v = v / (60 * 1000);
		result = parseInt(v) + "�� ��";
	} else {
		result =  "1�ð� ��";
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
    <!-- ���ʸ޴� ���� -->
	<td width="33%" valign="top">
	    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
        <tr valign="top">
            <td>
        	    <!-- �ֹ������˻� -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("menubar") %>">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
                        <tr height="25">
            			    <td>
            			    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>�ֹ����� �˻�</b>
            			    </td>
            			    <td align="right">
            			        <a href="javascript:PopOrderMasterWithCallRing();"> �ֹ��������� <img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a>
            			    </td>
            			</tr>
            	        </table>
            	    </td>
            	</tr>
            	</table>
        	    <!-- �ֹ������˻� -->
        	</td>
        </tr>
        <tr valign="top">
            <td height="10"></td>
        </tr>
        <tr valign="top">
            <td>
        	    <!-- �Խ��� ���� ����-->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("menubar") %>">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
                        <tr height="25">
            			    <td style="border-bottom:1px solid #BABABA">
            			    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>1:1 ���Խ��� ����</b>
            			    	(<span id="objOneToOneBoard"></span>) <a href="javascript:RefreshData('csTimeOneToOneBoard')"><img src="/images/icon_reload.gif" border="0"></a>
            			    </td>
            			    <td align="right" style="border-bottom:1px solid #BABABA">
            			        <a href="javascript:PopMyQnaList('', '', '');">
        				        �ٷΰ���
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td>����ں� ��ó�� ����</td>
            			    <td align="right"></td>
            			</tr>
            			<%
            			'// ù��°�� ��ŵ�Ѵ�.(��������)
            			for i = 1 to UBound(csBoardUserArr)
            			%>
            			<tr height="25">
            			    <td>&nbsp;&nbsp;- <%= csBoardUserArr(i) %></td>
            			    <td align="right">
            			        <a href="javascript:PopMyQnaListByChargeId('<%= csBoardUserArr(i) %>', 'N');">
        				        <b><%= csBoardChargecntArr(i) %></b> ��
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
            			</tr>
            			<% next %>
            			<tr height="25">
            			    <td>&nbsp;&nbsp;- �̺й� ����</td>
            			    <td align="right">
            			        <a href="javascript:PopMyQnaListByChargeId('all', 'N');">
        				        <b><%= csBoardNochargecnt %></b> ��
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
            			</tr>
            	        </table>
            	    </td>
            	</tr>
            	</table>
        	    <!-- �Խ��� ���� ��-->
        	</td>
        </tr>
        <tr valign="top">
            <td height="10"></td>
        </tr>
        <tr valign="top">
            <td width="50%">
        	    <!-- ��ü��� ���� ���� -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("menubar") %>">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
                        <tr height="25">
            			    <td style="border-bottom:1px solid #BABABA">
            			    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>�������Ʈ[��ü���]</b>
								(<span id="objMichulgoListUpche"></span>) <a href="javascript:RefreshData('csTimeMichulgoListUpche')"><img src="/images/icon_reload.gif" border="0"></a>
            			    </td>
            			    <td align="right" style="border-bottom:1px solid #BABABA">
            			        <a href="javascript:upchebeasongmaster();">
        				        �ٷΰ���
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td>2�� �̻� ��Ȯ�ΰ�</td>
            			    <td align="right"><b><%=FormatNumber(arrMiSend(0),0)%></b> ��<a href="javascript:popUpcheMisend('','2','2');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<tr height="25">
            			    <td>3�� �̻� �̹߼۰�</td>
            			    <td align="right"><b><%=FormatNumber(arrMiSend(1),0)%></b> ��<a href="javascript:popUpcheMisend('','3','3');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<tr height="25">
            			    <td>����ں� D+3 �̹߼۰�(������,����� ����)</td>
            			    <td align="right"></td>
            			</tr>
            			<%
            			'// ù��°�� ��ŵ�Ѵ�.(��������)
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
		        				        <%= csMichulgoBrandcntArr(j) %> ��
		        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
		        				        </a>
		            			    </td>
		            			</tr>
	            				<% end if %>
	            			<% next %>
            			<% next %>
            			<tr height="25">
            			    <td>&nbsp;&nbsp;- �̺й� �̹߼۰�</td>
            			    <td align="right">
        				        <b><%= nochargeMichulgoBrandcnt %></b>
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td>ǰ����ҿ�û��</td>
            			    <td align="right"><b><%=FormatNumber(arrMiSend(2),0)%></b> ��<a href="javascript:popUpcheMisend('05','','');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<!--
            			<tr height="25">
            			    <td colspan="2"></td>
            			 	too Many
            			    <td>�������� ���Է°� (D+2�̻�)</td>
            			    <td align="right"><b>??</b> ��<img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></td>
            			</tr>
            			-->

            	        </table>
            	    </td>
            	</tr>
            	</table>
        	    <!-- ��ü��� ���� ��-->
        	</td>
		</tr>


        <tr valign="top">
            <td height="10"></td>
        </tr>

        <tr valign="top">
            <td width="50%">
        	    <!-- ��ü��� ���� ���� -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("menubar") %>">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
                        <tr height="25">
            			    <td style="border-bottom:1px solid #BABABA">
            			    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>������ �������Ʈ[�ٹ�+����]</b>
								(<span id="objIpjumMichulgo"></span>) <a href="javascript:RefreshData('csTimeIpjumMichulgo')"><img src="/images/icon_reload.gif" border="0"></a>
            			    </td>
            			    <td align="right" style="border-bottom:1px solid #BABABA">
            			    	&nbsp;
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td>��������(����� ����)</td>
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
		        				        <%= csMichulgoOtherSiteOrderdetailCntcntArr(j) %> ��
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
        	    <!-- ��ü��� ���� ��-->
        	</td>
        </tr>

        <tr valign="top">
            <td height="10"></td>
        </tr>
        <tr valign="top">
            <td width="50%">
        	    <!-- �ٹ����� ������ ����-->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("menubar") %>">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
                        <tr height="25">
            			    <td style="border-bottom:1px solid #BABABA">
            			    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>�������Ʈ[�ٹ�����]</b>
								(<span id="objMichulgoListTenTen"></span>) <a href="javascript:RefreshData('csTimeMichulgoListTenTen')"><img src="/images/icon_reload.gif" border="0"></a>
            			    </td>
            			    <td align="right" style="border-bottom:1px solid #BABABA">
                    			<a href="javascript:misendmaster('Y');">
        				        �ٷΰ���
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td>���� ���Է°�</td>
            			    <td align="right"><a href="javascript:misendmaster('N');"><b><%=FormatNumber(arrTenMiSend(0),0)%></b> ��<img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<tr height="25">
            			    <td>�� �̾ȳ���</td>
            			    <td align="right"><a href="javascript:misendmaster('N');"><b><%=FormatNumber(arrTenMiSend(1),0)%></b> ��<img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<tr height="25">
            			    <td>SMS/Mail/��ȭ�Ϸ��</td>
            			    <td align="right"><a href="javascript:misendmaster('4');"><b><%=FormatNumber(arrTenMiSend(2),0)%></b> ��<img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<tr height="25">
            			    <td>CSó���Ϸ�(������ó����û)��</td>
            			    <td align="right"><a href="javascript:misendmaster('6');"><b><%=FormatNumber(arrTenMiSend(3),0)%></b> ��<img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            	        </table>
            	    </td>
            	</tr>
            	</table>
        	    <!--  csó����û ��-->
        	</td>
        </tr>

        </table>
    </td>
    <!-- ���ʸ޴� �� -->

    <td width="10"></td>

    <!-- ����޴� ���� -->
    <td width="33%" valign="top">
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">

        <tr valign="top">
        	<td>
        	    <!-- CSó������Ʈ ����-->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("menubar") %>">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
                        <tr height="25">
            			    <td style="border-bottom:1px solid #BABABA">
            			    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>CSó������Ʈ ����</b>
            			    	(<span id="objCSList"></span>) <a href="javascript:RefreshData('csTimeCSList')"><img src="/images/icon_reload.gif" border="0"></a>
            			    </td>
            			    <td align="right" style="border-bottom:1px solid #BABABA">
            			        <a href="javascript:Cscenter_Action_List2('');">
        				        �ٷΰ���
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td>ȯ�ҹ�ó��(����)</td>
            			    <td align="right"><b><%= csRefundRequestRegCount %></b> ��<a href="javascript:Cscenter_Action_List2('norefund');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<tr height="25">
            			    <td>ȯ�ҹ�ó��(Ȯ�ο�û)</td>
            			    <td align="right"><b><%= csRefundRequestConfirmCount %></b> ��<a href="javascript:Cscenter_Action_List2('confirm');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<tr height="25">
            			    <td>���ϸ���/��ġ�� ȯ�ҹ�ó��</td>
            			    <td align="right"><b><%= csNotFinMileRefund %></b> ��<a href="javascript:Cscenter_Action_List2('norefundmile');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<tr height="25">
            			    <td>ī����� ��ó��</td>
            			    <td align="right"><b><%= csRequestCardCancelCount %></b> ��<a href="javascript:Cscenter_Action_List2('cardnocheck');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<tr height="25">
            			    <td>�ֹ���� ��ó��</td>
            			    <td align="right"><b><%= csNotFinA008 %></b> ��<a href="javascript:Cscenter_Action_List2('cancelnofinish');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<tr height="25">
            			    <td>ȸ����û ��ó��</td>
            			    <td align="right"><b><%= csReturnNotFinish %></b> ��<a href="javascript:Cscenter_Action_List2('returnmifinish');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<!--
            			<tr height="25">
            			    <td>��ü ��ó��</td>
            			    <td align="right"><b><%= csUpcheNotFin %></b> ��<a href="javascript:Cscenter_Action_List2('upchemifinish');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			-->
            			<tr height="25">
            			    <td>��ǰ����(��ü���) D+7 ��ó��</td>
            			    <td align="right"><b><%= CSUpcheReturnNotFinish7 %></b> ��<a href="javascript:Cscenter_Action_MiFinishReturnListDPlus(7);"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>

            			<tr height="25">
            			    <td>��üó���Ϸ�</td>
            			    <td align="right"><b><%= csUpcheFinished %></b> ��<a href="javascript:Cscenter_Action_List2('upchefinish');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<!--
            			<tr height="25">
            			    <td>�ܺθ�ȯ�� ��ó��</td>
            			    <td align="right"><b><%= csNotFinA005 %></b> ��<a href="javascript:Cscenter_Action_List2('norefundetc');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			-->
            	        </table>
            	    </td>
            	</tr>
            	</table>
        	    <!--  CSó������Ʈ ��-->
        	</td>
        </tr>

        <tr valign="top">
            <td height="10"></td>
        </tr>

        <tr valign="top">
            <td width="50%">
        	    <!-- ��ǰ����(��ü���) D+3 ��ó�� ���� -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("menubar") %>">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
                        <tr height="25">
            			    <td style="border-bottom:1px solid #BABABA">
            			    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>��ǰ����(����) D+3 ��ó��</b>
            			    	(<span id="objUpcheReturn3"></span>) <a href="javascript:RefreshData('csTimeUpcheReturn')"><img src="/images/icon_reload.gif" border="0"></a>
            			    </td>
            			    <td align="right" style="border-bottom:1px solid #BABABA">
            			        <a href="javascript:Cscenter_Action_MiFinishReturnListDPlus(3);">
        				        �ٷΰ���
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td>��ǰ����(��ü���) D+3 ��ó��</td>
            			    <td align="right"><b><%= CSUpcheReturnNotFinish3 %></b> ��<a href="javascript:Cscenter_Action_MiFinishReturnListDPlus(3);"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<tr height="25">
            			    <td>����ں� D+3 ��ó����(�����Է� ����)</td>
            			    <td align="right"></td>
            			</tr>
            			<%
            			'// ù��°�� ��ŵ�Ѵ�.(��������)
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
		        				        <%= csReturn3BrandcntArr(j) %> ��
		        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
		        				        </a>
		            			    </td>
		            			</tr>
	            				<% end if %>
	            			<% next %>
            			<% next %>
            			<tr height="25">
            			    <td>&nbsp;&nbsp;- �̺й� ��ó����</td>
            			    <td align="right">
        				        <b><%= nochargeReturn3Brandcnt %></b>
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
            			    </td>
            			</tr>
            	        </table>
            	    </td>
            	</tr>
            	</table>
        	    <!-- ��ǰ����(��ü���) D+3 ��ó�� ��-->
        	</td>
        </tr>

        <tr valign="top">
            <td height="10"></td>
        </tr>

        <tr valign="top">
            <td width="50%">
        	    <!-- ��ǰ����(��ü���) D+7 ��ó�� ���� -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("menubar") %>">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
                        <tr height="25">
            			    <td style="border-bottom:1px solid #BABABA">
            			    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>��ǰ����(����) D+7 ��ó��</b>
            			    	(<span id="objUpcheReturn7"></span>) <a href="javascript:RefreshData('csTimeUpcheReturn')"><img src="/images/icon_reload.gif" border="0"></a>
            			    </td>
            			    <td align="right" style="border-bottom:1px solid #BABABA">
            			        <a href="javascript:Cscenter_Action_MiFinishReturnListDPlus(7);">
        				        �ٷΰ���
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td>��ǰ����(��ü���) D+7 ��ó��</td>
            			    <td align="right"><b><%= CSUpcheReturnNotFinish7 %></b> ��<a href="javascript:Cscenter_Action_MiFinishReturnListDPlus(7);"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<tr height="25">
            			    <td>����ں� D+7 ��ó����(��ǰ���� ����)</td>
            			    <td align="right"></td>
            			</tr>
            			<%
            			'// ù��°�� ��ŵ�Ѵ�.(��������)
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
		        				        <%= csReturn7BrandcntArr(j) %> ��
		        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
		        				        </a>
		            			    </td>
		            			</tr>
	            				<% end if %>
	            			<% next %>
            			<% next %>
            			<tr height="25">
            			    <td>&nbsp;&nbsp;- �̺й� ��ó����</td>
            			    <td align="right">
        				        <b><%= nochargeReturn7Brandcnt %></b>
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
            			    </td>
            			</tr>
            	        </table>
            	    </td>
            	</tr>
            	</table>
        	    <!-- ��ǰ����(��ü���) D+7 ��ó�� ��-->
        	</td>
        </tr>

        <tr valign="top">
            <td height="10"></td>
        </tr>

        <tr valign="top">
            <td width="50%">
        	    <!-- ��ü��� ���� ���� -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("menubar") %>">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
                        <tr height="25">
            			    <td style="border-bottom:1px solid #BABABA">
            			    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>������ ǰ����ҿ�û��</b>
								(<span id="objIpjumStockOut"></span>) <a href="javascript:RefreshData('csTimeIpjumStockOut')"><img src="/images/icon_reload.gif" border="0"></a>
            			    </td>
            			    <td align="right" style="border-bottom:1px solid #BABABA">
            			    	&nbsp;
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td>��������(�� �ȳ�����)</td>
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
		        				        <%= csStockoutOtherSiteOrderdetailCntcntArr(j) %> ��
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
        	    <!-- ��ü��� ���� ��-->
        	</td>
        </tr>
        <tr valign="top">
            <td height="10"></td>
        </tr>
        <tr valign="top">
        	<td>
        	    <!-- ��ü��� ���� ���� -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("menubar") %>">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
                        <tr height="25">
            			    <td style="border-bottom:1px solid #BABABA">
            			    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>ǰ����ҿ�û��[�ٹ�+����]</b>
								(<span id="objTenTenStockOut"></span>) <a href="javascript:RefreshData('csTimeTenTenStockOut')"><img src="/images/icon_reload.gif" border="0"></a>
            			    </td>
            			    <td align="right" style="border-bottom:1px solid #BABABA">
            			    	&nbsp;
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td>����ں�(����������, �� �ȳ�����)</td>
            			    <td align="right">
            			    </td>
            			</tr>
            			<%
            			'// ù��°�� ��ŵ�Ѵ�.(��������)
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
		        				        <%= csStockoutOrderdetailCntcntArr(j) %> ��
		        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
		        				        </a>
		            			    </td>
		            			</tr>
	            				<% end if %>
	            			<% next %>
            			<% next %>
            			<tr height="25">
            			    <td>&nbsp;&nbsp;- �̺й�</td>
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
	            			        <input type="button" class="button" value="�й�" onclick="setUpcheStockout10x10ByOrderserial(frm<%= j %>)">
	            			    </td>
	            			</tr>
	            			</form>
            				<% end if %>
            			<% next %>
            	        </table>
            	    </td>
            	</tr>
            	</table>
        	    <!-- ��ü��� ���� ��-->
        	</td>
        </tr>

      	</table>
    </td>
    <!-- ����޴� �� -->

    <td width="10"></td>

    <!-- �����ʸ޴� ���� -->
    <td valign="top">
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
        <tr valign="top">
            <td>
                <!-- ���ΰ�ħ ���� -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("tabletop") %>">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
                        <tr height="25">
                        	<td>
            			    	<img src="/images/icon_star.gif" align="absbottom">
								<b>ID : </b>
								<input type="text" class="text" id="csMainUserID" value="<%=csMainUserID%>" size="10">
								<input type="button" class="button" value="�˻�" onclick="location.href = 'cscenter_main.asp?menupos=757&csMainUserID=' + document.getElementById('csMainUserID').value;">
								<!-- �ʱ�α��ν� �α��� ���̵�� ���� / �ٸ����̵�ε� �˻������ϵ��� -->
            			    </td>
            			    <td align="right">
            			    	<a href="javascript:document.location.reload();">
        				        ���ΰ�ħ
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
            			</tr>
            	        </table>
            	    </td>
            	</tr>
            	</table>
            	<!-- ���ΰ�ħ �� -->
            </td>
        </tr>

        <tr valign="top">
            <td height="10"></td>
        </tr>

        <tr valign="top">
            <td>
                <!-- ���̵� Ȯ�λ��� -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("tabletop") %>">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
            			<tr height="25">
            			    <td style="border-bottom:1px solid #BABABA">
            			        <img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>���̵� Ȯ�λ���</b>
            			    </td>
            			    <td align="right" style="border-bottom:1px solid #BABABA">
            			        &nbsp;
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td>��ó���޸�</td>
            			    <td align="right">
            			        <b><%= CSMemoNotFinish %></b> ��
        				    	<a href="javascript:cscenter_memo_list('','','N', '<%=csMainUserID%>');">
                    		    <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
                    		    </a>
            			    </td>
            			</tr>
            	        </table>
            	    </td>
            	</tr>
            	</table>
        	    <!--  ���̵� Ȯ�λ��� ��-->
            </td>
        </tr>

        <tr valign="top">
            <td height="10"></td>
        </tr>
        <%
        	Dim NewCoop
        	Set NewCoop = new CCooperate
        	NewCoop.FDoc_Id = session("ssBctId")
        	NewCoop.fnGetCooperateCount			' ���� �ø����� Ȯ�� �� �ּ� ����


        %>
        <tr valign="top">
            <td>
                <!-- ������ ���� -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("tabletop") %>">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
            			<tr height="25">
            			    <td>���� ������ (��ó��)</td>
            			    <td align="right">
            			        <a href="javascript:popCooperate();">
            			        <%
            			        	If NewCoop.FComeCnt = 0 Then
            			        		Response.Write "[" & NewCoop.FComeCnt & "] ��"
            			        	Else
            			        		Response.Write "[<b>" & NewCoop.FComeCnt & "</b>] ��...<img src='http://fiximage.10x10.co.kr/web2009/main/news_icon_new.gif' border='0'>"
            			        	End If
            			        %>
            			        </a>
        				    	<a href="javascript:popCooperate();">
                    		    <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
                    		    </a>
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td>���� ������ (��ó��)</td>
            			    <td align="right">
            			        <a href="javascript:popCooperate();">
            			        <%
            			        	If NewCoop.FSendCnt = 0 Then
            			        		Response.Write "[" & NewCoop.FSendCnt & "] ��"
            			        	Else
            			        		Response.Write "[<b>" & NewCoop.FSendCnt & "</b>] ��...<img src='http://fiximage.10x10.co.kr/web2009/main/news_icon_new.gif' border='0'>"
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
        	    <!--  ������ ��-->
            </td>
        </tr>


        <tr valign="top">
            <td height="10"></td>
        </tr>

        <tr valign="top">
            <td>
        		<!-- ���� ���� ���� ���� -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("tabletop") %>">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
            			<tr height="25">
            			    <td style="border-bottom:1px solid #BABABA">
            			        <img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>��ȯ��ó��</b>
            			    </td>
            			    <td align="right" style="border-bottom:1px solid #BABABA">
            			        <a href="javascript:cscenter_refund_list();">
        				        �ٷΰ���
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td style="border-bottom:1px solid #BABABA">
            			        <img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>���ݿ����� ����</b>
            			    </td>
            			    <td align="right" style="border-bottom:1px solid #BABABA">
            			        <a href="javascript:cscenter_cashreceipt_list();">
        				        �ٷΰ���
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
				                </a>
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td style="border-bottom:1px solid #BABABA">
            			        <img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>���ݰ�꼭 ����</b>
            			    </td>
            			    <td align="right" style="border-bottom:1px solid #BABABA">
            			        <a href="javascript:cscenter_tax_list();">
        				        �ٷΰ���
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td style="border-bottom:1px solid #BABABA">
            			        <img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>�������Ա� ����</b>
            			    </td>
            			    <td align="right" style="border-bottom:1px solid #BABABA">
            			        <a href="javascript:cscenter_payment_list(1);">
        				        �ٷΰ���
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
            			</tr>
            	        </table>
            	    </td>
            	</tr>
            	</table>
        	    <!--  ���� ���� ���� ��-->
        	</td>
        </tr>

        <tr valign="top">
            <td height="10"></td>
        </tr>

		<tr valign="top">
            <td>
            	<!-- ������ȸ -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("tabletop") %>">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
                        <tr height="25">
            			    <td>
            			    	<table width="100%" style="border:1px solid #BABABA" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#FFFFFF">
		                        <tr height="20">
		            			    <td align="center">
		            			    	<a href="javascript:cscenter_mileage_list('');">���ϸ�����ȸ</a>
		            			    </td>
		            			</tr>
		            	        </table>
            			    </td>
            			    <td width="5"></td>
            			    <td>
            			    	<table width="100%" style="border:1px solid #BABABA" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#FFFFFF">
		                        <tr height="20">
		            			    <td align="center">
		            			    	<a href="javascript:cscenter_deposit_list('');">��ġ����ȸ</a>
		            			    </td>
		            			</tr>
		            	        </table>
            			    </td>
            			    <td width="5"></td>
            			    <td>
            			    	<table width="100%" style="border:1px solid #BABABA" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#FFFFFF">
		                        <tr height="20">
		            			    <td align="center">
		            			    	<a href="javascript:cscenter_coupon_list('');">������ȸ</a>
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
		            			    	<a href="javascript:cscenter_eventjoin_list('');">�����̺�Ʈ��ȸ</a>
		            			    </td>
		            			</tr>
		            	        </table>
            			    </td>
            			    <td width="5"></td>
            			    <td>
            			    	<table width="100%" style="border:1px solid #BABABA" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#FFFFFF">
		                        <tr height="20">
		            			    <td align="center">
		            			    	<a href="javascript:cscenter_member_list('');">ȸ����ȸ</a>
		            			    </td>
		            			</tr>
		            	        </table>
            			    </td>
            			    <td width="5"></td>
            			    <td>
            			    	<table width="100%" style="border:1px solid #BABABA" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#FFFFFF">
		                        <tr height="20">
		            			    <td align="center">
		            			    	<a href="javascript:cscenter_eventprize_list('');">��÷��ȸ</a>
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
        	    <!-- ������ȸ -->
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
		            			    	<a href="javascript:PopCSSMSSend('','','','');">SMS�߼�</a>
		            			    </td>
		            			</tr>
		            	        </table>
            			    </td>
            			    <td width="5"></td>
            			    <td>
            			    	<table width="100%" style="border:1px solid #BABABA" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#FFFFFF">
		                        <tr height="20">
		            			    <td align="center">
		            			    	<a href="javascript:PopCSMailSend('','');">���Ϲ߼�</a>
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
    <!-- �����ʸ޴� �� -->

</tr>
</table>

<% Call checkAndWriteElapsedTime("017") %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
