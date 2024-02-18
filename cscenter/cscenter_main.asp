<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : cs����
' History : 2009.04.17 �̻� ����
'			2016.03.25 �ѿ�� ����
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
'// ����� ���
If Trim(application("csTimeUserList")) = "" Or Trim(application("csVipBoardUserArr")) = "" Or Trim(application("csBoardUserArr")) = "" Or DateDiff("s", application("csTimeUserList"), Now() ) > 1800 Then						'// 30��(1800��) �ʰ��� ���� ������
	application("csTimeUserList") = Now()
	IsUpdateUserListNeed = True
end if

'// 1:1 ���Խ���
If Trim(application("csTimeOneToOneBoard")) = "" Or DateDiff("s", application("csTimeOneToOneBoard"), Now() ) > 600 Then			'// 10��(600��) �ʰ��� ���� ������
	application("csTimeOneToOneBoard") = Now()
	IsUpdateOneToOneBoardNeed = True
end if

'// �������Ʈ[��ü���]
If Trim(application("csTimeMichulgoListUpcheNew")) = "" Or Trim(application("csMiSend")) = "" Or DateDiff("s", application("csTimeMichulgoListUpcheNew"), Now() ) > 1800 Then	'// 30��(1800��) �ʰ��� ���� ������
	application("csTimeMichulgoListUpcheNew") = Now()
	IsUpdateMichulgoListUpcheNeed = True
end if

'// ��ǰ���� �亯����
If Trim(application("csTimeMiAnswerList")) = "" Or DateDiff("s", application("csTimeMiAnswerList"), Now() ) > 7200 Then	'// 120��(7200��) �ʰ��� ���� ������
	application("csTimeMiAnswerList") = Now()
	IsUpdateMiAnswerUpcheNeed = True
end if

'// �������Ʈ[�ٹ�����]
If Trim(application("csTimeMichulgoListTenTen")) = "" Or Trim(application("csTenMiSend")) = "" Or DateDiff("s", application("csTimeMichulgoListTenTen"), Now() ) > 900 Then	'// 15��(900��) �ʰ��� ���� ������
	application("csTimeMichulgoListTenTen") = Now()
	IsUpdateMichulgoListTenTenNeed = True
end if

'// �������Ʈ[������]
If Trim(application("csTimeIpjumMichulgo")) = "" Or Trim(application("csMiSend")) = "" Or DateDiff("s", application("csTimeIpjumMichulgo"), Now() ) > 1800 Then	'// 30��(1800��) �ʰ��� ���� ������
	application("csTimeIpjumMichulgo") = Now()
	IsUpdateIpjumMichulgoNeed = True
end if

'// CSó������Ʈ
If Trim(application("csTimeCSList")) = "" Or DateDiff("s", application("csTimeCSList"), Now() ) > 900 Then	'// 15��(900��) �ʰ��� ���� ������
	application("csTimeCSList") = Now()
	IsUpdateCSListNeed = True

	application("csTimeUpcheReturn") = Now()
	IsUpdateUpcheReturnListNeed = True
end if

'// ��ǰ����(��ü���)
If Trim(application("csTimeUpcheReturn")) = "" Or DateDiff("s", application("csTimeUpcheReturn"), Now() ) > 1800 Then	'// 30��(1800��) �ʰ��� ���� ������
	application("csTimeCSList") = Now()
	IsUpdateCSListNeed = True

	application("csTimeUpcheReturn") = Now()
	IsUpdateUpcheReturnListNeed = True
end if

'// ������ ǰ����ҿ�û��
If Trim(application("csTimeIpjumStockOut")) = "" Or DateDiff("s", application("csTimeIpjumStockOut"), Now() ) > 900 Then	'// 15��(900��) �ʰ��� ���� ������
	application("csTimeIpjumStockOut") = Now()
	IsUpdateIpjumStockOutNeed = True
end if

'// ������ ȯ�ҹ�ó����
If Trim(application("csTimeIpjumRefund")) = "" Or DateDiff("n", application("csTimeIpjumRefund"), Now() ) > 120 Then	'// 120�� �ʰ��� ���� ������
	application("csTimeIpjumRefund") = Now()
	IsUpdateIpjumRefundNeed = True
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
dim csVipBoardUserArr, csVVipBoardUserArr, csBoardUserArr, csMichulgoStockoutBoardUserArr, csReturnBoardUserArr

if (IsUpdateUserListNeed = True) then
	'// YN �� N �� �ƴѰ��̾�� �Ѵ�.
	'// �й��� ���� YN �� Y �ΰ�, �й������ ǥ���Ҷ��� N �� �ƴѰ�!!
	'// YN = T �� ��� : �й�����(�й�������� ǥ�õǰ�, ���̻� �й���� �ʴ´�.)
	sqlStr = " exec [db_cs].[dbo].[usp_Ten_Cs_ChargeUserList] "
	''response.write sqlStr
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

	'// ��������(����ڰ� �Ѹ� ������ ������ �߻��Ѵ�.)
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
'// 1:1 ���Խ��� ����(����ں�)
dim csVipBoardChargecntArr, csVipBoardNochargecnt, csVVipBoardChargecntArr, csVVipBoardNochargecnt
dim csBoardChargecntArr, csBoardNochargecnt, csrecommendBoardcnt, csstaffBoardcnt, csStaffStockoutOrderCount, csStaffStockoutOrderArr
dim csExtSiteCntArr

if (IsUpdateOneToOneBoardNeed = True) then

	'// ���/����ھ��̵�/ī��Ʈ
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
	sqlStr = sqlStr & "	and isnull(qadiv,'')<>'26' "	'�����ǻ��� ���Ǵ� ����
	sqlStr = sqlStr & "	and userlevel not in (7)"	'��������
	sqlStr = sqlStr & "	and IsNull(sitename, '10x10') = '10x10'"	'���޸�CS ����
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

			'// �Ϲݰ�
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

	'//�����ǻ��׹��� ī��Ʈ		'/2016.03.28 �ѿ�� �߰�
	sqlStr = " select count(*) as cnt "
	sqlStr = sqlStr + "	from [db_cs].[dbo].tbl_myqna "
	sqlStr = sqlStr + "	where isusing = 'Y' "
	sqlStr = sqlStr + "	and replydate is NULL "
	sqlStr = sqlStr + "	and isnull(qadiv,'')='26' "	'�����ǻ��� ����
	sqlStr = sqlStr + "	and userlevel not in (7)"	'��������
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

	'//�������ǻ���  ī��Ʈ		'/2016.04.15 �ѿ�� �߰�
	sqlStr = " select count(*) as cnt "
	sqlStr = sqlStr + "	from [db_cs].[dbo].tbl_myqna "
	sqlStr = sqlStr + "	where isusing = 'Y' "
	sqlStr = sqlStr + "	and replydate is NULL "
	sqlStr = sqlStr + "	and userlevel in (7)"	'��������
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

	'// ���޸��� �亯���� �Ǽ�
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

	'### STAFF ǰ����ҿ�û��
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
' �������Ʈ[��ü���] - ���
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
'��ǰ��ó��[��ü���,����������] - �ٹ��ϼ� ���� D+4 ��
dim upcheReturnBaseDate

tmpSql = " exec [db_cs].[dbo].[usp_getDayMinusWorkday] '" & Left(now(), 10) & "', 4 " & VbCRLF
rsget.CursorLocation = adUseClient
rsget.Open tmpSql, dbget, adOpenForwardOnly
if Not rsget.Eof then
    '// �ٹ��ϼ� ���� D+4 ��
    upcheReturnBaseDate = rsget("minusworkday")
end if
rsget.close

'==============================================================================
'// ��ǰ���� �亯����
'// -- �ٹ� : 1���� �̳� ��ü
'// -- ���� : D+3 �� �ʰ�, 1���� �̳� ��ü

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
'// ��ǰ���� �亯����
'// -- �ٹ� : 1���� �̳� ��ü
'// -- ���� : D+3 �� �ʰ�, 1���� �̳� ��ü

if (IsUpdateMiAnswerUpcheNeed = True) then
	'// ��ǰ���� �亯����[�ٹ�]
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

	'// ��ǰ���� �亯����[����]
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
'�������Ʈ[��ü���]
dim tmpcsMichulgoList, tmpcsMichulgoListByBrand
dim csMichulgoList, csMichulgoListByBrand
dim csMichulgoItem, csMichulgoItemByBrand
dim nochargeMichulgoBrandcnt

if (IsUpdateMichulgoListUpcheNeed = True) then
	tmpcsMichulgoList = ""
	tmpcsMichulgoListByBrand = ""

	'// ����ں� ���Ӹ�
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

	'// �귣�庰 ���Ӹ�
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

'// �̺й� �귣��
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
'��ǰ ��ó��[��ü���,����������]
dim tmpcsUpcheReturnList, tmpcsUpcheReturnListByBrand
dim csUpcheReturnList, csUpcheReturnListByBrand
dim csUpcheReturnItem, csUpcheReturnItemByBrand
dim nochargeUpcheReturnBrandcnt

if (IsUpdateUpcheReturnListNeed = True) then

	tmpcsUpcheReturnList = ""
	tmpcsUpcheReturnListByBrand = ""

	'// ����ں� ���Ӹ�
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

	'// �귣�庰 ���Ӹ�
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

'// �̺й� �귣��
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
'�������Ʈ[������,����] - �귣�庰 ���Ӹ�
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
	sqlStr = sqlStr + " 	and m.sitename <> '10x10' "																		'// �������ֹ�
	sqlStr = sqlStr + " 	and d.cancelyn <> 'Y' "
	''sqlStr = sqlStr + " 	and datediff(d,m.baljudate,getdate())>=4 "
	sqlStr = sqlStr + " 	and datediff(d,m.baljudate,'" + CStr(michulgoBaseDate) + "')>=0 "								'// �ٹ��ϼ� ���� D+3 ��
	sqlStr = sqlStr + " 	and not ((T.ipgodate is not null) and (datediff(d, T.ipgodate, getdate()) <= 0)) "				'// ������� ���� �ֹ� ����
	sqlStr = sqlStr + " 	and IsNULL(T.code,'00') not in ('05','06')"																	'// ǰ�����Ұ�/�ù��ľ� ����
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
'�������Ʈ[������,����] - ������ ����Ʈ : �ֹ��� ���Ӹ�
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
	sqlStr = sqlStr + " 	and m.sitename <> '10x10' "																		'// �������ֹ�
	sqlStr = sqlStr + " 	and d.isupchebeasong='Y' "
	sqlStr = sqlStr + " 	and d.cancelyn <> 'Y' "
	''sqlStr = sqlStr + " 	and datediff(d,m.baljudate,getdate())>=4 "
	sqlStr = sqlStr + " 	and datediff(d,m.baljudate,'" + CStr(michulgoBaseDate) + "')>=0 "								'// �ٹ��ϼ� ���� D+3 ��
	sqlStr = sqlStr + " 	and not ((T.ipgodate is not null) and (datediff(d, T.ipgodate, getdate()) <= 0)) "				'// ������� ���� �ֹ� ����
	sqlStr = sqlStr + " 	and IsNULL(T.code,'00') not in ('05','06') "																	'// ǰ�����Ұ�/�ù��ľ� ����
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
' �������Ʈ �ٹ�����
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
'CSó������Ʈ ����
''ȯ�� ��ó��(A003), ���ϸ��� ȯ�� ��ó��, ī����ҹ�ó��(A007), �ֹ���ҹ�ó��(A008),
''�������ǻ���, ��ü��ó��, ��üó���Ϸ�, ȸ����û��ó��, Ȯ�ο�û, �ܺθ�ȯ�ҹ�ó��
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
	sqlStr = sqlStr + " and regdate>'"&treemonthbefore&"'"        				''�Ѵ����� �������� �⿡�� ����X =>��ǰ�����Ǹ�.
	sqlStr = sqlStr + " and id > " & (maxCSMasterIdx - 200000) & " "			'// �ֱ� 20������ �˻��Ѵ�.(�ӵ�����)

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
    '// �ٹ��ϼ� ���� D+7 ��
    upReturnMiFinishBaseDate7 = rsget("minusworkday")
end if
rsget.close

'' tmpSql = " exec [db_cs].[dbo].[usp_getDayMinusWorkday] '" & Left(now(), 10) & "', 3 " & VbCRLF
'' rsget.CursorLocation = adUseClient
'' rsget.Open tmpSql, dbget, adOpenForwardOnly
'' if Not rsget.Eof then
''     '// �ٹ��ϼ� ���� D+3 ��
''     upReturnMiFinishBaseDate3 = rsget("minusworkday")
'' end if
'' rsget.close

'==============================================================================
'��ǰ����(��ü���) D+4 ��ó��
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
'��ǰ����(��ü���, �������ֹ�) D+7 ��ó��(�귣�庰 ���Ӹ�)
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
		sqlStr = sqlStr + " 	and m.id >= 1200000 "		'// �ӵ�����
	else
		sqlStr = sqlStr + " 	and m.id >= 600000 "		'// �ӵ�����
	end if
	sqlStr = sqlStr + " 	and m.deleteyn = 'N' "
	sqlStr = sqlStr + " 	and m.currstate < 'B006' "
	sqlStr = sqlStr + " 	and d.itemid <> 0 "
	sqlStr = sqlStr + " 	and d.isupchebeasong='Y' "
	sqlStr = sqlStr + " 	and m.extsitename <> '10x10' "																		'// �������ֹ�
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
''��üó���Ϸ��, ����ó���Ϸ��
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
''���ϸ���/��ġ�� ȯ�ҹ�ó��
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
	sqlStr = sqlStr + " 	and m.sitename <> '10x10' "																'// �������ֹ�
	''sqlStr = sqlStr + " 	and d.isupchebeasong='Y' "																'// �ٹ�+����
	sqlStr = sqlStr + " 	and d.cancelyn <> 'Y' "
	sqlStr = sqlStr + " 	and IsNull(T.code,'00') in ('05','06') "															'// ǰ�����Ұ�/�ù��ľ�
	sqlStr = sqlStr + " 	and IsNull(T.state, '0') in ('0', '4') "												'// ���ȳ� ����+���ȳ�
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
'������ ȯ�ҹ�ó�� : ������ ����Ʈ
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
	sqlStr = sqlStr + " 	and m.sitename <> '10x10' "							'// �������ֹ�
	sqlStr = sqlStr + " 	and a.divcd = 'A005' "								'// ������ȯ��
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
'������ ȯ�ҹ�ó��(�ֹ��� ���Ӹ�)
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
	'sqlStr = sqlStr + " 	and a.regdate >= DateAdd(m, -2, getdate()) "	' cs�ϼҶ��û(�Ⱓ������� 2�� �Ϸ� �ȵȰ� ����������)
	sqlStr = sqlStr + " 	and m.sitename <> '10x10' "							'// �������ֹ�
	sqlStr = sqlStr + " 	and a.divcd = 'A005' "								'// ������ȯ��
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
	sqlStr = sqlStr + " 			and m.sitename = '10x10' "																'// �������ֹ� ����
	''sqlStr = sqlStr + " 			and d.isupchebeasong='Y' "																'// �ٹ�+����
	sqlStr = sqlStr + " 			and d.cancelyn <> 'Y' "
	sqlStr = sqlStr + " 			and IsNull(T.code,'00') in ('05','06') "															'// ǰ�����Ұ�/�ù��ľ�
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
' ���̵� Ȯ�λ���
Dim csMainUserID, csMainUserName
csMainUserID	= req("csMainUserID", session("ssBctId") )
csMainUserName	= session("ssBctCname")

'// ��������
''if (csMainUserID = "raider7942") then
''	csMainUserName = "������B"
''end if

'��ó���޸�
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
' ������ �Ա� ��Ȯ��
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
' �μ� ��������
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
'// �� ����ǰ����� �� ����� �ֹ����
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
					"	alert('����Ǿ����ϴ�.'); location.href = 'cscenter_main.asp?menupos=" & menupos & "'; " &_
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

//��������
function popCooperate(){
	 var winCooperate = window.open("/admin/cooperate/popIndex.asp","popCooperate","width="+screen.availWidth+", height="+ screen.availHeight +",resizable=yes, scrollbars=yes");
	 winCooperate.focus();
}

function jsSearchEvent(frm) {
	if(frm.selEvt.value == "evt_code" && frm.sEtxt.value != "") {
		frm.sEtxt.value = frm.sEtxt.value.replace(/\s/g, "");
		if(!IsDigit(frm.sEtxt.value)) {
			alert("�̺�Ʈ�ڵ�� ���ڸ� �����մϴ�.");
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
	objUpcheReturn7_2.innerHTML = GetDateDiffString(v.getTime() - csTimeUpcheReturn.getTime());		// Ÿ�̸� ���� ����

	objIpjumStockOut.innerHTML = GetDateDiffString(v.getTime() - csTimeIpjumStockOut.getTime());
	objIpjumRefund.innerHTML = GetDateDiffString(v.getTime() - csTimeIpjumRefund.getTime());
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
            			    <td> <font color="blue">�����ǻ��� ��ó�� ����</font></td>
            			    <td align="right">
            			        <a href="javascript:PopMyQna('', '', 'N', '26', '', '', '', '', '');">
        				        <b><%= csrecommendBoardcnt %></b> ��
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
						</tr>
            			<tr height="25">
            			    <td> VVIP�� ����ں� ��ó�� ����</td>
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
        				        	<b><%= tmpVal(1) %></b> ��
									<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
						</tr>
						<%
							End If
						Next
						%>
            			<tr height="25">
            			    <td>&nbsp;&nbsp;- VVIP�� �̺й� ����</td>
            			    <td align="right">
            			        <a href="javascript:PopMyQna('', '', 'VV', '', '', '', '', '', '');">
        							<b><%= csVVipBoardNochargecnt %></b> ��
        							<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
						</tr>
            			<tr height="25">
            			    <td> VIP�� ����ں� ��ó�� ����</td>
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
        				        	<b><%= tmpVal(1) %></b> ��
									<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
						</tr>
						<%
							End If
						Next
						%>
            			<tr height="25">
            			    <td>&nbsp;&nbsp;- VIP�� �̺й� ����</td>
            			    <td align="right">
            			        <a href="javascript:PopMyQna('', '', 'V', '', '', '', '', '', '');">
        				        <b><%= csVipBoardNochargecnt %></b> ��
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
						</tr>
            			<tr height="25">
            			    <td> �Ϲݰ� ����ں� ��ó�� ����</td>
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
        				        	<b><%= tmpVal(1) %></b> ��
									<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
						</tr>
						<%
							End If
						Next
						%>
            			<tr height="25">
            			    <td>&nbsp;&nbsp;- �Ϲݰ� �̺й� ����</td>
            			    <td align="right">
            			        <a href="javascript:PopMyQnaListByChargeId('', 'N');">
        				        <b><%= csBoardNochargecnt %></b> ��
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td> <font color="blue">STAFF ��ó�� ����</font></td>
            			    <td align="right">
            			        <a href="javascript:PopMyQna('', '', 'N', '', '', '', '', '7', '');">
        				        <b><%= csstaffBoardcnt %></b> ��
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
						</tr>
            			<tr height="25">
            			    <td> <font color="blue">STAFF ǰ����ҿ�û��</font></td>
            			    <td align="right">
            			        <a href="javascript:showHideBrand('csstaffstockoutordertr');">
        				        <b><%= csStaffStockoutOrderCount %></b> ��
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
		        				        Ȯ��
		        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
		        				        </a>
		            			    </td>
		            			</tr>
	            			<% next %>
            			<tr height="25">
            			    <td>���޸� ��ó�� ����</td>
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
        				        	<b><%= tmpVal(1) %></b> ��
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
        	    <!-- �Խ��� ���� ��-->
        	</td>
        </tr>
        <tr valign="top">
            <td height="10"></td>
		</tr>
        <!-- CS ����� ��û���� ����ó��(2013-07-24 skyer9)<tr valign="top">
            <td>
        	    <!-aa- ��ǰ���� �亯���� ���� ����-aa->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("menubar") %>">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
                        <tr height="25">
            			    <td style="border-bottom:1px solid #BABABA">
            			    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>��ǰ���� �亯���� ����</b>
            			    	(<span id="objMiAnswerList"></span>) <a href="javascript:RefreshData('csTimeMiAnswerList')"><img src="/images/icon_reload.gif" border="0"></a>
            			    </td>
            			    <td align="right" style="border-bottom:1px solid #BABABA">
								&nbsp;
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td>��ǰ���� �亯����[1����, ����, D+3���̳� ����]</td>
            			    <td align="right"><a href="javascript:showHideBrand('csMiAnswerListByBrandUpche');"><b><%= FormatNumber(csTotalCountMiAnswerUpche, 0) %></b> ��</a></td>
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
        				        <%= tmpVal(1) %> ��
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
            			</tr>
							<% end if %>
            			<% next %>
            			<tr height="25">
            			    <td>��ǰ���� �亯����[1����, �ٹ�]</td>
            			    <td align="right"><a href="javascript:showHideBrand('csMiAnswerListByBrandTen');"><b><%= FormatNumber(csTotalCountMiAnswerTen, 0) %></b> ��</a></td>
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
        				        <%= tmpVal(1) %> ��
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
        	    <!-aa- ��ǰ���� �亯���� ���� ��-aa->
        	</td>
        </tr>
        <tr valign="top">
            <td height="10"></td>
		</tr>-->
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
								(<span id="objMichulgoListUpche"></span>) <a href="javascript:RefreshData('csTimeMichulgoListUpcheNew')"><img src="/images/icon_reload.gif" border="0"></a>
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
            			    <td align="right"><b><%=FormatNumber(arrMiSend(0),0)%></b> ��<a href="javascript:popUpcheMisendNEW('Y', '', '', '2', '', '', 'Y', 'topN', '');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<tr height="25">
            			    <td>3�� �̻� �̹߼۰�</td>
            			    <td align="right"><b><%=FormatNumber(arrMiSend(1),0)%></b> ��<a href="javascript:popUpcheMisendNEW('Y', '', '', '3', '', '', '', 'topN', '');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
						<% if (UBound(arrMiSend) > 2) then %>
            			<tr height="25">
            			    <td>&nbsp;&nbsp;- ���� ���Է�</td>
            			    <td align="right">
								<b><%=FormatNumber(arrMiSend(3),0)%></b> ��<a href="javascript:popUpcheMisendNEW('Y', '', '', '3', '00', '', '', 'all', '');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a>
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td>&nbsp;&nbsp;- ������� ����</td>
            			    <td align="right">
								<b><%=FormatNumber(arrMiSend(4),0)%></b> ��<a href="javascript:popUpcheMisendNEW('Y', '', '', '3', '', '', '', 'all', 'Y');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a>
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td>&nbsp;&nbsp;- ǰ��</td>
            			    <td align="right">
								<b><%=FormatNumber(arrMiSend(5),0)%></b> ��<a href="javascript:popUpcheMisendNEW('Y', '', '', '3', '05', '', '', 'all', '');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a>
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td>&nbsp;&nbsp;- �ֹ�����</td>
            			    <td align="right">
								<b><%=FormatNumber(arrMiSend(6),0)%></b> ��<a href="javascript:popUpcheMisendNEW('Y', '', '', '3', '02', '', '', 'all', '');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a>
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td>&nbsp;&nbsp;- �������</td>
            			    <td align="right">
								<b><%=FormatNumber(arrMiSend(7),0)%></b> ��<a href="javascript:popUpcheMisendNEW('Y', '', '', '3', '09', '', '', 'all', '');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a>
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td>&nbsp;&nbsp;- �������</td>
            			    <td align="right">
								<b><%=FormatNumber(arrMiSend(8),0)%></b> ��<a href="javascript:popUpcheMisendNEW('Y', '', '', '3', '03', '', '', 'all', '');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a>
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td>&nbsp;&nbsp;- ��Ÿ����</td>
            			    <td align="right">
								<b><%=FormatNumber(arrMiSend(9),0)%></b> ��<a href="javascript:popUpcheMisendNEW('Y', '', '', '3', '', '', '', 'all', '');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a>
            			    </td>
            			</tr>
						<% end if %>
            			<tr height="25">
            			    <td>����ں� D+3 �̹߼۰�(������,�������,����� ����)</td>
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
														<%= csMichulgoItemByBrand(2) %> ��
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
            			    <td>&nbsp;&nbsp;- �̺й� �̹߼۰�</td>
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
											<%= csMichulgoItemByBrand(2) %> ��
											<img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
										</a>
									</td>
								</tr>
								<% end if %>
							<% end if %>
						<% next %>
            			<tr height="25">
            			    <td>ǰ����ҿ�û��</td>
            			    <td align="right"><b><%=FormatNumber(arrMiSend(2),0)%></b> ��<a href="javascript:popUpcheMisendNEW('Y', '', '', '', '05', '0', '', 'topN', '');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<!--<tr height="25">
            			    <td colspan="2"></td>
            			 	too Many
            			    <td>�������� ���Է°� (D+2�̻�)</td>
            			    <td align="right"><b>??</b> ��<img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></td>
            			</tr>-->
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
        	    <!-- �������� ����� ���� ���� -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="#B2CCFF">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
                        <tr height="25">
            			    <td style="border-bottom:1px solid #BABABA">
            			    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>������ �������Ʈ[����]</b>
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
						prevSiteName = ""
            			for i = 0 to UBound(csMichulgoOtherSiteBrandSiteNameArr)
            			%>
							<% if (prevSiteName <> csMichulgoOtherSiteBrandSiteNameArr(i)) then %>
								<%
								prevSiteName = csMichulgoOtherSiteBrandSiteNameArr(i)

								'// ����Ʈ�� �հ� ���ϱ�
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
											<%= csMichulgoOtherSiteBrandcntArr(j) %> ��
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
        	    <!-- �������� ����� ���� ��-->
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
            			    <td align="right"><a href="javascript:popUpcheMisendNEW('N', '', '', '', '00', '', '', 'topN', '');"><b><%=FormatNumber(arrTenMiSend(0),0)%></b> ��<img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<tr height="25">
            			    <td>�� �̾ȳ���</td>
            			    <td align="right"><a href="javascript:popUpcheMisendNEW('N', '', '', '', '', '0', '', 'topN', '');"><b><%=FormatNumber(arrTenMiSend(1),0)%></b> ��<img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<tr height="25">
            			    <td>SMS/Mail/��ȭ�Ϸ��</td>
            			    <td align="right"><a href="javascript:misendmaster('4');"><b>-</b> ��<img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<tr height="25">
            			    <td>CSó���Ϸ�(������ó����û)��</td>
            			    <td align="right"><a href="javascript:misendmaster('6');"><b>-</b> ��<img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
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
            <td width="50%">
        	    <!-- ��ǰ����(��ü���, ����������) D+4 ��ó�� ���� -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("menubar") %>">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
                        <tr height="25">
            			    <td style="border-bottom:1px solid #BABABA">
            			    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>��ǰ����(����) D+4 ��ó��</b>
            			    	(<span id="objUpcheReturn"></span>) <a href="javascript:RefreshData('csTimeUpcheReturn')"><img src="/images/icon_reload.gif" border="0"></a>
            			    </td>
            			    <td align="right" style="border-bottom:1px solid #BABABA">
            			        <a href="javascript:Cscenter_Action_MiFinishReturnListDPlus(7);">
        				        �ٷΰ���
        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
        				        </a>
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td>��ǰ����(��ü���) D+4 ��ó��</td>
            			    <td align="right"><b><%= upcheReturnNotFinish %></b> ��<a href="javascript:Cscenter_Action_MiFinishReturnListDPlus(4);"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<tr height="25">
            			    <td>����ں� D+4 ��ó����(��ǰ���� ����, ����������)</td>
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
														<%= csUpcheReturnItemByBrand(2) %> ��
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
            			    <td>&nbsp;&nbsp;- �̺й� ��ó����</td>
            			    <td align="right">
        				        <b><%= nochargeUpcheReturnBrandcnt %></b>
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
        	    <!-- �������� ��ǰ���� ���� ���� -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="#B2CCFF">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
                        <tr height="25">
            			    <td style="border-bottom:1px solid #BABABA">
            			    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>������ ��ǰ���� D+7 ��ó��</b>
								(<span id="objUpcheReturn7_2"></span>) <a href="javascript:RefreshData('csTimeUpcheReturn')"><img src="/images/icon_reload.gif" border="0"></a>
            			    </td>
            			    <td align="right" style="border-bottom:1px solid #BABABA">
            			    	&nbsp;
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td>��������(��ǰ���� ����)</td>
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

								'// ����Ʈ�� �հ� ���ϱ�
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
											<%= csReturn7OtherSiteBrandcntArr(j) %> ��
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
        	    <!-- �������� ����� ���� ��-->
        	</td>
        </tr>
        <tr valign="top">
            <td height="10"></td>
        </tr>
        <tr valign="top">
            <td width="50%">
        	    <!-- ��ü��� ���� ���� -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="#B2CCFF">
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
            			    <td>��������(�� �ȳ�����+���ȳ�)</td>
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
								����Ÿ�� ������ �ֽ��ϴ�. ���ΰ�ħ�ϼ���.(<%= application("csStockoutOrderUseridArr") %>)
							</td>
            			    <td align="right">
            			    </td>
            			</tr>
                        <%
else
                        '// ù��°�� ��ŵ�Ѵ�.(��������)
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
									'// ���⼭ ������
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
		        				        <%= csStockoutOrderdetailCntcntArr(j) %> ��
		        				        <img src="/images/icon_arrow_right.gif" align="absbottom" border="0">
		        				        </a>
		            			    </td>
		            			</tr>
                            	<% end if %>
	            			<% next %>
            			<% next %>
<% end if %>
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
	            			    	<% 'if Not IsNull(csStockoutOrderUseridArr(i)) and csStockoutOrderUseridArr(i) <> "" then %>
	            			    		(<%'= csStockoutOrderUseridArr(i) %>)
	            			    	<% 'end if %>
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

				<p />

        	    <!-- ǰ����ǰ ����� �� ����� �ֹ���� -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("menubar") %>">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
                        <tr height="25">
            			    <td style="border-bottom:1px solid #BABABA" colspan="3">
            			    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>ǰ����ǰ ����� �� ����� �ֹ�</b>
								(����������Ʈ : <%= chulgoAbleOrderRegdate %>)
								<a href="javascript:jsUpdateChulgoAbleOrder()"><img src="/images/icon_reload.gif" border="0"></a>
            			    </td>
            			    <td align="right" style="border-bottom:1px solid #BABABA">
            			    	&nbsp;
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td align="center" width="170">�� �������</td>
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
						<tr height="25">
            			    <td>��ü��޹��� ��ó��</td>
            			    <td align="right">
            			        <b><%= CSA060NotFinish %></b> ��
                                <a href="javascript:Cscenter_Action_List2_A060('<%= csMainUserID %>');">
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
        <tr valign="top">
            <td>
                <!-- �̺�Ʈ�˻� -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("tabletop") %>">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
            			<tr height="25">
            			    <td style="border-bottom:1px solid #BABABA">
            			        <img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>�̺�Ʈ�˻�</b>
            			    </td>
            			    <td align="right" style="border-bottom:1px solid #BABABA">
            			        &nbsp;
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td colspan="2">
								<form name="frmEvent" method="post" action="" onSubmit="return false;">
									<select class="select" name="selEvt">
			    						<option value="evt_code" >�̺�Ʈ �ڵ�</option>
			    						<option value="evt_name">�̺�Ʈ��</option>
			    					</select>
									<input type="text" class="text" name="sEtxt" value="" size="20" maxlength="60" onKeyPress="if (event.keyCode == 13) jsSearchEvent(frmEvent);">
									<input type="button" class="button" value="�˻�" onClick="jsSearchEvent(frmEvent)">
								</form>
							</td>
            			</tr>
            	        </table>
            	    </td>
            	</tr>
            	</table>
        	    <!--  �̺�Ʈ�˻� ��-->
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
            			    <td style="border-bottom:1px solid #BABABA" colspan="2">
        						<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
            						<%
									if (oCTaxRequest.FResultCount < 1) then
										%>
									<tr height="25">
            							<td>&nbsp;&nbsp; ��û����</td>
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
            			        <img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>�������Ա� ����</b>
            			    </td>
            			    <td align="right" style="border-bottom:1px solid #BABABA">
								��Ȯ�� [<b><%= csNotFinishTodayIpkum %></b>] / ���� ��ü : [<%= csTotalTodayIpkum %>]
								&nbsp;...&nbsp;
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
            			    <td width="5"></td>
            			    <td>
            			    	<table width="100%" style="border:1px solid #BABABA" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#FFFFFF">
		                        <tr height="20">
		            			    <td align="center">
		            			    	<a href="javascript:PopCSfileSend('','','','');">���������۰���</a>
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
            	<!-- ������ȣ -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="<%= adminColor("tabletop") %>">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
                        <tr height="25">
            			    <td width="33%">
            			    	<table width="100%" style="border:1px solid #BABABA" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#FFFFFF">
		                        <tr height="20">
		            			    <td align="center">
		            			    	<a href="javascript:PopCSSMSCertLog('','','','');">������ȣ�����̷�</a>
		            			    </td>
		            			</tr>
		            	        </table>
            			    </td>
            			    <td width="5"></td>
            			    <td width="33%">
            			    	<table width="100%" style="border:1px solid #BABABA" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#FFFFFF">
		                        <tr height="20">
		            			    <td align="center">
		            			    	<a href="javascript:PopCSKAKAOLog();">īī���������̷�</a>
		            			    </td>
		            			</tr>
		            	        </table>
            			    </td>
            			    <td width="5"></td>
            			    <td>
								<table width="100%" style="border:1px solid #BABABA" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#FFFFFF">
									<tr height="20">
		            					<td align="center">
		            			    		<a href="javascript:PopCSReqWork();">�μ���������<%= CHKIIF(myPartWorkCnt>0, "<font color='#FF1493'><b>((( " & myPartWorkCnt & " )))</b></font>", "(" & myPartWorkCnt & ")") %></a>
		            					</td>
		            				</tr>
		            	        </table>
            			    </td>
            			</tr>
            	        </table>
            	    </td>
            	</tr>
            	</table>
        	    <!-- ������ȣ -->
            </td>
        </tr>
        <tr valign="top">
            <td height="10"></td>
        </tr>
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
            			    <td>ī����� D+1 ��ó��</td>
            			    <td align="right"><b><%= csRequestCardCancelCount %></b> ��<a href="javascript:Cscenter_Action_List2('cardnocheckdp1');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<tr height="25">
            			    <td>�ֹ���� ��ó��</td>
            			    <td align="right"><b><%= csNotFinA008 %></b> ��<a href="javascript:Cscenter_Action_List2('cancelnofinish');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<tr height="25">
            			    <td>ȸ����û ��ó��</td>
            			    <td align="right"><b><%= csReturnNotFinish %></b> ��<a href="javascript:Cscenter_Action_List2('returnmifinish');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<!--<tr height="25">
            			    <td>��ü ��ó��</td>
            			    <td align="right"><b><%'= csUpcheNotFin %></b> ��<a href="javascript:Cscenter_Action_List2('upchemifinish');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>-->
            			<tr height="25">
            			    <td>��ǰ����(��ü���) D+4 ��ó��</td>
            			    <td align="right"><b><%= upcheReturnNotFinish %></b> ��<a href="javascript:Cscenter_Action_MiFinishReturnListDPlus(7);"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>

            			<tr height="25">
            			    <td>��üó���Ϸ�</td>
            			    <td align="right"><b><%= csUpcheFinished %></b> ��<a href="javascript:Cscenter_Action_List2('upchefinish');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<tr height="25">
            			    <td>����ó���Ϸ�</td>
            			    <td align="right"><b><%= csLogicsFinished %></b> ��<a href="javascript:Cscenter_Action_List2('logicsfinish');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
            			<tr height="25">
            			    <td>���߰�����</td>
            			    <td align="right"><b><%= csCustomerAddPayment %></b> ��<a href="javascript:Cscenter_Action_List2('customeraddpay');"><img src="/images/icon_arrow_right.gif" align="absbottom" border="0"></a></td>
            			</tr>
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
        	    <!-- ������ ȯ�ҹ�ó�� ���� -->
                <table width="100%" style="border:1px solid #BABABA" align="center" cellpadding="5" cellspacing="0" class="a">
        	    <tr bgcolor="#B2CCFF">
        	        <td>
        	            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
                        <tr height="25">
            			    <td style="border-bottom:1px solid #BABABA">
            			    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>������ ȯ�ҹ�ó��</b>
								(<span id="objIpjumRefund"></span>) <a href="javascript:RefreshData('csTimeIpjumRefund')"><img src="/images/icon_reload.gif" border="0"></a>
            			    </td>
            			    <td align="right" style="border-bottom:1px solid #BABABA">
            			    	&nbsp;
            			    </td>
            			</tr>
            			<tr height="25">
            			    <td>������ ȯ�ҹ�ó��</td>
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
		        				        <%= csRefundOtherSiteCSCntcntArr(j) %> ��
		            			    </td>
		            			</tr>
	            				<% end if %>
	            			<% next %>
            			<% next %>
            	        </table>
            	    </td>
            	</tr>
            	</table>
        	    <!-- ������ ȯ�ҹ�ó�� ��-->
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
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
