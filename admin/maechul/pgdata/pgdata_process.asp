<%@ language=vbscript %>
<% option explicit %>
<% Server.ScriptTimeOut = 900 %>
<%
'###########################################################
' Description : PG����γ���
' Hieditor : 2011.04.22 �̻� ����
'			 2023.03.28 �ѿ�� ����(Apple Pay�߰�)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionSTAdmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<!-- #include virtual="/cscenter/action/incKakaopayCommonNew.asp"-->
<%
'' /wAPI/autojob/pgdata_process.asp  '' �ǵ��� �������� �̰���
' �ݵ�� 3������ �ҽ��� �����ؾ� �մϴ�. �Ѱ��� ��ĥ��� ������ �ΰ��� ������ �ּ���.
' scm\admin\maechul\pgdata\pgdata_process.asp
' webadmin\admin\maechul\pgdata\pgdata_process.asp
' wapi\autojob\pgdata_process.asp

Dim StopWatch(19)

sub StartTimer(x)
	StopWatch(x) = Timer
end Sub

function StopTimer(x)
	Dim EndTime

	EndTime = Timer

	if EndTime < StopWatch(x) Then
		EndTime = EndTime + (86400)
	end if

	StopTimer = EndTime - StopWatch(x)
end function

dim refer
	refer = request.ServerVariables("HTTP_REFERER")

dim mode, reguserid, logidx, orderno, matchidx, IsMatched, objData, errMsg, objXML, xmlURL, objLine, xmlURLArr
dim PGkey, PGCSkey, appDate, cancelDate, yyyymmdd, subPgKey
dim appPrice, commPrice, commVatPrice, jungsanPrice, lastipkumdate, searchipkumdate
dim prevPGkey, prevPrevPGkey, prevAppDivCode, prevPrevAppDivCode, IsDuplicate
dim tmpStr, arrOrderSerial, orderserial, searchipkumdateMAX, force
dim objFSO, objOpenedFile, yyyymm, Status, resultStatus, targetFileName, sqlStr, addSqlStr, i, j, k, ix
dim orgpgkey, orgpgcskey, gubun
dim testAppPrice, testCommPrice, testCommVatPrice, testJungsanPrice, checkDate
dim feeAmount, feeTaxAmount, settlementAmount, lastPGkey, checkEndDate
'// �̴Ϸ�Ż �����Է°� �߰�
dim inirentalpgkey, inirentalgubun, inirentalconfirmdate, inirentalipkumdate, inirentalappprice, inirentalcommprice, inirentalcommvatprice, inirentaljungsanprice
dim inirentalreduplication
inirentalpgkey = requestCheckvar(request("inirentalpgkey"),64) '//pgŰ
inirentalgubun = requestCheckvar(request("inirentalgubun"),32) '//����(inirentalbuy - ����, inirentalcancel - ���)
inirentalconfirmdate = requestCheckvar(request("inirentalconfirmdate"),32) '//����,��� ����
inirentalipkumdate = requestCheckvar(request("inirentalipkumdate"),32) '//�Աݿ���(����)��
inirentalappprice = requestCheckvar(request("inirentalappprice"),64) '//�ݾ�
inirentalcommprice = requestCheckvar(request("inirentalcommprice"),64) '//������
inirentalcommvatprice = requestCheckvar(request("inirentalcommvatprice"),64) '//�ΰ���
inirentaljungsanprice = requestCheckvar(request("inirentaljungsanprice"),64) '//���꿹��(�Աݿ���)��
inirentalreduplication = 0

Const ForReading = 1
	mode = requestCheckvar(request("mode"),64)
	logidx = requestCheckvar(request("logidx"),32)
	orderno = requestCheckvar(request("orderno"),32)
	yyyymmdd = requestCheckvar(request("yyyymmdd"),32)
	yyyymm = requestCheckvar(request("yyyymm"),7)
	reguserid = session("ssBctId")
	orgpgkey = requestCheckvar(request("orgpgkey"),64)
	orgpgcskey = requestCheckvar(request("orgpgcskey"),64)
	gubun = requestCheckvar(request("gubun"),32)
	cancelDate = requestCheckvar(request("cancelDate"),32)

dim excmatchfinish, onlypricenotequal, yyyy1, mm1, dd1, yyyy2, mm2, dd2, yyyy3, mm3, dd3, yyyy4, mm4, dd4, selectreasonGubun
dim sitename, appDivCode, ipkumdate, PGuserid, appMethod, searchfield, searchtext, pggubun, reasonGubun, tmpDate
dim showjumunlog, showjumunlogNotMatch, chkSearchIpkumDate, chkSearchAppDate, fromDate, toDate, fromDate2, toDate2
	selectreasonGubun 	= requestCheckvar(request("selectreasonGubun"),32)
	excmatchfinish = requestCheckvar(request("excmatchfinish"),10)
	onlypricenotequal = requestCheckvar(request("onlypricenotequal"),10)

	yyyy1   = requestCheckvar(request("yyyy1"),32)
	mm1     = requestCheckvar(request("mm1"),32)
	dd1     = requestCheckvar(request("dd1"),32)
	yyyy2   = requestCheckvar(request("yyyy2"),32)
	mm2     = requestCheckvar(request("mm2"),32)
	dd2     = requestCheckvar(request("dd2"),32)

	yyyy3   = requestCheckvar(request("yyyy3"),32)
	mm3     = requestCheckvar(request("mm3"),32)
	dd3     = requestCheckvar(request("dd3"),32)
	yyyy4   = requestCheckvar(request("yyyy4"),32)
	mm4     = requestCheckvar(request("mm4"),32)
	dd4     = requestCheckvar(request("dd4"),32)

	sitename		= requestCheckvar(request("sitename"),32)
	appDivCode 		= requestCheckvar(request("appDivCode"),32)
	ipkumdate 		= requestCheckvar(request("ipkumdate"),32)
	PGuserid 		= requestCheckvar(request("PGuserid"),32)
	appMethod 		= requestCheckvar(request("appMethod"),32)

	searchfield 	= requestCheckvar(request("searchfield"),32)
	searchtext 		= Replace(Replace(requestCheckvar(request("searchtext"),64), "'", ""), Chr(34), "")

	pggubun 		= requestCheckvar(request("pggubun"),32)
	reasonGubun 	= requestCheckvar(request("reasonGubun"),32)

	showjumunlog 				= requestCheckvar(request("showjumunlog"),32)
	showjumunlogNotMatch 		= requestCheckvar(request("showjumunlogNotMatch"),32)
	chkSearchIpkumDate 			= requestCheckvar(request("chkSearchIpkumDate"),32)
	chkSearchAppDate 			= requestCheckvar(request("chkSearchAppDate"),32)

if (chkSearchIpkumDate="") then chkSearchAppDate = "Y"
if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) - 1), 1)
	toDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), 1)

	yyyy1 = Cstr(Year(fromDate))
	mm1 = Cstr(Month(fromDate))
	dd1 = Cstr(day(fromDate))

	tmpDate = DateAdd("d", -1, toDate)
	yyyy2 = Cstr(Year(tmpDate))
	mm2 = Cstr(Month(tmpDate))
	dd2 = Cstr(day(tmpDate))

	fromDate2 = fromDate
	toDate2 = toDate
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
	toDate = DateSerial(yyyy2, mm2, dd2+1)
end if
if (yyyy3="") then
	fromDate2 = DateSerial(Cstr(Year(now())), Cstr(Month(now()) - 1), 1)
	toDate2 = DateSerial(Cstr(Year(now())), Cstr(Month(now())), 1)

	yyyy3 = Cstr(Year(fromDate2))
	mm3 = Cstr(Month(fromDate2))
	dd3 = Cstr(day(fromDate2))

	tmpDate = DateAdd("d", -1, toDate2)
	yyyy4 = Cstr(Year(tmpDate))
	mm4 = Cstr(Month(tmpDate))
	dd4 = Cstr(day(tmpDate))
else
	fromDate2 = DateSerial(yyyy3, mm3, dd3)
	toDate2 = DateSerial(yyyy4, mm4, dd4+1)
end if

function GetAppFromInicis(searchipkumdate, ByRef objData, ByRef errMsg)
	dim xmlURL
	dim objXML

	ipkumdate = Replace(searchipkumdate, "-", "")

	xmlURL = "https://iniweb.inicis.com/mall/cr/urlsvc//UrlSendCalCulAll.jsp?urlid=Teenxt14GI&passwd=cube1010??&date=" + CStr(ipkumdate) + "&flgurl=31"
	''response.write xmlURL
	''response.end

	objData = ""
	errMsg = ""

	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")

	objXML.setTimeouts 5 * 000, 5 * 000, 15 * 000, 45 * 000
	objXML.Open "GET", xmlURL, false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXML.Send()

	if objXML.Status = "200" then
		''response.write objXML.ResponseBody
		''response.end
		if (Trim(objXML.ResponseBody)<>"") then
			objData = BinaryToText(objXML.ResponseBody, "euc-kr")
		else
			errMsg = "������ ����Ÿ�� �����ϴ�.[0]"
			Set objXML  = Nothing
			Exit Function
		end if
	end if

	''response.write objXML.Status
	''response.write objData
    ''response.end

	Set objXML  = Nothing

	if (InStr(objData, "NO DATA") > 0) then
		errMsg = "������ ����Ÿ�� �����ϴ�.[1]"
		Exit Function
	end if
end function

if (mode="matchoneorder") then

    sqlStr = " select isNULL(orderserial,'') as orderserial " & VbCRLF
    sqlStr = sqlStr & " from db_shop.dbo.tbl_shopjumun_cardApp_log " & VbCRLF
    sqlStr = sqlStr & " where idx="&logidx&VbCRLF

	IsMatched = True

    rsget.Open sqlStr,dbget,1
	if Not(rsget.EOF or rsget.BOF) then
	    IsMatched = Not (rsget("orderserial") = "")
	end if
	rsget.Close

	if IsMatched then
		response.write "<script>alert('�̹� ��Ī�� �����Դϴ�.');</script>"
		response.write "�̹� ��Ī�� �����Դϴ�."
		dbget.close()
		response.end
	end if

	sqlStr = " update l "
	sqlStr = sqlStr + " set l.shopJumunMasterIdx = m.idx, l.orderserial = m.orderno, l.shopid = m.shopid "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " db_shop.dbo.tbl_shopjumun_cardApp_log l "
	sqlStr = sqlStr + " join db_shop.dbo.tbl_shopjumun_master m "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and l.idx = " + CStr(logidx) + " "
	sqlStr = sqlStr + " 	and m.orderno = '" + CStr(orderno) + "' "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and l.idx = " + CStr(logidx) + " "
	sqlStr = sqlStr + " 	and l.shopJumunMasterIdx is NULL "
	dbget.Execute sqlStr

	response.write "<script>alert('����Ǿ����ϴ�.'); opener.location.reload(); opener.focus(); window.close();</script>"
	dbget.close()
	response.end

elseif (mode="matchsumipkum") then

	arrOrderSerial = Split(requestCheckvar(request("arrOrderSerial"),512), vbCrLf)
	tmpStr = ""

	for each orderserial in arrOrderSerial
		if (Len(orderserial) > 0) then
			if (Len(orderserial) <> 11) then
				response.write "<script>alert('�߸��� �ֹ���ȣ�Դϴ�.');</script>"
				response.write "�߸��� �ֹ���ȣ�Դϴ�." & orderserial
				dbget.close()
				response.end
			end if

			if (tmpStr = "") then
				tmpStr = " select '" + CStr(orderserial) + "' as orderserial " & vbCrLf
			else
				tmpStr = tmpStr + " union all " & vbCrLf & " select '" + CStr(orderserial) + "' as orderserial " & vbCrLf
			end if
		end if
	next

	if (tmpStr = "") then
		response.write "<script>alert('�Էµ� �ֹ���ȣ�� �����ϴ�.');</script>"
		response.write "�Էµ� �ֹ���ȣ�� �����ϴ�."
		dbget.close()
		response.end
	end if

	sqlStr = " insert into db_order.dbo.tbl_onlineApp_log(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, orderserial, PGuserid, orgPayDate, PGmeachulDate) " & vbCrLf
	sqlStr = sqlStr + " select l.PGgubun, l.PGkey, T.orderserial, l.sitename, l.appDivCode, l.appMethod, l.appDate, l.cancelDate, 0, 0, 0, 0, l.ipkumdate, T.orderserial, l.PGuserid, l.orgPayDate, l.PGmeachulDate " & vbCrLf
	sqlStr = sqlStr + " from " & vbCrLf
	sqlStr = sqlStr + "	db_order.dbo.tbl_onlineApp_log l " & vbCrLf
	sqlStr = sqlStr + "	join ( " & vbCrLf

	sqlStr = sqlStr + tmpStr

	sqlStr = sqlStr + "	) T " & vbCrLf
	sqlStr = sqlStr + "	on " & vbCrLf
	sqlStr = sqlStr + "		1 = 1 " & vbCrLf
	sqlStr = sqlStr + "	left join db_order.dbo.tbl_onlineApp_log l2 " & vbCrLf
	sqlStr = sqlStr + "	on " & vbCrLf
	sqlStr = sqlStr + "		1 = 1 " & vbCrLf
	sqlStr = sqlStr + "		and l.pggubun = l2.pggubun " & vbCrLf
	sqlStr = sqlStr + "		and l.pgkey = l2.pgkey " & vbCrLf
	sqlStr = sqlStr + "		and T.orderserial = l2.pgcskey " & vbCrLf
	sqlStr = sqlStr + "where " & vbCrLf
	sqlStr = sqlStr + "	1 = 1 " & vbCrLf
	sqlStr = sqlStr + "	and l.pggubun = 'bankipkum' " & vbCrLf
	sqlStr = sqlStr + "	and l.appDivCode = 'A' " & vbCrLf
	sqlStr = sqlStr + "	and l.idx = " + CStr(logidx) + " " & vbCrLf
	''sqlStr = sqlStr + "	and l.PGCSkey = '' " & vbCrLf
	sqlStr = sqlStr + "	and l2.idx is NULL " & vbCrLf
	''response.write sqlStr
	dbget.Execute sqlStr

	response.write "<script>alert('����Ǿ����ϴ�.'); opener.location.reload(); opener.focus(); window.close();</script>"
	dbget.close()
	response.end

elseif (mode="forcematchorderserial") then

	sqlStr = " update m "
	sqlStr = sqlStr & " set m.orderserial = '" & orderserial & "' "
	sqlStr = sqlStr & " from "
	sqlStr = sqlStr & " 	db_order.dbo.tbl_onlineApp_log m "
	sqlStr = sqlStr & " where m.idx = '" & logidx & "' "
	dbget.execute sqlStr

	response.write "<script>alert('����Ǿ����ϴ�.'); opener.location.reload(); opener.focus(); window.close();</script>"
	dbget.close()
	response.end

elseif (mode="matchorderserial") then

    orderserial = requestCheckvar(request("OrderSerial"),32)

    '// �ֹ���ȣ ����
	sqlStr = " update t "
	sqlStr = sqlStr & " set t.orderserial = '" & orderserial & "' "
	sqlStr = sqlStr & " from "
	sqlStr = sqlStr & " 	[db_order].[dbo].[tbl_order_temp] t "
	sqlStr = sqlStr & " 	join db_order.dbo.tbl_onlineApp_log m on m.pgkey = t.P_TID "
	sqlStr = sqlStr & " where m.idx = '" & logidx & "' and t.orderserial = '' "
	dbget.execute sqlStr

	sqlStr = " update t "
	sqlStr = sqlStr & " set t.orderserial = m.orderserial, t.sitename = (case when m.rdsite = 'mobile' or m.rdsite = 'app_wish2' then '10x10mobile' else '10x10' end) "
	sqlStr = sqlStr & " from "
	sqlStr = sqlStr & " 	db_order.dbo.tbl_onlineApp_log t "
	sqlStr = sqlStr & " 	join [db_order].[dbo].[tbl_order_temp] m on t.pgkey = m.P_TID "
	sqlStr = sqlStr & " where t.idx = '" & logidx & "' "
	dbget.execute sqlStr

	sqlStr = " update t "
	sqlStr = sqlStr & " set t.appDate = m.regdate "
	sqlStr = sqlStr & " from "
	sqlStr = sqlStr & " 	db_order.dbo.tbl_onlineApp_log t "
	sqlStr = sqlStr & " 	join [db_order].[dbo].[tbl_order_temp] m on t.pgkey = m.P_TID "
	sqlStr = sqlStr & " where t.idx = '" & logidx & "' and t.appDivCode = 'A' and t.pggubun = 'toss' "
	dbget.execute sqlStr

	sqlStr = " update T "
	sqlStr = sqlStr & " set T.csasid = a.id, T.cancelDate = a.finishdate "
	sqlStr = sqlStr & " from "
	sqlStr = sqlStr & " 	db_order.dbo.tbl_onlineApp_log T "
	sqlStr = sqlStr & " 	join [db_cs].[dbo].[tbl_new_as_list] a "
	sqlStr = sqlStr & " 	on "
	sqlStr = sqlStr & " 		1 = 1 "
	sqlStr = sqlStr & " 		and T.orderserial = a.orderserial "
	sqlStr = sqlStr & " 		and T.PGCSkey = a.orderserial + '_' + convert(varchar, a.id) "
	sqlStr = sqlStr & " where t.idx = '" & logidx & "' and t.appDivCode <> 'A' and t.pggubun = 'toss' "
	dbget.execute sqlStr

	sqlStr = " update T "
	sqlStr = sqlStr & " set T.csasid = a.id, T.cancelDate = a.finishdate "
	sqlStr = sqlStr & " from "
	sqlStr = sqlStr & " 	db_order.dbo.tbl_onlineApp_log T "
	sqlStr = sqlStr & " 	join [db_cs].[dbo].[tbl_new_as_list] a "
	sqlStr = sqlStr & " 	on "
	sqlStr = sqlStr & " 		1 = 1 "
	sqlStr = sqlStr & " 		and T.orderserial = a.orderserial "
	sqlStr = sqlStr & " 		and T.PGCSkey = 'CANCELALL' "
    sqlStr = sqlStr & " 		and a.divcd in ('A008', 'A010', 'A004') "
    sqlStr = sqlStr & " 		and a.currstate = 'B007' "
    sqlStr = sqlStr & " 		and a.deleteyn = 'N' "
	sqlStr = sqlStr & " where t.idx = '" & logidx & "' and t.appDivCode = 'C' and t.pggubun = 'toss' "
	dbget.execute sqlStr

	sqlStr = " update m "
	sqlStr = sqlStr & " set m.ipkumdate = T.appDate "
	sqlStr = sqlStr & " from "
	sqlStr = sqlStr & " 	db_order.dbo.tbl_onlineApp_log T "
	sqlStr = sqlStr & " 	join [db_order].[dbo].[tbl_order_master] m on T.orderserial = m.orderserial "
	sqlStr = sqlStr & " where "
	sqlStr = sqlStr & " 	1 = 1 "
	sqlStr = sqlStr & " 	and T.PGgubun = 'toss' "
    sqlStr = sqlStr & " 	and T.idx = '" & logidx & "' "
	sqlStr = sqlStr & " 	and T.appDivCode = 'A' "
	sqlStr = sqlStr & " 	and DateDiff(day, T.appDate, m.ipkumdate) <> 0 "
	dbget.execute sqlStr

	sqlStr = " update m "
	sqlStr = sqlStr & " set m.orderserial = '" & orderserial & "' "
	sqlStr = sqlStr & " from "
	sqlStr = sqlStr & " 	db_order.dbo.tbl_onlineApp_log m "
	sqlStr = sqlStr & " where m.idx = '" & logidx & "' and IsNull(m.orderserial, '') = '' "
	dbget.execute sqlStr

	response.write "<script>alert('����Ǿ����ϴ�.'); opener.location.reload(); opener.focus(); window.close();</script>"
	dbget.close()
	response.end

elseif (mode="regReasonGubun") then

	sqlStr = " update db_order.dbo.tbl_onlineApp_log "
	sqlStr = sqlStr + " set reasonGubun = '" + CStr(reasonGubun) + "' "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and idx = " + CStr(logidx) + " "
	sqlStr = sqlStr + " 	and IsNull(reasonGubun, '') not in ('030') "
	''response.write sqlStr
	dbget.Execute sqlStr

	response.write "<script>alert('����Ǿ����ϴ�.'); opener.location.reload(); opener.focus(); window.close();</script>"
	dbget.close()
	response.end

elseif (mode="matchetcpay") then

    sqlStr = " exec [db_log].[dbo].[usp_Ten_MakeEtcPaymentLog_ON] '" & Left(DateAdd("d", -7, Now()), 10) & "', '" & Left(Now(), 10) & "' "
    dbget.Execute sqlStr

	response.write "<script>alert('��Ī�Ǿ����ϴ�.'); location.replace('" + CStr(refer) + "');</script>"
	dbget.close()
	response.end

elseif (mode="regReasonGubunOff") then

	sqlStr = " update db_shop.dbo.tbl_shopjumun_cardApp_log "
	sqlStr = sqlStr + " set reasonGubun = '" + CStr(reasonGubun) + "' "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and idx = " + CStr(logidx) + " "
	sqlStr = sqlStr + " 	and IsNull(reasonGubun, '') not in ('030') "
	''response.write sqlStr
	dbget.Execute sqlStr

	response.write "<script>alert('����Ǿ����ϴ�.'); opener.location.reload(); opener.focus(); window.close();</script>"
	dbget.close()
	response.end

elseif (mode="RegReasonGubun025") then

	sqlStr = " update m set m.reasonGubun = '025' "
	sqlStr = sqlStr + " from db_order.dbo.tbl_onlineApp_log m "
	sqlStr = sqlStr + " join [db_cs].[dbo].[tbl_new_as_list] a on m.csasid = a.id "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1=1 "
	sqlStr = sqlStr + " 	and m.pggubun = 'bankrefund' "
	sqlStr = sqlStr + " 	and IsNull(m.cancelDate, m.appDate)>='" + CStr(fromDate) + "' "
	sqlStr = sqlStr + " 	and IsNull(m.cancelDate, m.appDate)<'" + CStr(toDate) + "' "
	sqlStr = sqlStr + " 	and m.PGuserid = 'bankrefund_10x10' "
	''sqlStr = sqlStr + " 	and IsNull(m.reasonGubun, '') = 'R02' "
	sqlStr = sqlStr + " 	and a.title = '��ġ���� ���������� ȯ��' "

	''response.write sqlStr & "<br>"
	''response.end
	dbget.Execute sqlStr

	response.write "<script type='text/javascript'>alert('������ �ϰ� ����Ǿ����ϴ�.'); location.replace('" & CStr(refer) & "');</script>"
	dbget.close() : response.end

elseif (mode="RegReasonGubunarr") then
	'// ====================================================================
	addSqlStr = ""

	if (PGGubun <> "") then
		addSqlStr = addSqlStr + " and m.pggubun = '" + CStr(PGGubun) + "' "
	end if

	if (ExcMatchFinish <> "") then
		addSqlStr = addSqlStr + " and ( "
		addSqlStr = addSqlStr + " 	(m.appDivCode = 'A' and m.orderserial is NULL) "
		addSqlStr = addSqlStr + " 	or "
		addSqlStr = addSqlStr + " 	(m.appDivCode <> 'A' and m.csasid is NULL) "
		addSqlStr = addSqlStr + " ) "
		''addSqlStr = addSqlStr + " and not (m.appDivCode = 'C' and m.pgcskey = 'CANCELALL' and m.orderserial is not NULL) "
	end if

	'// ��������
	if (chkSearchAppDate = "Y") then
		if fromDate <> "" then
			addSqlStr = addSqlStr + " and IsNull(m.cancelDate, m.appDate)>='" + CStr(fromDate) + "'"
		end if
		if toDate <> "" then
			addSqlStr = addSqlStr + " and IsNull(m.cancelDate, m.appDate)<'" + CStr(toDate) + "'"
		end if
	end if

	'// �Աݿ�����
	if (chkSearchIpkumDate = "Y") then
		if fromDate2 <> "" then
			addSqlStr = addSqlStr + " and m.ipkumdate>='" + CStr(fromDate2) + "'"
		end if
		if toDate2 <> "" then
			addSqlStr = addSqlStr + " and m.ipkumdate<'" + CStr(toDate2) + "'"
		end if
	end if

	if (SiteName <> "") then
		addSqlStr = addSqlStr + " and m.sitename = '" + CStr(SiteName) + "' "
	end if

	if (AppDivCode <> "") then
		addSqlStr = addSqlStr + " and m.appDivCode = '" + CStr(AppDivCode) + "' "
	end if

	if (Ipkumdate <> "") then
		addSqlStr = addSqlStr + " and m.ipkumdate = '" + CStr(Ipkumdate) + "' "
	end if

	if (SearchField <> "") and (SearchText <> "") then
		if (SearchField = "orderserial") then
			addSqlStr = addSqlStr + " and Left(m.orderserial, 11) = '" + CStr(Left(SearchText, 11)) + "' "
		else
			addSqlStr = addSqlStr + " and m." + CStr(SearchField) + " = '" + CStr(SearchText) + "' "
		end if
	end if

	if (PGuserid <> "") then
		addSqlStr = addSqlStr + " and m.PGuserid = '" + CStr(PGuserid) + "' "
	end if

	if (AppMethod <> "") then
		addSqlStr = addSqlStr + " and m.appMethod = '" + CStr(AppMethod) + "' "
	end if

	if (ReasonGubun <> "") then
		if (ReasonGubun = "XXX") then
			addSqlStr = addSqlStr + " and IsNull(m.reasonGubun, '') not in ('001', '002', '020', '025', '030', '035', '040', '950', '999', '900', '901', '800') "
		else
			addSqlStr = addSqlStr + " and IsNull(m.reasonGubun, '') = '" + CStr(ReasonGubun) + "' "
		end if
	end if

	if (OnlyPriceNotEqual <> "") then
		addSqlStr = addSqlStr + " and m.appdivcode = 'A' "
		addSqlStr = addSqlStr + " and e.acctamount <> m.appprice "
	end if

	if (ShowJumunLogNotMatch = "Y") then
		ShowJumunLog = "Y"
		addSqlStr = addSqlStr + " and p.pggubun is NULL "
		''addSqlStr = addSqlStr + " and m.sitename not in ('fingers', '10x10gift') "
	end if
	'// ====================================================================

	sqlStr = "update m set m.reasonGubun='"& selectreasonGubun &"'"
	sqlStr = sqlStr + " from db_order.dbo.tbl_onlineApp_log m "

	if (OnlyPriceNotEqual <> "") then
		sqlStr = sqlStr + " join db_order.dbo.tbl_order_PaymentEtc e "
		sqlStr = sqlStr + " 	on m.orderserial = e.orderserial "
		sqlStr = sqlStr + " 	and e.acctdiv = m.appmethod "
	end if

	if (ShowJumunLog = "Y") then
		sqlStr = sqlStr + " left join db_datamart.dbo.tbl_order_payment_log p "
		sqlStr = sqlStr + " 	on m.pggubun = p.pggubun "
		sqlStr = sqlStr + " 	and m.pgkey = p.pgkey "
		sqlStr = sqlStr + " 	and m.pgcskey = p.pgcskey "
		sqlStr = sqlStr + " 	and m.appprice = p.realPayPrice "
	end if

	sqlStr = sqlStr + " where 1=1"
	sqlStr = sqlStr + addSqlStr

	'response.write sqlStr & "<br>"
	dbget.Execute sqlStr

	response.write "<script type='text/javascript'>alert('������ �ϰ� ����Ǿ����ϴ�.'); location.replace('" & CStr(refer) & "');</script>"
	dbget.close() : response.end

elseif (mode="delmatchone") then

	sqlStr = " update l "
	sqlStr = sqlStr + " set l.shopJumunMasterIdx = NULL, l.orderserial = NULL "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " db_shop.dbo.tbl_shopjumun_cardApp_log l "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and l.idx = " + CStr(logidx) + " "
	''sqlStr = sqlStr + " 	and l.shopJumunMasterIdx is not NULL "
	dbget.Execute sqlStr

	response.write "<script>alert('�����Ǿ����ϴ�.'); location.replace('" + CStr(refer) + "');</script>"
	dbget.close()
	response.end

elseif (mode="matchcancel") then

	sqlStr = " select top 1 a.idx "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_shop.dbo.tbl_shopjumun_cardApp_log c "
	sqlStr = sqlStr + " 	join db_shop.dbo.tbl_shopjumun_cardApp_log a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and c.cardAppNo = a.cardAppNo "
	''sqlStr = sqlStr + " 		and convert(VARCHAR(10), c.appDate, 127) = convert(VARCHAR(10), a.appDate, 127) "
	sqlStr = sqlStr + " 		and DateDiff(d,a.appDate,c.appDate) <= 7 "
	''sqlStr = sqlStr + " 		and c.shopid = a.shopid "
	sqlStr = sqlStr + " 		and ((c.shopid = a.shopid) or (a.shopid is NULL and c.cardReaderID = a.cardReaderID)) "
	sqlStr = sqlStr + " 		and c.cardPrice*-1 = a.cardPrice "
	sqlStr = sqlStr + " 		and c.appDivCode in ('C','P') "
	sqlStr = sqlStr + " 		and a.appDivCode = 'A' "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and c.idx = " + CStr(logidx) + " "
	sqlStr = sqlStr + " 	and a.orderserial is NULL "
	sqlStr = sqlStr + " 	and c.orderserial is NULL "

	matchidx = -1

    rsget.Open sqlStr,dbget,1
	if Not(rsget.EOF or rsget.BOF) then
	    matchidx = rsget("idx")
	end if
	rsget.Close

	if matchidx = -1 then
		response.write "<script>alert('����!!\n\n��Ī������ �����ϴ�[0].');</script>"
		response.write "��Ī������ �����ϴ�."
		dbget.close()
		response.end
	end if

	sqlStr = " update db_shop.dbo.tbl_shopjumun_cardApp_log "
	sqlStr = sqlStr + " set shopJumunMasterIdx = -1, orderserial = '��Ҹ�Ī' "
	sqlStr = sqlStr + " where idx in (" + CStr(logidx) + ", " + CStr(matchidx) + ") "
	dbget.Execute sqlStr

	response.write "<script>alert('��Ī�Ǿ����ϴ�.'); location.replace('" + CStr(refer) + "');</script>"
	dbget.close()
	response.end

elseif (mode="addActLog") then

	sqlStr = " select count(*) as cnt "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " db_order.dbo.tbl_onlineApp_log o "
	sqlStr = sqlStr + " join db_order.dbo.tbl_onlineApp_log a "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and o.pggubun = a.pggubun "
	sqlStr = sqlStr + " 	and o.pgkey = a.pgkey "
	sqlStr = sqlStr + " 	and o.pgcskey = Left(a.pgcskey, len(o.pgcskey)) "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 1 = 1 "
	sqlStr = sqlStr + " and o.idx = " + CStr(logidx) + " "

	PGCSkey = ""

    rsget.Open sqlStr,dbget,1
	if Not(rsget.EOF or rsget.BOF) then
	    PGCSkey = "-" + Format00(3, rsget("cnt"))
	end if
	rsget.Close

	sqlStr = " insert into db_order.dbo.tbl_onlineApp_log(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid, PGmeachulDate, orderserial) "
	sqlStr = sqlStr + " select top 1 t.PGgubun, t.PGkey, t.PGCSkey + '" + CStr(PGCSkey) + "', t.sitename, t.appDivCode, t.appMethod, t.appDate, t.cancelDate, 0, 0, 0, 0, t.ipkumdate, t.PGuserid, t.PGmeachulDate, t.orderserial "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log t "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and t.idx = " + CStr(logidx) + " "
	sqlStr = sqlStr + " 	and t.appPrice <> 0 "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	response.write "<script>alert('�߰��Ǿ����ϴ�.'); location.replace('" + CStr(refer) + "');</script>"
	dbget.close()
	response.end

elseif (mode="matchcancelOnline") then

	PGkey = requestCheckvar(request("PGkey"),64)
	force = requestCheckvar(request("force"),1)

	sqlStr = " select top 1 a.idx "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " db_order.dbo.tbl_onlineApp_log c "
	sqlStr = sqlStr + " join db_order.dbo.tbl_onlineApp_log a "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and c.pgkey = a.pgkey "
	if (force = "Y") then
		sqlStr = sqlStr + " 	and c.pgkey = '" & PGkey & "' "
	else
		sqlStr = sqlStr + " 	and (convert(VARCHAR(10), IsNull(c.appDate,c.cancelDate), 127) = convert(VARCHAR(10), a.appDate, 127) or a.pggubun = 'bankipkum') "		'// �������ڿ� ������ڰ� �ٸ� ���, �ּ�ó�� �� ��Ī�Ѵ�.
	end if
	sqlStr = sqlStr + " 	and IsNull(c.sitename, '') = IsNull(a.sitename, '') "
	sqlStr = sqlStr + " 	and c.appPrice*-1 = a.appPrice "
	sqlStr = sqlStr + " 	and c.appDivCode = 'C' "
	sqlStr = sqlStr + " 	and a.appDivCode = 'A' "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 1 = 1 "
	sqlStr = sqlStr + " and c.idx = " + CStr(logidx) + " "
    if (force <> "Y") then
	    sqlStr = sqlStr + " and c.orderserial is NULL "
    end if
    sqlStr = sqlStr + " and a.orderserial is NULL "
	''rw sqlStr

	matchidx = -1

    rsget.Open sqlStr,dbget,1
	if Not(rsget.EOF or rsget.BOF) then
	    matchidx = rsget("idx")
	end if
	rsget.Close

	if matchidx = -1 then
		response.write "<script>alert('����!!\n\n��Ī������ �����ϴ�[1]. �������ڿ� ������ڰ� �ٸ� ��� �����ּ���.');</script>"
		response.write "��Ī������ �����ϴ�. �������ڿ� ������ڰ� �ٸ� ��� �����ּ���."
		dbget.close()
		response.end
	end if

	sqlStr = " update db_order.dbo.tbl_onlineApp_log "
	sqlStr = sqlStr + " set orderserial = '��Ҹ�Ī' "
	sqlStr = sqlStr + " where idx in (" + CStr(logidx) + ", " + CStr(matchidx) + ") "
	dbget.Execute sqlStr

	sqlStr = " update db_order.dbo.tbl_onlineApp_log "
	sqlStr = sqlStr + " set csasid = -1 "
	sqlStr = sqlStr + " where idx = " + CStr(logidx) + " "
	dbget.Execute sqlStr

	response.write "<script>alert('��Ī�Ǿ����ϴ�.'); location.replace('" + CStr(refer) + "');</script>"
	dbget.close()
	response.end

elseif (mode="matchcancelOnlineDup") then

	sqlStr = " select top 1 a.idx "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " db_order.dbo.tbl_onlineApp_log c "
	sqlStr = sqlStr + " join db_order.dbo.tbl_onlineApp_log a "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and c.pgkey = a.pgkey "
	''sqlStr = sqlStr + " 	and convert(VARCHAR(10), IsNull(c.appDate, c.cancelDate), 127) = convert(VARCHAR(10), a.appDate, 127) "
	sqlStr = sqlStr + " 	and abs(datediff(d, convert(VARCHAR(10), IsNull(c.appDate, c.cancelDate), 127), convert(VARCHAR(10), a.appDate, 127))) <= 1 "
	sqlStr = sqlStr + " 	and c.sitename = a.sitename "
	sqlStr = sqlStr + " 	and c.appPrice*-1 = a.appPrice "
	sqlStr = sqlStr + " 	and c.appDivCode = 'C' "
	sqlStr = sqlStr + " 	and a.appDivCode = 'A' "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 1 = 1 "
	sqlStr = sqlStr + " and c.idx = " + CStr(logidx) + " "
	sqlStr = sqlStr + " and c.csasid is NULL "
	sqlStr = sqlStr + " and c.orderserial is NULL "		'// �ֹ���ȣ ���� ���
	sqlStr = sqlStr + " and c.idx > a.idx "
	''response.write sqlStr

	matchidx = -1

    rsget.Open sqlStr,dbget,1
	if Not(rsget.EOF or rsget.BOF) then
	    matchidx = rsget("idx")
	end if
	rsget.Close

	if matchidx = -1 then
		'// �ֹ���ȣ �ִ� ���
		sqlStr = " select top 1 a.idx "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " db_order.dbo.tbl_onlineApp_log c "
		sqlStr = sqlStr + " join db_order.dbo.tbl_onlineApp_log a "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and c.pgkey = a.pgkey "
		sqlStr = sqlStr + " 	and abs(datediff(d, convert(VARCHAR(10), IsNull(c.appDate, c.cancelDate), 127), convert(VARCHAR(10), a.appDate, 127))) <= 15 "
		sqlStr = sqlStr + " 	and c.sitename = a.sitename "
		sqlStr = sqlStr + " 	and c.appPrice*-1 = a.appPrice "
		sqlStr = sqlStr + " 	and c.appDivCode = 'C' "
		sqlStr = sqlStr + " 	and a.appDivCode = 'A' "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 1 = 1 "
		sqlStr = sqlStr + " and c.idx = " + CStr(logidx) + " "
		sqlStr = sqlStr + " and c.csasid is NULL "
		sqlStr = sqlStr + " and c.orderserial is not NULL "
		sqlStr = sqlStr + " and c.idx > a.idx "
		sqlStr = sqlStr + " and a.orderserial = c.orderserial "

		''response.write sqlStr
		matchidx = -1

		rsget.Open sqlStr,dbget,1
		if Not(rsget.EOF or rsget.BOF) then
			matchidx = rsget("idx")
		end if
		rsget.Close
	end if

	if matchidx = -1 then
		response.write "<script>alert('����!!\n\n��Ī������ �����ϴ�[2].');</script>"
		response.write "��Ī������ �����ϴ�."
		dbget.close()
		response.end
	end if

	sqlStr = " update db_order.dbo.tbl_onlineApp_log "
	sqlStr = sqlStr + " set csasid = -1, reasonGubun = NULL "
	sqlStr = sqlStr + " where idx = " + CStr(logidx) + " "
	dbget.Execute sqlStr

	sqlStr = " update db_order.dbo.tbl_onlineApp_log "
	sqlStr = sqlStr + " set reasonGubun = NULL "
	sqlStr = sqlStr + " where idx = " + CStr(matchidx) + " "
	dbget.Execute sqlStr

	response.write "<script>alert('��Ī�Ǿ����ϴ�.'); location.replace('" + CStr(refer) + "');</script>"
	dbget.close()
	response.end

elseif (mode="getonpgdata") then

	'// ========================================================================
	'// INICIS
	if (yyyymmdd = "") then

		searchipkumdateMAX = ""

		'// �ٹ��ϼ� ���� 4��(���� ����)
		sqlStr = " exec [db_cs].[dbo].[usp_getDayPlusWorkday_Inc_CurrDate_V2] '" & Left(now(), 10) & "', " & 4 & VbCRLF
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly
		if Not rsget.Eof then
			'// �ٹ��ϼ� ���� D+5 ��
			searchipkumdateMAX = rsget("plusworkday")
		end if
		rsget.close

		lastipkumdate = searchipkumdateMAX
		searchipkumdate = Left(DateSerial(Left(lastipkumdate, 4), Right(Left(lastipkumdate, 7), 2), (CLng(Right(Left(lastipkumdate, 10), 2)))), 10)

		Call GetAppFromInicis(searchipkumdate, objData, errMsg)

		if (errMsg <> "") then
			if  (Not IsAutoScript) then
				response.write "<script>alert('" & errMsg & "');</script>"
			end if
			response.write errMsg
			response.write objData
			dbget.close()
			response.end
		end if

		ipkumdate = Replace(searchipkumdate, "-", "")
		lastipkumdate = searchipkumdate
	else
		lastipkumdate = yyyymmdd

		searchipkumdate = Left(DateSerial(Left(lastipkumdate, 4), Right(Left(lastipkumdate, 7), 2), (CLng(Right(Left(lastipkumdate, 10), 2)))), 10)

		Call GetAppFromInicis(searchipkumdate, objData, errMsg)

		ipkumdate = Replace(searchipkumdate, "-", "")

		if (errMsg <> "") then
			if  (Not IsAutoScript) then
				response.write "<script>alert('" & errMsg & "');</script>"
			end if
			response.write errMsg
			response.write objData
			dbget.close()
			response.end
		end if
	end if
	''response.write objData
	''response.end

	objData = Split(objData, "<br>")

	'// ========================================================================
	sqlStr = " delete from db_temp.dbo.tbl_onlineApp_log_tmp "
	sqlStr = sqlStr & " where PGgubun in ('inicis','inirental') " & VbCRLF
	''response.write sqlStr
	dbget.execute sqlStr

	for i = 0 to UBound(objData)
		objLine = objData(i)
		objLine = Split(objLine, "|")

		if (objLine(0) = "B") then

			PGgubun			= "inicis"

			PGuserid = objLine(4)

			if (objLine(4) = "teenxteen3") then
				''sitename = "fingers"
                sitename = "wholesale"					'// 2022-04-13
			elseif (objLine(4) = "teenxteen4") then
				sitename = "10x10"
			elseif (objLine(4) = "teenxteen5") then
				sitename = "10x10"
			elseif (objLine(4) = "teenxteen6") then
				sitename = "10x10"
			elseif (objLine(4) = "teenxteen8") then
				sitename = "10x10gift"
			elseif (objLine(4) = "teenxteen9") then
				sitename = "10x10mobile"
            elseif (objLine(4) = "teenteensp") then
				sitename = "10x10"
			elseif (objLine(4) = "teenxteenr") then
				sitename = "10x10"
				'// ��Ż�� pgbun���� inirental�� �ٲ�
				PGgubun	= "inirental"
            elseif (objLine(4) = "teenteenap") then		' Apple Pay
				sitename = "10x10"
			else
				sitename = "XXX"
			end if

			if (objLine(11) = "A") then
				'// ==============================
				PGkey		= objLine(8)
				appDivCode	= "A"
				PGCSkey		= ""

				appDate			= objLine(12)
				cancelDate		= "NULL"
			elseif (objLine(11) = "C") then
				'// ==============================
				PGkey		= objLine(8)
				appDivCode	= "C"
				PGCSkey		= "CANCELALL"

				appDate			= objLine(12)
				cancelDate		= objLine(13)
			elseif (objLine(11) = "P") then
				'// ==============================
				'// �κ����
				PGkey		= objLine(9)
				appDivCode	= "R"
				PGCSkey		= objLine(8)

				appDate			= "NULL"
				cancelDate		= objLine(13)
			else
				'// ==============================
				PGkey		= objLine(8)
				appDivCode = "E"
				PGCSkey		= "ERROR"
			end if

			''appMethod		= objLine(3)

			if (objLine(3) = "CC") then
				appMethod = "100"
			elseif (objLine(3) = "AC") then
				appMethod = "20"
			elseif (objLine(3) = "VA") then
				appMethod = "7"
			elseif (objLine(3) = "RT") then
				appMethod = "150" '// �̴Ϸ�Ż
			else
				appMethod = objLine(3)
			end if

			appPrice		= objLine(16)
			commPrice		= objLine(17)
			commVatPrice	= objLine(18)
			jungsanPrice	= objLine(20)

			ipkumdate		= objLine(5)

			'// 20130503000623
			'// (2013-05-03 00:06:23)
			if (appDate <> "NULL") then
				appDate = "'" + Left(appDate, 4) + "-" + Right(Left(appDate, 6), 2) + "-" + Right(Left(appDate, 8), 2) + " " + Right(Left(appDate, 10), 2) + ":" + Right(Left(appDate, 12), 2) + ":" + Right(Left(appDate, 14), 2) + "'"
			end if

			if (cancelDate <> "NULL") then
				cancelDate = "'" + Left(cancelDate, 4) + "-" + Right(Left(cancelDate, 6), 2) + "-" + Right(Left(cancelDate, 8), 2) + " " + Right(Left(cancelDate, 10), 2) + ":" + Right(Left(cancelDate, 12), 2) + ":" + Right(Left(cancelDate, 14), 2) + "'"
			end if

			'// 20130510
			'// (2013-05-10)
			ipkumdate = Left(ipkumdate, 4) + "-" + Right(Left(ipkumdate, 6), 2) + "-" + Right(ipkumdate, 2)

			sqlStr = " insert into db_temp.dbo.tbl_onlineApp_log_tmp(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid) "
			sqlStr = sqlStr + " values('" + CStr(PGgubun) + "', '" + CStr(PGkey) + "', '" + CStr(PGCSkey) + "', '" + CStr(sitename) + "', '" + CStr(appDivCode) + "', '" + CStr(appMethod) + "', " + CStr(appDate) + ", " + CStr(cancelDate) + ", '" + CStr(appPrice) + "', '" + CStr(commPrice) + "', '" + CStr(commVatPrice) + "', '" + CStr(jungsanPrice) + "', '" + CStr(ipkumdate) + "', '" + CStr(PGuserid) + "') "
			''response.write sqlStr + "<br>"
			dbget.execute sqlStr

		end if
	next

	sqlStr = " insert into db_order.dbo.tbl_onlineApp_log(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid, PGmeachulDate) "
	sqlStr = sqlStr + " select t.PGgubun, t.PGkey, t.PGCSkey, t.sitename, t.appDivCode, t.appMethod, t.appDate, t.cancelDate, t.appPrice, t.commPrice, t.commVatPrice, t.jungsanPrice, t.ipkumdate, t.PGuserid, convert(varchar(10), isnull(t.cancelDate, t.appDate), 121) "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp t "
	sqlStr = sqlStr + " 	left join db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and l.PGgubun = t.PGgubun "
	sqlStr = sqlStr + " 		and l.PGkey = t.PGkey "
	sqlStr = sqlStr + " 		and l.PGCSkey = t.PGCSkey "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and l.idx is NULL "
	'sqlStr = sqlStr + " 	and t.PGgubun = 'inicis' "
	sqlStr = sqlStr + " 	and t.PGgubun in ('inicis','inirental') "
	sqlStr = sqlStr + " order by "
	sqlStr = sqlStr + " 	IsNull(t.cancelDate, t.appDate) "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	if  (Not IsAutoScript) then
	response.write "<script>alert('�Ա����� : " + CStr(searchipkumdate) + "');</script>"
	end if

elseif (mode="getpaycoT") Then

	'// ========================================================================
	'// ������ ���γ���
	'// ========================================================================

	''yyyymmdd = "2017-06-11"

	if (yyyymmdd = "") Then
		yyyymmdd = Left(DateAdd("d", -1, Now()),10)
	End If

	'// ���� : https://apis.krp.toastoven.net/paycobo/paycotrade/pgTradeFileDownload
	'// �׼� : https://dev-apis.krp.toastoven.net/paycobo/paycotrade/pgTradeFileDownload
	'// CSV ������ Response ����
	'// ?serviceCode=PAY_D&mrcCode=78NUHJ&token=78NUHJ-1C81DFC-ABA23-9EB95AE&version=1.0&ymd=20150101
	'// ?serviceCode=ST_PAY_D&mrcCode=78NUHJ&token=78NUHJ-1C81DFC-ABA23-9EB95AE&version=1.0&ymd=20150101

	ReDim xmlURLArr(2)
	xmlURLArr(0) = "https://apis.krp.toastoven.net/paycobo/paycotrade/pgTradeFileDownload?serviceCode=PAY_D&mrcCode=78NUHJ&token=78NUHJ-1C81DFC-ABA23-9EB95AE&version=1.0&ymd=" & Replace(yyyymmdd, "-", "")
	xmlURLArr(1) = "https://apis.krp.toastoven.net/paycobo/paycotrade/pgTradeFileDownload?serviceCode=PAY_D&mrcCode=RR0VR3&token=RR0VR3-8EA5C0D-768CA-5F33225&version=1.0&ymd=" & Replace(yyyymmdd, "-", "")
	xmlURLArr(2) = "https://apis.krp.toastoven.net/paycobo/paycotrade/pgTradeFileDownload?serviceCode=PAY_D&mrcCode=8973MQ&token=8973MQ-5CBF5E4-7B1A9-D8FD548&version=1.0&ymd=" & Replace(yyyymmdd, "-", "")

	objData = ""

	For Each xmlURL In xmlURLArr
		Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")

		objXML.setTimeouts 5 * 000, 5 * 000, 15 * 000, 45 * 000
		objXML.Open "GET", xmlURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()

		if objXML.Status = "200" And Len(objXML.ResponseText) > 0 Then
			objData = objData & vbLf & BinaryToText(objXML.ResponseBody, "UTF-8")
		else
		    response.write "NODATA:"&xmlURL
		end if

		Set objXML  = Nothing
	Next

	''response.write objData
	''response.end

	if (objData = "") then
		if  (Not IsAutoScript) then
			response.write "<script>alert('������ ����Ÿ�� �����ϴ�.[1]');</script>"
		end if
		response.write "������ ����Ÿ�� �����ϴ�[1]"
		dbget.close()
		response.end
	end If

	''response.Write objData

	objData = Split(objData, vbLf)

	''response.Write UBound(objData)

	'// ========================================================================
	sqlStr = " delete from db_temp.dbo.tbl_onlineApp_log_tmp "
	sqlStr = sqlStr & " where PGgubun = 'payco' " & VbCRLF
	''response.write sqlStr
	dbget.execute sqlStr

	for i = 0 to UBound(objData)
		objLine = objData(i)
		objLine = Split(objLine, Chr(9))		'// �ǹ���

		If (UBound(objLine) > 0) Then
			If (IsNumeric(objLine(0))) Then
				''response.Write objData(i) & "<br />"


				PGgubun			= "payco"
				PGuserid 		= "payco"			'// PGuserid, sitename �� tbl_order_PaymentEtc ���� �����;� ��
				sitename 		= "10x10"

				'// ���� : �ſ�ī��/������ ����/PAYCO ����Ʈ �� �ɰ����� ���´�. ���������� ���ľ� �Ѵ�.
				if (objLine(7) = "����") then
					'// ==============================
					PGkey		= objLine(1)
					appDivCode	= "A"
					PGCSkey		= ""

					appDate		= objLine(0)
					cancelDate	= "NULL"
				else
					'// ==============================
					'// �κ����(���/�κ���Ҵ� ���γ������� �ݾ׺񱳷� ã�ƾ� �Ѵ�.)
					PGkey		= objLine(1)
					appDivCode	= "R"
					PGCSkey		= objLine(3)

					appDate		= "NULL"
					cancelDate	= objLine(0)
				end If

				appMethod = "100"			'// �ſ�ī�常 �ִ�.

				appPrice		= objLine(5)
				commPrice		= 0
				commVatPrice	= 0
				jungsanPrice	= 0

				ipkumdate		= ""

				'// 20130503000623
				'// (2013-05-03 00:06:23)
				if (appDate <> "NULL") then
					appDate = "'" + Left(appDate, 4) + "-" + Right(Left(appDate, 6), 2) + "-" + Right(Left(appDate, 8), 2) + " " + Right(Left(appDate, 10), 2) + ":" + Right(Left(appDate, 12), 2) + ":" + Right(Left(appDate, 14), 2) + "'"
				end if

				if (cancelDate <> "NULL") then
					cancelDate = "'" + Left(cancelDate, 4) + "-" + Right(Left(cancelDate, 6), 2) + "-" + Right(Left(cancelDate, 8), 2) + " " + Right(Left(cancelDate, 10), 2) + ":" + Right(Left(cancelDate, 12), 2) + ":" + Right(Left(cancelDate, 14), 2) + "'"
				end If

				sqlStr = " if exists( "
				sqlStr = sqlStr + " 	select 1 "
				sqlStr = sqlStr + " 	from db_temp.dbo.tbl_onlineApp_log_tmp "
				sqlStr = sqlStr + " 	where PGgubun = 'payco' and PGkey = '" & PGkey & "' and appDivCode = '" & appDivCode & "' and ((cancelDate is not NULL and cancelDate = " & cancelDate & ") or (appDate is not NULL and appDate = " & appDate & ")) "
				sqlStr = sqlStr + " ) "
				sqlStr = sqlStr + " begin "
				sqlStr = sqlStr + " 	update db_temp.dbo.tbl_onlineApp_log_tmp "
				sqlStr = sqlStr + " 	set appPrice = appPrice + '" & appPrice & "' "
				sqlStr = sqlStr + " 	where PGgubun = 'payco' and PGkey = '" & PGkey & "' and appDivCode = '" & appDivCode & "' and ((cancelDate is not NULL and cancelDate = " & cancelDate & ") or (appDate is not NULL and appDate = " & appDate & ")) "
				sqlStr = sqlStr + " end "
				sqlStr = sqlStr + " else "
				sqlStr = sqlStr + " begin "
				sqlStr = sqlStr + " 	insert into db_temp.dbo.tbl_onlineApp_log_tmp(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid) "
				sqlStr = sqlStr + " 	values('" + CStr(PGgubun) + "', '" + CStr(PGkey) + "', '" + CStr(PGCSkey) + "', '" + CStr(sitename) + "', '" + CStr(appDivCode) + "', '" + CStr(appMethod) + "', " + CStr(appDate) + ", " + CStr(cancelDate) + ", '" + CStr(appPrice) + "', '" + CStr(commPrice) + "', '" + CStr(commVatPrice) + "', '" + CStr(jungsanPrice) + "', '" + CStr(ipkumdate) + "', '" + CStr(PGuserid) + "') "
				sqlStr = sqlStr + " End "
				''response.Write sqlStr & "<br />"
				dbget.execute sqlStr

			End If
		End If
	Next

	sqlStr = " insert into db_order.dbo.tbl_onlineApp_log(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid, PGmeachulDate) "
	sqlStr = sqlStr + " select t.PGgubun, t.PGkey, t.PGCSkey, t.sitename, t.appDivCode, t.appMethod, t.appDate, t.cancelDate, t.appPrice, t.commPrice, t.commVatPrice, t.jungsanPrice, t.ipkumdate, t.PGuserid, convert(varchar(10), isnull(t.cancelDate, t.appDate), 121) "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp t "
	sqlStr = sqlStr + " 	left join db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and l.PGgubun = t.PGgubun "
	sqlStr = sqlStr + " 		and l.PGkey = t.PGkey "
	sqlStr = sqlStr + " 		and l.PGCSkey = t.PGCSkey "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and l.idx is NULL "
	sqlStr = sqlStr + " 	and t.PGgubun = 'payco' "
	sqlStr = sqlStr + " 	and t.appDivCode = 'A' "				'// ���γ�����
	sqlStr = sqlStr + " order by "
	sqlStr = sqlStr + " 	IsNull(t.cancelDate, t.appDate) "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	sqlStr = " update r "
	sqlStr = sqlStr + " set r.appDivCode = 'C', r.PGCSkey = 'CANCELALL' "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log a "
	sqlStr = sqlStr + " 	join db_temp.dbo.tbl_onlineApp_log_tmp r "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and a.PGgubun = r.PGgubun "
	sqlStr = sqlStr + " 		and a.PGkey = r.PGkey "
	sqlStr = sqlStr + " 		and a.appDivCode = 'A' "
	sqlStr = sqlStr + " 		and r.appDivCode <> 'A' "
	sqlStr = sqlStr + " 		and a.appPrice = r.appPrice*-1 "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	a.PGgubun = 'payco' "
	dbget.execute sqlStr

    '// ��ҳ����� ���� ���⵵ ��
	sqlStr = " update a "
	sqlStr = sqlStr + " set a.appDivCode = 'C', a.PGCSkey = 'CANCELALL' "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log a "
	sqlStr = sqlStr + " 	join db_temp.dbo.tbl_onlineApp_log_tmp r "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and a.PGgubun = r.PGgubun "
	sqlStr = sqlStr + " 		and a.PGkey = r.PGkey "
	sqlStr = sqlStr + " 		and a.appDivCode = 'R' "
	sqlStr = sqlStr + " 		and r.appDivCode = 'C' "
	sqlStr = sqlStr + " 		and a.appPrice = r.appPrice "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	a.PGgubun = 'payco' "
	dbget.execute sqlStr

	sqlStr = " insert into db_order.dbo.tbl_onlineApp_log(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid, PGmeachulDate) "
	sqlStr = sqlStr + " select t.PGgubun, t.PGkey, t.PGCSkey, t.sitename, t.appDivCode, t.appMethod, t.appDate, t.cancelDate, t.appPrice, t.commPrice, t.commVatPrice, t.jungsanPrice, t.ipkumdate, t.PGuserid, convert(varchar(10), isnull(t.cancelDate, t.appDate), 121) "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp t "
	sqlStr = sqlStr + " 	left join db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and l.PGgubun = t.PGgubun "
	sqlStr = sqlStr + " 		and l.PGkey = t.PGkey "
	sqlStr = sqlStr + " 		and ((l.PGCSkey = t.PGCSkey) or (l.PGCSkey = 'CANCELALL')) "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and l.idx is NULL "
	sqlStr = sqlStr + " 	and t.PGgubun = 'payco' "
	sqlStr = sqlStr + " order by "
	sqlStr = sqlStr + " 	IsNull(t.cancelDate, t.appDate) "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	if  (Not IsAutoScript) then
	response.write "<script>alert('�ŷ����� : " + CStr(yyyymmdd) + " [9]');</script>"
	end If


elseif (mode="getpaycoS") Then

	'// ========================================================================
	'// ������ ���곻��
	'// ========================================================================

	''yyyymmdd = "2017-06-13" ''�ּ�ó��..;;

	if (yyyymmdd = "") Then
        if (Hour(Now()) >= 8) then
            '// ������ ����
            yyyymmdd = Left(DateAdd("d", -1, Now()),10)
        else
            '// �������� ����
            yyyymmdd = Left(DateAdd("d", -2, Now()),10)   ''2016/12/23 d-2�� ���� ���� 4�ÿ� ������ ���µ���.
        end if
	End If

	'// ���� : https://apis.krp.toastoven.net/paycobo/paycotrade/pgTradeFileDownload
	'// �׼� : https://dev-apis.krp.toastoven.net/paycobo/paycotrade/pgTradeFileDownload
	'// CSV ������ Response ����
	'// ?serviceCode=PAY_D&mrcCode=78NUHJ&token=78NUHJ-1C81DFC-ABA23-9EB95AE&version=1.0&ymd=20150101
	'// ?serviceCode=ST_PAY_D&mrcCode=78NUHJ&token=78NUHJ-1C81DFC-ABA23-9EB95AE&version=1.0&ymd=20150101

	ReDim xmlURLArr(2)
	xmlURLArr(0) = "https://apis.krp.toastoven.net/paycobo/paycotrade/pgTradeFileDownload?serviceCode=SB_PAY_D&mrcCode=78NUHJ&token=78NUHJ-1C81DFC-ABA23-9EB95AE&version=1.0&ymd=" & Replace(yyyymmdd, "-", "")
	xmlURLArr(1) = "https://apis.krp.toastoven.net/paycobo/paycotrade/pgTradeFileDownload?serviceCode=SB_PAY_D&mrcCode=RR0VR3&token=RR0VR3-8EA5C0D-768CA-5F33225&version=1.0&ymd=" & Replace(yyyymmdd, "-", "")
	xmlURLArr(2) = "https://apis.krp.toastoven.net/paycobo/paycotrade/pgTradeFileDownload?serviceCode=SB_PAY_D&mrcCode=8973MQ&token=8973MQ-5CBF5E4-7B1A9-D8FD548&version=1.0&ymd=" & Replace(yyyymmdd, "-", "")


	For Each xmlURL In xmlURLArr
		Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")

		objXML.setTimeouts 5 * 000, 5 * 000, 15 * 000, 45 * 000
		objXML.Open "GET", xmlURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()

		if objXML.Status = "200" And Len(objXML.ResponseText) > 0 Then
			objData = objData & vbLf & BinaryToText(objXML.ResponseBody, "UTF-8")
		else
		    response.write "NODATA:"&xmlURL
		end if

		Set objXML  = Nothing
	Next

	if (objData = "") then
		if  (Not IsAutoScript) then
			response.write "<script>alert('������ ����Ÿ�� �����ϴ�.[1]');</script>"
		end if
		response.write "������ ����Ÿ�� �����ϴ�[1]"
		dbget.close()
		response.end
	end If

	''response.Write objData
	''response.End

	objData = Split(objData, vbLf)


	'// ========================================================================
	sqlStr = " delete from db_temp.dbo.tbl_onlineApp_log_tmp "
	sqlStr = sqlStr & " where PGgubun = 'payco' " & VbCRLF
	''response.write sqlStr
	dbget.execute sqlStr

	for i = 0 to UBound(objData)
		objLine = objData(i)
		objLine = Split(objLine, Chr(9))		'// �ǹ���

		If (UBound(objLine) > 0) Then
			If (IsNumeric(objLine(0))) Then
				''response.Write objData(i) & "<br />"

				PGgubun			= "payco"
				PGuserid 		= "payco"			'// PGuserid, sitename �� tbl_order_PaymentEtc ���� �����;� ��
				sitename 		= "10x10"

				'// ���� : �ſ�ī��/������ ����/PAYCO ����Ʈ �� �ɰ����� ���´�. ���������� ���ľ� �Ѵ�.
				if (objLine(14) = "����") then
					'// ==============================
					PGkey		= objLine(10)
					appDivCode	= "A"
					PGCSkey		= ""

					appDate		= objLine(1)
					cancelDate	= "NULL"
				else
					'// ==============================
					'// �κ����(���/�κ���Ҵ� ���γ������� �ݾ׺񱳷� ã�ƾ� �Ѵ�.)
					PGkey		= objLine(10)
					appDivCode	= "R"
					PGCSkey		= objLine(12)

					appDate		= "NULL"
					cancelDate	= objLine(1)
				end If

				appMethod = "100"			'// �ſ�ī�常 �ִ�.

				appPrice		= objLine(16)
				commPrice		= objLine(17)
				commVatPrice	= objLine(20)
				jungsanPrice	= objLine(21)

				ipkumdate		= objLine(0)

				'// 20130510
				'// (2013-05-10)
				ipkumdate = Left(ipkumdate, 4) + "-" + Right(Left(ipkumdate, 6), 2) + "-" + Right(ipkumdate, 2)

				'// 20130503000623
				'// (2013-05-03 00:06:23)
				if (appDate <> "NULL") then
					appDate = "'" + Left(appDate, 4) + "-" + Right(Left(appDate, 6), 2) + "-" + Right(Left(appDate, 8), 2) + " " + Right(Left(appDate, 10), 2) + ":" + Right(Left(appDate, 12), 2) + ":" + Right(Left(appDate, 14), 2) + "'"
				end if

				if (cancelDate <> "NULL") then
					cancelDate = "'" + Left(cancelDate, 4) + "-" + Right(Left(cancelDate, 6), 2) + "-" + Right(Left(cancelDate, 8), 2) + " " + Right(Left(cancelDate, 10), 2) + ":" + Right(Left(cancelDate, 12), 2) + ":" + Right(Left(cancelDate, 14), 2) + "'"
				end If

				sqlStr = " if exists( "
				sqlStr = sqlStr + " 	select 1 "
				sqlStr = sqlStr + " 	from db_temp.dbo.tbl_onlineApp_log_tmp "
				sqlStr = sqlStr + " 	where PGgubun = 'payco' and PGkey = '" & PGkey & "' and appDivCode = '" & appDivCode & "' and ((cancelDate is not NULL and cancelDate = " & cancelDate & ") or (appDate is not NULL and appDate = " & appDate & ")) "
				sqlStr = sqlStr + " ) "
				sqlStr = sqlStr + " begin "
				sqlStr = sqlStr + " 	update db_temp.dbo.tbl_onlineApp_log_tmp "
				sqlStr = sqlStr + " 	set appPrice = appPrice + '" & appPrice & "', commPrice = commPrice + '" & commPrice & "', commVatPrice = commVatPrice + '" & commVatPrice & "', jungsanPrice = jungsanPrice + '" & jungsanPrice & "' "
				sqlStr = sqlStr + " 	where PGgubun = 'payco' and PGkey = '" & PGkey & "' and appDivCode = '" & appDivCode & "' and ((cancelDate is not NULL and cancelDate = " & cancelDate & ") or (appDate is not NULL and appDate = " & appDate & ")) "
				sqlStr = sqlStr + " end "
				sqlStr = sqlStr + " else "
				sqlStr = sqlStr + " begin "
				sqlStr = sqlStr + " 	insert into db_temp.dbo.tbl_onlineApp_log_tmp(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid) "
				sqlStr = sqlStr + " 	values('" + CStr(PGgubun) + "', '" + CStr(PGkey) + "', '" + CStr(PGCSkey) + "', '" + CStr(sitename) + "', '" + CStr(appDivCode) + "', '" + CStr(appMethod) + "', " + CStr(appDate) + ", " + CStr(cancelDate) + ", '" + CStr(appPrice) + "', '" + CStr(commPrice) + "', '" + CStr(commVatPrice) + "', '" + CStr(jungsanPrice) + "', '" + CStr(ipkumdate) + "', '" + CStr(PGuserid) + "') "
				sqlStr = sqlStr + " End "
				''response.Write sqlStr & "<br />"
				dbget.execute sqlStr

			End If
		End If
	Next

	'// ������ ���곻�� ��, �κ���� ������ �Ǿ� ��ü��ҵǸ� ��ü��� ���곻�� �ѰǸ� �´�.
	sqlStr = " update r "
	sqlStr = sqlStr + " set r.appDivCode = 'C', r.PGCSkey = 'CANCELALL' "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log a "
	sqlStr = sqlStr + " 	join db_temp.dbo.tbl_onlineApp_log_tmp r "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and a.PGgubun = r.PGgubun "
	sqlStr = sqlStr + " 		and a.PGkey = r.PGkey "
	sqlStr = sqlStr + " 		and a.appDivCode = 'A' "
	sqlStr = sqlStr + " 		and r.appDivCode <> 'A' "
	sqlStr = sqlStr + " 		and a.appPrice = r.appPrice*-1 "
	''sqlStr = sqlStr + " 		and IsNull(a.cancelDate,a.appDate) = IsNull(r.cancelDate,r.appDate) "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	a.PGgubun = 'payco' "
	dbget.execute sqlStr

	sqlStr = " update l "
	sqlStr = sqlStr + " set l.commPrice = t.commPrice*-1, l.commVatPrice = t.commVatPrice*-1, l.jungsanPrice = t.jungsanPrice, l.ipkumdate = t.ipkumdate "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " db_temp.dbo.tbl_onlineApp_log_tmp t "
	sqlStr = sqlStr + " join db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and l.PGgubun = t.PGgubun "
	sqlStr = sqlStr + " 	and l.PGkey = t.PGkey "
	sqlStr = sqlStr + " 	and ((l.appDivCode = t.appDivCode) or (l.appDivCode = 'R' and t.appDivCode = 'C')) " '// ���� : 17020632889
	sqlStr = sqlStr + " 	and IsNull(l.cancelDate,l.appDate) = IsNull(t.cancelDate,t.appDate) "
	''sqlStr = sqlStr + " 	and l.appPrice = t.appPrice "			'// �ݾ��� �޶� �Է��Ѵ�.
	sqlStr = sqlStr + " where t.PGgubun = 'payco' "
	dbget.execute sqlStr

	sqlStr = " delete from db_temp.dbo.tbl_onlineApp_log_tmp "
	sqlStr = sqlStr & " where PGgubun = 'payco' " & VbCRLF
	''response.write sqlStr
	dbget.execute sqlStr

	if  (Not IsAutoScript) then
		response.write "<script>alert('�������� : " + CStr(yyyymmdd) + " [9]');</script>"
		''dbget.close()
		''response.end
	end If

	''response.Write "aaa"
	''response.end

elseif (mode="getonpgdatahppre") then

	'// ========================================================================
	'// INICIS �ڵ���(�����۾�)

	'// ========================================================================
	sqlStr = " delete from db_temp.dbo.tbl_onlineApp_log_tmp "
	sqlStr = sqlStr & " where PGgubun = 'inicis' " & VbCRLF
	''response.write sqlStr
	dbget.execute sqlStr

	if  (Not IsAutoScript) then
	response.write "<script>alert('OK');</script>"
	end If

elseif (mode="getonpgdatahp") Then

	Call StartTimer(0)

	if (gubun = "prevmonth") then
		'// ���� ���γ���
		i = 1
	else
		i = 2
	end if

	'// ========================================================================
	'// INICIS �ڵ���
	if (yyyymmdd = "") Then
		'// �ٴ����� ������ ������
		sqlStr = " exec [db_cs].[dbo].[usp_getLastWorkDay] '" & Left(DateAdd("m", i, now()), 10) & "'" & VbCRLF
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly
		if Not rsget.Eof then
			'// �ٴ����� ������ ������
			yyyymmdd = rsget("workday")
		end if
		rsget.close
	end if

	ipkumdate = Replace(yyyymmdd, "-", "")

	xmlURL = "https://iniweb.inicis.com/mall/cr/urlsvc/UrlSendExtraDc.jsp?urlid=teenteen10&passwd=cube1010??&date=" & ipkumdate & "&flgdate=P"

	objData = ""

	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")

	objXML.setTimeouts 5 * 1000, 5 * 1000, 90 * 1000, 90 * 1000
	objXML.Open "GET", xmlURL, false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXML.Send()

	if objXML.Status = "200" then
		objData = BinaryToText(objXML.ResponseBody, "euc-kr")
	end if

	Set objXML  = Nothing

	if (InStr(objData, "NO DATA") > 0) then
		if  (Not IsAutoScript) then
			response.write "<script>alert('������ ����Ÿ�� �����ϴ�.[1]');</script>"
		end if
		response.write "������ ����Ÿ�� �����ϴ�[1]"
		response.write objData
		response.end
	end if

	''Response.Write "Elapsed time was: " & StopTimer(0)
	''dbget.Close()
	''Response.End

	objData = Split(objData, "<br>")

	'// ========================================================================
	sqlStr = " delete from db_temp.dbo.tbl_onlineApp_log_tmp "
	sqlStr = sqlStr & " where PGgubun = 'inicis' " & VbCRLF
	''response.write sqlStr
	dbget.execute sqlStr

	sqlStr = ""
	for i = 0 to UBound(objData)
		objLine = objData(i)
		objLine = Split(objLine, "|")

		if (objLine(0) = "B") then

			PGgubun			= "inicis"

			PGuserid = objLine(4)

			if (objLine(4) = "teenxteen3") then
				''sitename = "fingers"
                sitename = "wholesale"					'// 2022-04-21
			elseif (objLine(4) = "teenxteen4") then
				sitename = "10x10"
			elseif (objLine(4) = "teenxteen5") then
				sitename = "10x10"
			elseif (objLine(4) = "teenxteen6") then
				sitename = "10x10"
			elseif (objLine(4) = "teenxteen8") then
				sitename = "10x10gift"
			elseif (objLine(4) = "teenxteen9") then
				sitename = "10x10mobile"
			elseif (objLine(4) = "teenteen10") then
				if (Left(objLine(8),6) = "INIMX_") Then
					sitename = "10x10mobile"
				Else
					sitename = "10x10"
				End If
			elseif (objLine(4) = "teenteenr") then
				if (Left(objLine(8),6) = "INIMX_") Then
					sitename = "10x10mobile"
				Else
					sitename = "10x10"
				End If
			elseif (objLine(4) = "teenteenap") then		' Apple Pay
				sitename = "10x10"
			else
				sitename = "XXX"
			end if

			if (objLine(11) = "A") then
				'// ==============================
				PGkey		= objLine(8)
				appDivCode	= "A"
				PGCSkey		= ""

				appDate			= objLine(12)
				cancelDate		= "NULL"
			elseif (objLine(11) = "C") then
				'// ==============================
				PGkey		= objLine(8)
				appDivCode	= "C"
				PGCSkey		= "CANCELALL"

				appDate			= objLine(12)
				cancelDate		= objLine(13)
			elseif (objLine(11) = "P") then
				'// ==============================
				'// �κ����
				PGkey		= objLine(9)
				appDivCode	= "R"
				PGCSkey		= objLine(8)

				appDate			= "NULL"
				cancelDate		= objLine(13)
			else
				'// ==============================
				PGkey		= objLine(8)
				appDivCode = "E"
				PGCSkey		= "ERROR"
			end if

			''appMethod		= objLine(3)

			if (objLine(3) = "CC") then
				appMethod = "100"
			elseif (objLine(3) = "AC") then
				appMethod = "20"
			elseif (objLine(3) = "VA") then
				appMethod = "7"
			elseif (objLine(3) = "MO") then
				appMethod = "400"
			else
				appMethod = objLine(3)
			end if

			appPrice		= objLine(16)
			commPrice		= objLine(17)
			commVatPrice	= objLine(18)
			jungsanPrice	= objLine(20)

			ipkumdate		= objLine(5)

			'// 20130503000623
			'// (2013-05-03 00:06:23)
			if (appDate <> "NULL") then
				appDate = "'" + Left(appDate, 4) + "-" + Right(Left(appDate, 6), 2) + "-" + Right(Left(appDate, 8), 2) + " " + Right(Left(appDate, 10), 2) + ":" + Right(Left(appDate, 12), 2) + ":" + Right(Left(appDate, 14), 2) + "'"
			end if

			if (cancelDate <> "NULL") then
				cancelDate = "'" + Left(cancelDate, 4) + "-" + Right(Left(cancelDate, 6), 2) + "-" + Right(Left(cancelDate, 8), 2) + " " + Right(Left(cancelDate, 10), 2) + ":" + Right(Left(cancelDate, 12), 2) + ":" + Right(Left(cancelDate, 14), 2) + "'"
			end if

			'// 20130510
			'// (2013-05-10)
			ipkumdate = Left(ipkumdate, 4) + "-" + Right(Left(ipkumdate, 6), 2) + "-" + Right(ipkumdate, 2)


			If (sqlStr = "") Then
				sqlStr = " insert into db_temp.dbo.tbl_onlineApp_log_tmp(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid) "
				sqlStr = sqlStr + " values('" + CStr(PGgubun) + "', '" + CStr(PGkey) + "', '" + CStr(PGCSkey) + "', '" + CStr(sitename) + "', '" + CStr(appDivCode) + "', '" + CStr(appMethod) + "', " + CStr(appDate) + ", " + CStr(cancelDate) + ", '" + CStr(appPrice) + "', '" + CStr(commPrice) + "', '" + CStr(commVatPrice) + "', '" + CStr(jungsanPrice) + "', '" + CStr(ipkumdate) + "', '" + CStr(PGuserid) + "') "
			Else
				sqlStr = sqlStr + ", ('" + CStr(PGgubun) + "', '" + CStr(PGkey) + "', '" + CStr(PGCSkey) + "', '" + CStr(sitename) + "', '" + CStr(appDivCode) + "', '" + CStr(appMethod) + "', " + CStr(appDate) + ", " + CStr(cancelDate) + ", '" + CStr(appPrice) + "', '" + CStr(commPrice) + "', '" + CStr(commVatPrice) + "', '" + CStr(jungsanPrice) + "', '" + CStr(ipkumdate) + "', '" + CStr(PGuserid) + "') "
			End If

			If (i <> 0) And ((i mod 500) = 0) Then
				''response.write sqlStr + "<br>"
				dbget.execute sqlStr

				sqlStr = ""
			End If
		end if
	Next

	If (sqlStr <> "") Then
		''response.write sqlStr + "<br>"
		dbget.execute sqlStr

		sqlStr = ""
	End If

	''rw "aaa" & Now()
	''dbget.close()
	''response.end

	sqlStr = " insert into db_order.dbo.tbl_onlineApp_log(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid, PGmeachulDate) "
	sqlStr = sqlStr + " select t.PGgubun, t.PGkey, t.PGCSkey, t.sitename, t.appDivCode, t.appMethod, t.appDate, t.cancelDate, t.appPrice, t.commPrice, t.commVatPrice, t.jungsanPrice, t.ipkumdate, t.PGuserid, convert(varchar(10), isnull(t.cancelDate, t.appDate), 121) "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp t "
	sqlStr = sqlStr + " 	left join db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and l.PGgubun = t.PGgubun "
	sqlStr = sqlStr + " 		and l.PGkey = t.PGkey "
	sqlStr = sqlStr + " 		and l.PGCSkey = t.PGCSkey "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and l.idx is NULL "
	sqlStr = sqlStr + " 	and t.PGgubun = 'inicis' "
	sqlStr = sqlStr + " order by "
	sqlStr = sqlStr + " 	IsNull(t.cancelDate, t.appDate) "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	if  (Not IsAutoScript) then
	response.write "<script>alert('�Ա����� : " + CStr(yyyymmdd) + " [" & StopTimer(0) & " sec]');</script>"
	end If

elseif (mode="getonpgdatakakaopayT") then
	'// ========================================================================
	'// īī��PAY(�ŷ����)

	'// C:/KMPay_jungsan/Report/cnstest22mT20150323.csv
	'// C:/KMPay_jungsan/Report/KCTEN0001gT20150818.csv

	''yyyymmdd = "20170309"

	If (yyyymmdd = "") Then
		'// ����
		yyyymmdd = Left(DateAdd("d", -1, Now()), 10)
	End If

	yyyymmdd = Replace(yyyymmdd, "-", "")
	yyyymmdd = Replace(yyyymmdd, ".", "")		'// ��ŷ���

	''yyyymmdd = "20150819"

	targetFileName = "C:/KMPay_jungsan/Report/KCTEN0001gT" & yyyymmdd & ".csv"
	''response.write targetFileName
	''targetFileName = "C:/KMPay_jungsan/Report/cnstest22mS20150323.csv"

	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

	If objFSO.FileExists(targetFileName) Then
		Set objOpenedFile = objFSO.OpenTextFile(targetFileName, ForReading)

		sqlStr = " delete from db_temp.dbo.tbl_onlineApp_log_tmp "
		sqlStr = sqlStr & " where PGgubun = 'kakaopay' " & VbCRLF
		''response.write sqlStr
		dbget.execute sqlStr

		Do Until objOpenedFile.AtEndOfStream
			objLine = objOpenedFile.ReadLine
			objLine = Split(objLine, ",")

			if (objLine(0) = "D") Then

				PGgubun			= "kakaopay"

				PGuserid = objLine(1)

				If False Then
					'// ���� ���� �����
					sitename = "10x10"
				Else
					sitename = "10x10mobile"
				End If

				'// A :����, C : ���(��ü��� or �κ����)
				Select Case objLine(3)
					Case "A"
						'// ==============================
						PGkey		= objLine(5)
						appDivCode	= "A"
						PGCSkey		= ""

						appDate		= objLine(2)
						cancelDate		= "NULL"
					Case "C"
						'// ==============================
						PGkey		= objLine(5)
						appDivCode	= "C"
						PGCSkey		= "UNKNOWN"

						appDate		= "NULL"
						cancelDate		= objLine(2)
					Case Else
						'// ==============================
						PGkey		= objLine(5)
						appDivCode = "E"
						PGCSkey		= "ERROR"
				End Select

				If True Then
					'// ���� ī�������
					appMethod = "100"
				Else
					appMethod = "ERR"
				End If

				appPrice		= objLine(8)
				commPrice		= 0
				commVatPrice	= 0
				jungsanPrice	= 0

				If appDivCode <> "A" Then
					appPrice = appPrice * -1
				End If

				ipkumdate		= ""

				'// 20130503
				'// (2013-05-03)
				if (appDate <> "NULL") then
					appDate = "'" + Left(appDate, 4) + "-" + Right(Left(appDate, 6), 2) + "-" + Right(Left(appDate, 8), 2) + "'"
				end if

				if (cancelDate <> "NULL") then
					cancelDate = "'" + Left(cancelDate, 4) + "-" + Right(Left(cancelDate, 6), 2) + "-" + Right(Left(cancelDate, 8), 2) + "'"
				end if

				'// 20130510
				'// (2013-05-10)
				''ipkumdate = Left(ipkumdate, 4) + "-" + Right(Left(ipkumdate, 6), 2) + "-" + Right(ipkumdate, 2)

				sqlStr = " insert into db_temp.dbo.tbl_onlineApp_log_tmp(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid) "
				sqlStr = sqlStr + " values('" + CStr(PGgubun) + "', '" + CStr(PGkey) + "', '" + CStr(PGCSkey) + "', '" + CStr(sitename) + "', '" + CStr(appDivCode) + "', '" + CStr(appMethod) + "', " + CStr(appDate) + ", " + CStr(cancelDate) + ", '" + CStr(appPrice) + "', '" + CStr(commPrice) + "', '" + CStr(commVatPrice) + "', '" + CStr(jungsanPrice) + "', '" + CStr(ipkumdate) + "', '" + CStr(PGuserid) + "') "
				''response.write sqlStr + "<br>"
				dbget.execute sqlStr
			End If
		Loop

		objOpenedFile.Close
		Set objOpenedFile = Nothing

		if application("Svr_Info") <> "Dev" Then
			'// �׼��� ����Ÿ �����Ƿ� ���۵�

			'// ��ü��� or �κ����
			sqlStr = " update T "
			sqlStr = sqlStr + " set "
			sqlStr = sqlStr + " T.PGCSkey = (case when l.clogIdx is NULL then 'CANCELALL' else T.pgkey end) "
			sqlStr = sqlStr + " , T.appDivCode = (case when l.clogIdx is NULL then 'C' else 'R' End) "
			sqlStr = sqlStr + " , T.orderserial = (case when l.clogIdx is NULL then NULL else l.orderserial End) "
			sqlStr = sqlStr + " , T.cancelDate = (case when l.clogIdx is NULL then T.cancelDate else l.regdate end) "
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp T "
			sqlStr = sqlStr + " 	left join [db_order].[dbo].[tbl_card_cancel_log] l "
			sqlStr = sqlStr + " 	on "
			sqlStr = sqlStr + " 		T.pgkey = l.newtid "
			sqlStr = sqlStr + " where "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and T.pggubun = 'kakaopay' "
			sqlStr = sqlStr + " 	and T.appDivCode = 'C' "
			sqlStr = sqlStr + " 	and T.PGCSkey = 'UNKNOWN' "
			dbget.execute sqlStr

			'// �ֹ���ȣ, ��������
			sqlStr = " update T "
			sqlStr = sqlStr + " set T.orderserial = o.orderserial, T.appDate = (case when T.appDivCode in ('A', 'C') then o.ipkumdate else T.appDate end) "
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp T "
			sqlStr = sqlStr + " 	join db_order.dbo.tbl_order_master o "
			sqlStr = sqlStr + " on "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and T.PGgubun = 'kakaopay' "
			sqlStr = sqlStr + " 	and o.PGgubun = 'KA' "
			sqlStr = sqlStr + " 	and o.paygatetid = T.PGkey "
			sqlStr = sqlStr + " 	and o.jumundiv not in  ('6', '9') "
			sqlStr = sqlStr + " 	and o.ipkumdiv > 3 "
			sqlStr = sqlStr + " 	and T.orderserial is NULL "
			dbget.execute sqlStr

			'// ���ų���
			sqlStr = " update T "
			sqlStr = sqlStr + " set T.orderserial = o.orderserial, T.appDate = (case when T.appDivCode in ('A', 'C') then o.ipkumdate else T.appDate end) "
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp T "
			sqlStr = sqlStr + " 	join db_log.dbo.tbl_old_order_master_2003 o "
			sqlStr = sqlStr + " on "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and T.PGgubun = 'kakaopay' "
			sqlStr = sqlStr + " 	and o.PGgubun = 'KA' "
			sqlStr = sqlStr + " 	and o.paygatetid = T.PGkey "
			sqlStr = sqlStr + " 	and o.jumundiv not in  ('6', '9') "
			sqlStr = sqlStr + " 	and T.orderserial is NULL "
			dbget.execute sqlStr

			'// ��ü�������
			sqlStr = " update T "
			sqlStr = sqlStr + " set T.cancelDate = a.finishdate "
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp T "
			sqlStr = sqlStr + " 	join [db_cs].[dbo].[tbl_new_as_list] a "
			sqlStr = sqlStr + " 	on "
			sqlStr = sqlStr + " 		1 = 1 "
			sqlStr = sqlStr + " 		and T.orderserial = a.orderserial "
			sqlStr = sqlStr + " 		and T.appDivCode = 'C' "
			sqlStr = sqlStr + " 		and a.divcd = 'A007' "
			sqlStr = sqlStr + " 		and a.currstate = 'B007' "
			sqlStr = sqlStr + " 		and a.deleteyn <> 'Y' "
			dbget.execute sqlStr

			sqlStr = " update T "
			sqlStr = sqlStr + " set T.PGkey = o.paygatetid "
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp T "
			sqlStr = sqlStr + " 	join db_order.dbo.tbl_order_master o "
			sqlStr = sqlStr + " on "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and T.PGgubun = 'kakaopay' "
			sqlStr = sqlStr + " 	and o.PGgubun = 'KA' "
			sqlStr = sqlStr + " 	and o.orderserial = T.orderserial "
			sqlStr = sqlStr + " 	and o.jumundiv not in  ('6', '9') "
			sqlStr = sqlStr + " 	and T.PGkey = T.PGCSkey "
			dbget.execute sqlStr

			sqlStr = " update T "
			sqlStr = sqlStr + " set T.PGkey = o.paygatetid "
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp T "
			sqlStr = sqlStr + " 	join db_log.dbo.tbl_old_order_master_2003 o "
			sqlStr = sqlStr + " on "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and T.PGgubun = 'kakaopay' "
			sqlStr = sqlStr + " 	and o.PGgubun = 'KA' "
			sqlStr = sqlStr + " 	and o.orderserial = T.orderserial "
			sqlStr = sqlStr + " 	and o.jumundiv not in  ('6', '9') "
			sqlStr = sqlStr + " 	and T.PGkey = T.PGCSkey "
			dbget.execute sqlStr

		End If

		sqlStr = " insert into db_order.dbo.tbl_onlineApp_log(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid, PGmeachulDate, orderserial) "
		sqlStr = sqlStr + " select t.PGgubun, t.PGkey, t.PGCSkey, t.sitename, t.appDivCode, t.appMethod, t.appDate, t.cancelDate, t.appPrice, t.commPrice, t.commVatPrice, t.jungsanPrice, t.ipkumdate, t.PGuserid, convert(varchar(10), isnull(t.cancelDate, t.appDate), 121), t.orderserial "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp t "
		sqlStr = sqlStr + " 	left join db_order.dbo.tbl_onlineApp_log l "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and l.PGgubun = t.PGgubun "
		sqlStr = sqlStr + " 		and l.PGkey = t.PGkey "
		sqlStr = sqlStr + " 		and l.PGCSkey = t.PGCSkey "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and l.idx is NULL "
		sqlStr = sqlStr + " 	and t.PGgubun = 'kakaopay' "
		sqlStr = sqlStr + " order by "
		sqlStr = sqlStr + " 	IsNull(t.cancelDate, t.appDate) "
		''response.write sqlStr + "<br>"
		dbget.execute sqlStr

		' dbget.Close()
		' Response.end

		if  (Not IsAutoScript) then
			response.write "<script>alert('�ŷ����� : " + CStr(yyyymmdd) + "');</script>"
		end If

	Else
		if  (Not IsAutoScript) then
			response.write "<script>alert('�ŷ���� ������ �����ϴ�.[0]');</script>"
		end if
		response.write "�ŷ���� ������ �����ϴ�[0]" & targetFileName
		dbget.Close
		response.end
	End If

	Set objFSO = Nothing

elseif (mode="getonpgdatakakaopayS") then
	'// ========================================================================
	'// īī��PAY(�ŷ����)

	'// C:/KMPay_jungsan/Report/cnstest22mS20150323.csv
	'// C:/KMPay_jungsan/Report/KCTEN0001gS20150818.csv

	''yyyymmdd = "20170309"

	If (yyyymmdd = "") Then
		'// ����
		yyyymmdd = Left(DateAdd("d", -1, Now()), 10)
	End If

	yyyymmdd = Replace(yyyymmdd, "-", "")
	yyyymmdd = Replace(yyyymmdd, ".", "")		'// ��ŷ���

	''yyyymmdd = "20150827"

	targetFileName = "C:/KMPay_jungsan/Report/KCTEN0001gS" & yyyymmdd & ".csv"
	''targetFileName = "C:/KMPay_jungsan/Report/cnstest22mS20150323.csv"

	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

	If objFSO.FileExists(targetFileName) Then
		Set objOpenedFile = objFSO.OpenTextFile(targetFileName, ForReading)

		sqlStr = " delete from db_temp.dbo.tbl_onlineApp_log_tmp "
		sqlStr = sqlStr & " where PGgubun = 'kakaopay' " & VbCRLF
		''response.write sqlStr
		dbget.execute sqlStr

		Do Until objOpenedFile.AtEndOfStream
			objLine = objOpenedFile.ReadLine
			''rw objLine
			objLine = Split(objLine, ",")

			if (objLine(0) = "D") Then

				PGgubun			= "kakaopay"

				PGuserid = objLine(1)

				If False Then
					'// ���� ���� �����
					sitename = "10x10"
				Else
					sitename = "10x10mobile"
				End If

				'// A : ����, C : ���, P: �κ����
				Select Case objLine(2)
					Case "A"
						'// ==============================
						PGkey		= objLine(8)
						appDivCode	= "A"
						PGCSkey		= ""

						'// 20150303,160405
						'// 20130503000623
						'// (2013-05-03 00:06:23)
						appDate		= objLine(3) & objLine(4)
						''appDate = "'" + Left(appDate, 4) + "-" + Right(Left(appDate, 6), 2) + "-" + Right(Left(appDate, 8), 2) + " " + Right(Left(appDate, 10), 2) + ":" + Right(Left(appDate, 12), 2) + ":" + Right(Left(appDate, 14), 2) + "'"

						cancelDate		= "NULL"
					Case "C"
						'// ==============================
						PGkey		= objLine(8)
						appDivCode	= "C"
						PGCSkey		= "CANCELALL"

						appDate		= objLine(3) & objLine(4)
						''appDate = "'" + Left(appDate, 4) + "-" + Right(Left(appDate, 6), 2) + "-" + Right(Left(appDate, 8), 2) + " " + Right(Left(appDate, 10), 2) + ":" + Right(Left(appDate, 12), 2) + ":" + Right(Left(appDate, 14), 2) + "'"

						cancelDate		= objLine(5) & objLine(6)
						''cancelDate = "'" + Left(cancelDate, 4) + "-" + Right(Left(cancelDate, 6), 2) + "-" + Right(Left(cancelDate, 8), 2) + " " + Right(Left(cancelDate, 10), 2) + ":" + Right(Left(cancelDate, 12), 2) + ":" + Right(Left(cancelDate, 14), 2) + "'"
					Case "P"
						'// ==============================
						'// �κ����
						PGkey		= objLine(17)
						appDivCode	= "R"
						PGCSkey		= objLine(8)

						appDate			= "NULL"
						cancelDate		= objLine(5) & objLine(6)
						''cancelDate = "'" + Left(cancelDate, 4) + "-" + Right(Left(cancelDate, 6), 2) + "-" + Right(Left(cancelDate, 8), 2) + " " + Right(Left(cancelDate, 10), 2) + ":" + Right(Left(cancelDate, 12), 2) + ":" + Right(Left(cancelDate, 14), 2) + "'"
					Case Else
						'// ==============================
						PGkey		= objLine(8)
						appDivCode = "E"
						PGCSkey		= "ERROR"
				End Select

				If True Then
					'// ���� ī�������
					appMethod = "100"
				Else
					appMethod = "ERR"
				End If

				appPrice		= objLine(11)
				If (appDivCode <> "A") Then
					appPrice = appPrice * -1
				End If

				commPrice		= objLine(13)
				commVatPrice	= Round(1.0 * commPrice * (1.0/11))

				commPrice = commPrice - commVatPrice

				If (appDivCode = "A") Then
					commPrice = commPrice * -1
					commVatPrice = commVatPrice * -1
				End If

				jungsanPrice	= appPrice + (commPrice + commVatPrice)

				ipkumdate		= objLine(14)

				'// 20130503000623
				'// (2013-05-03 00:06:23)
				if (appDate <> "NULL") Then
					If (appDate = "") Then
						appDate = "NULL"
					Else
						appDate = "'" + Left(appDate, 4) + "-" + Right(Left(appDate, 6), 2) + "-" + Right(Left(appDate, 8), 2) + " " + Right(Left(appDate, 10), 2) + ":" + Right(Left(appDate, 12), 2) + ":" + Right(Left(appDate, 14), 2) + "'"
					End If
				end if

				if (cancelDate <> "NULL") then
					cancelDate = "'" + Left(cancelDate, 4) + "-" + Right(Left(cancelDate, 6), 2) + "-" + Right(Left(cancelDate, 8), 2) + " " + Right(Left(cancelDate, 10), 2) + ":" + Right(Left(cancelDate, 12), 2) + ":" + Right(Left(cancelDate, 14), 2) + "'"
				end if

				'// 20130510
				'// (2013-05-10)
				ipkumdate = Left(ipkumdate, 4) + "-" + Right(Left(ipkumdate, 6), 2) + "-" + Right(ipkumdate, 2)

				sqlStr = " insert into db_temp.dbo.tbl_onlineApp_log_tmp(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid) "
				sqlStr = sqlStr + " values('" + CStr(PGgubun) + "', '" + CStr(PGkey) + "', '" + CStr(PGCSkey) + "', '" + CStr(sitename) + "', '" + CStr(appDivCode) + "', '" + CStr(appMethod) + "', " + CStr(appDate) + ", " + CStr(cancelDate) + ", '" + CStr(appPrice) + "', '" + CStr(commPrice) + "', '" + CStr(commVatPrice) + "', '" + CStr(jungsanPrice) + "', '" + CStr(ipkumdate) + "', '" + CStr(PGuserid) + "') "
				''response.write sqlStr + "<br>"
				dbget.execute sqlStr
			End If
		Loop

		objOpenedFile.Close
		Set objOpenedFile = Nothing

		sqlStr = " update l "
		sqlStr = sqlStr + " set l.commPrice = T.commPrice, l.commVatPrice = T.commVatPrice, l.jungsanPrice = T.jungsanPrice, l.ipkumdate = T.ipkumdate "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp T "
		sqlStr = sqlStr + " 	join db_order.dbo.tbl_onlineApp_log l "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and T.PGgubun = 'kakaopay' "
		sqlStr = sqlStr + " 		and T.PGgubun = l.PGgubun "
		sqlStr = sqlStr + " 		and T.PGkey = l.PGkey "
		sqlStr = sqlStr + " 		and T.appDivCode = l.appDivCode "
		sqlStr = sqlStr + " 		and T.PGCSkey = l.PGCSkey "
		''response.write sqlStr + "<br>"
		dbget.execute sqlStr

		'// ���� ��ü��Ҵ� ������ �ȿ´�.
		sqlStr = " update db_order.dbo.tbl_onlineApp_log "
		sqlStr = sqlStr + " set jungsanPrice = appPrice, ipkumdate = convert(varchar(10), IsNull(cancelDate,appDate), 127) "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and PGgubun = 'kakaopay' "
		sqlStr = sqlStr + " 	and PGkey in ( "
		sqlStr = sqlStr + " 		select a.PGkey "
		sqlStr = sqlStr + " 		from "
		sqlStr = sqlStr + " 			db_order.dbo.tbl_onlineApp_log a "
		sqlStr = sqlStr + " 			join db_order.dbo.tbl_onlineApp_log c "
		sqlStr = sqlStr + " 			on "
		sqlStr = sqlStr + " 				1 = 1 "
		sqlStr = sqlStr + " 				and a.PGgubun = 'kakaopay' "
		sqlStr = sqlStr + " 				and a.PGgubun = c.PGgubun "
		sqlStr = sqlStr + " 				and a.PGkey = c.PGkey "
		sqlStr = sqlStr + " 				and a.appDivCode = 'A' "
		sqlStr = sqlStr + " 				and c.appDivCode = 'C' "
		sqlStr = sqlStr + " 				and a.PGCSkey = '' "
		sqlStr = sqlStr + " 				and c.PGCSkey = 'CANCELALL' "
		sqlStr = sqlStr + " 				and convert(varchar(10), a.appDate, 127) = convert(varchar(10), c.cancelDate, 127) "
		sqlStr = sqlStr + " 				and a.ipkumdate = '' "
		sqlStr = sqlStr + " 				and a.ipkumdate = c.ipkumdate "
		sqlStr = sqlStr + " 	) "
		sqlStr = sqlStr + " 	and ipkumdate = '' "
		''response.write sqlStr + "<br>"
		dbget.execute sqlStr

		if  (Not IsAutoScript) then
			response.write "<script>alert('�������� : " + CStr(yyyymmdd) + "');</script>"
		end If

	Else
		if  (Not IsAutoScript) then
			response.write "<script>alert('������ ������ �����ϴ�.[0]');</script>"
		end if
		response.write "������ ������ �����ϴ�[0]" & targetFileName
		dbget.Close
		response.end
	End If

	Set objFSO = Nothing

	''dbget.Close
	''response.end

elseif (mode="getonpgdatanewkakaopayT") then
	'// ========================================================================
	'// īī��PAY(�ŷ����)

	'// C:/KMPay_jungsan/Report/cnstest22mT20150323.csv
	'// C:/KMPay_jungsan/Report/KCTEN0001gT20150818.csv

	''yyyymmdd = "20170309"
	dim objKakaoPay, objKakaoPayFile

	ix=0
	If (yyyymmdd = "") Then
		'// ����
        response.write "���� 8�� �����̸� �������� ����Ÿ ����<br />"
        response.write "���� 8�� ���Ŀ� ������ ����Ÿ ���Ű���<br />"
		if (Hour(Now) < 8) then
			'// 8�� �����̸� ������
			yyyymmdd = Left(DateAdd("d", -2, Now()), 10)
		else
			yyyymmdd = Left(DateAdd("d", -1, Now()), 10)
		end if
	End If

	yyyymmdd = Replace(yyyymmdd, "-", "")
	yyyymmdd = Replace(yyyymmdd, ".", "")		'// ��ŷ���

	'�������� API (���� URL ��������)
	Set objKakaoPay = fnCallKakaoPayFileUrl(yyyymmdd, Status)
	if Status="200" then

		Set objKakaoPayFile = fnCallKakaoPayCheckList(objKakaoPay.url, resultStatus)
		if resultStatus="200" then
			'���� �б� �����̸� ������ ����
			sqlStr = "delete from db_temp.dbo.tbl_onlineApp_log_tmp" & VbCRLF
			sqlStr = sqlStr & " where PGgubun='newkakaopay'"
			dbget.execute sqlStr

			For ix=0 To objKakaoPayFile.data.length - 1

				PGgubun	= "newkakaopay"
				PGuserid = "newkakaopay"
				If False Then
					'// ���� ���� �����
					sitename = "10x10"
				Else
					sitename = "10x10mobile"
				End If
				'payment_detail_action_type
				'�ŷ����� ���� �󼼸� �ȳ��մϴ�.
				'PAYMENT_LUMP_SUM - �ϽúҰ���
				'PAYMENT_INSTALLMENT - �Һΰ���
				'ISSUE_SID - ������� ���
				'PAYMENT_ISSUE_SID - ������� ��� �� ����
				'PAYMENT_SUBSCRIPTION -�������
				'ALL_CANCELED - ��ü���
				'PART_CANCELED - �κ����
				'ALL_PART_CANCELED - ��ü�κ����
				'// A :����, C : ���(��ü��� or �κ����)
				Select Case objKakaoPayFile.data.get(ix).payment_action_type			'//payment_action_type ����Ÿ��(PAYMENT/CANCEL/ISSUED_SID)
					Case "PAYMENT"
						'// ==============================
						PGkey		= objKakaoPayFile.data.get(ix).tid					'//tid ���ι�ȣ
						appDivCode	= "A"
						PGCSkey		= ""

						appDate		= objKakaoPayFile.data.get(ix).approved_at			'//approved_at ���� �Ͻ�
						cancelDate		= "NULL"
					Case "CANCEL"
						'// ==============================
						PGkey		= objKakaoPayFile.data.get(ix).tid
						if objKakaoPayFile.data.get(ix).payment_detail_action_type="ALL_CANCELED" then
							appDivCode	= "C"
							PGCSkey		= "CANCELALL"
						else
							appDivCode	= "R"
							PGCSkey		= objKakaoPayFile.data.get(ix).aid
						end if
						appDate		= "NULL"
						cancelDate		= objKakaoPayFile.data.get(ix).approved_at			'//approved_at ��� �Ͻ�

					Case "ISSUED_SID"
						'// ==============================
						PGkey		= objKakaoPayFile.data.get(ix).tid					'//tid ���ι�ȣ
						appDivCode	= "A"
						PGCSkey		= ""

						appDate		= objKakaoPayFile.data.get(ix).approved_at			'//approved_at ���� �Ͻ�
						cancelDate		= "NULL"
					Case Else
						'// ==============================
						PGkey		= objKakaoPayFile.data.get(ix).tid
						appDivCode = "E"
						PGCSkey		= "ERROR"
				End Select
				If True Then
					'// ���� ī�������
					appMethod = "100"
				Else
					appMethod = "ERR"
				End If
				appPrice = objKakaoPayFile.data.get(ix).amount		'//amount ����or��ұݾ�
				commPrice		= 0
				commVatPrice	= 0
				jungsanPrice	= 0
				If appDivCode <> "A" Then
					appPrice = appPrice * -1
				End If
				ipkumdate		= ""
				'// 20130503000623
				'// (2013-05-03 00:06:23)
				if (appDate <> "NULL") then
					appDate = "'" & appDate & "'"
				end if
				if (cancelDate <> "NULL") then
					cancelDate = "'" & cancelDate & "'"
				end if
				sqlStr = " insert into db_temp.dbo.tbl_onlineApp_log_tmp(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid) "
				sqlStr = sqlStr + " values('" + CStr(PGgubun) + "', '" + CStr(PGkey) + "', '" + CStr(PGCSkey) + "', '" + CStr(sitename) + "', '" + CStr(appDivCode) + "', '" + CStr(appMethod) + "', " + CStr(appDate) + ", " + CStr(cancelDate) + ", '" + CStr(appPrice) + "', '" + CStr(commPrice) + "', '" + CStr(commVatPrice) + "', '" + CStr(jungsanPrice) + "', '" + CStr(ipkumdate) + "', '" + CStr(PGuserid) + "') "
				dbget.execute sqlStr

			Next

			if application("Svr_Info") <> "Dev" Then
				'// �׼��� ����Ÿ �����Ƿ� ���۵�

				'// ��ü��� or �κ����
				sqlStr = " update T "
				sqlStr = sqlStr + " set "
				sqlStr = sqlStr + " T.PGCSkey = (case when l.clogIdx is NULL then 'CANCELALL' else T.pgkey end) "
				sqlStr = sqlStr + " , T.appDivCode = (case when l.clogIdx is NULL then 'C' else 'R' End) "
				sqlStr = sqlStr + " , T.orderserial = (case when l.clogIdx is NULL then NULL else l.orderserial End) "
				sqlStr = sqlStr + " , T.cancelDate = (case when l.clogIdx is NULL then T.cancelDate else l.regdate end) "
				sqlStr = sqlStr + " from "
				sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp T "
				sqlStr = sqlStr + " 	left join [db_order].[dbo].[tbl_card_cancel_log] l "
				sqlStr = sqlStr + " 	on "
				sqlStr = sqlStr + " 		T.pgkey = l.newtid "
				sqlStr = sqlStr + " where "
				sqlStr = sqlStr + " 	1 = 1 "
				sqlStr = sqlStr + " 	and T.pggubun = 'newkakaopay' "
				sqlStr = sqlStr + " 	and T.appDivCode = 'C' "
				sqlStr = sqlStr + " 	and T.PGCSkey = 'UNKNOWN' "
				dbget.execute sqlStr

				'// �ֹ���ȣ, ��������
				sqlStr = " update T "
				sqlStr = sqlStr + " set T.orderserial = o.orderserial, T.appDate = (case when T.appDivCode in ('A', 'C') then o.ipkumdate else T.appDate end) "
				sqlStr = sqlStr + " from "
				sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp T "
				sqlStr = sqlStr + " 	join db_order.dbo.tbl_order_master o "
				sqlStr = sqlStr + " on "
				sqlStr = sqlStr + " 	1 = 1 "
				sqlStr = sqlStr + " 	and T.PGgubun = 'newkakaopay' "
				sqlStr = sqlStr + " 	and o.PGgubun = 'KK' "
				sqlStr = sqlStr + " 	and o.paygatetid = T.PGkey "
				sqlStr = sqlStr + " 	and o.jumundiv not in  ('6', '9') "
				sqlStr = sqlStr + " 	and o.ipkumdiv > 3 "
				sqlStr = sqlStr + " 	and T.orderserial is NULL "
				dbget.execute sqlStr

				orderserial = requestCheckvar(request("orderserial"),32)

				'// ���ų���
				sqlStr = " update T "
				sqlStr = sqlStr + " set T.orderserial = o.orderserial, T.appDate = (case when T.appDivCode in ('A', 'C') then o.ipkumdate else T.appDate end) "
				sqlStr = sqlStr + " from "
				sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp T "
				sqlStr = sqlStr + " 	join db_log.dbo.tbl_old_order_master_2003 o "
				sqlStr = sqlStr + " on "
				sqlStr = sqlStr + " 	1 = 1 "
				sqlStr = sqlStr + " 	and T.PGgubun = 'newkakaopay' "
				sqlStr = sqlStr + " 	and o.PGgubun = 'KK' "
				sqlStr = sqlStr + " 	and o.paygatetid = T.PGkey "
				sqlStr = sqlStr + " 	and o.jumundiv not in  ('6', '9') "
				sqlStr = sqlStr + " 	and T.orderserial is NULL "
				sqlStr = sqlStr + " 	and o.orderserial = '" & orderserial & "' "
				if (orderserial <> "") then
					dbget.execute sqlStr
				end if

				'// ��ü�������
				sqlStr = " update T "
				sqlStr = sqlStr + " set T.cancelDate = a.finishdate "
				sqlStr = sqlStr + " from "
				sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp T "
				sqlStr = sqlStr + " 	join [db_cs].[dbo].[tbl_new_as_list] a "
				sqlStr = sqlStr + " 	on "
				sqlStr = sqlStr + " 		1 = 1 "
				sqlStr = sqlStr + " 		and T.orderserial = a.orderserial "
				sqlStr = sqlStr + " 		and T.appDivCode = 'C' "
				sqlStr = sqlStr + " 		and a.divcd = 'A007' "
				sqlStr = sqlStr + " 		and a.currstate = 'B007' "
				sqlStr = sqlStr + " 		and a.deleteyn <> 'Y' "
				dbget.execute sqlStr

				sqlStr = " update T "
				sqlStr = sqlStr + " set T.PGkey = o.paygatetid "
				sqlStr = sqlStr + " from "
				sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp T "
				sqlStr = sqlStr + " 	join db_order.dbo.tbl_order_master o "
				sqlStr = sqlStr + " on "
				sqlStr = sqlStr + " 	1 = 1 "
				sqlStr = sqlStr + " 	and T.PGgubun = 'newkakaopay' "
				sqlStr = sqlStr + " 	and o.PGgubun = 'KK' "
				sqlStr = sqlStr + " 	and o.orderserial = T.orderserial "
				sqlStr = sqlStr + " 	and o.jumundiv not in  ('6', '9') "
				sqlStr = sqlStr + " 	and T.PGkey = T.PGCSkey "
				dbget.execute sqlStr

				sqlStr = " update T "
				sqlStr = sqlStr + " set T.PGkey = o.paygatetid "
				sqlStr = sqlStr + " from "
				sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp T "
				sqlStr = sqlStr + " 	join db_log.dbo.tbl_old_order_master_2003 o "
				sqlStr = sqlStr + " on "
				sqlStr = sqlStr + " 	1 = 1 "
				sqlStr = sqlStr + " 	and T.PGgubun = 'newkakaopay' "
				sqlStr = sqlStr + " 	and o.PGgubun = 'KK' "
				sqlStr = sqlStr + " 	and o.orderserial = T.orderserial "
				sqlStr = sqlStr + " 	and o.jumundiv not in  ('6', '9') "
				sqlStr = sqlStr + " 	and T.PGkey = T.PGCSkey "
				sqlStr = sqlStr + " 	and o.orderserial = '" & orderserial & "' "
				if (orderserial <> "") then
					dbget.execute sqlStr
				end if

			End If

			sqlStr = " insert into db_order.dbo.tbl_onlineApp_log(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid, PGmeachulDate, orderserial) "
			sqlStr = sqlStr + " select t.PGgubun, t.PGkey, t.PGCSkey, t.sitename, t.appDivCode, t.appMethod, t.appDate, t.cancelDate, t.appPrice, t.commPrice, t.commVatPrice, t.jungsanPrice, t.ipkumdate, t.PGuserid, convert(varchar(10), isnull(t.cancelDate, t.appDate), 121), t.orderserial "
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp t "
			sqlStr = sqlStr + " 	left join db_order.dbo.tbl_onlineApp_log l "
			sqlStr = sqlStr + " 	on "
			sqlStr = sqlStr + " 		1 = 1 "
			sqlStr = sqlStr + " 		and l.PGgubun = t.PGgubun "
			sqlStr = sqlStr + " 		and l.PGkey = t.PGkey "
			sqlStr = sqlStr + " 		and l.PGCSkey = t.PGCSkey "
			sqlStr = sqlStr + " where "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and l.idx is NULL "
			sqlStr = sqlStr + " 	and t.PGgubun = 'newkakaopay' "
			sqlStr = sqlStr + " order by "
			sqlStr = sqlStr + " 	IsNull(t.cancelDate, t.appDate) "
			''response.write sqlStr + "<br>"
			dbget.execute sqlStr

			if  (Not IsAutoScript) then
				response.write "<script>alert('�������� : " + CStr(yyyymmdd) + "');</script>"
				response.write "<script>location.replace('" + refer + "');</script>"
				dbget.close : response.End
			else
				response.write "�������� : " + CStr(yyyymmdd) + ""
				dbget.close : response.End
			end if
		else
			if  (Not IsAutoScript) then
				response.write "<script>alert('" & objKakaoPayFile.message & "');</script>"
				response.write "<script>location.replace('" + refer + "');</script>"
				dbget.close : response.End
			else
				response.write "" & objKakaoPayFile.message & ""
				dbget.close : response.End
			end if
		end if
	else
		if  (Not IsAutoScript) then
			response.write "<script>alert('" & objKakaoPay.message & "');</script>"
			response.write "<script>location.replace('" + refer + "');</script>"
    		response.End
		else
			response.write "" & objKakaoPay.message & ""
			dbget.close : response.End
		end if
	end if
	Set objKakaoPay = Nothing
	Set objKakaoPayFile = Nothing

elseif (mode="getonpgdatanewkakaopayS") then
	'// ========================================================================
	'// īī��PAY(�ŷ����)

	'// C:/KMPay_jungsan/Report/cnstest22mT20150323.csv
	'// C:/KMPay_jungsan/Report/KCTEN0001gT20150818.csv

	''yyyymmdd = "20170309"
	dim objKakaoPayJS, objKakaoPayFileJS

	ix=0
	If (yyyymmdd = "") Then
		'// ����
        response.write "���� 8�� �����̸� �������� ����Ÿ ����<br />"
        response.write "���� 8�� ���Ŀ� ������ ����Ÿ ���Ű���<br />"
		if (Hour(Now) < 8) then
			'// 8�� �����̸� ������
			yyyymmdd = Left(DateAdd("d", -2, Now()), 10)
		else
			yyyymmdd = Left(DateAdd("d", -1, Now()), 10)
		end if
	End If

	yyyymmdd = Replace(yyyymmdd, "-", "")
	yyyymmdd = Replace(yyyymmdd, ".", "")		'// ��ŷ���

	'�������� API (���� URL ��������)
	Set objKakaoPayJS = fnCallKakaoPaySettlementsFileUrl(yyyymmdd, Status)
	if Status="200" then

		Set objKakaoPayFileJS = fnCallKakaoPaySettlementsCheckList(objKakaoPayJS.url, resultStatus)
		if resultStatus="200" then
			'���� �б� �����̸� ������ ����
			sqlStr = "delete from db_temp.dbo.tbl_onlineApp_log_tmp" & VbCRLF
			sqlStr = sqlStr & " where PGgubun='newkakaopay'"
			dbget.execute sqlStr

			For ix=0 To objKakaoPayFileJS.data.length - 1

				PGgubun	= "newkakaopay"
				PGuserid = "newkakaopay"
				If False Then
					'// ���� ���� �����
					sitename = "10x10"
				Else
					sitename = "10x10mobile"
				End If
				'payment_detail_action_type
				'�ŷ����� ���� �󼼸� �ȳ��մϴ�.
				'PAYMENT_LUMP_SUM - �ϽúҰ���
				'PAYMENT_INSTALLMENT - �Һΰ���
				'ISSUE_SID - ������� ���
				'PAYMENT_ISSUE_SID - ������� ��� �� ����
				'PAYMENT_SUBSCRIPTION -�������
				'ALL_CANCELED - ��ü���
				'PART_CANCELED - �κ����
				'ALL_PART_CANCELED - ��ü�κ����
				'// A :����, C : ���(��ü��� or �κ����)
				Select Case objKakaoPayFileJS.data.get(ix).payment_action_type			'//payment_action_type ����Ÿ��(PAYMENT/CANCEL/ISSUED_SID)
					Case "PAYMENT"
						'// ==============================
						PGkey		= objKakaoPayFileJS.data.get(ix).tid					'//tid ���ι�ȣ
						appDivCode	= "A"
						PGCSkey		= ""
						appDate		= objKakaoPayFileJS.data.get(ix).approved_at			'//approved_at ���� �Ͻ�
						cancelDate		= "NULL"
					Case "CANCEL"
						Select Case objKakaoPayFileJS.data.get(ix).payment_detail_action_type
							Case "ALL_CANCELED"
								'// ��ü ���
								PGkey		= objKakaoPayFileJS.data.get(ix).tid
								appDivCode	= "C"
								PGCSkey		= "CANCELALL"
								appDate		= "NULL"
								cancelDate		= objKakaoPayFileJS.data.get(ix).approved_at			'//approved_at ��� �Ͻ�
							Case Else
								'// �κ� ���
								PGkey		= objKakaoPayFileJS.data.get(ix).tid
								appDivCode	= "R"
								PGCSkey		= objKakaoPayFileJS.data.get(ix).aid
								appDate		= "NULL"
								cancelDate	= objKakaoPayFileJS.data.get(ix).approved_at
							End Select
					Case Else
						'// ==============================
						PGkey		= objKakaoPayFileJS.data.get(ix).tid
						appDivCode = "E"
						PGCSkey		= "ERROR"
				End Select
				If True Then
					'// ���� ī�������
					appMethod = "100"
				Else
					appMethod = "ERR"
				End If
				appPrice = objKakaoPayFileJS.data.get(ix).amount					'//amount ����or��ұݾ�
				commPrice		= objKakaoPayFileJS.data.get(ix).fee				'//fee PG������ݾ�
				commVatPrice	= objKakaoPayFileJS.data.get(ix).fee_vat			'//fee_vat PG������ΰ����ݾ�
				jungsanPrice	= objKakaoPayFileJS.data.get(ix).amount_payable	'//amount_payable �Աݿ����ݾ�
				ipkumdate		= objKakaoPayFileJS.data.get(ix).deposit_date		'//deposit_date �Աݿ�����
				'// 20130503000623
				'// (2013-05-03 00:06:23)
				if (appDate <> "NULL") then
					appDate = "'" & appDate & "'"
				end if
				if (cancelDate <> "NULL") then
					cancelDate = "'" & cancelDate & "'"
				end if
				sqlStr = " insert into db_temp.dbo.tbl_onlineApp_log_tmp(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid) "
				sqlStr = sqlStr + " values('" + CStr(PGgubun) + "', '" + CStr(PGkey) + "', '" + CStr(PGCSkey) + "', '" + CStr(sitename) + "', '" + CStr(appDivCode) + "', '" + CStr(appMethod) + "', " + CStr(appDate) + ", " + CStr(cancelDate) + ", '" + CStr(appPrice) + "', '" + CStr(commPrice) + "', '" + CStr(commVatPrice) + "', '" + CStr(jungsanPrice) + "', '" + CStr(ipkumdate) + "', '" + CStr(PGuserid) + "') "
				dbget.execute sqlStr

			Next

				sqlStr = " update l "
				sqlStr = sqlStr + " set l.commPrice = T.commPrice, l.commVatPrice = T.commVatPrice, l.jungsanPrice = T.jungsanPrice, l.ipkumdate = T.ipkumdate "
				sqlStr = sqlStr + " from "
				sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp T "
				sqlStr = sqlStr + " 	join db_order.dbo.tbl_onlineApp_log l "
				sqlStr = sqlStr + " 	on "
				sqlStr = sqlStr + " 		1 = 1 "
				sqlStr = sqlStr + " 		and T.PGgubun = 'newkakaopay' "
				sqlStr = sqlStr + " 		and T.PGgubun = l.PGgubun "
				sqlStr = sqlStr + " 		and T.PGkey = l.PGkey "
				sqlStr = sqlStr + " 		and T.appDivCode = l.appDivCode "
				sqlStr = sqlStr + " 		and T.PGCSkey = l.PGCSkey "
				''response.write sqlStr + "<br>"
				dbget.execute sqlStr

				'// ���� ��ü��Ҵ� ������ �ȿ´�.
				sqlStr = " update db_order.dbo.tbl_onlineApp_log "
				sqlStr = sqlStr + " set jungsanPrice = appPrice, ipkumdate = convert(varchar(10), IsNull(cancelDate,appDate), 127) "
				sqlStr = sqlStr + " where "
				sqlStr = sqlStr + " 	1 = 1 "
				sqlStr = sqlStr + " 	and PGgubun = 'newkakaopay' "
				sqlStr = sqlStr + " 	and PGkey in ( "
				sqlStr = sqlStr + " 		select a.PGkey "
				sqlStr = sqlStr + " 		from "
				sqlStr = sqlStr + " 			db_order.dbo.tbl_onlineApp_log a "
				sqlStr = sqlStr + " 			join db_order.dbo.tbl_onlineApp_log c "
				sqlStr = sqlStr + " 			on "
				sqlStr = sqlStr + " 				1 = 1 "
				sqlStr = sqlStr + " 				and a.PGgubun = 'newkakaopay' "
				sqlStr = sqlStr + " 				and a.PGgubun = c.PGgubun "
				sqlStr = sqlStr + " 				and a.PGkey = c.PGkey "
				sqlStr = sqlStr + " 				and a.appDivCode = 'A' "
				sqlStr = sqlStr + " 				and c.appDivCode = 'C' "
				sqlStr = sqlStr + " 				and a.PGCSkey = '' "
				sqlStr = sqlStr + " 				and c.PGCSkey = 'CANCELALL' "
				sqlStr = sqlStr + " 				and convert(varchar(10), a.appDate, 127) = convert(varchar(10), c.cancelDate, 127) "
				sqlStr = sqlStr + " 				and a.ipkumdate = '' "
				sqlStr = sqlStr + " 				and a.ipkumdate = c.ipkumdate "
				sqlStr = sqlStr + " 	) "
				sqlStr = sqlStr + " 	and ipkumdate = '' "
				''response.write sqlStr + "<br>"
				dbget.execute sqlStr

			if  (Not IsAutoScript) then
				response.write "<script>alert('�������� : " + CStr(yyyymmdd) + "');</script>"
				response.write "<script>location.replace('" + refer + "');</script>"
				response.End
			else
				response.write "�������� : " + CStr(yyyymmdd) + ""
				dbget.close : response.End
			end if
		else
			if  (Not IsAutoScript) then
				response.write "<script>alert('" & objKakaoPayFileJS.message & "');</script>"
				response.write "<script>location.replace('" + refer + "');</script>"
				response.End
			else
				response.write "" & objKakaoPayFileJS.message & ""
				dbget.close : response.End
			end if
		end if
	else
		if  (Not IsAutoScript) then
			response.write "<script>alert('" & objKakaoPayJS.message & "');</script>"
			response.write "<script>location.replace('" + refer + "');</script>"
    		response.End
		else
			response.write "" & objKakaoPayJS.message & ""
			dbget.close : response.End
		end if
	end if
	Set objKakaoPayJS = Nothing
	Set objKakaoPayFileJS = Nothing

elseif (mode="getonpgdatachaipayT") then
	'// ========================================================================
	'// ����PAY(������)
	'// ========================================================================



	if (yyyymmdd = "") Then
		yyyymmdd = Left(DateAdd("d", -1, Now()),10)
		checkDate = yyyymmdd
		yyyymmdd = replace(yyyymmdd,"-","")
	else
		checkDate = yyyymmdd
		yyyymmdd = replace(yyyymmdd,"-","")
	End If
	dim settlementKey
	if (application("Svr_Info")="Dev") then
		settlementKey = "14zuGq"
	else
		settlementKey = "14zuGq"
	end if

	xmlURL= "https://settlement.chai.finance/prod/" & settlementKey & "/" & yyyymmdd & ".txt"
	'xmlURL = "https://settlement.chai.finance/staging/6bfe11f5-23ac-4028-a9f0-e1f5ba02f0ff/20190722.txt"
	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.setTimeouts 5 * 000, 5 * 000, 15 * 000, 45 * 000
	objXML.Open "GET", xmlURL, false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXML.Send()

	if objXML.Status = "200" And Len(objXML.ResponseText) > 0 Then
		objData = objData & vbLf & BinaryToText(objXML.ResponseBody, "UTF-8")
	else
		response.write "NODATA:"&xmlURL
	end if

	Set objXML  = Nothing

	if (objData = "") then
		if  (Not IsAutoScript) then
			response.write "<script>alert('������ ����Ÿ�� �����ϴ�.[1]');</script>"
		end if
		response.write "������ ����Ÿ�� �����ϴ�[1]"
		dbget.close()
		response.end
	end If

	'response.Write UBound(objData)
	'response.End

	objData = Split(objData,vbLf)

	'// ========================================================================
	sqlStr = " delete from db_temp.dbo.tbl_onlineApp_log_tmp "
	sqlStr = sqlStr & " where PGgubun = 'chai' " & VbCRLF
	dbget.execute sqlStr

	for i = 1 to UBound(objData)-1
		objLine = objData(i)
		objLine = Split(objLine,"|")		'// | ���й��� �и�

		If (UBound(objLine) > 0) Then
			PGgubun = "chai"
			PGuserid = "chai"
			sitename = "10x10"
			'�������� �ϴ�� ���� �����
			'publicAPIKey | �ŷ��߻��ð� | ������� | idempotencyKey | paymentId | �������� | �����ݾ�(���밪) | ���������ð� | ����ڵ� | ��� ������ | �����ݾ� | ���θ�� �д�� | �κ���ҿ��� | �ŷ����̵� | ���� ���̵�

			if (objLine(5) = "P") then 'P: ����, C: ���
				'// ==============================
				PGkey		= objLine(4)		'paymentId
				appDivCode	= "A"
				PGCSkey		= ""
				appDate		= objLine(1)
				cancelDate	= "NULL"
				lastPGkey=PGkey
			else
				PGkey		= objLine(4)		'paymentId
				appDate		= "NULL"
				cancelDate	= objLine(1)
				if objLine(12) then '// �κ���ҿ���
					appDivCode	= "R"
					PGCSkey		= objLine(13)	'�κ���� �ŷ����̵�
				else
					appDivCode	= "C"
					PGCSkey		= "CANCELALL"
				end if

			end If

			appMethod = "20"			'// �ſ�ī�常 �ִ�.

			appPrice = objLine(10)								   '// +,- �� ǥ��� ���� or ��ұݾ�
			commPrice = round(appPrice * 0.015)					 'PG������ݾ�
			commVatPrice = round(commPrice*0.1)	 				'PG������ΰ����ݾ�
			jungsanPrice = appPrice - (commPrice+commVatPrice)	'//�Աݿ����ݾ�
			ipkumdate = objLine(9)								   '//������ ������

			'20200527
			ipkumdate = Left(ipkumdate, 4) + "-" + Right(Left(ipkumdate, 6), 2) + "-" + Right(ipkumdate, 2)
			'20200513232556
			if (appDate <> "NULL") then
				appDate = "'" + Left(appDate, 4) + "-" + Right(Left(appDate, 6), 2) + "-" + Right(Left(appDate, 8), 2) + " " + Right(Left(appDate, 10), 2) + ":" + Right(Left(appDate, 12), 2) + ":" + Right(Left(appDate, 14), 2) + "'"
			end if
			if (cancelDate <> "NULL") then
				cancelDate = "'" + Left(cancelDate, 4) + "-" + Right(Left(cancelDate, 6), 2) + "-" + Right(Left(cancelDate, 8), 2) + " " + Right(Left(cancelDate, 10), 2) + ":" + Right(Left(cancelDate, 12), 2) + ":" + Right(Left(cancelDate, 14), 2) + "'"
			end If

			sqlStr = "insert into db_temp.dbo.tbl_onlineApp_log_tmp(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid) "
			sqlStr = sqlStr + " 	values('" + CStr(PGgubun) + "', '" + CStr(PGkey) + "', '" + CStr(PGCSkey) + "', '" + CStr(sitename) + "', '" + CStr(appDivCode) + "', '" + CStr(appMethod) + "', " + CStr(appDate) + ", " + CStr(cancelDate) + ", '" + CStr(appPrice) + "', '" + CStr(commPrice) + "', '" + CStr(commVatPrice) + "', '" + CStr(jungsanPrice) + "', '" + CStr(ipkumdate) + "', '" + CStr(PGuserid) + "') "
			''response.Write sqlStr & "<br />"
			dbget.execute sqlStr

		End If
	Next

	if application("Svr_Info") <> "Dev" Then
		'// �׼��� ����Ÿ �����Ƿ� ���۵�

		'// ��ü��� or �κ����
		sqlStr = " update T "
		sqlStr = sqlStr + " set "
		sqlStr = sqlStr + " T.PGCSkey = (case when l.clogIdx is NULL then 'CANCELALL' else T.pgkey end) "
		sqlStr = sqlStr + " , T.appDivCode = (case when l.clogIdx is NULL then 'C' else 'R' End) "
		sqlStr = sqlStr + " , T.orderserial = (case when l.clogIdx is NULL then NULL else l.orderserial End) "
		sqlStr = sqlStr + " , T.cancelDate = (case when l.clogIdx is NULL then T.cancelDate else l.regdate end) "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp T "
		sqlStr = sqlStr + " 	left join [db_order].[dbo].[tbl_card_cancel_log] l "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		T.pgkey = l.newtid "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and T.pggubun = 'chai' "
		sqlStr = sqlStr + " 	and T.appDivCode = 'C' "
		sqlStr = sqlStr + " 	and T.PGCSkey = 'UNKNOWN' "
		dbget.execute sqlStr

		'// �ֹ���ȣ, ��������
		sqlStr = " update T "
		sqlStr = sqlStr + " set T.orderserial = o.orderserial, T.appDate = (case when T.appDivCode in ('A', 'C') then o.ipkumdate else T.appDate end) "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp T "
		sqlStr = sqlStr + " 	join db_order.dbo.tbl_order_master o "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and T.PGgubun = 'chai' "
		sqlStr = sqlStr + " 	and o.PGgubun = 'CH' "
		sqlStr = sqlStr + " 	and o.paygatetid = T.PGkey "
		sqlStr = sqlStr + " 	and o.jumundiv not in  ('6', '9') "
		sqlStr = sqlStr + " 	and o.ipkumdiv > 3 "
		sqlStr = sqlStr + " 	and T.orderserial is NULL "
		dbget.execute sqlStr

		orderserial = requestCheckvar(request("orderserial"),32)

		'// ���ų���
		sqlStr = " update T "
		sqlStr = sqlStr + " set T.orderserial = o.orderserial, T.appDate = (case when T.appDivCode in ('A', 'C') then o.ipkumdate else T.appDate end) "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp T "
		sqlStr = sqlStr + " 	join db_log.dbo.tbl_old_order_master_2003 o "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and T.PGgubun = 'chai' "
		sqlStr = sqlStr + " 	and o.PGgubun = 'CH' "
		sqlStr = sqlStr + " 	and o.paygatetid = T.PGkey "
		sqlStr = sqlStr + " 	and o.jumundiv not in  ('6', '9') "
		sqlStr = sqlStr + " 	and T.orderserial is NULL "
		sqlStr = sqlStr + " 	and o.orderserial = '" & orderserial & "' "
		if (orderserial <> "") then
			dbget.execute sqlStr
		end if

		'// ��ü�������
		sqlStr = " update T "
		sqlStr = sqlStr + " set T.cancelDate = a.finishdate "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp T "
		sqlStr = sqlStr + " 	join [db_cs].[dbo].[tbl_new_as_list] a "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and T.orderserial = a.orderserial "
		sqlStr = sqlStr + " 		and T.appDivCode = 'C' "
		sqlStr = sqlStr + " 		and a.divcd = 'A007' "
		sqlStr = sqlStr + " 		and a.currstate = 'B007' "
		sqlStr = sqlStr + " 		and a.deleteyn <> 'Y' "
		dbget.execute sqlStr

		sqlStr = " update T "
		sqlStr = sqlStr + " set T.PGkey = o.paygatetid "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp T "
		sqlStr = sqlStr + " 	join db_order.dbo.tbl_order_master o "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and T.PGgubun = 'chai' "
		sqlStr = sqlStr + " 	and o.PGgubun = 'CH' "
		sqlStr = sqlStr + " 	and o.orderserial = T.orderserial "
		sqlStr = sqlStr + " 	and o.jumundiv not in  ('6', '9') "
		sqlStr = sqlStr + " 	and T.PGkey = T.PGCSkey "
		dbget.execute sqlStr

		sqlStr = " update T "
		sqlStr = sqlStr + " set T.PGkey = o.paygatetid "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp T "
		sqlStr = sqlStr + " 	join db_log.dbo.tbl_old_order_master_2003 o "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and T.PGgubun = 'chai' "
		sqlStr = sqlStr + " 	and o.PGgubun = 'CH' "
		sqlStr = sqlStr + " 	and o.orderserial = T.orderserial "
		sqlStr = sqlStr + " 	and o.jumundiv not in  ('6', '9') "
		sqlStr = sqlStr + " 	and T.PGkey = T.PGCSkey "
		sqlStr = sqlStr + " 	and o.orderserial = '" & orderserial & "' "
		if (orderserial <> "") then
			dbget.execute sqlStr
		end if

	End If

	sqlStr = " insert into db_order.dbo.tbl_onlineApp_log(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid, PGmeachulDate, orderserial) "
	sqlStr = sqlStr + " select t.PGgubun, t.PGkey, t.PGCSkey, t.sitename, t.appDivCode, t.appMethod, t.appDate, t.cancelDate, t.appPrice, t.commPrice, t.commVatPrice, t.jungsanPrice, t.ipkumdate, t.PGuserid, convert(varchar(10), isnull(t.cancelDate, t.appDate), 121), t.orderserial "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp t "
	sqlStr = sqlStr + " 	left join db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and l.PGgubun = t.PGgubun "
	sqlStr = sqlStr + " 		and l.PGkey = t.PGkey "
	sqlStr = sqlStr + " 		and l.PGCSkey = t.PGCSkey "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and l.idx is NULL "
	sqlStr = sqlStr + " 	and t.PGgubun = 'chai' "
	sqlStr = sqlStr + " order by "
	sqlStr = sqlStr + " 	IsNull(t.cancelDate, t.appDate) "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

'response.write testAppPrice &"<br>"
'response.write testCommPrice &"<br>"
'response.write testCommVatPrice &"<br>"
'response.write testJungsanPrice &"<br>"
'response.end

	if  (Not IsAutoScript) then
		response.write "<script>alert('�������� : " + CStr(yyyymmdd) + " [9]');</script>"
	end If

	''response.Write "aaa"
	''response.end
elseif (mode="getonpgdatachaipayS") then
	testAppPrice=0
	testCommPrice=0
	testCommVatPrice=0
	testJungsanPrice=0
	if (yyyymmdd = "") Then
		yyyymmdd = Left(DateAdd("d", -1, Now()),10)
		checkDate = yyyymmdd
		yyyymmdd = replace(yyyymmdd,"-","")
	else
		checkDate = yyyymmdd
		yyyymmdd = replace(yyyymmdd,"-","")
	End If
	checkEndDate = Left(DateAdd("d", 1, checkDate),10)
	'����ݾ� üũ
	sqlStr = "select PGkey, appPrice"
	sqlStr = sqlStr + " from [db_order].[dbo].[tbl_onlineApp_log]"
	sqlStr = sqlStr + " where isnull(cancelDate,appDate)>='" & checkDate & "'"
	sqlStr = sqlStr + " and isnull(cancelDate,appDate)<'" & checkEndDate & "'"
	rsget.Open sqlStr,dbget,1
	if Not(rsget.EOF or rsget.BOF) Then
		do until rsget.eof
			appPrice = rsget("appPrice")						   '// +,- �� ǥ��� ���� or ��ұݾ�
			commPrice = round(appPrice * 0.015)					 'PG������ݾ�
			commVatPrice = round(commPrice*0.1)	 				'PG������ΰ����ݾ�
			jungsanPrice = appPrice - (commPrice+commVatPrice)	'//�Աݿ����ݾ�
			'�ݾ� �հ� ���
			testAppPrice = testAppPrice + appPrice
			testCommPrice = testCommPrice + commPrice
			testCommVatPrice = testCommVatPrice + commVatPrice
			testJungsanPrice = testJungsanPrice + jungsanPrice
			lastPGkey = rsget("PGkey")
			rsget.moveNext
		loop
	end if
	rsget.Close

	'����ݾ� üũ
	sqlStr = "select 1 feeAmount, feeTaxAmount, settlementAmount"
	sqlStr = sqlStr + " from [db_temp].[dbo].[tbl_chai_Jungsan_temp]"
	sqlStr = sqlStr + " where referenceDate='" & checkDate & "'"
	rsget.Open sqlStr,dbget,1
	if Not(rsget.EOF or rsget.BOF) Then
		feeAmount = rsget("feeAmount")
		feeTaxAmount = rsget("feeTaxAmount")
		settlementAmount = rsget("settlementAmount")
	end if
	rsget.Close

	'������ üũ
	if testCommPrice > feeAmount then
		testCommPrice = (testCommPrice - feeAmount)*-1
	elseif feeAmount > testCommPrice then
		testCommPrice = feeAmount - testCommPrice
	else
		testCommPrice=0
	end if
	if testCommVatPrice > feeTaxAmount then
		testCommVatPrice = (testCommVatPrice - feeTaxAmount)*-1
	elseif feeTaxAmount > testCommVatPrice then
		testCommVatPrice = feeTaxAmount - testCommVatPrice
	else
		testCommVatPrice=0
	end if
	if testJungsanPrice > settlementAmount then
		testJungsanPrice = (testJungsanPrice - settlementAmount)*-1
	elseif settlementAmount > testJungsanPrice then
		testJungsanPrice = settlementAmount - testJungsanPrice
	else
		testJungsanPrice=0
	end if
	sqlStr = "update db_order.dbo.tbl_onlineApp_log"
	sqlStr = sqlStr + "	set commPrice = commPrice + '" & testCommPrice & "'"
	sqlStr = sqlStr + ", commVatPrice = commVatPrice + '" & testCommVatPrice & "'"
	sqlStr = sqlStr + ", jungsanPrice = jungsanPrice + '" & testJungsanPrice & "'"
	sqlStr = sqlStr + " where PGgubun = 'chai' and PGkey = '" & lastPGkey & "'"
	dbget.execute sqlStr

	if  (Not IsAutoScript) then
		response.write "<script>alert('�������� : " + CStr(yyyymmdd) + "');</script>"
		response.write "<script>location.replace('" + refer + "');</script>"
		response.End
	else
		response.write "�������� : " + CStr(yyyymmdd) + ""
		dbget.close : response.End
	end if

elseif (mode="getonpgdatakakaopay") then
	'// ========================================================================
	'// īī��PAY

	'// C:/KMPay_jungsan/Report/cnstest22mT20150323.csv
	'// C:/KMPay_jungsan/Report/KCTEN0001gT20150818.csv

	yyyymmdd = Replace(yyyymmdd, "-", "")
	yyyymmdd = Replace(yyyymmdd, ".", "")		'// ��ŷ���

	If (yyyymmdd = "") Then
		'// ����
		yyyymmdd = Left(DateAdd("d", -1, Now()), 10)
	End If

	targetFileName = "C:/KMPay_jungsan/Report/KCTEN0001gT" & yyyymmdd & ".csv"
	''targetFileName = "C:/KMPay_jungsan/Report/cnstest22mS20150323.csv"

	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

	If objFSO.FileExists(targetFileName) Then
		Set objOpenedFile = objFSO.OpenTextFile(targetFileName, ForReading)

		sqlStr = " delete from db_temp.dbo.tbl_onlineApp_log_tmp "
		sqlStr = sqlStr & " where PGgubun = 'kakaopay' " & VbCRLF
		''response.write sqlStr
		dbget.execute sqlStr

		Do Until objOpenedFile.AtEndOfStream
			objLine = objOpenedFile.ReadLine
			objLine = Split(objLine, ",")

			if (objLine(0) = "D") Then

				PGgubun			= "kakaopay"

				PGuserid = objLine(1)

				If False Then
					'// ���� ���� �����
					sitename = "10x10"
				Else
					sitename = "10x10mobile"
				End If

				'// A : ����, C : ���, P: �κ����
				Select Case objLine(2)
					Case "A"
						'// ==============================
						PGkey		= objLine(8)
						appDivCode	= "A"
						PGCSkey		= ""

						'// 20150303,160405
						'// 20130503000623
						'// (2013-05-03 00:06:23)
						appDate		= objLine(3) & objLine(4)
						''appDate = "'" + Left(appDate, 4) + "-" + Right(Left(appDate, 6), 2) + "-" + Right(Left(appDate, 8), 2) + " " + Right(Left(appDate, 10), 2) + ":" + Right(Left(appDate, 12), 2) + ":" + Right(Left(appDate, 14), 2) + "'"

						cancelDate		= "NULL"
					Case "C"
						'// ==============================
						PGkey		= objLine(8)
						appDivCode	= "C"
						PGCSkey		= "CANCELALL"

						appDate		= objLine(3) & objLine(4)
						''appDate = "'" + Left(appDate, 4) + "-" + Right(Left(appDate, 6), 2) + "-" + Right(Left(appDate, 8), 2) + " " + Right(Left(appDate, 10), 2) + ":" + Right(Left(appDate, 12), 2) + ":" + Right(Left(appDate, 14), 2) + "'"

						cancelDate		= objLine(5) & objLine(6)
						''cancelDate = "'" + Left(cancelDate, 4) + "-" + Right(Left(cancelDate, 6), 2) + "-" + Right(Left(cancelDate, 8), 2) + " " + Right(Left(cancelDate, 10), 2) + ":" + Right(Left(cancelDate, 12), 2) + ":" + Right(Left(cancelDate, 14), 2) + "'"
					Case "P"
						'// ==============================
						'// �κ����
						PGkey		= objLine(17)
						appDivCode	= "R"
						PGCSkey		= objLine(8)

						appDate			= "NULL"
						cancelDate		= objLine(5) & objLine(6)
						''cancelDate = "'" + Left(cancelDate, 4) + "-" + Right(Left(cancelDate, 6), 2) + "-" + Right(Left(cancelDate, 8), 2) + " " + Right(Left(cancelDate, 10), 2) + ":" + Right(Left(cancelDate, 12), 2) + ":" + Right(Left(cancelDate, 14), 2) + "'"
					Case Else
						'// ==============================
						PGkey		= objLine(8)
						appDivCode = "E"
						PGCSkey		= "ERROR"
				End Select

				If True Then
					'// ���� ī�������
					appMethod = "100"
				Else
					appMethod = "ERR"
				End If

				appPrice		= objLine(11)
				commPrice		= objLine(13)
				commVatPrice	= Round(1.0 * commPrice * (1.0/11))
				jungsanPrice	= appPrice - commPrice

				ipkumdate		= objLine(14)

				'// 20130503000623
				'// (2013-05-03 00:06:23)
				if (appDate <> "NULL") then
					appDate = "'" + Left(appDate, 4) + "-" + Right(Left(appDate, 6), 2) + "-" + Right(Left(appDate, 8), 2) + " " + Right(Left(appDate, 10), 2) + ":" + Right(Left(appDate, 12), 2) + ":" + Right(Left(appDate, 14), 2) + "'"
				end if

				if (cancelDate <> "NULL") then
					cancelDate = "'" + Left(cancelDate, 4) + "-" + Right(Left(cancelDate, 6), 2) + "-" + Right(Left(cancelDate, 8), 2) + " " + Right(Left(cancelDate, 10), 2) + ":" + Right(Left(cancelDate, 12), 2) + ":" + Right(Left(cancelDate, 14), 2) + "'"
				end if

				'// 20130510
				'// (2013-05-10)
				ipkumdate = Left(ipkumdate, 4) + "-" + Right(Left(ipkumdate, 6), 2) + "-" + Right(ipkumdate, 2)

				sqlStr = " insert into db_temp.dbo.tbl_onlineApp_log_tmp(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid) "
				sqlStr = sqlStr + " values('" + CStr(PGgubun) + "', '" + CStr(PGkey) + "', '" + CStr(PGCSkey) + "', '" + CStr(sitename) + "', '" + CStr(appDivCode) + "', '" + CStr(appMethod) + "', " + CStr(appDate) + ", " + CStr(cancelDate) + ", '" + CStr(appPrice) + "', '" + CStr(commPrice) + "', '" + CStr(commVatPrice) + "', '" + CStr(jungsanPrice) + "', '" + CStr(ipkumdate) + "', '" + CStr(PGuserid) + "') "
				''response.write sqlStr + "<br>"
				dbget.execute sqlStr
			End If
		Loop

		objOpenedFile.Close
		Set objOpenedFile = Nothing

		sqlStr = " insert into db_order.dbo.tbl_onlineApp_log(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid, PGmeachulDate) "
		sqlStr = sqlStr + " select t.PGgubun, t.PGkey, t.PGCSkey, t.sitename, t.appDivCode, t.appMethod, t.appDate, t.cancelDate, t.appPrice, t.commPrice, t.commVatPrice, t.jungsanPrice, t.ipkumdate, t.PGuserid, convert(varchar(10), isnull(t.cancelDate, t.appDate), 121) "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp t "
		sqlStr = sqlStr + " 	left join db_order.dbo.tbl_onlineApp_log l "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and l.PGgubun = t.PGgubun "
		sqlStr = sqlStr + " 		and l.PGkey = t.PGkey "
		sqlStr = sqlStr + " 		and l.PGCSkey = t.PGCSkey "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and l.idx is NULL "
		sqlStr = sqlStr + " 	and t.PGgubun = 'kakaopay' "
		sqlStr = sqlStr + " order by "
		sqlStr = sqlStr + " 	IsNull(t.cancelDate, t.appDate) "
		''response.write sqlStr + "<br>"
		dbget.execute sqlStr

		if  (Not IsAutoScript) then
			response.write "<script>alert('�ŷ����� : " + CStr(yyyymmdd) + "');</script>"
		end If

	Else
		if  (Not IsAutoScript) then
			response.write "<script>alert('������ ������ �����ϴ�.[0]');</script>"
		end if
		response.write "������ ������ �����ϴ�[0]" & targetFileName
		dbget.Close
		response.end
	End If

	Set objFSO = Nothing

elseif (mode="getonpgdatauplus") then

	'// ========================================================================
	'// UPLUS

	'// ����/�������
	 ''yyyymmdd = "2017-10-30"

	if (yyyymmdd = "") then
		lastipkumdate = "2012-12-31"

		'// ��������
		sqlStr = " select top 1 PGmeachulDate as lastipkumdate " & VbCRLF
		sqlStr = sqlStr & " from db_order.dbo.tbl_onlineApp_log " & VbCRLF
		sqlStr = sqlStr & " where PGgubun = 'uplus' " & VbCRLF
		sqlStr = sqlStr & " order by idx desc " & VbCRLF
		''response.write sqlStr

		rsget.Open sqlStr,dbget,1
		if Not(rsget.EOF or rsget.BOF) then
			if Not IsNull(rsget("lastipkumdate")) then
				lastipkumdate = rsget("lastipkumdate")
			end if
		end if
		rsget.Close

		''lastipkumdate = "2017-10-01"

        response.write "�������� �������� ���� ù ������(������ ��� ������ ������)<br />"
        response.write "���Գ����� ��ȸ�� ������ �ִ� 20��ġ ��ȸ<br />"
		for i = 0 to 20
			'// TODO : 20�� �̻� �Աݾ��� ������ ����
			searchipkumdate = Left(DateSerial(Left(lastipkumdate, 4), Right(Left(lastipkumdate, 7), 2), (CLng(Right(Left(lastipkumdate, 10), 2)) + 1)), 10)

			if False and (searchipkumdate >= Left(now, 10)) then
				if  (Not IsAutoScript) then
					response.write "<script>alert('������ ����Ÿ�� �����ϴ�.[" & i & "]');</script>"
				end if
				response.write "������ ����Ÿ�� �����ϴ�[00]" & searchipkumdate
				response.end
			end if

			ipkumdate = Replace(searchipkumdate, "-", "")

			'// ========================================================================
			'// �¶��� ���� ���곻��
			response.write "������(��������X)"&CStr(ipkumdate) & "<br />"
			xmlURL = "http://pgweb.uplus.co.kr/pg/wmp/outerpage/trxdown.jsp?mertid=tenbyten01&servicecode=ADJ&trxdate=" + CStr(ipkumdate) + "&key=2beec91670e1f2840a7ac80adde00e49"

			objData = ""

			Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")

			objXML.Open "GET", xmlURL, false
			objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			if (request.serverVariables("REMOTE_ADDR")="110.93.128.113") then  ''�߰� ��ġ�۾��� timeout �ø�..
			    objXML.setTimeouts 30000,60000,60000,60000 ''2016/08/21 �߰�
		    end if
			objXML.Send()

			if objXML.Status = "200" then
			    if (Trim(objXML.ResponseBody)<>"") then  ''�ƿ� ���ΰ�� 2016/09/13 �߰�
				    objData = BinaryToText(objXML.ResponseBody, "euc-kr")
			    end if
			end if

			Set objXML  = Nothing

			if (Replace(Trim(objData), vbCrLf, "") <> "") then
				exit for
			end if

			lastipkumdate = searchipkumdate

		next

		if (i >= 20) then
			if  (Not IsAutoScript) then
				response.write "<script>alert('������ ����Ÿ�� �����ϴ�.[" + CStr(i) + "a]');</script>"
			end if
			response.write "������ ����Ÿ�� �����ϴ�[1a]"
			response.end
		end if
	else
		'// ========================================================================
		'// �¶��� ���� ���곻��
		response.write "������(��������X):::"&CStr(Replace(yyyymmdd, "-", ""))
		xmlURL = "http://pgweb.uplus.co.kr/pg/wmp/outerpage/trxdown.jsp?mertid=tenbyten01&servicecode=ADJ&trxdate=" + CStr(Replace(yyyymmdd, "-", "")) + "&key=2beec91670e1f2840a7ac80adde00e49"
		objData = ""
		Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")

		objXML.Open "GET", xmlURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()

		if objXML.Status = "200" then
		    if (Trim(objXML.ResponseBody)="") then  ''2016/08/22 �߰�
		        response.write "NO_DATA"
		    else
			    objData = BinaryToText(objXML.ResponseBody, "euc-kr")
		    end if
		end if

		Set objXML  = Nothing

		''response.write "aaa" & Trim(objData)

		if (Replace(Trim(objData), vbCrLf, "") = "") then
			if  (Not IsAutoScript) then
				response.write "<script>alert('������ ����Ÿ�� �����ϴ�.[--]');</script>"
			end if
			response.write "������ ����Ÿ�� �����ϴ�[--]"
			response.end
		end if

		searchipkumdate = yyyymmdd
	end if

	''Response.Write objData + "<br>"
	''response.end

	objData = Split(objData, vbCrLf)

	'// ========================================================================
	sqlStr = " delete from db_temp.dbo.tbl_onlineApp_log_tmp "
	sqlStr = sqlStr & " where PGgubun = 'uplus' " & VbCRLF
	''response.write sqlStr
	dbget.execute sqlStr

 '' response.write xmlURL
 '' response.end

 	orderserial = requestCheckvar(request("orderserial"),32)
	if (orderserial = "") then
		'// �ߺ� �ֹ���ȣ
		orderserial = "XXXXXXXXX"
	end if

	prevPGkey = ""
	prevPrevPGkey = ""
	prevAppDivCode = ""
	prevPrevAppDivCode = ""
	IsDuplicate = False
	for i = 0 to UBound(objData)
		objLine = objData(i)
		objLine = Split(objLine, ";")

		if (UBound(objLine) >= 0) then
			if (objLine(0) = "D") then

				PGgubun		= "uplus"
				PGkey		= objLine(3)
				PGuserid 	= objLine(2)

				if (PGuserid = "dacomtest") then
					sitename = "dacomtest"
				elseif (PGuserid = "tenbyten01") or (PGuserid = "tenbyten02") then
					'// PC MOBILE ���� ����(�ֹ��������� �и�)
					sitename = "10x10"
				else
					sitename = "XXX"
				end if

				if (objLine(6) = "CA01") or (objLine(6) = "CS01") or (objLine(6) = "WR01") then
					'// ==============================
					appDivCode	= "A"
					PGCSkey		= ""

					appDate			= objLine(9)

					cancelDate		= "NULL"
				elseif (objLine(6) = "CA02") or (objLine(6) = "CS02") or (objLine(6) = "WR02") then
					'// ==============================
					appDivCode	= "C"
					PGCSkey		= "CANCELALL"

					appDate			= "NULL"
					cancelDate		= objLine(9)
				elseif (objLine(6) = "CA11") or (objLine(6) = "CS03") or (objLine(6) = "WR06") then
					'// ==============================
					'// �κ����
					'// �������ȯ���� �κ���ҿ� ��ü��Ҹ� ���� �ݾ����� �����ؾ��Ѵ�.
					appDivCode	= "R"
					PGCSkey		= objLine(9) + "-" + objLine(1)			'// �������� + �Ϸù�ȣ

					appDate			= "NULL"
					cancelDate		= objLine(9)
				else
					'// ==============================
					appDivCode = "E"
					PGCSkey		= "ERROR"
				end if

				if (Left(objLine(6), 2) = "CS") then
					appMethod = "7"
				elseif (Left(objLine(6), 2) = "WR") then
					appMethod = "400"
				elseif (Left(objLine(6), 2) = "CA") then
					appMethod = "100"
				else
					appMethod = Left(objLine(6), 2)
				end if

				appPrice		= objLine(7)
				commVatPrice	= round(1.0 * objLine(8) * (1.0/11))
				commPrice		= objLine(8) - commVatPrice
				jungsanPrice	= objLine(7) - objLine(8)

				commPrice = commPrice * -1
				commVatPrice = commVatPrice * -1

				ipkumdate		= objLine(10)

				'// 20130510
				'// (2013-05-10)
				if (appDate <> "NULL") then
					appDate = Left(appDate, 4) + "-" + Right(Left(appDate, 6), 2) + "-" + Right(appDate, 2)
					appDate = "'" + appDate + "'"
				end if

				if (cancelDate <> "NULL") then
					cancelDate = Left(cancelDate, 4) + "-" + Right(Left(cancelDate, 6), 2) + "-" + Right(cancelDate, 2)
					cancelDate = "'" + cancelDate + "'"
				end if

				'// 20130510
				'// (2013-05-10)
				ipkumdate = Left(ipkumdate, 4) + "-" + Right(Left(ipkumdate, 6), 2) + "-" + Right(ipkumdate, 2)

				''prevPGkey, prevAppDivCode, IsDuplicate

				if (i >= 1) then
					'// �ߺ� ����ó��(���� : 13020397762)
					'// TODO : ������ �ֹ���ȣ ������ ���ĵǾ� �ִٰ� �����Ѵ�.
					'// ���� : 13020397762, 13050293886, 13080752741, 16010731214, 16010731454

					IsDuplicate = False
					If (PGkey = prevPGkey) Then
						if (objLine(6) = "CS01") and (prevAppDivCode = "CS01") Then
							''�ߺ�����
							IsDuplicate = True
						elseif (objLine(6) = "CS02") and (prevAppDivCode = "CS02") Then
							''�ߺ����
							IsDuplicate = True
						elseif (prevPGkey = prevPrevPGkey) Then
							''3���̻�
							IsDuplicate = True
						End If
					End If

					if (prevPGkey <> "") then
						prevPrevPGkey = prevPGkey
						prevPrevAppDivCode = prevAppDivCode
					end if

					prevPGkey = PGkey
					prevAppDivCode = objLine(6)

					if (IsDuplicate = True) Or (PGkey = "21091168661") Or (PGkey = "16010512377") Or (PGkey = "16010731258") Or (PGkey = "20041296896") Or (PGkey = orderserial) then
						sqlStr = " select count(*) as cnt "
						sqlStr = sqlStr + " from "
                        if (PGkey = "21091168661") then
                            '// �ٸ� ��¥ �ߺ�
                            sqlStr = sqlStr + " [db_order].[dbo].[tbl_onlineApp_log] "
                        else
                            '// ���� ��¥ �ߺ�
                            sqlStr = sqlStr + " db_temp.dbo.tbl_onlineApp_log_tmp "
                        end if
						sqlStr = sqlStr + " where "
						sqlStr = sqlStr + " 1 = 1 "
						sqlStr = sqlStr + " and PGkey like '" + CStr(PGkey) + "%' and appDivCode = '" + appDivCode + "' "
						''response.write sqlStr

						subPgKey = ""
						rsget.Open sqlStr,dbget,1
						if Not(rsget.EOF or rsget.BOF) Then
							If rsget("cnt") > 0 Then
								subPgKey = "-" & Format00(2, rsget("cnt"))
							End If
						end if
						rsget.Close

						PGkey = PGkey + subPgKey
					end if
				end if

				sqlStr = " if not exists (select 1 from db_temp.dbo.tbl_onlineApp_log_tmp where pggubun = '" + CStr(PGgubun) + "' and pgkey = '" + CStr(PGkey) + "' and pgcskey = '" + CStr(PGCSkey) + "') "
				sqlStr = sqlStr + " begin "
				sqlStr = sqlStr + " 	insert into db_temp.dbo.tbl_onlineApp_log_tmp(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid) "
				sqlStr = sqlStr + " 	values('" + CStr(PGgubun) + "', '" + CStr(PGkey) + "', '" + CStr(PGCSkey) + "', '" + CStr(sitename) + "', '" + CStr(appDivCode) + "', '" + CStr(appMethod) + "', " + CStr(appDate) + ", " + CStr(cancelDate) + ", '" + CStr(appPrice) + "', '" + CStr(commPrice) + "', '" + CStr(commVatPrice) + "', '" + CStr(jungsanPrice) + "', '" + CStr(ipkumdate) + "', '" + CStr(PGuserid) + "') "
				sqlStr = sqlStr + " end "
				sqlStr = sqlStr + " else if not exists (select 1 from db_temp.dbo.tbl_onlineApp_log_tmp where pggubun = '" + CStr(PGgubun) + "' and pgkey = '" + CStr(PGkey) + "-01' and pgcskey = '" + CStr(PGCSkey) + "') "
				sqlStr = sqlStr + " begin "
				sqlStr = sqlStr + " 	insert into db_temp.dbo.tbl_onlineApp_log_tmp(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid) "
				sqlStr = sqlStr + " 	values('" + CStr(PGgubun) + "', '" + CStr(PGkey) + "-01', '" + CStr(PGCSkey) + "', '" + CStr(sitename) + "', '" + CStr(appDivCode) + "', '" + CStr(appMethod) + "', " + CStr(appDate) + ", " + CStr(cancelDate) + ", '" + CStr(appPrice) + "', '" + CStr(commPrice) + "', '" + CStr(commVatPrice) + "', '" + CStr(jungsanPrice) + "', '" + CStr(ipkumdate) + "', '" + CStr(PGuserid) + "') "
				sqlStr = sqlStr + " end "
				sqlStr = sqlStr + " else "
                sqlStr = sqlStr + " begin "
				sqlStr = sqlStr + " 	insert into db_temp.dbo.tbl_onlineApp_log_tmp(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid) "
				sqlStr = sqlStr + " 	values('" + CStr(PGgubun) + "', '" + CStr(PGkey) + "-02', '" + CStr(PGCSkey) + "', '" + CStr(sitename) + "', '" + CStr(appDivCode) + "', '" + CStr(appMethod) + "', " + CStr(appDate) + ", " + CStr(cancelDate) + ", '" + CStr(appPrice) + "', '" + CStr(commPrice) + "', '" + CStr(commVatPrice) + "', '" + CStr(jungsanPrice) + "', '" + CStr(ipkumdate) + "', '" + CStr(PGuserid) + "') "
				sqlStr = sqlStr + " end "

				''sqlStr = " insert into db_temp.dbo.tbl_onlineApp_log_tmp(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid) "
				''sqlStr = sqlStr + " values('" + CStr(PGgubun) + "', '" + CStr(PGkey) + "', '" + CStr(PGCSkey) + "', '" + CStr(sitename) + "', '" + CStr(appDivCode) + "', '" + CStr(appMethod) + "', " + CStr(appDate) + ", " + CStr(cancelDate) + ", '" + CStr(appPrice) + "', '" + CStr(commPrice) + "', '" + CStr(commVatPrice) + "', '" + CStr(jungsanPrice) + "', '" + CStr(ipkumdate) + "', '" + CStr(PGuserid) + "') "

				''if (PGkey <> "16010512377") and (PGkey <> "16010512377-01") then
				if PGkey <> "18122572222-11" and PGkey <> "17021377452" then
					''response.write sqlStr + "<br>"
					dbget.execute sqlStr
				end if
				''end if
			end if
		end if
	Next

	''response.end

	'// ���� : 16010731214
	sqlStr = " update t3 "
	sqlStr = sqlStr + " set t3.PGkey = t1.PGkey "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp t1 "
	sqlStr = sqlStr + " 	left join db_temp.dbo.tbl_onlineApp_log_tmp t2 "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and t1.pggubun = t2.pggubun "
	sqlStr = sqlStr + " 		and t1.PGkey = t2.PGkey "
	sqlStr = sqlStr + " 		and t2.PGCSkey = 'CANCELALL' "
	sqlStr = sqlStr + " 	join db_temp.dbo.tbl_onlineApp_log_tmp t3 "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and t1.pggubun = t3.pggubun "
	sqlStr = sqlStr + " 		and Left(t1.PGkey, 11) = t3.PGkey "
	sqlStr = sqlStr + " 		and t3.PGCSkey = 'CANCELALL' "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and t1.PGgubun = 'uplus' "
	sqlStr = sqlStr + " 	and Len(t1.PGkey) > 11 "
	sqlStr = sqlStr + " 	and t1.PGCSkey = '' "
	sqlStr = sqlStr + " 	and t2.PGkey is NULL "
	dbget.execute sqlStr

	sqlStr = " update db_temp.dbo.tbl_onlineApp_log_tmp "
	sqlStr = sqlStr + " set orderserial = pgkey "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and pggubun = 'uplus' "
	sqlStr = sqlStr + " 	and len(pgkey) < 20 "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'// �ӽ��ֹ���ȣ => �ֹ���ȣ
	sqlStr = " update t set t.orderserial = o.orderserial "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp t "
	sqlStr = sqlStr + " 	join ( "
	sqlStr = sqlStr + " 		select o.orderserial, Left(o.paygatetid, (charindex('|', o.paygatetid) - 1)) as paygatetid "
	sqlStr = sqlStr + " 		from db_order.dbo.tbl_order_master o "
	sqlStr = sqlStr + " 		where "
	sqlStr = sqlStr + " 			1 = 1 "
	sqlStr = sqlStr + " 			and o.accountdiv = '400' "
	sqlStr = sqlStr + " 			and o.jumundiv not in ('6', '9') "
	sqlStr = sqlStr + " 			and o.paygatetid is not NULL "
	sqlStr = sqlStr + " 			and charindex('|', o.paygatetid) > 0 "										'// ������ '|'
	sqlStr = sqlStr + " 			and datediff(m, o.ipkumdate, '" + CStr(searchipkumdate) + "') <= 2 "		'// ������ ���� �Ǵ� �̹���
	sqlStr = sqlStr + " 	) o "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		t.pgkey = o.paygatetid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and t.pggubun = 'uplus' "
	sqlStr = sqlStr + " 	and len(t.pgkey) >= 6 "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'// �ӽ��ֹ���ȣ => �ֹ���ȣ
	sqlStr = " update t set t.orderserial = o.orderserial "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp t "
	sqlStr = sqlStr + " 	join ( "
	sqlStr = sqlStr + " 		select o.orderserial, Left(o.paygatetid, (charindex(';', o.paygatetid) - 1)) as paygatetid "
	sqlStr = sqlStr + " 		from db_order.dbo.tbl_order_master o "
	sqlStr = sqlStr + " 		where "
	sqlStr = sqlStr + " 			1 = 1 "
	sqlStr = sqlStr + " 			and o.accountdiv = '400' "
	sqlStr = sqlStr + " 			and o.jumundiv not in ('6', '9') "
	sqlStr = sqlStr + " 			and o.paygatetid is not NULL "
	sqlStr = sqlStr + " 			and charindex(';', o.paygatetid) > 0 "										'// ������ ';'
	sqlStr = sqlStr + " 			and datediff(m, o.ipkumdate, '" + CStr(searchipkumdate) + "') <= 2 "		'// ������ ���� �Ǵ� �̹���
	sqlStr = sqlStr + " 	) o "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		t.pgkey = o.paygatetid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and t.pggubun = 'uplus' "
	sqlStr = sqlStr + " 	and len(t.pgkey) >= 6 "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'// �����
	sqlStr = " update t set t.sitename = '10x10mobile' "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp t "
	sqlStr = sqlStr + " 	join ( "
	sqlStr = sqlStr + " 		select o.orderserial "
	sqlStr = sqlStr + " 		from db_order.dbo.tbl_order_master o "
	sqlStr = sqlStr + " 		where "
	sqlStr = sqlStr + " 			1 = 1 "
	''sqlStr = sqlStr + " 			and o.accountdiv = '400' "
	sqlStr = sqlStr + " 			and o.jumundiv not in ('6', '9') "
	sqlStr = sqlStr + " 			and o.paygatetid is not NULL "
	''sqlStr = sqlStr + " 			and o.rdsite = 'mobile' "													'// �����
	sqlStr = sqlStr + " 			and o.beadaldiv in (4,5,7,8) "												'// �����(4:mobile, 5:mobile_link, 7:APP, 8:between ), 2015-07-13, skyer9
	sqlStr = sqlStr + " 			and datediff(m, o.ipkumdate, '" + CStr(searchipkumdate) + "') <= 2 "		'// ������ ���� �Ǵ� �̹���
	sqlStr = sqlStr + " 	) o "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		t.orderserial = o.orderserial "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and t.pggubun = 'uplus' "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'// PG�� ������
	sqlStr = " update db_temp.dbo.tbl_onlineApp_log_tmp "
	sqlStr = sqlStr + " set PGmeachulDate = convert(varchar(10), IsNull(cancelDate, appdate), 127) "
	sqlStr = sqlStr + " where pggubun = 'uplus' and PGmeachulDate is NULL "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'// �ǰ�����
	sqlStr = " update t set t.appdate = IsNull(o.ipkumdate, t.appdate) "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp t "
	sqlStr = sqlStr + " 	join db_order.dbo.tbl_order_master o "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		t.orderserial = o.orderserial "
	sqlStr = sqlStr + " where t.pggubun = 'uplus' and appDivCode = 'A' and o.jumundiv not in ('6', '9') "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'// �������
	sqlStr = " update t set t.cancelDate = a.finishdate "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp t "
	sqlStr = sqlStr + " 	join [db_cs].[dbo].tbl_new_as_list a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and t.orderserial = a.orderserial "
	sqlStr = sqlStr + " 		and t.appDivCode = 'C' "						'// ���
	sqlStr = sqlStr + " 		and a.divcd = 'A007' "
	sqlStr = sqlStr + " 	join [db_cs].dbo.tbl_as_refund_info r "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		a.id = r.asid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and a.deleteyn = 'N' "
	sqlStr = sqlStr + " 	and a.currstate = 'B007' "
	sqlStr = sqlStr + " 	and r.refundresult = (t.appPrice * -1) "
	sqlStr = sqlStr + " 	and t.PGgubun = 'uplus' "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'// �������(��ȯ�ֹ� ��ǰ)
	sqlStr = " update t set t.cancelDate = a.finishdate "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log t "
	sqlStr = sqlStr + " 	join db_order.dbo.tbl_change_order c "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		t.orderserial = c.orgorderserial "
	sqlStr = sqlStr + " 	join [db_cs].[dbo].tbl_new_as_list a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and c.chgorderserial = a.orderserial "
	sqlStr = sqlStr + " 		and t.appDivCode = 'C' "
	sqlStr = sqlStr + " 		and a.divcd = 'A007' "
	sqlStr = sqlStr + " 	join [db_cs].dbo.tbl_as_refund_info r "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		a.id = r.asid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and a.deleteyn = 'N' "
	sqlStr = sqlStr + " 	and a.currstate = 'B007' "
	sqlStr = sqlStr + " 	and r.refundresult = (t.appPrice * -1) "
	sqlStr = sqlStr + " 	and t.PGgubun = 'uplus' "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	''sqlStr = " delete from db_order.dbo.tbl_onlineApp_log where PGmeachulDate = '" + CStr(searchipkumdate) + "' "
	''response.write sqlStr + "<br>"
	''dbget.execute sqlStr

	sqlStr = " delete l "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp t "
	sqlStr = sqlStr + " 	join db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and t.PGgubun = 'uplus' "
	sqlStr = sqlStr + " 		and t.PGgubun = l.PGgubun "
	sqlStr = sqlStr + " 		and t.PGkey = l.PGkey "
	sqlStr = sqlStr + " 		and t.PGCSkey = l.PGCSkey "
	sqlStr = sqlStr + " 		and l.PGmeachulDate = '" + CStr(searchipkumdate) + "' "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	sqlStr = " insert into db_order.dbo.tbl_onlineApp_log(PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, PGuserid, PGmeachulDate, orderserial) "
	sqlStr = sqlStr + " select t.PGgubun, t.PGkey, t.PGCSkey, t.sitename, t.appDivCode, t.appMethod, t.appDate, t.cancelDate, t.appPrice, t.commPrice, t.commVatPrice, t.jungsanPrice, t.ipkumdate, t.PGuserid, t.PGmeachulDate, t.orderserial "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_temp.dbo.tbl_onlineApp_log_tmp t "
	sqlStr = sqlStr + " 	left join db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and l.PGgubun = t.PGgubun "
	sqlStr = sqlStr + " 		and l.PGkey = t.PGkey "
	sqlStr = sqlStr + " 		and l.PGCSkey = t.PGCSkey "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and l.idx is NULL "
	sqlStr = sqlStr + " 	and t.PGgubun = 'uplus' "
	sqlStr = sqlStr + " order by "
	sqlStr = sqlStr + " 	IsNull(t.cancelDate, t.appDate) "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'// ��Ī
	sqlStr = " update l set l.orderserial = o.orderserial "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	join db_order.dbo.tbl_order_master o "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		l.pgkey = o.orderserial "
	sqlStr = sqlStr + " where l.PGgubun = 'uplus' "
	''response.write sqlStr + "<br>"
	''dbget.execute sqlStr

	if  (Not IsAutoScript) then
		response.write "<script>alert('�ŷ����� : " + CStr(searchipkumdate) + "');</script>"
	end If

' �ٹ�����pg�ڵ���Ī.
elseif (mode = "matchpgdata") then

	sqlStr = " update l set l.orderserial = o.orderserial "
	sqlStr = sqlStr & " , l.sitename = (case when isnull(l.sitename,'')='' then (case when o.rdsite = 'mobile' or o.rdsite = 'app_wish2' then '10x10mobile' else '10x10' end) else l.sitename end)"
	sqlStr = sqlStr & " from db_order.dbo.tbl_order_master o "
	sqlStr = sqlStr & " join db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr & " 	on o.paygatetid = l.PGkey "
	sqlStr = sqlStr & " 	and l.PGgubun in ('inicis', 'payco', 'chai', 'inirental', 'convinienspay', 'naverpay') "
	sqlStr = sqlStr & " 	and o.paygatetid is not NULL "
	sqlStr = sqlStr & " 	and o.ipkumdate is not NULL "
	sqlStr = sqlStr & " 	and o.jumundiv <> '6' "			'// ��ȯ�ֹ� ����
	''sqlStr = sqlStr & " 	and ((l.appDivCode = 'A' and abs(DateDiff(mi, o.ipkumdate, l.appDate)) < 3) or (l.appDivCode <> 'A'))  "		'// 3��
	sqlStr = sqlStr & " 	and ((Left(o.paygatetid, 9) = 'IniTechPG') or (Left(o.paygatetid, 5) = 'INIMX') or (Left(o.paygatetid, 6) = 'INIswt') or (Left(o.paygatetid, 6) = 'Stdpay') or (Left(o.paygatetid, 5) = 'StdMX') or (Left(o.paygatetid, 10) = 'INIAPICARD') or (l.PGgubun = 'payco') or (l.PGgubun = 'chai') or (l.PGgubun = 'convinienspay') or (l.PGgubun = 'naverpay')) "
	sqlStr = sqlStr & " where l.orderserial is NULL "

	'response.write sqlStr & "<br>"
	dbget.execute sqlStr

	'' 6���� ���� �������� ��Ī
	sqlStr = "     select distinct top 1 l.PGkey  "
	sqlStr = sqlStr + " 	from "
	sqlStr = sqlStr + " 		db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	where "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and l.cancelDate >= DateAdd(day, -3, getdate()) "
	sqlStr = sqlStr + " 		and l.orderserial is NULL "
	sqlStr = sqlStr + " 		and l.appMethod <> '77' "

    PGkey = ""
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if Not(rsget.EOF or rsget.BOF) then
	    PGkey = rsget("PGkey")
	end if
	rsget.Close

	sqlStr = " update l set l.orderserial = o.orderserial "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " db_log.dbo.tbl_old_order_master_2003 o "
	sqlStr = sqlStr + " join db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and o.paygatetid = l.PGkey "
	sqlStr = sqlStr + " 	and l.PGgubun in ('inicis', 'payco', 'chai', 'convinienspay', 'naverpay') "
	sqlStr = sqlStr + " 	and o.paygatetid is not NULL "
	sqlStr = sqlStr + " 	and o.ipkumdate is not NULL "
	sqlStr = sqlStr + " 	and o.jumundiv <> '6' "				'// ��ȯ�ֹ� ����
	sqlStr = sqlStr + " 	and (l.appDivCode <> 'A')  "		'// ���γ����� 6���� �ڿ� �������� ���� ��Ҹ� �����Ƿ� ��Ҹ� ��Ī
    sqlStr = sqlStr + " 	and ((Left(o.paygatetid, 9) = 'IniTechPG') or (Left(o.paygatetid, 5) = 'INIMX') or (Left(o.paygatetid, 6) = 'INIswt') or (Left(o.paygatetid, 6) = 'Stdpay') or (Left(o.paygatetid, 5) = 'StdMX') or (Left(o.paygatetid, 10) = 'INIAPICARD') or (l.PGgubun = 'payco') or (l.PGgubun = 'chai') or (l.PGgubun = 'convinienspay') or (l.PGgubun = 'naverpay')) "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and l.orderserial is NULL "
	sqlStr = sqlStr + " 	and l.PGkey in ('" & PGkey & "') "

	'' �ϴ� ���� �ʿ��� �� ����� ��������(2014-09-05, skyer9)
	''response.write sqlStr + "<br>"
    if (PGkey <> "") then
	    dbget.execute sqlStr
    end if

	'// ���ֹ� �������
	sqlStr = " update l set l.csasid = a.id "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	join [db_cs].[dbo].tbl_new_as_list a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and l.orderserial = a.orderserial "
	sqlStr = sqlStr + " 		and l.appDivCode = 'C' "								'// ��Ҹ�
	sqlStr = sqlStr + " 		and a.divcd = 'A007' "
	sqlStr = sqlStr + " 		and a.finishdate >= DateAdd(d, -50, getdate()) "
	sqlStr = sqlStr + " 	join [db_cs].dbo.tbl_as_refund_info r "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		a.id = r.asid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and a.deleteyn = 'N' "
	sqlStr = sqlStr + " 	and a.currstate = 'B007' "
	sqlStr = sqlStr + " 	and abs(DateDiff(d, a.finishdate, l.cancelDate)) <= 1 "		'// ��Ҵ� �Ѱ��̹Ƿ� �Ϸ� ���̳��� ��Ī
	sqlStr = sqlStr + " 	and r.refundresult = (l.appPrice * -1) "
    sqlStr = sqlStr + " 	and Right(l.pgkey,2) <> '_1' "								'// ���̹�����Ʈ ����
	sqlStr = sqlStr + " 	and l.csasid is NULL "
	sqlStr = sqlStr + " 	and l.PGgubun in ('inicis', 'uplus', 'kakaopay', 'newkakaopay', 'naverpay', 'payco', 'chai', 'inirental') "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'// ���ֹ� �������(OK+�ſ�)
	sqlStr = " update l set l.csasid = a.id "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	join [db_cs].[dbo].tbl_new_as_list a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and l.orderserial = a.orderserial "
	sqlStr = sqlStr + " 		and l.appDivCode = 'C' "							'// ���
	sqlStr = sqlStr + " 		and a.divcd = 'A007' "
	sqlStr = sqlStr + " 		and a.finishdate >= DateAdd(d, -7, getdate()) "
	sqlStr = sqlStr + " 	join [db_cs].dbo.tbl_as_refund_info r "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		a.id = r.asid "
	sqlStr = sqlStr + " join db_order.dbo.tbl_order_master o "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	a.orderserial = o.orderserial "
	sqlStr = sqlStr + " join db_order.dbo.tbl_order_PaymentEtc e "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	o.orderserial = e.orderserial and e.acctdiv = '100' "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and a.deleteyn = 'N' "
	sqlStr = sqlStr + " 	and a.currstate = 'B007' "
	sqlStr = sqlStr + " 	and abs(DateDiff(d, a.finishdate, l.cancelDate)) <= 1 "		'// ��Ҵ� �Ѱ��̹Ƿ� �Ϸ� ���̳��� ��Ī
	sqlStr = sqlStr + " 	and e.realPayedsum = (l.appPrice * -1) "
	sqlStr = sqlStr + " 	and l.csasid is NULL "
	sqlStr = sqlStr + " 	and l.PGgubun in ('inicis', 'payco') "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'// ���ֹ� ���&��ǰ
	sqlStr = " update l set l.csasid = a.id "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	join [db_cs].[dbo].tbl_new_as_list a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and l.orderserial = a.orderserial "
	sqlStr = sqlStr + " 		and l.appDivCode in ('C', 'R') "						'// ���, �κ����
	sqlStr = sqlStr + " 		and a.divcd = 'A007' "
	sqlStr = sqlStr + " 		and a.finishdate >= DateAdd(d, -7, getdate()) "
	sqlStr = sqlStr + " 	join [db_cs].dbo.tbl_as_refund_info r "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		a.id = r.asid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and a.deleteyn = 'N' "
	sqlStr = sqlStr + " 	and a.currstate = 'B007' "
	sqlStr = sqlStr + " 	and abs(DateDiff(mi, a.finishdate, l.cancelDate)) < 2 "		'// 2��

	'/2���� ������� �ؿ��� 4��¥�� �ּ� ���� Ǯ�� �ְ� ������.
	'sqlStr = sqlStr + " 	and abs(DateDiff(d, a.finishdate, l.cancelDate)) < 4 "		'// 4��
	'sqlStr = sqlStr + " 	and l.canceldate >= '2016-12-01' "		'/��¥�� �ٲ��ְ�

	sqlStr = sqlStr + " 	and r.refundresult = (l.appPrice * -1) "
	sqlStr = sqlStr + " 	and l.csasid is NULL "
	sqlStr = sqlStr + " 	and l.PGgubun in ('inicis', 'uplus', 'kakaopay', 'newkakaopay', 'naverpay', 'payco', 'chai') "

	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'// ���ֹ� ���&��ǰ(OK+�ſ�)
	sqlStr = " update l set l.csasid = a.id "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	join [db_cs].[dbo].tbl_new_as_list a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and l.orderserial = a.orderserial "
	sqlStr = sqlStr + " 		and l.appDivCode = 'C' "							'// ���
	sqlStr = sqlStr + " 		and a.divcd = 'A007' "
	sqlStr = sqlStr + " 		and a.finishdate >= DateAdd(d, -7, getdate()) "
	sqlStr = sqlStr + " 	join [db_cs].dbo.tbl_as_refund_info r "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		a.id = r.asid "
	sqlStr = sqlStr + " join db_order.dbo.tbl_order_master o "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	a.orderserial = o.orderserial "
	sqlStr = sqlStr + " join db_order.dbo.tbl_order_PaymentEtc e "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	o.orderserial = e.orderserial and e.acctdiv = '100' "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and a.deleteyn = 'N' "
	sqlStr = sqlStr + " 	and a.currstate = 'B007' "
	sqlStr = sqlStr + " 	and abs(DateDiff(mi, a.finishdate, l.cancelDate)) < 2 "		'// 2��
	sqlStr = sqlStr + " 	and e.realPayedsum = (l.appPrice * -1) "
	sqlStr = sqlStr + " 	and l.csasid is NULL "
	sqlStr = sqlStr + " 	and l.PGgubun in ('inicis', 'payco') "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'// ��ȯ�ֹ� ��ǰ
	sqlStr = " update l set l.csasid = a.id "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	join db_order.dbo.tbl_change_order c "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		l.orderserial = c.orgorderserial and c.deldate is NULL "
	sqlStr = sqlStr + " 	join [db_cs].[dbo].tbl_new_as_list a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and c.chgorderserial = a.orderserial "
	sqlStr = sqlStr + " 		and l.appDivCode in ('C', 'R') "
	sqlStr = sqlStr + " 		and a.divcd = 'A007' "
	sqlStr = sqlStr + " 		and a.finishdate >= DateAdd(d, -30, getdate()) "
	sqlStr = sqlStr + " 	join [db_cs].dbo.tbl_as_refund_info r "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		a.id = r.asid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and a.deleteyn = 'N' "
	sqlStr = sqlStr + " 	and a.currstate = 'B007' "
	sqlStr = sqlStr + " 	and abs(DateDiff(mi, a.finishdate, l.cancelDate)) < 2 "
	sqlStr = sqlStr + " 	and r.refundresult = (l.appPrice * -1) "
	sqlStr = sqlStr + " 	and l.csasid is NULL "
	sqlStr = sqlStr + " 	and l.PGgubun in ('inicis', 'payco', 'naverpay') "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	''�ߺ���Ī Ȯ��
	'' select orderserial, count(*) as cnt
	'' from db_order.dbo.tbl_onlineApp_log
	'' where appDivCode = 'A'
	'' group by orderserial
	'' having count(*) > 1

	'' select orderserial, count(*) as cnt
	'' from db_order.dbo.tbl_onlineApp_log
	'' where appDivCode = 'C'
	'' group by orderserial
	'' having count(*) > 1

	'' select orderserial, csasid, count(*) as cnt
	'' from db_order.dbo.tbl_onlineApp_log
	'' where appDivCode <> 'A' and csasid is not NULL
	'' group by orderserial, csasid
	'' having count(*) > 1


	'// �κ�����̸鼭 �������� ����� ���
	'// cancelDate �� ������ ���� ���ڷ� �����ǰ� �ð��븸 �����ϰ� �����ȴ�.
	'// ���� �ð��븸 ���ؼ� ��Ī���ش�.
	sqlStr = " update l set l.csasid = a.id "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	join [db_cs].[dbo].tbl_new_as_list a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and l.orderserial = a.orderserial "
	sqlStr = sqlStr + " 		and l.appDivCode = 'R' "
	sqlStr = sqlStr + " 		and a.divcd = 'A007' "
	sqlStr = sqlStr + " 		and a.finishdate >= DateAdd(d, -30, getdate()) "
	sqlStr = sqlStr + " 	left join [db_cs].dbo.tbl_as_refund_info r "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		a.id = r.asid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and a.deleteyn = 'N' "
	sqlStr = sqlStr + " 	and a.currstate = 'B007' "
	sqlStr = sqlStr + " 	and abs(DateDiff(mi, a.finishdate, l.cancelDate) % (24 * 60)) < 2 "			'// ���� �ð���
	sqlStr = sqlStr + " 	and r.refundresult = (l.appPrice * -1) "
	sqlStr = sqlStr + " 	and l.csasid is NULL "
	sqlStr = sqlStr + " 	and l.PGgubun in ('inicis', 'payco', 'naverpay', 'chai') "
    ''sqlStr = sqlStr + " 	and l.orderserial in ('21040845030', '21040737318') "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

' �ٹ�����pg�ڵ���Ī. 6��������
elseif (mode = "matchpgdata6month") then
	'// 6�������� ���� ��Ī(PG Key �ִ� ��츸)

	PGkey = requestCheckvar(request("PGkey"),64)

	'' 6���� ���� �������� ��Ī
	sqlStr = " update l set l.orderserial = o.orderserial "
	sqlStr = sqlStr & " , l.sitename = (case when isnull(l.sitename,'')='' then (case when o.rdsite = 'mobile' or o.rdsite = 'app_wish2' then '10x10mobile' else '10x10' end) else l.sitename end)"
	sqlStr = sqlStr & " from db_log.dbo.tbl_old_order_master_2003 o "
	sqlStr = sqlStr & " join db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr & " 	on o.paygatetid = l.PGkey "
	sqlStr = sqlStr & " 	and l.PGgubun = 'inicis' "
	sqlStr = sqlStr & " 	and o.paygatetid is not NULL "
	sqlStr = sqlStr & " 	and o.ipkumdate is not NULL "
	sqlStr = sqlStr & " 	and o.jumundiv <> '6' "			'// ��ȯ�ֹ� ����
	sqlStr = sqlStr & " 	and ((l.appDivCode = 'A' and abs(DateDiff(mi, o.ipkumdate, l.appDate)) < 3) or (l.appDivCode <> 'A'))  "		'// 3��
	sqlStr = sqlStr & " 	and ((Left(o.paygatetid, 9) = 'IniTechPG') or (Left(o.paygatetid, 5) = 'INIMX') or (Left(o.paygatetid, 6) = 'INIswt') or (Left(o.paygatetid, 6) = 'Stdpay') or (Left(o.paygatetid, 5) = 'StdMX') or (Left(o.paygatetid, 10) = 'INIAPICARD')) "
	sqlStr = sqlStr & " where l.orderserial is NULL "
	sqlStr = sqlStr & " and l.PGkey = '" & PGkey & "' "

	'response.write sqlStr & "<br>"
	dbget.execute sqlStr

	'// �ֹ����
	sqlStr = " update l set l.csasid = a.id "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	join [db_cs].[dbo].tbl_new_as_list a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and l.orderserial = a.orderserial "
	sqlStr = sqlStr + " 		and l.appDivCode = 'C' "
	sqlStr = sqlStr + " 		and a.divcd = 'A007' "
	sqlStr = sqlStr + " 	join [db_cs].dbo.tbl_as_refund_info r "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		a.id = r.asid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and l.PGkey = '" & PGkey & "' "
	sqlStr = sqlStr + " 	and a.deleteyn = 'N' "
	sqlStr = sqlStr + " 	and a.currstate = 'B007' "
	sqlStr = sqlStr + " 	and abs(DateDiff(d, a.finishdate, l.cancelDate)) <= 1 "
	sqlStr = sqlStr + " 	and r.refundresult = (l.appPrice * -1) "
	sqlStr = sqlStr + " 	and l.csasid is NULL "
	sqlStr = sqlStr + " 	and l.PGgubun = 'inicis' "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'// �κ����
	sqlStr = " update l set l.csasid = a.id "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	left join [db_cs].[dbo].tbl_new_as_list a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and l.orderserial = a.orderserial "
	sqlStr = sqlStr + " 		and l.appDivCode = 'R' "
	sqlStr = sqlStr + " 		and a.divcd = 'A007' "
	sqlStr = sqlStr + " 	left join [db_cs].dbo.tbl_as_refund_info r "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		a.id = r.asid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and l.PGkey = '" & PGkey & "' "
	sqlStr = sqlStr + " 	and a.deleteyn = 'N' "
	sqlStr = sqlStr + " 	and a.currstate = 'B007' "
	sqlStr = sqlStr + " 	and abs(DateDiff(mi, a.finishdate, l.cancelDate) % (24 * 60)) < 2 "			'// ���� �ð���
	sqlStr = sqlStr + " 	and r.refundresult = (l.appPrice * -1) "
	sqlStr = sqlStr + " 	and l.csasid is NULL "
	sqlStr = sqlStr + " 	and l.PGgubun = 'inicis' "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

elseif (mode = "matchfingerspgdata") then

	'' sqlStr = " update l set l.orderserial = o.orderserial "
	'' sqlStr = sqlStr + " from "
	'' sqlStr = sqlStr + " [ACADEMYDB].[db_academy].[dbo].tbl_academy_order_master o "
	'' sqlStr = sqlStr + " join db_order.dbo.tbl_onlineApp_log l "
	'' sqlStr = sqlStr + " on "
	'' sqlStr = sqlStr + " 	1 = 1 "
	'' sqlStr = sqlStr + " 	and o.paygatetid = l.PGkey "
	'' sqlStr = sqlStr + " 	and l.PGgubun = 'inicis' "
	'' sqlStr = sqlStr + " 	and o.paygatetid is not NULL "
	'' sqlStr = sqlStr + " 	and o.ipkumdate is not NULL "
	'' sqlStr = sqlStr + " 	and o.jumundiv <> '6' "			'// ��ȯ�ֹ� ����
	'' sqlStr = sqlStr + " 	and ((l.appDivCode = 'A' and abs(DateDiff(mi, o.ipkumdate, l.appDate)) < 3) or (l.appDivCode <> 'A'))  "		'// 3��
	'' sqlStr = sqlStr + " 	and ((Left(o.paygatetid, 9) = 'IniTechPG') or (Left(o.paygatetid, 5) = 'INIMX') or (Left(o.paygatetid, 6) = 'INIswt') or (Left(o.paygatetid, 6) = 'Stdpay') or (Left(o.paygatetid, 5) = 'StdMX') or (Left(o.paygatetid, 10) = 'INIAPICARD')) "
	'' sqlStr = sqlStr + " where "
	'' sqlStr = sqlStr + " 	1 = 1 "
	'' sqlStr = sqlStr + " 	and l.orderserial is NULL "
	sqlStr = " exec [db_order].[dbo].[usp_TEN_PGData_Match_FingersOrder] "
	''response.write sqlStr + "<br>"
	''response.end
	dbget.execute sqlStr

	'// ���ֹ� ���&��ǰ
	sqlStr = " update l set l.csasid = a.id "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	join [ACADEMYDB].[db_academy].[dbo].tbl_academy_as_list a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and l.orderserial = a.orderserial "
	sqlStr = sqlStr + " 		and l.appDivCode in ('C', 'R') "						'// ���, �κ����
	sqlStr = sqlStr + " 		and a.divcd = 'A007' "
	sqlStr = sqlStr + " 	join [ACADEMYDB].[db_academy].dbo.tbl_academy_as_refund_info r "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		a.id = r.asid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and a.deleteyn = 'N' "
	sqlStr = sqlStr + " 	and a.currstate = 'B007' "
	sqlStr = sqlStr + " 	and abs(DateDiff(mi, a.finishdate, l.cancelDate)) < 2 "		'// 2��

	'/2���� ������� �ؿ��� 4��¥�� �ּ� ���� Ǯ�� �ְ� ������.
	'sqlStr = sqlStr + " 	and abs(DateDiff(mi, a.finishdate, l.cancelDate)) < 4 "		'// 4��
	'sqlStr = sqlStr + " 	and l.canceldate >= '2016-12-01' "		'/��¥�� �ٲ��ְ�

	sqlStr = sqlStr + " 	and r.refundresult = (l.appPrice * -1) "
	sqlStr = sqlStr + " 	and l.csasid is NULL "
	sqlStr = sqlStr + " 	and l.PGgubun in ('inicis', 'kcp') "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

elseif (mode = "matchgiftcardpgdata") then
    ''�̴Ͻý� ��������� ��� �Աݿ�û TID��  �ԱݿϷ�TID�� �ٸ�����  tbl_onlineApp_log ���� �Ա�TID�� ��
    ''TID ����ġ �����Ϸ�

	sqlStr = " update l "
	sqlStr = sqlStr + " set l.orderserial = o.giftOrderSerial "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	join db_order.dbo.tbl_giftcard_order o on o.paydateid = l.PGkey "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and l.PGgubun = 'inicis' "
	sqlStr = sqlStr + " 	and l.PGuserid = 'teenxteen8' "
	sqlStr = sqlStr + " 	and l.orderserial is NULL "
    dbget.execute sqlStr

	'// ���ֹ� ���&��ǰ
	sqlStr = " update l set l.csasid = a.id "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log l "
	sqlStr = sqlStr + " 	join [db_cs].[dbo].tbl_new_as_list a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and l.orderserial = a.orderserial "
	sqlStr = sqlStr + " 		and l.appDivCode in ('C', 'R') "						'// ���, �κ����
	sqlStr = sqlStr + " 		and a.divcd = 'A007' "
	sqlStr = sqlStr + " 		and a.finishdate >= DateAdd(d, -7, getdate()) "
	sqlStr = sqlStr + " 	join [db_cs].dbo.tbl_as_refund_info r "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		a.id = r.asid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and a.deleteyn = 'N' "
	sqlStr = sqlStr + " 	and a.currstate = 'B007' "
	sqlStr = sqlStr + " 	and abs(DateDiff(mi, a.finishdate, l.cancelDate)) <= 5 "		'// 5��
	sqlStr = sqlStr + " 	and r.refundresult = (l.appPrice * -1) "
	sqlStr = sqlStr + " 	and l.csasid is NULL "
	sqlStr = sqlStr + " 	and l.PGgubun = 'inicis' "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	'// ��ü���� �� ������ ������Ʈ
	sqlStr = " update "
	sqlStr = sqlStr + " db_order.dbo.tbl_onlineApp_log "
	sqlStr = sqlStr + " set orgPayDate = convert(VARCHAR(10), appDate, 127) "
	sqlStr = sqlStr + " where appDate is not NULL and orgPayDate is NULL "
	''response.write sqlStr + "<br>"
	dbget.execute sqlStr

	sqlStr = " update r "
	sqlStr = sqlStr + " set r.orgPayDate = convert(VARCHAR(10), a.appDate, 127) "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " db_order.dbo.tbl_onlineApp_log r "
	sqlStr = sqlStr + " join db_order.dbo.tbl_onlineApp_log a "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and r.PGgubun = a.PGgubun "
	sqlStr = sqlStr + " 	and r.PGkey = a.PGkey "
	sqlStr = sqlStr + " 	and r.appDivCode = 'R' "
	sqlStr = sqlStr + " 	and a.appDivCode = 'A' "
	sqlStr = sqlStr + " where r.appDate is NULL and a.appDate is not NULL and r.orgPayDate is NULL "
	''response.write sqlStr + "<br>"
    ''������ �Ʒ� ������ ��ü
	''dbget.execute sqlStr

	sqlStr = " update r "
	sqlStr = sqlStr + " set r.orgPayDate = convert(VARCHAR(10), a.appDate, 127) "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " db_order.dbo.tbl_onlineApp_log r "
	sqlStr = sqlStr + " join db_order.dbo.tbl_onlineApp_log a "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and r.PGgubun = a.PGgubun "
	sqlStr = sqlStr + " 	and r.PGkey = a.PGkey "
	sqlStr = sqlStr + " 	and a.orderserial = r.orderserial "
	sqlStr = sqlStr + " 	and r.appDivCode = 'R' "
	sqlStr = sqlStr + " 	and a.appDivCode = 'A' "
	sqlStr = sqlStr + " join ( "
	sqlStr = sqlStr + " 	select r.idx "
	sqlStr = sqlStr + " 	from "
	sqlStr = sqlStr + " 		db_order.dbo.tbl_onlineApp_log r "
	sqlStr = sqlStr + " 	where "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and r.orgPayDate is NULL "
	sqlStr = sqlStr + " 		and r.appDate is NULL "
	sqlStr = sqlStr + " 		and r.appDivCode = 'R' "
	sqlStr = sqlStr + " 		and r.cancelDate >= DateAdd(day, -7, getdate()) "
	sqlStr = sqlStr + " ) T on r.idx = T.idx "
	sqlStr = sqlStr + " where r.appDate is NULL and a.appDate is not NULL and r.orgPayDate is NULL "
    ''response.write sqlStr + "<br>"
	dbget.execute sqlStr

elseif (mode = "makeadvprc") then

	sqlStr = " select PGgubun, PGuserid, PGkey, PGCSkey "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	[DBDATAMART].db_datamart.dbo.tbl_order_payment_log "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and paydate is not NULL "
	sqlStr = sqlStr + " 	and pgkey is not NULL "
	sqlStr = sqlStr + " 	and paydate >= '" + Left(DateAdd("m", -1, Now), 7) + "-01' "
	sqlStr = sqlStr + " 	and paydate < '" + Left(DateAdd("m", 1, Now), 7) + "-01' "
	sqlStr = sqlStr + " 	and payDivCode not in ('mil', 'dep', 'gif', '0', 'XXX') "
	sqlStr = sqlStr + " 	and not (payDivCode in ('rde') and realPayPrice = 0) "
	sqlStr = sqlStr + " 	and PGgubun <> 'KICC' "

	'// ���� ��¥ ���� �ݾ��� ȯ�Ұ��� �ִ� ��� �߸� ��Ī�� �� �ִ�.
	'// sqlStr = sqlStr + " 	and PGkey<>'15062692753'"  ''�ϴ�.����. �� �ذ�.

	sqlStr = sqlStr + " group by "
	sqlStr = sqlStr + " 	PGgubun, PGuserid, PGkey, PGCSkey "
	sqlStr = sqlStr + " having count(*) > 1 "
	''response.write sqlStr

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

	if not rsget.EOF  then
		response.write "����α� �ߺ� ERROR : PGKey = " & rsget("pgkey")
		rsget.close
		dbget.close()
		response.end
	end if
	rsget.close

	sqlStr = " exec [db_summary].[dbo].[usp_Ten_appPrc_advPrc_SumMake] '" + CStr(yyyymm) + "' "
	''rw sqlStr : response.end
	rsget.Open sqlStr, dbget, 1

	'response.write	"<script language='javascript'>" &_
	'				"	alert('�ۼ��Ǿ����ϴ�.'); " &_
	'				"	history.back(); " &_
	'				"</script>"

elseif (mode = "makeadvprc01") then

	sqlStr = " select PGgubun, PGuserid, PGkey, PGCSkey "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_datamart.dbo.tbl_order_payment_log "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and paydate is not NULL "
	sqlStr = sqlStr + " 	and pgkey is not NULL "
	sqlStr = sqlStr + " 	and paydate >= '" + Left(DateAdd("m", -1, Now), 7) + "-01' "
	sqlStr = sqlStr + " 	and paydate < '" + Left(DateAdd("m", 1, Now), 7) + "-01' "
	sqlStr = sqlStr + " 	and payDivCode not in ('mil', 'dep', 'gif', '0', 'XXX') "
	sqlStr = sqlStr + " 	and not (payDivCode in ('rde') and realPayPrice = 0) "
	sqlStr = sqlStr + " 	and PGgubun <> 'KICC' "

	'// ���� ��¥ ���� �ݾ��� ȯ�Ұ��� �ִ� ��� �߸� ��Ī�� �� �ִ�.
	'// sqlStr = sqlStr + " 	and PGkey<>'15062692753'"  ''�ϴ�.����. �� �ذ�.

	sqlStr = sqlStr + " group by "
	sqlStr = sqlStr + " 	PGgubun, PGuserid, PGkey, PGCSkey "
	sqlStr = sqlStr + " having count(*) > 1 "
	''response.write sqlStr

	db3_rsget.CursorLocation = adUseClient
	db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly, adLockReadOnly

	if not db3_rsget.EOF  then
		response.write "����α� �ߺ� ERROR : PGKey = " & db3_rsget("pgkey")
		db3_rsget.close
		db3_dbget.close()
		response.end
	end if
	db3_rsget.close

elseif (mode = "makeadvprc02") then
	if (yyyymm = "AUTO") then
		if (Day(Now()) = 1) then
			'// ����
			yyyymm = Left(DateAdd("m", -1, Now()), 7)
		else
			'// ���ϱ���
			yyyymm = Left(DateAdd("m", 0, Now()), 7)
		end if
	end if

	sqlStr = " exec [db_summary].[dbo].[usp_Ten_appPrc_advPrc_SumMake_01] '" + CStr(yyyymm) + "' "
	''rw sqlStr : response.end
	rsget.Open sqlStr, dbget, 1

	'response.write	"<script language='javascript'>" &_
	'				"	alert('�ۼ��Ǿ����ϴ�.'); " &_
	'				"	history.back(); " &_
	'				"</script>"

elseif (mode = "makeadvprc03") then
	if (yyyymm = "AUTO") then
		if (Day(Now()) = 1) then
			'// ����
			yyyymm = Left(DateAdd("m", -1, Now()), 7)
		else
			'// ���ϱ���
			yyyymm = Left(DateAdd("m", 0, Now()), 7)
		end if
	end if

	sqlStr = " exec [db_summary].[dbo].[usp_Ten_appPrc_advPrc_SumMake_02] '" + CStr(yyyymm) + "' "
	''rw sqlStr : response.end
	rsget.Open sqlStr, dbget, 1

	'response.write	"<script language='javascript'>" &_
	'				"	alert('�ۼ��Ǿ����ϴ�.'); " &_
	'				"	history.back(); " &_
	'				"</script>"

elseif (mode = "addhand") then

	select case gubun
		case "cancel"
			sqlStr = " insert into [db_shop].[dbo].[tbl_shopjumun_cardApp_log]( "
			sqlStr = sqlStr + "	PGgubun, PGkey, appDivCode, appDate, cardReaderID, cardPrice, cardAppNo, cardGubun, cardComp, cardAffiliateNo, ipkumdate, shopid, shopJumunMasterIdx, orderserial, matchtype, cardChargePrice, ipkumPrice, reasonGubun "
			sqlStr = sqlStr + ") "
			sqlStr = sqlStr + "select top 1 'HAND', PGkey, 'C', '" + CStr(canceldate) + "', cardReaderID, cardPrice*-1, cardAppNo, cardGubun, cardComp, cardAffiliateNo, '" + CStr(ipkumdate) + "', shopid, NULL, NULL, NULL, cardChargePrice*-1, ipkumPrice*-1, reasonGubun "
			sqlStr = sqlStr + "from [db_shop].[dbo].[tbl_shopjumun_cardApp_log] "
			sqlStr = sqlStr + "where PGkey = '" + CStr(orgpgkey) + "' and PGgubun <> 'HAND' and appDivCode = 'A' "
			rsget.Open sqlStr, dbget, 1
		case "dup"
			orderserial = requestCheckvar(request("orderserial"),32)

			sqlStr = " insert into [db_shop].[dbo].[tbl_shopjumun_cardApp_log]( "
			sqlStr = sqlStr + "	PGgubun, PGkey, appDivCode, appDate, cardReaderID, cardPrice, cardAppNo, cardGubun, cardComp, cardAffiliateNo, ipkumdate, shopid, shopJumunMasterIdx, orderserial, matchtype, cardChargePrice, ipkumPrice, reasonGubun "
			sqlStr = sqlStr + ") "
			sqlStr = sqlStr + "select top 1 'HAND', PGkey + '_1', 'A', appDate, cardReaderID, 0, cardAppNo, cardGubun, cardComp, cardAffiliateNo, '" + CStr(ipkumdate) + "', shopid, NULL, '" & orderserial & "', NULL, 0, 0, reasonGubun "
			sqlStr = sqlStr + "from [db_shop].[dbo].[tbl_shopjumun_cardApp_log] "
			sqlStr = sqlStr + "where PGkey = '" + CStr(orgpgkey) + "' and PGgubun <> 'HAND' and appDivCode = 'A' "
			rsget.Open sqlStr, dbget, 1
		case "del"
			sqlStr = " delete from [db_shop].[dbo].[tbl_shopjumun_cardApp_log] "
			sqlStr = sqlStr + " where PGkey = '" + CStr(orgpgkey) + "' and PGgubun = 'HAND' "
			rsget.Open sqlStr, dbget, 1
		case else
			response.write "�߸��� �����Դϴ�."
			dbget.close : response.end
	end select

elseif (mode = "addhandOn") then

	select case gubun
		case "partcancel"
			sqlStr = " insert into db_order.dbo.tbl_onlineApp_log( "
			sqlStr = sqlStr + "	PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, orderserial, csasid, regdate, PGuserid, orgPayDate, PGmeachulDate, reasonGubun, etcPoint "
			sqlStr = sqlStr + ") "
			sqlStr = sqlStr + "select top 1 PGgubun, PGkey, PGCSkey + '_1', sitename, appDivCode, appMethod, appDate, cancelDate, 0, 0, 0, 0, ipkumdate, orderserial, csasid, regdate, PGuserid, orgPayDate, PGmeachulDate, reasonGubun, 0 "
			sqlStr = sqlStr + "from db_order.dbo.tbl_onlineApp_log "
			sqlStr = sqlStr + "where PGkey = '" + CStr(orgpgkey) + "' and PGCSkey = '" + CStr(orgpgcskey) + "' and PGgubun <> 'HAND' and appDivCode in ('R', 'C') "		'// ��ü����� ���̽� ����
            ''response.write sqlStr
			rsget.Open sqlStr, dbget, 1
		case "del"
			sqlStr = " delete from db_order.dbo.tbl_onlineApp_log "
			sqlStr = sqlStr + " where PGkey = '" + CStr(orgpgkey) + "' and PGCSkey like '" + CStr(orgpgcskey) + "_%' "
			rsget.Open sqlStr, dbget, 1
		case else
			response.write "�߸��� �����Դϴ�."
			dbget.close : response.end
	end select

elseif (mode = "ModiAppDate") then

    appdate = requestCheckvar(request("appdate"), 10)

    sqlStr = " update db_order.dbo.tbl_onlineApp_log "
	sqlStr = sqlStr + " set "
	sqlStr = sqlStr + " 	appDate = (case when appDivCode = 'A' then '" & appdate & "' else appDate end), "
	sqlStr = sqlStr + " 	cancelDate = (case when appDivCode <> 'A' then '" & appdate & "' else cancelDate end) "
	sqlStr = sqlStr + " where idx = " & logidx
    rsget.Open sqlStr, dbget, 1

elseif (mode = "ModiIpkumDate") then

    appdate = requestCheckvar(request("appdate"), 10)

    sqlStr = " update db_order.dbo.tbl_onlineApp_log "
	sqlStr = sqlStr + " set "
	sqlStr = sqlStr + " 	ipkumdate = '" & appdate & "' "
	sqlStr = sqlStr + " where idx = " & logidx
    rsget.Open sqlStr, dbget, 1

elseif (mode = "getonpgdatacappMethod6") then

    sqlStr = " exec [db_log].[dbo].[usp_Ten_MakeEtcPaymentLog_ON] '" & yyyymmdd & "', '" & yyyymmdd & "' "
    ''response.write sqlStr
    ''dbget.close() : response.end
    dbget.execute sqlStr

elseif (mode = "delapplog") then
	'// ������ ������ �����͸� db_log.dbo.tbl_onlineApp_Delete_Log�� ������ �� ����
	sqlStr = " insert into db_log.dbo.tbl_onlineApp_Delete_Log( "
	sqlStr = sqlStr + "	oidx, PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, orderserial, csasid, regdate, PGuserid, orgPayDate, PGmeachulDate, reasonGubun, etcPoint, logRegdate, delAdminId "
	sqlStr = sqlStr + ") "
	sqlStr = sqlStr + "select top 1 idx, PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPRice, ipkumdate, orderserial, csasid, regdate, PGuserid, orgPayDate, PGmeachulDate, reasonGubun, etcPoint, getdate(), '"&session("ssBctId")&"' "
	sqlStr = sqlStr + "from db_order.dbo.tbl_onlineApp_log "
	sqlStr = sqlStr + "where idx= " & logidx
	''response.write sqlStr
	rsget.Open sqlStr, dbget, 1

	sqlStr = " delete from db_order.dbo.tbl_onlineApp_log "
	sqlStr = sqlStr + " where idx = " & logidx
	rsget.Open sqlStr, dbget, 1

elseif (mode = "addIniRentalManualWrite") then
	select case Trim(inirentalgubun)
		case "inirentalbuy"
			sqlStr = " select count(*) "
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log "
			sqlStr = sqlStr + " where "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and PGkey = '"&Trim(inirentalpgkey)&"' "
			sqlStr = sqlStr + " 	and appDivCode = 'A' "
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
			inirentalreduplication = rsget(0)
			rsget.close

			If inirentalreduplication > 0 Then
				response.write "�̹� ��ϵ� ���系���� �ֽ��ϴ�. : PGKey = " & inirentalpgkey
				response.end
			End If

			sqlStr = " insert into db_order.dbo.tbl_onlineApp_log( "
			sqlStr = sqlStr + "	PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, orderserial, regdate, PGuserid, PGmeachulDate "
			sqlStr = sqlStr + ") "
			sqlStr = sqlStr + "values ('inirental','"&Trim(inirentalpgkey)&"','','10x10','A','150','"&inirentalconfirmdate&"',NULL,'"&inirentalappprice&"','"&commPrice&"','"&inirentalcommvatprice&"','"&inirentaljungsanprice&"','"&inirentalipkumdate&"',NULL,getdate(),'teenxteenr','"&inirentalipkumdate&"') "
			rsget.Open sqlStr, dbget, 1
		case "inirentalcancel"
			sqlStr = " select count(*) "
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log "
			sqlStr = sqlStr + " where "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and PGkey = '"&Trim(inirentalpgkey)&"' "
			sqlStr = sqlStr + " 	and appDivCode = 'C' "
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
			inirentalreduplication = rsget(0)
			rsget.close

			If inirentalreduplication > 0 Then
				response.write "�̹� ��ϵ� ��ҳ����� �ֽ��ϴ�. : PGKey = " & inirentalpgkey
				response.end
			End If

			sqlStr = " insert into db_order.dbo.tbl_onlineApp_log( "
			sqlStr = sqlStr + "	PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, orderserial, regdate, PGuserid, PGmeachulDate "
			sqlStr = sqlStr + ") "
			sqlStr = sqlStr + "values ('inirental','"&Trim(inirentalpgkey)&"','CANCELALL','10x10','C','150',NULL,'"&inirentalconfirmdate&"','-"&inirentalappprice&"','-"&commPrice&"','-"&inirentalcommvatprice&"','-"&inirentaljungsanprice&"','"&inirentalipkumdate&"',NULL,getdate(),'teenxteenr','"&inirentalipkumdate&"') "
			rsget.Open sqlStr, dbget, 1
		case else
			response.write "�߸��� �����Դϴ�."
			dbget.close : response.end
	end select
elseif (mode = "inirentalcancel") then

	inirentalpgkey = requestCheckvar(request("PGkey"),64)

	sqlStr = " select count(*) "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and PGkey = '"&Trim(inirentalpgkey)&"' "
	sqlStr = sqlStr + " 	and appDivCode = 'C' "
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
	inirentalreduplication = rsget(0)
	rsget.close

	If inirentalreduplication > 0 Then
		response.write "�̹� ��ϵ� ��ҳ����� �ֽ��ϴ�. : PGKey = " & inirentalpgkey
		response.end
	End If


	sqlStr = " select Top 1 * "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_order.dbo.tbl_onlineApp_log "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and PGkey = '"&Trim(inirentalpgkey)&"' "
	sqlStr = sqlStr + " 	and appDivCode = 'A' "
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
	if not rsget.EOF  then
		inirentalappprice = rsget("appPrice")
		commPrice = rsget("commPrice")
		inirentalcommvatprice = rsget("commVatPrice")
		inirentaljungsanprice = rsget("jungsanPrice")
		inirentalipkumdate = rsget("ipkumdate")
	Else
		response.write "��ϵ� ���������� �����ϴ�. ��Ҵ� ��ϵ� ���������� �־�� �����մϴ�. PGKey = " & Trim(inirentalpgkey)
		response.end
	end if
	rsget.close

	sqlStr = " insert into db_order.dbo.tbl_onlineApp_log( "
	sqlStr = sqlStr + "	PGgubun, PGkey, PGCSkey, sitename, appDivCode, appMethod, appDate, cancelDate, appPrice, commPrice, commVatPrice, jungsanPrice, ipkumdate, orderserial, regdate, PGuserid, PGmeachulDate "
	sqlStr = sqlStr + ") "
	sqlStr = sqlStr + "values ('inirental','"&Trim(inirentalpgkey)&"','CANCELALL','10x10','C','150',NULL,getdate(),'-"&inirentalappprice&"','-"&commPrice&"','-"&inirentalcommvatprice&"','-"&inirentaljungsanprice&"','"&inirentalipkumdate&"',NULL,getdate(),'teenxteenr','"&inirentalipkumdate&"') "
	rsget.Open sqlStr, dbget, 1

end if

%>
<% if  (IsAutoScript) then  %>
<% rw "OK" %>
<% else %>
<script language='javascript'>
alert('����Ǿ����ϴ�.');
location.replace('<%= refer %>');
</script>
<% end if %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
