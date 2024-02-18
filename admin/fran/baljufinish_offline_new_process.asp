<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 오프라인 출고 처리
' Hieditor : 2011.03.09 이상구 생성
'			 2012.08.14 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/db_LogisticsOpen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/logisticsbaljuofflinecls.asp"-->
<%

dim lastPageTime, pageElapsedTime
lastPageTime = Timer

'// Call checkAndWriteElapsedTime("001")
function checkAndWriteElapsedTime(memo)
	pageElapsedTime = Timer - lastPageTime
	lastPageTime = Timer
	response.write "<!-- Page Execute Time Check : " & FormatNumber(pageElapsedTime, 4) & " : " & memo & " -->" & vbCrLf
end function

dim mode , baljuid, baljudate, itemgubun, itemid, itemoption, comment
dim i,cnt,sqlStr, errstring ,masteridxlist, baljuname, baljucodelist ,songjangdiv
dim divcode, vatinclude, targetid, targetname, baljucode, brandlist, obaljucode
dim siteseq, companyid ,masteridx, ordercode ,isorgorder, orgordercode ,remasteridx, reordercode
dim itemexists, iid, newbaljucode, itemAlreadyExists, tmp
dim currencyUnit, IsFinished, isWait, loginsite
Dim baljuKey, siteBaljuKey
Dim ordercodelistNoBeasongDate, ordercodelistNoSiteInsertDate

	mode        = RequestCheckVar(request("mode"),32)
	baljukey    = RequestCheckVar(request("baljunum"),32)
	baljuid     = RequestCheckVar(request("baljuid"),32)
	itemgubun   = RequestCheckVar(request("itemgubun"),200)
	itemid      = RequestCheckVar(request("itemid"),840)
	itemoption  = RequestCheckVar(request("itemoption"),440)
	comment     = RequestCheckVar(request("comment"),1280)
	isWait      = RequestCheckVar(request("isWait"),32)


dim IsWriteReOrderSheet : IsWriteReOrderSheet = False			'// 재주문서 작성
dim IsWriteChulgoSheet : IsWriteChulgoSheet = False				'// 출고내역 작성
dim errMsg :errMsg = ""

sqlStr = " select IsFinished, siteBaljuid as siteBaljukey, songjangdiv " + VbCrLf
sqlStr = sqlStr + " from db_aLogistics.dbo.tbl_Logistics_offline_baljumaster " + VbCrLf
sqlStr = sqlStr + " where baljuKey = " & baljuKey & " " + VbCrLf
'response.write sqlStr & "<Br>"
rsget_Logistics.Open sqlStr,dbget_Logistics,1

if  not rsget_Logistics.EOF  then
	IsFinished 		= rsget_Logistics("IsFinished")
	siteBaljuKey 	= rsget_Logistics("siteBaljuKey")
	songjangdiv 	= rsget_Logistics("songjangdiv")
end if
rsget_Logistics.Close

Select Case IsFinished
	Case "N"
		'// 출고작업중
		IsWriteReOrderSheet = True
		if (isWait = "N") then
			IsWriteChulgoSheet = True
		end if
	Case "W"
		'// 출고대기
		if (isWait = "N") then
			IsWriteChulgoSheet = True
		else
			errMsg = "에러 : 출고작업중 상태가 아닙니다."
		end if
	Case "Y"
		'// 에러
		errMsg = "에러 : 이미 출고완료된 내역입니다."
	Case Else
		errMsg = "에러 : 알 수 없는 에러"
End Select

if (errMsg <> "") then
	response.write errMsg
	dbget_Logistics.Close
	dbget.Close
	response.end
end If


companyid = session("ssBctID")

dim refer
	refer = request.ServerVariables("HTTP_REFERER")

siteseq = GetLogicsSiteSeq		'/lib/classes/order/logisticsbaljuofflinecls.asp

function GetFromWhere(siteseq, baljuKey, baljuid)
	dim tmpsql

    tmpsql = " FROM " + VbCrLf
    tmpsql = tmpsql + " 	db_aLogistics.dbo.tbl_Logistics_offline_baljumaster bm " + VbCrLf
    tmpsql = tmpsql + " 	, [db_aLogistics].[dbo].tbl_Logistics_offline_baljudetail b " + VbCrLf
    tmpsql = tmpsql + " 	, [db_aLogistics].[dbo].tbl_Logistics_offline_order_master m " + VbCrLf
    tmpsql = tmpsql + " 	, [db_aLogistics].[dbo].tbl_Logistics_offline_order_detail d " + VbCrLf
    tmpsql = tmpsql + " 	LEFT JOIN [db_aLogistics].[dbo].tbl_Logistics_offline_item i " + VbCrLf
    tmpsql = tmpsql + " 	ON " + VbCrLf
    tmpsql = tmpsql + " 		1 = 1 " + VbCrLf
    tmpsql = tmpsql + " 		and d.siteseq = i.siteseq " + VbCrLf
    tmpsql = tmpsql + " 		and d.itemgubun = i.siteitemgubun " + VbCrLf
    tmpsql = tmpsql + " 		and d.itemid = i.siteitemid " + VbCrLf
    tmpsql = tmpsql + " 		and d.itemoption = i.siteitemoption " + VbCrLf
    tmpsql = tmpsql + " 	LEFT JOIN [db_aLogistics].[dbo].tbl_Logistics_offline_tmppacking p " + VbCrLf
    tmpsql = tmpsql + " 	ON " + VbCrLf
    tmpsql = tmpsql + " 		1 = 1 " + VbCrLf
    tmpsql = tmpsql + " 		and d.siteseq = p.siteseq " + VbCrLf
    tmpsql = tmpsql + " 		and d.ordercode = p.ordercode " + VbCrLf
    tmpsql = tmpsql + " 		and d.itemgubun = p.itemgubun " + VbCrLf
    tmpsql = tmpsql + " 		and d.itemid = p.itemid " + VbCrLf
    tmpsql = tmpsql + " 		and d.itemoption = p.itemoption " + VbCrLf
    tmpsql = tmpsql + " WHERE " + VbCrLf
    tmpsql = tmpsql + " 	1 = 1 " + VbCrLf
    tmpsql = tmpsql + " 	and bm.baljukey = b.baljukey " + VbCrLf
    tmpsql = tmpsql + " 	and m.siteseq = d.siteseq " + VbCrLf
    tmpsql = tmpsql + " 	and m.ordercode = d.ordercode " + VbCrLf
    tmpsql = tmpsql + " 	and b.ordercode = m.ordercode " + VbCrLf
    tmpsql = tmpsql + " 	and d.cancelyn <> 'Y' " + VbCrLf
    tmpsql = tmpsql + " 	and m.SiteSeq = " & siteseq & " " + VbCrLf
    tmpsql = tmpsql + "     and b.baljukey = '" + CStr(baljuKey) + "' " + VbCrLf
    tmpsql = tmpsql + " 	and m.shopid = '" & baljuid & "' " + VbCrLf

    GetFromWhere = tmpsql
end Function

Call checkAndWriteElapsedTime("001")
''dbget.close() : response.end

if mode="chulgoproc" Then

	'// ========================================================================
    '에러체크 : 잘못된 입력(박스번호가 0 이면서 송장번호가 있는경우 or realitemno 가 있으면서, 박스번호가 없는경우)체크
    sqlStr = " select d.itemname,d.itemoptionname " + VbCrLf
    sqlStr = sqlStr + GetFromWhere(siteseq, baljuKey, baljuid)
    sqlStr = sqlStr + " 	and m.beasongdate is null " + VbCrLf
    sqlStr = sqlStr + " 	and (((isnull(d.packingstate,0) = 0) and (isnull(d.songjangno,'0') <> '0')) or ((d.fixedno > 0) and (isnull(d.songjangno,'0') = '0'))) " + VbCrLf

	'response.write sqlStr & "<Br>"
    rsget_Logistics.Open sqlStr, dbget_Logistics, 1
    if  not rsget_Logistics.EOF  then
        do until rsget_Logistics.eof
            if (trim(errstring) = "") then
				errstring = rsget_Logistics("itemname") + "(" + rsget_Logistics("itemoptionname") + ")"
            else
                errstring = errstring + ", " + rsget_Logistics("itemname") + "(" + rsget_Logistics("itemoptionname") + ")"
            end if

            rsget_Logistics.MoveNext
        loop
    else
		errstring = ""
    end if
    rsget_Logistics.close

    if (errstring <> "") then
        response.write "<script>alert('잘못된 입력이 있습니다. 송장번호 또는 상품코드를 삭제후 다시 입력하세요.\n\n" + errstring + "');</script>"
        response.write "<script>history.back();</script>"

        dbget.close()
        dbget_Logistics.close()
        response.End
    end If


	'// ========================================================================
	ordercodelistNoSiteInsertDate = ""
	ordercodelistNoBeasongDate = ""

	'// siteinsertdate 가 널이 아니면 재주문서 작성완료
	sqlStr = " select distinct d.ordercode " + VbCrLf
	sqlStr = sqlStr + GetFromWhere(siteseq, baljuKey, baljuid) + VbCrLf
	sqlStr = sqlStr + " 	and m.siteinsertdate is null " + VbCrLf
	'response.write sqlStr & "<Br>"
	rsget_Logistics.Open sqlStr,dbget_Logistics,1

	if  not rsget_Logistics.EOF  then
		do until rsget_Logistics.eof
			if (ordercodelistNoSiteInsertDate <> "") then
				ordercodelistNoSiteInsertDate = ordercodelistNoSiteInsertDate + ",'" + CStr(rsget_Logistics("ordercode")) & "'"
			else
				ordercodelistNoSiteInsertDate = "'" & CStr(rsget_Logistics("ordercode")) & "'"
			end if

			rsget_Logistics.MoveNext
		loop
	end if
	rsget_Logistics.close

	'// beasongdate 가 널이 아니면 출고완료
	sqlStr = " select distinct d.ordercode "
	sqlStr = sqlStr + GetFromWhere(siteseq, baljuKey, baljuid)
	sqlStr = sqlStr + " 	and m.beasongdate is null " + VbCrLf
	'response.write sqlStr & "<Br>"
	rsget_Logistics.Open sqlStr,dbget_Logistics,1

	if  not rsget_Logistics.EOF  then
		do until rsget_Logistics.eof
			if (ordercodelistNoBeasongDate <> "") then
				ordercodelistNoBeasongDate = ordercodelistNoBeasongDate + ",'" + CStr(rsget_Logistics("ordercode")) & "'"
			else
				ordercodelistNoBeasongDate = "'" & CStr(rsget_Logistics("ordercode")) & "'"
			end if

			rsget_Logistics.MoveNext
		loop
	end if
	rsget_Logistics.Close

	'해당 출고지시코드/출고지시아이디에 대한 masteridx 를 구한다.
	sqlStr = " select distinct d.masteridx "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_shopbalju b "
	sqlStr = sqlStr + " 	join [db_storage].[dbo].tbl_ordersheet_master m "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		b.baljucode = m.baljucode "
	sqlStr = sqlStr + " 	join [db_storage].[dbo].tbl_ordersheet_detail d "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		m.idx = d.masteridx "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	b.baljunum = " & siteBaljuKey
	sqlStr = sqlStr + " 	and m.deldt is null "
	sqlStr = sqlStr + " 	and d.deldt is null "
	sqlStr = sqlStr + " 	and m.divcode in ('501','502','503') "
	sqlStr = sqlStr + " 	and m.statecd < '7' "
	'response.write sqlStr & "<Br>"
	rsget.Open sqlStr,dbget,1

	masteridxlist = ""
	if  not rsget.EOF  then
		do until rsget.eof
			if (masteridxlist <> "") then
				masteridxlist = masteridxlist + "," + CStr(rsget("masteridx"))
			else
				masteridxlist = CStr(rsget("masteridx"))
			end if

			rsget.MoveNext
		loop
	end if
	rsget.Close

	IF (masteridxlist="") then masteridxlist="-1"       ''2011-07-04 추가

	Call checkAndWriteElapsedTime("002")
	''dbget.close() : response.end

	'// ========================================================================
	'// 로직스 작업정보(미배송 사유, 실제 상품수, 송장정보 등) 입력 및 복사
	if IsWriteReOrderSheet And (ordercodelistNoSiteInsertDate <> "") Then
		'로직스 미배송 주문상품에 대한 코맨트 입력
		if (Trim(comment) <> "") then
			itemgubun = split(itemgubun,"|")
			itemid = split(itemid,"|")
			itemoption = split(itemoption,"|")
			comment = split(comment,"|")
			cnt = ubound(itemgubun)

			for i=0 to cnt
				if (Trim(comment(i)) <> "") then
					sqlstr = " update [db_aLogistics].[dbo].tbl_Logistics_offline_order_detail "
					sqlstr = sqlstr + " set comment = '" + Trim(comment(i)) + "' "
					sqlstr = sqlstr + " where itemgubun = '" + Trim(itemgubun(i)) + "' "
					sqlstr = sqlstr + " and itemid = " + Trim(itemid(i)) + " "
					sqlstr = sqlstr + " and itemoption = '" + Trim(itemoption(i)) + "' "
					sqlstr = sqlstr + " and siteseq = " & siteseq & " "
					sqlstr = sqlstr + " and ordercode in (" + CStr(ordercodelistNoSiteInsertDate) + ") "
					rsget_Logistics.Open sqlStr, dbget_Logistics, 1
				end if
			next
		end If

		Call checkAndWriteElapsedTime("003")
		''dbget.close() : response.end

		if (CStr(siteseq) = "10") Then
			'// 어드민에 작업정보 복사
			sqlStr = " exec [db_storage].[dbo].[usp_Ten_LogicsOffChulgo2SCM] " & baljuKey
			dbget.Execute sqlStr

			Call checkAndWriteElapsedTime("004")
			''dbget.close() : response.end

			sqlStr = " update "
			sqlStr = sqlStr + " 	td "
			sqlStr = sqlStr + " set "
			sqlStr = sqlStr + " 	td.realitemno = ld.fixedno "
			sqlStr = sqlStr + " 	, td.packingstate = ld.packingstate "
			sqlStr = sqlStr + " 	, td.boxsongjangno = ld.songjangno "
			sqlStr = sqlStr + " 	, td.comment = ld.comment "
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " 	[db_storage].[dbo].[tbl_Logistics_offline_order_detail_COPY] ld "
			sqlStr = sqlStr + " 	join [db_storage].[dbo].tbl_ordersheet_detail td "
			sqlStr = sqlStr + " 	on "
			sqlStr = sqlStr + " 		ld.sitedetailidx = td.idx "
			sqlStr = sqlStr + " 		and ld.baljuKey = " & baljuKey
			sqlstr = sqlstr + " 		and ld.ordercode in (" + CStr(ordercodelistNoSiteInsertDate) + ") "
			''response.write sqlStr & "<Br>"
			rsget.Open sqlStr, dbget, 1

			Call checkAndWriteElapsedTime("005")
			''dbget.close() : response.end

			'어드민 마스터정보 업데이트
			sqlStr = " update m " + vbCrLf
			sqlStr = sqlStr + " set jumunsellcash=IsNULL(T.totsell,0)" + vbCrLf
			sqlStr = sqlStr + " ,jumunsuplycash=IsNULL(T.totsupp,0)" + vbCrLf
			sqlStr = sqlStr + " ,jumunbuycash=IsNULL(T.totbuy,0)" + vbCrLf
			sqlStr = sqlStr + " ,totalsellcash=IsNULL(T.realtotsell,0)" + vbCrLf
			sqlStr = sqlStr + " ,totalsuplycash=IsNULL(T.realtotsupp,0)" + vbCrLf
			sqlStr = sqlStr + " ,totalbuycash=IsNULL(T.realtotbuy,0)" + vbCrLf
			sqlStr = sqlStr + " ,jumunforeign_sellcash=IsNULL(T.totforeign_sellcash,0)" + vbCrLf
			sqlStr = sqlStr + " ,jumunforeign_suplycash=IsNULL(T.totforeign_suplycash,0)" + vbCrLf
			sqlStr = sqlStr + " ,totalforeign_sellcash=IsNULL(T.realforeign_sellcash,0)" + vbCrLf
			sqlStr = sqlStr + " ,totalforeign_suplycash	=IsNULL(T.realforeign_suplycash,0)" + vbCrLf
			sqlStr = sqlStr + " from " + vbCrLf
			sqlStr = sqlStr + " 	( " + vbCrLf
			sqlStr = sqlStr + " 		select m.baljucode, sum(sellcash*baljuitemno) as totsell " + vbCrLf
			sqlStr = sqlStr + " 		,sum(suplycash*baljuitemno) as totsupp " + vbCrLf
			sqlStr = sqlStr + " 		,sum(buycash*baljuitemno) as totbuy " + vbCrLf
			sqlStr = sqlStr + " 		,sum(sellcash*realitemno) as realtotsell " + vbCrLf
			sqlStr = sqlStr + " 		,sum(suplycash*realitemno) as realtotsupp " + vbCrLf
			sqlStr = sqlStr + " 		,sum(buycash*realitemno) as realtotbuy " + vbCrLf
			sqlStr = sqlStr + " 		,sum(IsNull(foreign_sellcash,0)*baljuitemno) as totforeign_sellcash " + vbCrLf
			sqlStr = sqlStr + " 		,sum(IsNull(foreign_suplycash,0)*baljuitemno) as totforeign_suplycash " + vbCrLf
			sqlStr = sqlStr + " 		,sum(IsNull(foreign_sellcash,0)*realitemno) as realforeign_sellcash " + vbCrLf
			sqlStr = sqlStr + " 		,sum(IsNull(foreign_suplycash,0)*realitemno) as realforeign_suplycash " + vbCrLf
			sqlStr = sqlStr + " 		from " + vbCrLf
			sqlStr = sqlStr + " 			[db_storage].[dbo].tbl_ordersheet_master m " + vbCrLf
			sqlStr = sqlStr + " 			join [db_storage].[dbo].tbl_ordersheet_detail d " + vbCrLf
			sqlStr = sqlStr + " 			on " + vbCrLf
			sqlStr = sqlStr + " 				m.idx = d.masteridx " + vbCrLf
			sqlStr = sqlStr + " 		where " + vbCrLf
			sqlStr = sqlStr + " 			1 = 1 " + vbCrLf
			sqlStr = sqlStr + " 			and m.baljucode in (" + CStr(ordercodelistNoSiteInsertDate) + ") " + vbCrLf
			sqlStr = sqlStr + " 			and d.deldt is null" + vbCrLf
			sqlStr = sqlStr + " 		group by " + vbCrLf
			sqlStr = sqlStr + " 			m.baljucode " + vbCrLf
			sqlStr = sqlStr + " 	) as T" + vbCrLf
			sqlStr = sqlStr + " 	join [db_storage].[dbo].tbl_ordersheet_master m " + vbCrLf
			sqlStr = sqlStr + " 	on " + vbCrLf
			sqlStr = sqlStr + " 		m.baljucode = T.baljucode " + vbCrLf
			'response.write sqlStr & "<Br>"
			rsget.Open sqlStr, dbget, 1


			'TEN-5. 미배송주문 내역 체크(재주문 대상 상품 검색)
			sqlStr = " select count(d.idx) as cnt  from [db_storage].[dbo].tbl_ordersheet_detail d "
			sqlStr = sqlStr + " where d.masteridx in (" + CStr(masteridxlist) + ") "
			sqlStr = sqlStr + " and d.baljuitemno <> d.realitemno "
			sqlStr = sqlStr + " and d.comment='5일내출고' "
			sqlStr = sqlStr + " and deldt is null "
			'response.write sqlStr & "<Br>"
			rsget.Open sqlStr, dbget, 1
    			itemexists = (rsget("cnt")>0)
			rsget.Close

			sqlStr = " select count(idx) as cnt from  [db_storage].[dbo].tbl_ordersheet_master"
			sqlStr = sqlStr + " where idx in (" + CStr(masteridxlist) + ") "
			sqlStr = sqlStr + " and clinkcode  is not null "
			sqlStr = sqlStr + " and clinkcode<>'' "

			Call checkAndWriteElapsedTime("006")
			''dbget.close() : response.end

			'response.write sqlStr & "<Br>"
			rsget.Open sqlStr, dbget, 1
				itemAlreadyExists = (rsget("cnt")>0)
			rsget.Close

			if Not itemexists then
				'response.write "<script>alert('재 주문할 내역이 없습니다.');</script>"
			elseif itemAlreadyExists then
				'response.write "<script>alert('재 주문서가 이미 작성되어 있습니다. 작성할 수 없습니다.');</script>"
			Else
				'미배송 주문서 작성
				'여러개의 주문서를 하나로 묶는다.
				sqlStr = " select top 1 * from [db_storage].[dbo].tbl_ordersheet_master"
				sqlStr = sqlStr + " where idx in (" + CStr(masteridxlist) + ") "

				'response.write sqlStr & "<Br>"
				rsget.Open sqlStr, dbget, 1
					targetid = rsget("targetid")
					targetname = rsget("targetname")
					divcode = rsget("divcode")
					vatinclude = rsget("vatinclude")
					currencyUnit = rsget("currencyUnit")
					loginsite = rsget("sitename")
				rsget.Close

				'해당 출고지시코드/출고지시아이디에 대한 기본정보를 구한다.
				sqlStr = " select distinct m.baljuname, m.baljucode "
				sqlStr = sqlStr + " from "
				sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_shopbalju b "
				sqlStr = sqlStr + " 	join [db_storage].[dbo].tbl_ordersheet_master m "
				sqlStr = sqlStr + " 	on "
				sqlStr = sqlStr + " 		b.baljucode = m.baljucode "
				sqlStr = sqlStr + " 	join [db_storage].[dbo].tbl_ordersheet_detail d "
				sqlStr = sqlStr + " 	on "
				sqlStr = sqlStr + " 		m.idx = d.masteridx "
				sqlStr = sqlStr + " where "
				sqlStr = sqlStr + " 	b.baljunum = " & siteBaljuKey
				sqlStr = sqlStr + " 	and m.deldt is null "
				sqlStr = sqlStr + " 	and d.deldt is null "
				sqlStr = sqlStr + " 	and m.divcode in ('501','502','503') "
				sqlStr = sqlStr + " 	and m.statecd < '7' "
				sqlStr = sqlStr + " 	and d.baljuitemno <> d.realitemno "
				sqlStr = sqlStr + " 	and d.comment='5일내출고' "
				'response.write sqlStr & "<Br>"
				rsget.Open sqlStr,dbget,1

				baljuname = ""
				baljucode = ""
				baljucodelist = ""
				if  not rsget.EOF  then
					baljuname = CStr(rsget("baljuname"))
					baljucode = CStr(rsget("baljucode"))
					baljucodelist = CStr(rsget("baljucode"))

					rsget.MoveNext
					do until rsget.eof
                        baljucodelist = baljucodelist + "," + CStr(rsget("baljucode"))
                        rsget.MoveNext
					loop
				end if
				rsget.Close

				sqlStr = " select * from [db_storage].[dbo].tbl_ordersheet_master where 1=0 "
				'response.write sqlStr & "<Br>"
				rsget.Open sqlStr,dbget,1,3
				rsget.AddNew
				rsget("targetid") = targetid
				rsget("targetname") = targetname
				rsget("baljuid") = baljuid
				rsget("baljuname") = baljuname

				if loginsite = "WSLWEB"	 then
					rsget("currencyUnit") = currencyUnit
					rsget("foreign_statecd") = "0"
					rsget("sitename") = loginsite
				end if

				rsget("reguser") = session("ssBctId")
				rsget("regname") = session("ssBctCname")
				rsget("divcode") = divcode
				rsget("vatinclude") = vatinclude
				rsget("scheduledate") = Left(now(), 10)
				rsget("statecd") = "0"
				rsget("comment") = baljucodelist + " 미배송건 재작성"

				rsget.update
        			iid = rsget("idx")
				rsget.close

				baljucode = "RJ" + Format00(6,Right(CStr(iid),6))

				''디테일 저장
				sqlStr = " insert into [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
				sqlStr = sqlStr + " (masteridx,itemgubun,makerid,itemid,itemoption," + vbCrlf
				sqlStr = sqlStr + " itemname,itemoptionname,sellcash,suplycash,buycash,foreign_sellcash, foreign_suplycash," + vbCrlf
				sqlStr = sqlStr + " baljuitemno,realitemno,baljudiv)"  + vbCrlf
				sqlStr = sqlStr + " select " + CStr(iid) + ",itemgubun,makerid,itemid,itemoption," + vbCrlf
				sqlStr = sqlStr + " itemname,itemoptionname,sellcash,suplycash,buycash,foreign_sellcash, foreign_suplycash," + vbCrlf
				sqlStr = sqlStr + " sum(baljuitemno-realitemno),sum(baljuitemno-realitemno),baljudiv" + vbCrlf
				sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
				sqlStr = sqlStr + " where masteridx in (" + CStr(masteridxlist) + ") "
				sqlStr = sqlStr + " and baljuitemno <> realitemno "
				sqlStr = sqlStr + " and comment='5일내출고'"
				sqlStr = sqlStr + " and deldt is null"
				sqlStr = sqlStr + " group by itemgubun,makerid,itemid,itemoption,itemname,itemoptionname,sellcash,suplycash,buycash,foreign_sellcash, foreign_suplycash,baljudiv "
				'response.write sqlStr & "<Br>"
				rsget.Open sqlStr, dbget, 1

				''서머리 저장
				sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
				sqlStr = sqlStr + " set jumunsellcash=IsNULL(T.totsell,0)" + vbCrlf
				sqlStr = sqlStr + " ,jumunsuplycash=IsNULL(T.totsupp,0)" + vbCrlf
				sqlStr = sqlStr + " ,jumunbuycash=IsNULL(T.totbuy,0)" + vbCrlf
				sqlStr = sqlStr + " ,totalsellcash=IsNULL(T.realtotsell,0)" + vbCrlf
				sqlStr = sqlStr + " ,totalsuplycash=IsNULL(T.realtotsupp,0)" + vbCrlf
				sqlStr = sqlStr + " ,totalbuycash=IsNULL(T.realtotbuy,0)" + vbCrlf
				sqlStr = sqlStr + " ,jumunforeign_sellcash=IsNULL(T.totsellforeign,0)" + vbCrlf
				sqlStr = sqlStr + " ,jumunforeign_suplycash=IsNULL(T.totsuppforeign,0)" + vbCrlf
				sqlStr = sqlStr + " ,totalforeign_sellcash=IsNULL(T.realtotsellforeign,0)" + vbCrlf
				sqlStr = sqlStr + " ,totalforeign_suplycash=IsNULL(T.realtotsuppforeign,0)" + vbCrlf
				sqlStr = sqlStr + " from (select sum(sellcash*baljuitemno) as totsell, " + vbCrlf
				sqlStr = sqlStr + " sum(suplycash*baljuitemno) as totsupp, " + vbCrlf
				sqlStr = sqlStr + " sum(buycash*baljuitemno) as totbuy, " + vbCrlf
				sqlStr = sqlStr + " sum(sellcash*realitemno) as realtotsell, " + vbCrlf
				sqlStr = sqlStr + " sum(suplycash*realitemno) as realtotsupp, " + vbCrlf
				sqlStr = sqlStr + " sum(buycash*realitemno) as realtotbuy " + vbCrlf
				sqlStr = sqlStr + " ,sum(foreign_sellcash * baljuitemno) as totsellforeign " + vbCrlf
				sqlStr = sqlStr + " ,sum(foreign_suplycash * baljuitemno) as totsuppforeign " + vbCrlf
				sqlStr = sqlStr + " ,sum(foreign_sellcash * realitemno) as realtotsellforeign " + vbCrlf
				sqlStr = sqlStr + " ,sum(foreign_suplycash * realitemno) as realtotsuppforeign " + vbCrlf
				sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
				sqlStr = sqlStr + " where masteridx="  + CStr(iid) + vbCrlf
				sqlStr = sqlStr + " and deldt is null" + vbCrlf
				sqlStr = sqlStr + " ) as T" + vbCrlf
				sqlStr = sqlStr + " where [db_storage].[dbo].tbl_ordersheet_master.idx=" + CStr(iid)
				'response.write sqlStr & "<Br>"
				rsget.Open sqlStr, dbget, 1


				''브랜드 리스트
				brandlist = ""
				sqlStr = " select distinct makerid from [db_storage].[dbo].tbl_ordersheet_detail"
				sqlStr = sqlStr + " where masteridx=" + CStr(iid)
				'response.write sqlStr & "<Br>"
				rsget.Open sqlStr, dbget, 1
        		do until rsget.eof
        			brandlist = brandlist + rsget("makerid") + ","
        			rsget.movenext
        		loop
				rsget.close

				if brandlist<>"" then
					brandlist = Left(brandlist,Len(brandlist)-1)
					brandlist = Left(brandlist,255)
				end if

				sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + VbCrlf
				sqlStr = sqlStr + " set baljucode='" + baljucode + "'" + VbCrlf
				sqlStr = sqlStr + " , brandlist='" + brandlist + "'"
				sqlStr = sqlStr + " where idx=" + CStr(iid)
				'response.write sqlStr & "<Br>"
				rsget.Open sqlStr, dbget, 1


				''원출고지시서에 링크코드 저장.
				sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + VbCrlf
				sqlStr = sqlStr + " set clinkcode='" + baljucode + "'" + VbCrlf
				sqlStr = sqlStr + " where idx in (" + CStr(masteridxlist) + ") "
				'response.write sqlStr & "<Br>"
				rsget.Open sqlStr, dbget, 1
			End If

			Call checkAndWriteElapsedTime("007")
			''dbget.close() : response.end

			'카툰박스정보
			sqlstr = " delete d "
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " db_storage.dbo.tbl_cartoonbox_detail d "
			if application("Svr_Info") <> "Dev" then
				sqlStr = sqlStr + " 	join [LOGISTICSDB].[db_aLogistics].[dbo].tbl_Logistics_offline_cartoonbox c " + VbCrLf
			else
				sqlStr = sqlStr + " 	join [db_aLogistics].[dbo].tbl_Logistics_offline_cartoonbox c " + VbCrLf
			end if
			sqlStr = sqlStr + " on " + VbCrLf
			sqlStr = sqlStr + " 	1 = 1 " + VbCrLf
			sqlStr = sqlStr + " 	and d.baljudate = c.baljudate " + VbCrLf
			sqlStr = sqlStr + " 	and d.shopid = c.shopid " + VbCrLf
			sqlStr = sqlStr + " 	and d.innerboxno = c.innerboxno " + VbCrLf
			'response.write sqlStr & "<Br>"
			dbget.Execute sqlStr

			sqlstr = " insert into db_storage.dbo.tbl_cartoonbox_detail(" + VbCrLf
			sqlStr = sqlStr + " baljudate" + VbCrLf
			sqlStr = sqlStr + " ,shopid" + VbCrLf
			sqlStr = sqlStr + " ,cartoonboxno" + VbCrLf
			sqlStr = sqlStr + " ,cartoonboxweight" + VbCrLf
			sqlStr = sqlStr + " ,cartonboxsongjangdiv" + VbCrLf
			sqlStr = sqlStr + " ,cartonboxsongjangno" + VbCrLf
			sqlStr = sqlStr + " ,innerboxno" + VbCrLf
			sqlStr = sqlStr + " ,innerboxweight" + VbCrLf
			sqlStr = sqlStr + " ,innerboxidx" + VbCrLf
			sqlStr = sqlStr + " ) " + VbCrLf
			sqlStr = sqlStr + " 	select"
			sqlStr = sqlStr + " 	baljudate"
			sqlStr = sqlStr + " 	, shopid"
			sqlStr = sqlStr + " 	, cartoonboxno"
			sqlStr = sqlStr + " 	, 0"
			sqlStr = sqlStr + " 	, cartoonboxsongjangdiv"
			sqlStr = sqlStr + " 	, cartoonboxsongjangno"
			sqlStr = sqlStr + " 	, innerboxno"
			sqlStr = sqlStr + " 	, innerboxweight"
			sqlStr = sqlStr + " 	, ctIDX"
			if application("Svr_Info") <> "Dev" then
				sqlStr = sqlStr + " 	from [LOGISTICSDB].[db_aLogistics].[dbo].tbl_Logistics_offline_cartoonbox " + VbCrLf
			else
				sqlStr = sqlStr + " 	from [db_aLogistics].[dbo].tbl_Logistics_offline_cartoonbox " + VbCrLf
			end if
			sqlStr = sqlStr + " where siteseq = " + CStr(siteseq) + " "
			'response.write sqlStr & "<Br>"
			dbget.Execute sqlStr

			sqlstr = " delete from "
			if application("Svr_Info") <> "Dev" then
				sqlStr = sqlStr + " [LOGISTICSDB].[db_aLogistics].[dbo].tbl_Logistics_offline_cartoonbox " + VbCrLf
			else
				sqlStr = sqlStr + " [db_aLogistics].[dbo].tbl_Logistics_offline_cartoonbox " + VbCrLf
			end if
			sqlStr = sqlStr + " where siteseq = " + CStr(siteseq) + " "
			'response.write sqlStr & "<Br>"
			dbget.Execute sqlStr


			''주문상태 : 출고대기
			sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + VbCrlf
			sqlStr = sqlStr + " set statecd='6'" + VbCrlf
			sqlStr = sqlStr + " where idx in (" + CStr(masteridxlist) + ") "
			'response.write sqlStr & "<Br>"
			rsget.Open sqlStr, dbget, 1
		End If


		'로직스 MASTER 정보 - 입력완료표시
		sqlStr = " update [db_aLogistics].[dbo].tbl_Logistics_offline_order_master " + vbCrlf
		sqlStr = sqlStr + " set siteinsertdate='" + Left(now(), 10) + "'" + vbCrlf
		sqlStr = sqlStr + " where siteseq = " & siteseq & " and ordercode in (" + ordercodelistNoSiteInsertDate + ")  "
		'response.write sqlStr & "<Br>"
		rsget_Logistics.Open sqlStr, dbget_Logistics, 1
	End If

	Call checkAndWriteElapsedTime("008")
	''dbget.close() : response.end

	'// ========================================================================
	'// 주문 마스터 정보수정
	if IsWriteChulgoSheet And ordercodelistNoBeasongDate <> ""Then

		if (CStr(siteseq) = "10") Then
			'각 주문코드별 처리(출고데이타 생성 등)
			tmp = split(masteridxlist,",")
			for i=0 to UBound(tmp)
				if (Trim(tmp(i)) <> "") Then
					'확정수량에 대한 합계금액 계산
					sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
					sqlStr = sqlStr + " set jumunsellcash=IsNull(T.totsell,0)" + vbCrlf
					sqlStr = sqlStr + " ,jumunsuplycash=IsNull(T.totsupp,0)" + vbCrlf
					sqlStr = sqlStr + " ,jumunbuycash=IsNull(T.totbuy,0)" + vbCrlf
					sqlStr = sqlStr + " ,totalsellcash=IsNull(T.realtotsell,0)" + vbCrlf
					sqlStr = sqlStr + " ,totalsuplycash=IsNull(T.realtotsupp,0)" + vbCrlf
					sqlStr = sqlStr + " ,totalbuycash=IsNull(T.realtotbuy,0)" + vbCrlf
					sqlStr = sqlStr + " from (select sum(sellcash*baljuitemno) as totsell, " + vbCrlf
					sqlStr = sqlStr + " sum(suplycash*baljuitemno) as totsupp, " + vbCrlf
					sqlStr = sqlStr + " sum(buycash*baljuitemno) as totbuy, " + vbCrlf
					sqlStr = sqlStr + " sum(sellcash*realitemno) as realtotsell, " + vbCrlf
					sqlStr = sqlStr + " sum(suplycash*realitemno) as realtotsupp, " + vbCrlf
					sqlStr = sqlStr + " sum(buycash*realitemno) as realtotbuy " + vbCrlf
					sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
					sqlStr = sqlStr + " where masteridx="  + CStr(Trim(tmp(i))) + vbCrlf
					sqlStr = sqlStr + " and deldt is null" + vbCrlf
					sqlStr = sqlStr + " ) as T" + vbCrlf
					sqlStr = sqlStr + " where [db_storage].[dbo].tbl_ordersheet_master.idx=" + CStr(Trim(tmp(i)))
					'response.write sqlStr & "<Br>"
					rsget.Open sqlStr, dbget, 1

					'해당 출고지시코드/출고지시아이디에 대한 기본정보를 구한다.
					sqlStr = " select distinct m.baljuname, m.baljucode "
					sqlStr = sqlStr + " from [db_storage].[dbo].tbl_shopbalju b, [db_storage].[dbo].tbl_ordersheet_master m, [db_storage].[dbo].tbl_ordersheet_detail d "
					sqlStr = sqlStr + " where 1 = 1 "
					sqlStr = sqlStr + " and m.idx = d.masteridx "
					sqlStr = sqlStr + " and m.deldt is null "
					sqlStr = sqlStr + " and d.deldt is null "
					sqlStr = sqlStr + " and b.baljucode = m.baljucode "
					sqlStr = sqlStr + " and m.divcode in ('501','502','503') "
					sqlStr = sqlStr + " and m.statecd < '7' "
					sqlStr = sqlStr + " and b.baljuid = '" + CStr(baljuid) + "' "
					sqlStr = sqlStr + " and b.baljunum = " + CStr(siteBaljuKey) + " "
					sqlStr = sqlStr + " and m.idx = " + CStr(Trim(tmp(i))) + " "
					'response.write sqlStr & "<Br>"
					rsget.Open sqlStr,dbget,1

					baljuname = ""
					baljucode = ""
					if  not rsget.EOF  then
						baljuname = CStr(rsget("baljuname"))
						baljucode = CStr(rsget("baljucode"))
					end if
					rsget.Close

					''출고 마스타에 입력. *-1
					sqlStr = "select count(idx) as cnt from [db_storage].[dbo].tbl_ordersheet_detail d "
					sqlStr = sqlStr + " where d.masteridx = " + CStr(Trim(tmp(i))) + " "
					sqlStr = sqlStr + " and d.deldt is null "
					sqlStr = sqlStr + " and d.realitemno <> 0 "

					'response.write sqlStr & "<Br>"
					rsget.Open sqlStr, dbget, 1
                    	itemexists = rsget("cnt")>0
					rsget.close

					if itemexists Then
						'1.온라인 출고 마스타
						sqlStr = " select * from [db_storage].[dbo].tbl_acount_storage_master where 1=0 "

						'response.write sqlStr & "<Br>"
						rsget.Open sqlStr,dbget,1,3
						rsget.AddNew
						rsget("code") = ""
						rsget("socid") = baljuid
						rsget("socname") = baljuname
						rsget("chargeid") = session("ssBctId")
						rsget("divcode") = "006"
						rsget("vatcode") = "008"
						rsget("comment") = baljucode + " 주문 출고지시 후 자동출고처리"
						rsget("chargename") = session("ssBctCname")
						rsget("ipchulflag") = "S"

						rsget.update
						iid = rsget("id")
						rsget.close

						newbaljucode = "SO" + Format00(6,Right(CStr(iid),6))

						sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
						sqlStr = sqlStr + " set code='" + newbaljucode + "'" + VBCrlf
						sqlStr = sqlStr + " where id=" + CStr(iid)
						'response.write sqlStr & "<Br>"
						rsget.Open sqlStr,dbget,1

						'2.온라인 출고 디테일 입력
						sqlStr = " insert into [db_storage].[dbo].tbl_acount_storage_detail "
						sqlStr = sqlStr + " (mastercode,itemid,itemoption,sellcash,suplycash,itemno, "
						sqlStr = sqlStr + " buycash,mwgubun,iitemgubun,iitemname,iitemoptionname,imakerid) "
						sqlStr = sqlStr + " select '" + newbaljucode + "',d.itemid, d.itemoption, d.sellcash, d.suplycash, "
						sqlStr = sqlStr + " sum(d.realitemno*-1) as itemno, d.buycash,d.ipgoflag,d.itemgubun,d.itemname,d.itemoptionname,d.makerid "
						sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_detail d "
						sqlStr = sqlStr + " where d.masteridx = " + CStr(Trim(tmp(i))) + " "
						sqlStr = sqlStr + " and deldt is null "
						sqlStr = sqlStr + " and d.realitemno<>0 "
						sqlStr = sqlStr + " group by d.itemid, d.itemoption, d.sellcash, d.suplycash, d.buycash,d.ipgoflag, "
						sqlStr = sqlStr + " d.itemgubun,d.itemname,d.itemoptionname,d.makerid "
						'response.write sqlStr & "<Br>"
						rsget.Open sqlStr,dbget,1

						'3.온라인 출고 마스타 업데이트
						sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
						sqlStr = sqlStr + " set executedt='" + Left(now(), 10) + "'" + VBCrlf
						sqlStr = sqlStr + " ,scheduledt='" + Left(now(), 10) + "'" + VBCrlf
						sqlStr = sqlStr + " ,totalsellcash=IsNULL(T.totsell,0)" + VBCrlf
						sqlStr = sqlStr + " ,totalsuplycash=IsNULL(T.totsupp,0)" + VBCrlf
						sqlStr = sqlStr + " ,totalbuycash=IsNULL(T.totbuy,0)" + VBCrlf
						sqlStr = sqlStr + " ,indt=getdate()" + VBCrlf
						sqlStr = sqlStr + " ,updt=getdate()" + VBCrlf
						sqlStr = sqlStr + " from (select sum(sellcash*itemno) as totsell, " + vbCrlf
						sqlStr = sqlStr + " sum(suplycash*itemno) as totsupp, " + vbCrlf
						sqlStr = sqlStr + " sum(buycash*itemno) as totbuy " + vbCrlf
						sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
						sqlStr = sqlStr + " where mastercode='"  + CStr(newbaljucode) + "'" + vbCrlf
						sqlStr = sqlStr + " and deldt is null" + vbCrlf
						sqlStr = sqlStr + " ) as T"
						sqlStr = sqlStr + " where id=" + CStr(iid)
						'response.write sqlStr & "<Br>"
						rsget.Open sqlStr,dbget,1

						'4.상태변경
						sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
						sqlStr = sqlStr + " set statecd='7'" + vbCrlf
						sqlStr = sqlStr + " ,ipgodate='" + Left(now(), 10) + "'" + vbCrlf
						sqlStr = sqlStr + " ,alinkcode='"  + CStr(newbaljucode) + "'" + vbCrlf
						sqlStr = sqlStr + " where idx = " + CStr(Trim(tmp(i))) + " "
						'response.write sqlStr & "<Br>"
						rsget.Open sqlStr, dbget, 1

						'' 입/출고 재고반영
						sqlStr = "exec db_summary.dbo.sp_ten_recentIpChul_Update '" & newbaljucode & "','','',0,'',''"
						'response.write sqlStr & "<Br>"
						dbget.Execute sqlStr

						'// 매장재고 반영
                        if (baljukey <> 74851) then
						    sqlStr = "exec [db_summary].[dbo].[sp_Ten_Shop_Stock_RecentLogicsIpChul_Update] '" & baljuid & "', '" & newbaljucode & "' "
						    'response.write sqlStr & "<Br>"
						    dbget.Execute sqlStr
                        end if
					else
						'// 출고가능한 상품없는 경우 : 그대로 출고완료로 전환

						'4.상태변경
						sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
						sqlStr = sqlStr + " set statecd='7'" + vbCrlf
						''sqlStr = sqlStr + " ,ipgodate='" + Left(now(), 10) + "'" + vbCrlf
						''sqlStr = sqlStr + " ,alinkcode='"  + CStr(newbaljucode) + "'" + vbCrlf
						sqlStr = sqlStr + " where idx = " + CStr(Trim(tmp(i))) + " "
						'response.write sqlStr & "<Br>"
						rsget.Open sqlStr, dbget, 1
					End If
				End If
			Next
		End If

		Call checkAndWriteElapsedTime("009")
		''dbget.close() : response.end

		'로직스 기본 MASTER 정보 수정
		sqlStr = " update [db_aLogistics].[dbo].tbl_Logistics_offline_order_master " + vbCrlf
		sqlStr = sqlStr + " set beasongdate='" + Left(now(), 10) + "'" + vbCrlf
		sqlStr = sqlStr + " where siteseq = " & siteseq & " and ordercode in (" + ordercodelistNoBeasongDate + ")  "
		'response.write sqlStr & "<Br>"
		rsget_Logistics.Open sqlStr, dbget_Logistics, 1

		if (CStr(siteseq) = "10") Then
			'어드민 기본 MASTER 정보 수정
			sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
			sqlStr = sqlStr + " set beasongdate='" + Left(now(), 10) + "'" + vbCrlf
			sqlStr = sqlStr + " where baljucode in (" + ordercodelistNoBeasongDate + ") "
			'response.write sqlStr & "<Br>"
			rsget.Open sqlStr, dbget, 1

			'// 매장재고 반영(배송중)
			sqlStr = " update d "
			sqlStr = sqlStr + " set d.shopReceive = 'N' "
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " [db_storage].[dbo].tbl_cartoonbox_detail d "
			sqlStr = sqlStr + " join ( "
			sqlStr = sqlStr + " 	select distinct b.baljuid as shopid, DATEADD(dd, DATEDIFF(dd, 0, b.baljudate), 0) as baljudate, d.packingstate as innerboxno "
			sqlStr = sqlStr + " 	from "
			sqlStr = sqlStr + " 		[db_storage].[dbo].tbl_shopbalju b "
			sqlStr = sqlStr + " 		join [db_storage].[dbo].tbl_ordersheet_master m "
			sqlStr = sqlStr + " 		on "
			sqlStr = sqlStr + " 			b.baljucode = m.baljucode "
			sqlStr = sqlStr + " 		join [db_storage].[dbo].tbl_ordersheet_detail d "
			sqlStr = sqlStr + " 		on "
			sqlStr = sqlStr + " 			m.idx = d.masteridx "
			sqlStr = sqlStr + " 	where "
			sqlStr = sqlStr + " 		1 = 1 "
			sqlStr = sqlStr + " 		and b.baljunum = " & siteBaljuKey
			sqlStr = sqlStr + " ) T "
			sqlStr = sqlStr + " on "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and d.shopid = T.shopid "
			sqlStr = sqlStr + " 	and d.baljudate = T.baljudate "
			sqlStr = sqlStr + " 	and d.innerboxno = T.innerboxno "
			'response.write sqlStr & "<Br>"
			dbget.Execute sqlStr

			'' 오프 접수수량 재계산
			''sqlstr = " exec [db_summary].[dbo].sp_ten_RealtimeStock_offjupsuAll"
            sqlstr = " exec [db_summary].[dbo].[sp_Ten_RealtimeStock_offjupsuAll_bySiteBaljuKey] " & siteBaljuKey
			'response.write sqlStr & "<Br>"
			dbget.Execute sqlStr


			''siteBaljukey   ''오프 출고 상품 한정 비교 재고로 한정 재설정.. 2011-06 추가.
			if (siteBaljukey<>0) then
				sqlstr = " exec [db_summary].[dbo].[sp_Ten_RealtimeStock_itemLimitByOffChulgo] "&siteBaljukey
				'response.write sqlStr & "<Br>"
				dbget.Execute sqlStr
			end If


			'' 매장 접수수량 재계산 2011-08 추가
			'기주문 재계산하지 않는다.(매장기준으로 의미가 없다.)
			'// 전체 재계산보다는 출고지시상품목록만 업데이트하는 것으로 수정 필요
			'sqlstr = " exec [db_summary].dbo.[sp_Ten_Shop_Stock_PreOrderUpdate_ALL]"
			'dbget.Execute sqlStr
		End If
	End If

	Call checkAndWriteElapsedTime("010")
	''dbget.close() : response.end

	'// ========================================================================
	'// 출고지시 마스터 정보수정
	if (isWait = "Y") then
		'로직스 출고지시마스터 IsFinished="W" 입력
		sqlStr = " update db_aLogistics.dbo.tbl_Logistics_offline_baljumaster " + VbCrLf
		sqlStr = sqlStr + " set IsFinished = 'W'  " + VbCrLf
		sqlStr = sqlStr + " where baljuKey = " & baljuKey & " and siteseq = " & siteseq & " " + VbCrLf
		'response.write sqlStr & "<Br>"
		rsget_Logistics.Open sqlStr, dbget_Logistics, 1
	elseif IsWriteChulgoSheet then
		'로직스 출고지시마스터 IsFinished="Y" 입력
		'로직스 오더마스터 beasongdate 입력
		sqlStr = " update db_aLogistics.dbo.tbl_Logistics_offline_baljumaster " + VbCrLf
		sqlStr = sqlStr + " set IsFinished = 'Y'  " + VbCrLf
		sqlStr = sqlStr + " where baljuKey = " & baljuKey & " and siteseq = " & siteseq & " " + VbCrLf
		'response.write sqlStr & "<Br>"
		rsget_Logistics.Open sqlStr, dbget_Logistics, 1
	end If


	If (CStr(siteseq) <> "10") Then
		'3PL
		dim STOCK_GUID
		response.write "<script>alert('에러 - 처리안함.');</script>"
		response.End
	End If
end if

Call checkAndWriteElapsedTime("011")
''dbget.close() : response.end

%>

<script language="javascript">
	alert('저장 되었습니다.');
	location.replace('baljulist_offline_new.asp');
</script>

<!-- #include virtual="/lib/db/db_LogisticsClose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
