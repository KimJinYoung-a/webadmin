<%@ language=vbscript %>
<% option explicit %>
<%
Server.ScriptTimeOut = 60*15		' 물류센터 재고재작성 15분
%>
<%
'###########################################################
' Description : 재고자산
' History : 이상구 생성
'			2023.05.03 한용민 수정(검색조건추가)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionSTAdmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stockclass/monthlystockcls.asp"-->
<%
dim mode, yyyymm
dim itemid, itemoption
mode = request("mode")
yyyymm = request("yyyymm")

itemid = request("itemid")
itemoption = request("itemoption")

dim yyyy1,mm1,isusing,sysorreal, research, newitem, vatyn, minusinc, bPriceGbn, i
dim mwgubun, buseo, itemgubun, stplace, purchasetype, showsuply, dtype, makerid, shopid, etcjungsantype, showDiff
dim brandUseYN
	yyyy1       = requestCheckvar(request("yyyy1"),10)
	mm1         = requestCheckvar(request("mm1"),10)
	isusing     = requestCheckvar(request("isusing"),10)
	sysorreal   = requestCheckvar(request("sysorreal"),10)
	research    = requestCheckvar(request("research"),10)
	newitem     = requestCheckvar(request("newitem"),10)
	mwgubun     = requestCheckvar(request("mwgubun"),10)
	vatyn       = requestCheckvar(request("vatyn"),10)
	minusinc   = requestCheckvar(request("minusinc"),10)
	bPriceGbn   = requestCheckvar(request("bPriceGbn"),10)
	buseo       = requestCheckvar(request("buseo"),10)
	itemgubun   = requestCheckvar(request("itemgubun"),10)
	purchasetype   = requestCheckvar(request("purchasetype"),10)
	stplace     = requestCheckvar(request("stplace"),10)
	showsuply   = requestCheckvar(request("showsuply"),10)
	dtype       = requestCheckvar(request("dtype"),10)
	makerid     = requestCheckvar(request("makerid"),32)
	shopid     = requestCheckvar(request("shopid"),32)
	etcjungsantype      = requestCheckvar(request("etcjungsantype"),10)
	showDiff      = requestCheckvar(request("showDiff"),10)
	brandUseYN      = requestCheckvar(request("brandUseYN"),10)

if (makerid<>"") then dtype=""
if (sysorreal="") then sysorreal="sys"  ''real
if (research="") and (bPriceGbn = "") then
    bPriceGbn="V"
end if
if (stplace="") then
    stplace="L"
	showDiff = "Y"
end if
if (research="") then
	if (itemgubun = "") then
		'itemgubun = "AA"
	end if
	if (buseo = "") then
		buseo = "3X"
	end if
end if

dim nowdate
if yyyy1="" then
	nowdate = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(nowdate),4)
	mm1 = Mid(CStr(nowdate),6,2)
end if

dim totno, totbuy, subTotno, subTotbuy '', totavgBuy, offtotavgBuy
dim totPreno, totPrebuy     , subPreno, subPrebuy
dim totIpno,totIpBuy        , subIpno, subIpBuy
dim totLossno, totLossBuy   , subLossno, subLossBuy
dim totSellno, totSellBuy   , subSellno, subSellBuy
dim totOffChulno, totOffChulBuy  , subOffChulno, subOffChulBuy
dim totEtcChulno, totEtcChulBuy  , subEtcChulno, subEtcChulBuy
dim totCsChulno, totCsChulBuy    , subCsChulno, subCsChulBuy
dim iURL, iURLEtc, nBusiName, diffStock, diffStockPrc, diffStockW
DIM isGroupByBrand : isGroupByBrand = (dtype="mk")
Dim isItemList : isItemList = (makerid<>"")

dim totErrBadItemno, totErrBadItemBuy, subErrBadItemno, subErrBadItemBuy
dim totMoveItemno, totMoveItemBuy, subMoveItemno, subMoveItemBuy
dim totErrRealCheckno, totErrRealCheckBuy, subErrRealCheckno, subErrRealCheckBuy
dim totRealStockno, totRealStockBuy, subRealStockno, subRealStockBuy
dim totErrRealCheckBuyPlus, totErrRealCheckBuyMinus


dim sqlStr, resultrows
dim diffMonth

' "[통계]재고자산>>재고자산-물류" 재고자산재작성 버튼
if mode="monthlystock" then
    ''tbl_monthly_logisstock_summary :: DailyLogisStockMaker_일_2시55 / DailyLogisStockMaker_ThisDate_일_7시55 스케줄에 포함됨.
    '' 전달 데이커 까지는 지운 후 재작서 .// 그 이전은 재고수량 재작성(업데이트)

    ''diffMonth = dateDiff("m",yyyymm+"-01",now())
 'rw "수정중"
 'response.end

	'// 누적재고 생성
    sqlStr = " db_summary.dbo.sp_Ten_monthly_Acc_LogisStockMake '"&yyyymm&"'"
    dbget.execute sqlStr,resultrows

    response.write "<br>누적재고 생성"

    '// 매입구분 입력
    sqlStr = "exec db_summary.dbo.sp_Ten_monthlyLogisstock_mwFlagUpdate '"&yyyymm&"'"
    dbget.execute sqlStr,resultrows

    response.write "<br>매입구분 입력"

    '// 출고플래그 생성
    sqlStr = "EXEC [db_summary].[dbo].[sp_Ten_monthly_ChulgoFlagUpdate] '"&yyyymm&"','3pl'" ' 3pl
    dbget.execute sqlStr,resultrows

    sqlStr = "EXEC [db_summary].[dbo].[sp_Ten_monthly_ChulgoFlagUpdate] '"&yyyymm&"','C'"  ' 출고정산
    dbget.execute sqlStr,resultrows

    sqlStr = "EXEC [db_summary].[dbo].[sp_Ten_monthly_ChulgoFlagUpdate] '"&yyyymm&"','W'"  ' 출고위탁
    dbget.execute sqlStr,resultrows

    sqlStr = "EXEC [db_summary].[dbo].[sp_Ten_monthly_ChulgoFlagUpdate] '"&yyyymm&"','M'"  ' 온라인매입.
    dbget.execute sqlStr,resultrows

    sqlStr = "EXEC [db_summary].[dbo].[sp_Ten_monthly_ChulgoFlagUpdate] '"&yyyymm&"','F'"  ' 오프매입.
    dbget.execute sqlStr,resultrows

    sqlStr = "EXEC [db_summary].[dbo].[sp_Ten_monthly_ChulgoFlagUpdate] '"&yyyymm&"','E'"  ' 기타출고.
    dbget.execute sqlStr,resultrows

    response.write "<br>출고플래그 생성"

	'// 입고내역 서머리 생성
    sqlStr = " db_summary.dbo.sp_Ten_monthlyLogisstock_ipgoSumMake '"&yyyymm&"', 'L' "
    dbget.execute sqlStr,resultrows

    response.write "<br>입고내역 서머리 생성"

	'// 평균매입가 계산
    sqlStr = " db_summary.dbo.sp_Ten_monthlyLogisstock_avgipgoPrice '"&yyyymm&"', 'L' "
    dbget.execute sqlStr,resultrows

    response.write "<br>평균매입가 계산"

''	''기 데이타 삭제
''	sqlStr = "delete from [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary"+ VbCrlf
''	sqlStr = sqlStr + " where yyyymm='" + yyyymm + "'"+ VbCrlf
''	dbget.execute sqlStr
''
''
''	''상품이 업는경우. 기본값 입력
''	sqlStr = "insert into [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary"+ VbCrlf
''	sqlStr = sqlStr + " (yyyymm,itemgubun,itemid,itemoption"+ VbCrlf
''	sqlStr = sqlStr + " ,ipgono,reipgono,totipgono,offchulgono,offrechulgono,etcchulgono"+ VbCrlf
''	sqlStr = sqlStr + " ,etcrechulgono,totchulgono,sellno,resellno,totsellno"+ VbCrlf
''	sqlStr = sqlStr + " ,errcsno,errbaditemno,errrealcheckno,erretcno,toterrno"+ VbCrlf
''	sqlStr = sqlStr + " ,offsellno,totsysstock,availsysstock,realstock,lossno)"+ VbCrlf
''
''	sqlStr = sqlStr + " 	select '" + yyyymm + "',itemgubun,itemid,itemoption,"+ VbCrlf
''	sqlStr = sqlStr + "  	sum(ipgono) as ipgono,sum(reipgono) as reipgono,sum(totipgono) as totipgono,"+ VbCrlf
''	sqlStr = sqlStr + "  	sum(offchulgono) as offchulgono,sum(offrechulgono) as offrechulgono,sum(etcchulgono) as etcchulgono,"+ VbCrlf
''	sqlStr = sqlStr + " 	sum(etcrechulgono) as etcrechulgono,sum(totchulgono) as totchulgono,sum(sellno) as sellno,"+ VbCrlf
''	sqlStr = sqlStr + "  	sum(resellno) as resellno,sum(totsellno) as totsellno,"+ VbCrlf
''	sqlStr = sqlStr + " 	sum(errcsno) as errcsno,sum(errbaditemno) as reipgono,sum(errrealcheckno) as reipgono,"+ VbCrlf
''	sqlStr = sqlStr + "  	sum(erretcno) as erretcno,sum(toterrno) as toterrno,"+ VbCrlf
''	sqlStr = sqlStr + " 	sum(offsellno) as offsellno,sum(totsysstock) as totsysstock,"+ VbCrlf
''	sqlStr = sqlStr + "  	sum(availsysstock) as availsysstock,sum(realstock) as realstock,sum(lossno) as lossno"+ VbCrlf
''	sqlStr = sqlStr + " 	from [db_summary].[dbo].tbl_monthly_logisstock_summary"+ VbCrlf
''	sqlStr = sqlStr + " 	where yyyymm<='" + yyyymm + "'"+ VbCrlf
''	sqlStr = sqlStr + " 	group by itemgubun,itemid,itemoption "+ VbCrlf
''
''	 dbget.execute sqlStr, resultrows

    '' 작성시 매입구분 입력 // 작성시 매입가// 부서구분// 작성시 makerid // 평균매입가 수정
    ''if (resultrows>0) then
    ''    sqlStr = "exec db_summary.dbo.sp_Ten_monthlyLogisstock_mwFlagUpdate '"&yyyymm&"'"
    ''    dbget.execute sqlStr
    ''end if
	response.write "<script type='text/javascript'>"
    response.write "    alert('작성 되었습니다.');"
	response.write "    opener.location.reload();"
    'response.write "    self.close();"
    response.write "</script>"
	dbget.close()	:	response.End

' "[통계]재고자산>>재고자산-물류" 일별입출고재작성STEP1 버튼
elseif mode="dailystock1" then
    '-- 일별 입출고 서머리 오늘날짜 온라인 판매집계
    sqlStr = "exec db_summary.[dbo].[sp_Ten_recentOnlineSell_Update_All]"
    dbget.execute sqlStr,resultrows

    response.write "<br>일별 입출고 서머리 오늘날짜 온라인 판매집계"

    '-- 일별 입출고 서머리 오늘날짜 입출데이터
    sqlStr = "exec db_summary.[dbo].[usp_TEN_daily_logisstock_summary_currentdate_make_ipchuldata]"
    dbget.execute sqlStr,resultrows

    response.write "<br>일별 입출고 서머리 오늘날짜 입출데이터"

    '-- 월별 물류재고 서머리 오늘날짜재계산
    sqlStr = "exec db_summary.[dbo].[usp_TEN_monthly_logisstock_summary_currentdate_make]"
    dbget.execute sqlStr,resultrows

    response.write "<br>월별 물류재고 서머리 오늘날짜재계산"

    '-- 현재물류재고 서머리 오늘날짜재계산
    sqlStr = "exec db_summary.[dbo].[usp_TEN_current_logisstock_summary_currentdate_make]"
    dbget.execute sqlStr,resultrows

    response.write "<br>현재물류재고 서머리 오늘날짜재계산"

    '-- 현재물류재고 서머리 오늘날짜재계산 온라인 판매집계
    sqlStr = "exec db_summary.[dbo].[usp_TEN_current_logisstock_summary_currentdate_make_on_ipkum_chulgo]"
    dbget.execute sqlStr,resultrows

    response.write "<br>현재물류재고 서머리 오늘날짜재계산 온라인 판매집계"

    '-- 3PL 온라인 7일 판매수량(usp_TEN_current_logisstock_summary_currentdate_make_on_ipkum_chulgo 에 포함시킴)
    ''sqlStr = "exec [db_summary].[dbo].[usp_TPL_7dayOnlineSell_Update]"
    ''dbget.execute sqlStr,resultrows

    response.write "<br>3PL 온라인 7일 판매수량"

    '-- 현재물류재고 서머리 오늘날짜재계산 오프라인7일판매수량 업데이트
    sqlStr = "exec db_summary.[dbo].[usp_TEN_current_logisstock_summary_currentdate_make_offchulgo7days]"
    dbget.execute sqlStr,resultrows

    response.write "<br>현재물류재고 서머리 오늘날짜재계산 오프라인7일판매수량 업데이트"

    '-- 현재물류재고 서머리 오늘날짜재계산 오프라인 입출고
    sqlStr = "exec db_summary.[dbo].[usp_TEN_current_logisstock_summary_currentdate_make_offipchul]"
    dbget.execute sqlStr,resultrows

    response.write "<br>현재물류재고 서머리 오늘날짜재계산 오프라인 입출고"

    '-- 현재물류재고 서머리 오늘날짜재계산 기타정보 업데이트
    sqlStr = "exec db_summary.[dbo].[usp_TEN_current_logisstock_summary_currentdate_make_etcinfo]"
    dbget.execute sqlStr,resultrows

    response.write "<br>현재물류재고 서머리 오늘날짜재계산 기타정보 업데이트"

    '-- 현재물류재고 서머리 오늘날짜재계산 기타정보 업데이트2
    sqlStr = "exec [db_summary].[dbo].[usp_Ten_Refresh_MakeItem_RequireNO] NULL, NULL, NULL"
    dbget.execute sqlStr,resultrows

    response.write "<br>현재물류재고 서머리 오늘날짜재계산 기타정보 업데이트2"

	response.write "<script type='text/javascript'>"
    response.write "    alert('작성 되었습니다.');"
	response.write "    opener.location.reload();"
    'response.write "    self.close();"
    response.write "</script>"
	dbget.close()	:	response.End

' "[통계]재고자산>>재고자산-물류" 일별입출고재작성STEP2 버튼
elseif mode="dailystock2" then
    '-- 일별 입출고 서머리 이번달 판매데이터
    sqlStr = "exec db_summary.[dbo].[usp_TEN_daily_logisstock_summary_currentmonth_make_online_selldata]"
    dbget.execute sqlStr,resultrows

    response.write "<br>일별 입출고 서머리 이번달 판매데이터"

    '-- 3pl stock STEP1
    sqlStr = "exec [db_summary].[dbo].[usp_TPL_recentOnlineSell_Update_All]"
    dbget.execute sqlStr,resultrows

    response.write "<br>3pl stock STEP1"

    '-- 현재물류재고 서머리 이번달 재계산 히치하이커 정기구독
    sqlStr = "exec db_summary.[dbo].sp_Ten_recentOnlineSell_Update_With_6MonthAgo_loop_item_standing"
   	dbget.CommandTimeout = 60*5   ' 5분
    dbget.execute sqlStr,resultrows

    response.write "<br>현재물류재고 서머리 이번달 재계산 히치하이커 정기구독"

    '-- 일별 입출고 서머리 이번달 입출데이터
    sqlStr = "exec db_summary.[dbo].[usp_TEN_daily_logisstock_summary_currentmonth_make_ipchuldata]"
    dbget.execute sqlStr,resultrows

    response.write "<br>일별 입출고 서머리 이번달 입출데이터"

    '-- 월별 물류재고 서머리 이번달 재계산
    sqlStr = "exec db_summary.[dbo].[usp_TEN_monthly_logisstock_summary_currentdate_make]"
    dbget.execute sqlStr,resultrows

    response.write "<br>월별 물류재고 서머리 이번달 재계산"

    '-- 현재물류재고 서머리 이번달 재계산
    sqlStr = "exec db_summary.[dbo].[usp_TEN_current_logisstock_summary_currentmonth_make]"
    dbget.execute sqlStr,resultrows

    response.write "<br>현재물류재고 서머리 이번달 재계산"

    '-- 현재물류재고 서머리 이번달 재계산 온라인 판매집계
    sqlStr = "exec db_summary.[dbo].[usp_TEN_current_logisstock_summary_currentdate_make_on_ipkum_chulgo]"
    dbget.execute sqlStr,resultrows

    response.write "<br>현재물류재고 서머리 이번달 재계산 온라인 판매집계"

    '-- 3pl stock STEP2
    sqlStr = "exec [db_summary].[dbo].[usp_TPL_recentOnlineJupsu_Update_All]"
    dbget.execute sqlStr,resultrows

    response.write "<br>3pl stock STEP2"

    '-- 3pl stock STEP3
    sqlStr = "exec [db_summary].[dbo].[usp_TPL_7dayOnlineSell_Update]"
    dbget.execute sqlStr,resultrows

    response.write "<br>3pl stock STEP3"

    '-- 현재물류재고 서머리 이번달 재계산 오프라인7일판매수량 업데이트
    sqlStr = "exec db_summary.[dbo].[usp_TEN_current_logisstock_summary_currentdate_make_offchulgo7days]"
    dbget.execute sqlStr,resultrows

    response.write "<br>현재물류재고 서머리 이번달 재계산 오프라인7일판매수량 업데이트"

    '-- 현재물류재고 서머리 이번달 재계산 오프라인 입출고
    sqlStr = "exec db_summary.[dbo].[usp_TEN_current_logisstock_summary_currentdate_make_offipchul]"
    dbget.execute sqlStr,resultrows

    response.write "<br>현재물류재고 서머리 이번달 재계산 오프라인 입출고"

    '-- 현재물류재고 서머리 이번달 재계산 기타정보 업데이트
    sqlStr = "exec db_summary.[dbo].[usp_TEN_current_logisstock_summary_currentdate_make_etcinfo]"
    dbget.execute sqlStr,resultrows

    response.write "<br>현재물류재고 서머리 이번달 재계산 기타정보 업데이트"

    '-- 핑거스 매입재고 작성
    sqlStr = "exec [db_summary].[dbo].[sp_Ten_recentOnlineSell_Update_ACA] 'thefingerscollabo'"
    dbget.execute sqlStr,resultrows

    response.write "<br>핑거스 매입재고 작성"

    '-- 매장 판매 상품 재고 요약 테이블 작성
    sqlStr = "exec db_summary.[dbo].[usp_Ten_ShopItem_Front_RealStockMake]"
    dbget.execute sqlStr,resultrows

    response.write "<br>매장 판매 상품 재고 요약 테이블 작성"

	response.write "<script type='text/javascript'>"
    response.write "    alert('작성 되었습니다.');"
	response.write "    opener.location.reload();"
    'response.write "    self.close();"
    response.write "</script>"
	dbget.close()	:	response.End

elseif mode="monthlystockexl" then

    sqlStr = "exec [db_datamart].[dbo].[sp_Ten_monthlystock_Asset_Make] '" & yyyy1 & "-" & mm1 & "', 'L','" & shopid & "','"&buseo&"','"&itemgubun&"','"&mwgubun&"','"&vatyn&"','"&purchasetype&"','"&CHKIIF(showsuply="on",1,0)&"','"&CHKIIF(dtype="mk",1,0)&"','"&etcjungsantype&"','" & brandUseYN & "',''"
    db3_dbget.CommandTimeout = 60*5   ' 5분
    db3_dbget.execute sqlStr

	response.write "<script>alert('생성 되었습니다.');</script>"
	response.write "<script>opener.location.reload();window.close();</script>"
	dbget.close()	:	response.End

elseif mode="monthlystockoneitem" then
	'// 누적재고, 서머리 재작성(물류, 위탁상품)

    sqlStr = " db_summary.[dbo].[sp_Ten_monthly_Acc_LogisStockMake_OneItem] '" + CStr(yyyymm) + "', '" + CStr(itemgubun) + "', " + CStr(itemid) + ", '" + CStr(itemoption) + "' "
    dbget.execute sqlStr,resultrows

    sqlStr = " db_summary.[dbo].[sp_Ten_monthlyLogisstock_mwFlagUpdate_OneItem] '" + CStr(yyyymm) + "', '" + CStr(itemgubun) + "', " + CStr(itemid) + ", '" + CStr(itemoption) + "' "
    dbget.execute sqlStr,resultrows

    sqlStr = " db_summary.[dbo].[sp_Ten_monthly_Stockledger_Make_OneItem] '" + CStr(yyyymm) + "','L', '" + CStr(itemgubun) + "', " + CStr(itemid) + ", '" + CStr(itemoption) + "' "
    dbget.execute sqlStr,resultrows

	response.write "<script>alert('작성 되었습니다.');</script>"
	response.write "<script>opener.location.reload();window.close();</script>"
	dbget.close()	:	response.End

elseif mode="monthlystocksum" then
	'// 서머리정보 재작성(물류+매장)

    sqlStr = " db_summary.dbo.sp_Ten_monthly_Stockledger_Make '"&yyyymm&"', 'L' "
    dbget.execute sqlStr,resultrows

    sqlStr = " db_summary.dbo.sp_Ten_monthly_Stockledger_Make '"&yyyymm&"', 'S' "
    dbget.execute sqlStr,resultrows

	response.write "<script>alert('작성 되었습니다.');</script>"
	response.write "<script>opener.location.reload();window.close();</script>"
	dbget.close()	:	response.End

elseif mode="shopmonthly" then

	''사용안함(3단계로 나눔)
	response.end

    sqlStr = "exec db_summary.dbo.sp_Ten_monthly_Acc_ShopstockMake '"&yyyymm&"'"
    dbget.execute sqlStr, resultrows

	'// 입고내역 서머리
    sqlStr = " db_summary.dbo.sp_Ten_monthlyLogisstock_ipgoSumMake '"&yyyymm&"', 'S' "
    dbget.execute sqlStr,resultrows

	'// 평균매입가(매장)
    sqlStr = " db_summary.dbo.sp_Ten_monthlyLogisstock_avgipgoPrice '"&yyyymm&"', 'S' "
    dbget.execute sqlStr,resultrows

    sqlStr = "exec db_summary.dbo.sp_Ten_monthlyShopstock_mwFlagUpdate '"&yyyymm&"'"
    dbget.execute sqlStr

    ''수익율분석>>오프라인수익서머리 의 기말재고작성
    sqlStr = " exec [db_summary].[dbo].[usp_Ten_Shop_Create_ShopStockSUM] '"&yyyymm&"' "
    dbget.Execute sqlStr

    response.write "<script>alert('작성 되었습니다. ');</script>"
	response.write "<script>opener.location.reload();window.close();</script>"
	dbget.close()	:	response.End

elseif mode="shopdailystock1" then
    ' 일별 매장재고 서머리 이번달 입출,판매 데이터
    sqlStr = "exec db_summary.dbo.sp_Ten_Shop_Stock_RecentUpdateALL"
   	dbget.CommandTimeout = 60*5   ' 5분
    dbget.execute sqlStr, resultrows

    response.write "<br>일별 매장재고 서머리 이번달 입출,판매 데이터"

    ' 현재 매장재고 서머리 이번달 판매 데이터
    sqlStr = "exec db_summary.dbo.sp_Ten_Shop_Stock_7daysSellUpdate"
    dbget.execute sqlStr, resultrows

    response.write "<br>현재 매장재고 서머리 이번달 판매 데이터"

    ' 현재 매장재고 서머리 이번달 입출고 기주문 업데이트
    sqlStr = "exec db_summary.[dbo].[sp_Ten_Shop_Stock_PreOrderUpdate_ALL]"
    dbget.execute sqlStr, resultrows

    response.write "<br>현재 매장재고 서머리 이번달 입출고 기주문 업데이트"

    ' 현재 매장재고 서머리 이번달 매장 입출고. 이동중수량.
    sqlStr = "exec db_summary.dbo.[usp_Ten_ShopChulgo_Update]"
    dbget.execute sqlStr, resultrows

    response.write "<br>현재 매장재고 서머리 이번달 매장 입출고. 이동중수량."

    ' 현재 매장재고 서머리 이번달 매장 입출고. 반품중수량.
    sqlStr = "exec db_summary.dbo.[usp_Ten_ShopReturn_Update]"
    dbget.execute sqlStr, resultrows

    response.write "<br>현재 매장재고 서머리 이번달 매장 입출고. 반품중수량."

	response.write "<script type='text/javascript'>"
    response.write "    alert('작성 되었습니다.');"
	response.write "    opener.location.reload();"
    'response.write "    self.close();"
    response.write "</script>"
	dbget.close()	:	response.End

elseif mode="shopmonthlystock" then
    ' 브랜드 마진. 브랜드 정산구분 작성
    sqlStr = "exec db_summary.[dbo].[sp_Ten_monthly_ShopDesigner_Make] '"&yyyymm&"'"
    dbget.execute sqlStr, resultrows

    response.write "<br>브랜드 마진. 브랜드 정산구분 작성"

    ' 월별 매장월말재고 서머리 이번달 매입설정
    sqlStr = "exec [db_summary].[dbo].[sp_Ten_monthly_Acc_ShopstockMake_SetMWDiv] '"&yyyymm&"'"
    dbget.execute sqlStr, resultrows

    response.write "<br>월별 매장월말재고 서머리 이번달 매입설정"

    ' 월별 매장재고 서머리 이번달 입출고
    sqlStr = "exec db_summary.dbo.sp_Ten_monthly_Acc_ShopstockMake '"&yyyymm&"'"
    dbget.execute sqlStr, resultrows

    response.write "<br>월별 매장재고 서머리 이번달 입출고"

    ' 월별 매장재고 서머리 이번달 입출고 매입구분,평균매입가
    sqlStr = "exec db_summary.dbo.sp_Ten_monthlyShopstock_mwFlagUpdate '"&yyyymm&"'"
    dbget.execute sqlStr, resultrows

    response.write "<br>월별 매장재고 서머리 이번달 입출고 매입구분,평균매입가"

    ' 입고내역 서머리 생성
    sqlStr = "exec db_summary.dbo.sp_Ten_monthlyLogisstock_ipgoSumMake '"&yyyymm&"', 'S'"
    dbget.execute sqlStr, resultrows

    response.write "<br>입고내역 서머리 생성"

    ' 평균매입가 계산
    sqlStr = "exec db_summary.dbo.sp_Ten_monthlyLogisstock_avgipgoPrice '"&yyyymm&"', 'S'"
    dbget.execute sqlStr, resultrows

    response.write "<br>평균매입가 계산"

    ' 기말재고작성.
    sqlStr = "exec [db_summary].[dbo].[usp_Ten_Shop_Create_ShopStockSUM] '"&yyyymm&"'"
    dbget.execute sqlStr, resultrows

    response.write "<br>기말재고작성"

	response.write "<script type='text/javascript'>"
    response.write "    alert('작성 되었습니다.');"
	response.write "    opener.location.reload();"
    'response.write "    self.close();"
    response.write "</script>"
	dbget.close()	:	response.End

elseif mode="shopmonthly10" then


    'sqlStr = "exec db_summary.[dbo].[sp_Ten_monthly_Acc_ShopstockMake_20150506_1] '"&yyyymm&"'"
    'dbget.execute sqlStr

    'dbget.close()	:	response.End

    sqlStr = "exec db_summary.dbo.sp_Ten_monthly_Acc_ShopstockMake '"&yyyymm&"'"
    dbget.execute sqlStr, resultrows
'rw sqlStr
    response.write "<script>alert('작성 되었습니다. ');</script>"
	response.write "<script>opener.location.reload();window.close();</script>"
	dbget.close()	:	response.End

elseif mode="shopmonthly101" then
    sqlStr = "exec [db_summary].[dbo].[sp_Ten_monthly_Acc_ShopstockMake_Update_monthlyTable] '"&yyyymm&"'"
    dbget.execute sqlStr, resultrows
	'rw sqlStr
    response.write "<script>alert('작성 되었습니다. ');</script>"
	response.write "<script>opener.location.reload();window.close();</script>"
	dbget.close()	:	response.End

elseif mode="shopmonthly102" then
    sqlStr = "exec [db_summary].[dbo].[sp_Ten_monthly_Acc_ShopstockMake_Update_accMonthlyTable] '"&yyyymm&"', '1'"
    dbget.execute sqlStr, resultrows
	'rw sqlStr
    response.write "<script>alert('작성 되었습니다. ');</script>"
	response.write "<script>opener.location.reload();window.close();</script>"
	dbget.close()	:	response.End

elseif mode="shopmonthly103" then
    sqlStr = "exec [db_summary].[dbo].[sp_Ten_monthly_Acc_ShopstockMake_Update_accMonthlyTable] '"&yyyymm&"', '2'"
    dbget.execute sqlStr, resultrows
	'rw sqlStr
    response.write "<script>alert('작성 되었습니다. ');</script>"
	response.write "<script>opener.location.reload();window.close();</script>"
	dbget.close()	:	response.End

elseif mode="shopmonthly104" then
    sqlStr = "exec [db_summary].[dbo].[sp_Ten_monthly_Acc_ShopstockMake_Update_accMonthlyTable] '"&yyyymm&"', '3'"
    dbget.execute sqlStr, resultrows
	'rw sqlStr
    response.write "<script>alert('작성 되었습니다. ');</script>"
	response.write "<script>opener.location.reload();window.close();</script>"
	dbget.close()	:	response.End

elseif mode="shopmonthly105" then
    sqlStr = "exec [db_summary].[dbo].[sp_Ten_monthly_Acc_ShopstockMake_Update_accMonthlyTable] '"&yyyymm&"', '4'"
    dbget.execute sqlStr, resultrows
	'rw sqlStr
    response.write "<script>alert('작성 되었습니다. ');</script>"
	response.write "<script>opener.location.reload();window.close();</script>"
	dbget.close()	:	response.End

elseif mode="shopmonthly11" then

    sqlStr = "exec db_summary.dbo.sp_Ten_monthlyShopstock_mwFlagUpdate '"&yyyymm&"'"
    dbget.execute sqlStr

    response.write "<script>alert('작성 되었습니다. ');</script>"
	response.write "<script>opener.location.reload();window.close();</script>"
	dbget.close()	:	response.End

elseif mode="shopmonthly20" then

	'// 입고내역 서머리
    sqlStr = " db_summary.dbo.sp_Ten_monthlyLogisstock_ipgoSumMake '"&yyyymm&"', 'S' "
    dbget.execute sqlStr,resultrows

    response.write "<script>alert('작성 되었습니다. ');</script>"
	response.write "<script>opener.location.reload();window.close();</script>"
	dbget.close()	:	response.End

elseif mode="shopmonthly21" then

	'// 평균매입가(매장)
    sqlStr = " db_summary.dbo.sp_Ten_monthlyLogisstock_avgipgoPrice '"&yyyymm&"', 'S' "
    dbget.execute sqlStr,resultrows

    response.write "<script>alert('작성 되었습니다. ');</script>"
	response.write "<script>opener.location.reload();window.close();</script>"
	dbget.close()	:	response.End

elseif mode="shopmonthly30" then

    ''수익율분석>>오프라인수익서머리 의 기말재고작성
    sqlStr = " exec [db_summary].[dbo].[usp_Ten_Shop_Create_ShopStockSUM] '"&yyyymm&"' "
    dbget.Execute sqlStr

    response.write "<script>alert('작성 되었습니다. ');</script>"
	response.write "<script>opener.location.reload();window.close();</script>"
	dbget.close()	:	response.End

' "[통계]재고자산>>재고월령 재고월령재작성 버튼
elseif mode="stockovervalue" then
	' 재고월령 마지막 입고일 생성. 물류
    sqlStr = "exec db_summary.dbo.usp_Ten_monthly_Acc_SetLastIpgoDate_Logis '"&yyyymm&"'"
    dbget.execute sqlStr, resultrows

    response.write "<br>재고월령 마지막 입고일 생성. 물류"

	' 재고월령 마지막 입고일 생성. 매장
    sqlStr = "exec db_summary.dbo.usp_Ten_monthly_Acc_SetLastIpgoDate_Shop '"&yyyymm&"'"
    dbget.execute sqlStr

    response.write "<br>재고월령 마지막 입고일 생성. 매장"

    response.write "<script type='text/javascript'>"
    response.write "    alert('작성 되었습니다.[" + CStr(yyyymm) + "]');"
	response.write "    opener.location.reload();"
    'response.write "    self.close();"
    response.write "</script>"
	dbget.close()	:	response.End

' "[통계]재고자산>>재고자산(월별)" 작성구분(매입재고) , 재고위치(전체) 생성 버튼
elseif mode="meaipsummake" then
    '매입재고/정산
    if (request("atype")="S") then

''        if (request("ptype")="A") then
''            sqlStr = "exec db_summary.[dbo].[sp_Ten_monthlyLogisstock_ipgoSumMake] '"&yyyymm&"','L'"
''            dbget.execute sqlStr
''
''            sqlStr = "exec db_summary.[dbo].[sp_Ten_monthlyLogisstock_ipgoSumMake] '"&yyyymm&"','S'"
''            dbget.execute sqlStr
''        else
''            sqlStr = "exec db_summary.[dbo].[sp_Ten_monthlyLogisstock_ipgoSumMake] '"&yyyymm&"','"&request("ptype")&"'"
''            dbget.execute sqlStr
''        end if

        '// 매입재고 수불대장. 매입상품 입고/출고내역 서머리
		sqlStr = "exec db_summary.[dbo].[sp_Ten_monthly_Maeip_Stockledger_Make] '"&yyyymm&"','"&CHKIIF(request("ptype")="A","",request("ptype"))&"'"
		''response.write sqlStr : dbget.close : response.end
        dbget.execute sqlStr

        response.write "<br>매입재고 수불대장. 매입상품 입고/출고내역 서머리"

		'// 재고자산 작성시 기타출고 평균매입가 업데이트
		sqlStr = " exec [db_summary].[dbo].[sp_Ten_monthly_EtcChulgoList_Apply_avgBuyPrice] '" & yyyymm & "' "
		dbget.execute sqlStr

        response.write "<br>재고자산 작성시 기타출고 평균매입가 업데이트"

        response.write "<script type='text/javascript'>"
        response.write "    alert('작성 되었습니다.[" + CStr(yyyymm) + "].');"
    	response.write "    opener.location.reload();"
        'response.write "    self.close();"
        response.write "</script>"
    	dbget.close()	:	response.End

    elseif (request("atype")="J") then
        sqlStr = "exec db_summary.[dbo].[sp_Ten_monthly_JungsanSum_Make] '"&yyyymm&"','"&CHKIIF(request("ptype")="A","",request("ptype"))&"'"
	'rw sqlStr
        dbget.execute sqlStr

        response.write "<script type='text/javascript'>"
        response.write "    alert('작성 되었습니다.[" + CStr(yyyymm) + "]..');"
    	response.write "    opener.location.reload();"
        'response.write "    self.close();"
        response.write "</script>"
    	dbget.close()	:	response.End
    else
        response.write "ERR:"&request("atype")
    end if
elseif mode="meaipsumcopy" then
	sqlStr = "exec db_summary.[dbo].[sp_Ten_monthly_Maeip_Stockledger_COPY] '" + CStr(yyyymm) + "'"
	dbget.execute sqlStr

	response.write "<script>alert('복사 되었습니다.[" + CStr(yyyymm) + "]');</script>"
	response.write "<script>opener.location.reload();window.close();</script>"
	dbget.close()	:	response.End
elseif mode="meaipsumdel" then
	sqlStr = "exec db_summary.[dbo].[sp_Ten_monthly_Maeip_Stockledger_DEL] '" + CStr(yyyymm) + "'"
	dbget.execute sqlStr

	response.write "<script>alert('삭제 되었습니다.[" + CStr(yyyymm) + "]');</script>"
	response.write "<script>opener.location.reload();window.close();</script>"
	dbget.close()	:	response.End
else
	response.write "mode=" + mode
	dbget.close()	:	response.End
end if
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
