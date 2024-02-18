<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  재고 처리
' History : 2009.04.07 서동석 생성
'			2017.10.18 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim mode,itemgubun,itemid,itemoption, makerid, realstock, errrealcheckno, errsampleitemno, shopid
dim refreshstartdate, Isexists_Standingorder
dim yyyymmdd, mayDay, errcsno, preyyyymmdd, BasicMonth, ThisDate
dim yyyymm
	mode	    = requestCheckvar(request.form("mode"),32)
	itemgubun   = requestCheckvar(request.form("itemgubun"),2)
	itemid      = requestCheckvar(request.form("itemid"),9)
	itemoption  = requestCheckvar(request.form("itemoption"),4)
	makerid     = requestCheckvar(request.form("makerid"),32)
	realstock   = requestCheckvar(request.form("realstock"),9)
	shopid      = requestCheckvar(request.form("shopid"),32)
	yyyymmdd    = requestCheckvar(request.form("yyyymmdd"),10)
	mayDay      = requestCheckvar(request.form("mayDay"),10)
	errcsno     = requestCheckvar(request.form("errcsno"),10)
	errrealcheckno = requestCheckvar(request.form("errrealcheckno"),10)
	errsampleitemno = requestCheckvar(request.form("errsampleitemno"),10)
	preyyyymmdd = requestCheckvar(request.form("preyyyymmdd"),10)
    yyyymm      = requestCheckvar(request.form("yyyymm"),7)

Isexists_Standingorder = false

dim refer
	refer = request.ServerVariables("HTTP_REFERER")

BasicMonth  = Left(CStr(DateSerial(Year(now()),Month(now())-1,1)),7)
ThisDate    = Left(CStr(now()),10)

dim sqlStr, AssignedRows
AssignedRows =0
if (mode="itemAccStock") then ''상품별 물류 기말재고 작성
    ''최근 내역업데이트
    sqlStr = "exec db_summary.[dbo].[usp_STOCK_ITEM_daily_logisstock_maker] '" & itemgubun & "'," & itemid & "," & CHKIIF(itemoption="" or itemoption="0000","NULL","'"&itemoption&"'") & ""
    dbget.Execute sqlStr

    sqlStr = "exec db_summary.[dbo].[usp_STOCK_ITEM_monthly_acc_logisstock_maker] '"&yyyymm&"','" & itemgubun & "'," & itemid & "," & CHKIIF(itemoption="" or itemoption="0000","NULL","'"&itemoption&"'") & ""
    dbget.Execute sqlStr

    sqlStr = "exec db_summary.[dbo].[sp_Ten_monthlyLogisstock_avgipgoPrice] '"&yyyymm&"', 'L','" & itemgubun & "'," & itemid & "," & CHKIIF(itemoption="" or itemoption="0000","NULL","'"&itemoption&"'") & ""
    dbget.Execute sqlStr

elseif (mode="itemAccStockShop") then ''상품별 매장 기말재고 작성
    sqlStr = "exec db_summary.[dbo].[usp_STOCK_ITEM_daily_shopstock_maker] "&CHKIIF(shopid="","NULL","'"&shopid&"'")&",'" & itemgubun & "'," & itemid & "," & CHKIIF(itemoption="" or itemoption="0000","NULL","'"&itemoption&"'") & ""
    dbget.Execute sqlStr

    ''매장아이디 불필요..
    sqlStr = "exec db_summary.[dbo].[usp_STOCK_ITEM_monthly_acc_shopstock_maker] '"&yyyymm&"','" & itemgubun & "'," & itemid & "," & CHKIIF(itemoption="" or itemoption="0000","NULL","'"&itemoption&"'") & ""
    dbget.Execute sqlStr
elseif (mode="itemrecentipchulrefreshv2") then
    ''최근 내역업데이트 v2
	if (IsTPLItemCode(itemgubun, itemid, itemoption) = True) then
		'' 3PL 온라인판매
        if (IsTPLIthinksoItemCode(itemgubun, itemid, itemoption) = True) then
            '// 띵소는 텐텐주문 3PL주문 합쳐야 함
            sqlStr = "exec db_summary.dbo.usp_TPL_recentOnlineSell_ITS_Update'" & itemgubun & "'," & itemid & ",'" & itemoption & "'"
			'response.write sqlStr & "<br>"
			dbget.Execute sqlStr
        else
            sqlStr = "exec db_summary.dbo.usp_TPL_recentOnlineSell_Update'" & itemgubun & "'," & itemid & ",'" & itemoption & "'"
			'response.write sqlStr & "<br>"
			dbget.Execute sqlStr
        end if
	else
		'' 온라인판매
        sqlStr = "exec db_summary.[dbo].[usp_STOCK_ITEM_daily_logisstock_maker] '" & itemgubun & "'," & itemid & "," & CHKIIF(itemoption="" or itemoption="0000" ,"NULL","'"&itemoption&"'") & ""
        dbget.Execute sqlStr
	end if

elseif (mode="itemRecentSellRefresh") then
    '' 상품별 최근[한달] 온라인 판매 내역 Update : 금일 출고내역, 최근 주문접수, 결제완료, 상품준비
    '' 반영 안되는 부분 : 출고후 Action 취소(업무상 불가)
    ''refreshstartdate = BasicMonth & "-01"
    ''itemrecentipchulrefresh

elseif (mode="itemrecentipchulrefresh") then
    '' 상품별 최근 입/출/판매 업데이트

	'/히치하이커 정기구독 인지 체크		'/2017.10.18 한용민 생성
	sqlStr = "select count(reserveidx) as cnt" & vbcrlf
	sqlStr = sqlStr & " from db_item.[dbo].[tbl_item_standing_order]" & vbcrlf
	sqlStr = sqlStr & " where reserveItemGubun = '" & itemgubun & "'" & vbcrlf
	sqlStr = sqlStr & " and reserveItemID = " & itemid & "" & vbcrlf
	sqlStr = sqlStr & " and reserveItemOption = '" & itemoption & "'" & vbcrlf

	'response.write sqlStr & "<br>"
	rsget.Open sqlStr,dbget,1
	if not rsget.EOF  then
		Isexists_Standingorder = rsget("cnt")>0
	end if
	rsget.close

    '' 입출고
    sqlStr = "exec db_summary.dbo.sp_ten_recentIpChul_Update '','','" & itemgubun & "'," & itemid & ",'" & itemoption & "',''"
	dbget.Execute sqlStr

	'/히치하이커 정기구독의 경우 기간을 6개월로 돌린다. 6개월 이전 6회차 기존 회차까지 전부 다 돌아가야함
	if Isexists_Standingorder then
		'' 온라인판매
	    sqlStr = "exec db_summary.dbo.sp_Ten_recentOnlineSell_Update_With_6MonthAgo '','" & itemgubun & "'," & itemid & ",'" & itemoption & "'"

		'response.write sqlStr & "<br>"
	    dbget.Execute sqlStr
	else
		if (IsTPLItemCode(itemgubun, itemid, itemoption) = True) then
			'' 3PL 온라인판매
            if (IsTPLIthinksoItemCode(itemgubun, itemid, itemoption) = True) then
                '// 띵소는 텐텐주문 3PL주문 합쳐야 함
                sqlStr = "exec db_summary.dbo.usp_TPL_recentOnlineSell_ITS_Update'" & itemgubun & "'," & itemid & ",'" & itemoption & "'"
			    'response.write sqlStr & "<br>"
			    dbget.Execute sqlStr
            else
                sqlStr = "exec db_summary.dbo.usp_TPL_recentOnlineSell_Update'" & itemgubun & "'," & itemid & ",'" & itemoption & "'"
			    'response.write sqlStr & "<br>"
			    dbget.Execute sqlStr
            end if
		else
			'' 온라인판매
			sqlStr = "exec db_summary.dbo.sp_ten_recentOnlineSell_Update '','" & itemgubun & "'," & itemid & ",'" & itemoption & "'"
			'response.write sqlStr & "<br>"
			dbget.Execute sqlStr

			'// 재고기준일수 별도 등록 상품만
			sqlStr = " exec [db_summary].[dbo].[usp_Ten_Refresh_MakeItem_RequireNO] '" + CStr(itemgubun) + "', " + CStr(itemid) + ", '" + CStr(itemoption) + "' "
			rsget.Open sqlStr,dbget,1

			'' 오차 업데이트. (last param 0 : 최근 1 : 전체) - 수정요
			''    sqlStr = "exec db_summary.dbo.sp_ten_recentCSipchul_Update '" & itemgubun & "'," & itemid & ",'" & itemoption & "'"
			''    dbget.Execute sqlStr
		end if
	end if

elseif (mode="ipchulallrefreshbyitemid") then
	''사은품(85) 출고내역 업데이트
	if (itemgubun = "85") then
		sqlStr = "exec db_summary.dbo.sp_Ten_On_Gift_Chulgo_Update_All '" & itemgubun & "'," & itemid & ",'" & itemoption & "'"
		dbget.Execute sqlStr
	end if

    ''입출내역 전체 새로고팀
    sqlStr = "exec db_summary.dbo.sp_ten_IpChul_Update_All '" & itemgubun & "'," & itemid & ",'" & itemoption & "'"
    dbget.Execute sqlStr
    ''rw sqlStr

    sqlStr = "exec db_summary.dbo.sp_Ten_recentCSipchul_Update '" & itemgubun & "'," & itemid & ",'" & itemoption & "'"

	'response.write sqlStr & "<br>"
    dbget.Execute sqlStr

'	'/히치하이커 정기구독 인지 체크		'/2017.10.18 한용민 생성
'	sqlStr = "select count(reserveidx) as cnt" & vbcrlf
'	sqlStr = sqlStr & " from db_item.[dbo].[tbl_item_standing_order]" & vbcrlf
'	sqlStr = sqlStr & " where reserveItemGubun = '" & itemgubun & "'" & vbcrlf
'	sqlStr = sqlStr & " and reserveItemID = " & itemid & "" & vbcrlf
'	sqlStr = sqlStr & " and reserveItemOption = '" & itemoption & "'" & vbcrlf
'
'	'response.write sqlStr & "<br>"
'	rsget.Open sqlStr,dbget,1
'	if not rsget.EOF  then
'		Isexists_Standingorder = rsget("cnt")>0
'	end if
'	rsget.close
'
'	'/히치하이커 정기구독의 경우 기간을 6개월로 돌린다.	6개월 이전 6회차 기존 회차까지 전부 다 돌아가야함
'	if Isexists_Standingorder then
'		sqlStr = "exec db_summary.dbo.sp_Ten_recentCSipchul_Update_With_6MonthAgo '" & itemgubun & "'," & itemid & ",'" & itemoption & "'"
'
'		'response.write sqlStr & "<br>"
'	    dbget.Execute sqlStr
'	else
'	    sqlStr = "exec db_summary.dbo.sp_Ten_recentCSipchul_Update '" & itemgubun & "'," & itemid & ",'" & itemoption & "'"
'
'		'response.write sqlStr & "<br>"
'	    dbget.Execute sqlStr
'	end if

elseif (mode="editdailyerrlog") then
    sqlStr = "update [db_summary].[dbo].tbl_erritem_daily_summary"& VbCRLF
    sqlStr = sqlStr + " set errbaditemno="&request("errbaditemno")& VbCRLF
    sqlStr = sqlStr + " , errrealcheckno="&request("errrealcheckno")& VbCRLF
    sqlStr = sqlStr + " where yyyymmdd='"&yyyymmdd&"'" & VbCRLF
    sqlStr = sqlStr + " and itemgubun='"&itemgubun&"'" & VbCRLF
    sqlStr = sqlStr + " and itemid='"&itemid&"'" & VbCRLF
    sqlStr = sqlStr + " and itemoption='"&itemoption&"'" & VbCRLF

    dbget.Execute sqlStr

    ''if (CDate(BasicMonth+"-01")>CDate(yyyymmdd)) then
        ''입출고 업데이트 ALL
    	sqlStr = "exec db_summary.dbo.sp_ten_IpChul_Update_All '" & itemgubun & "'," & itemid & ",'" & itemoption & "'"
        dbget.Execute sqlStr
	''end if

	 '' 최근 입출고
    sqlStr = "exec db_summary.dbo.sp_ten_recentIpChul_Update '','','" & itemgubun & "'," & itemid & ",'" & itemoption & "',''"
	dbget.Execute sqlStr
elseif (mode="editCsErr") then
    IF (LEN(mayDay)=7) then
        sqlStr = "update [db_summary].[dbo].tbl_monthly_logisstock_summary"& VbCRLF
        sqlStr = sqlStr + " set errcsno="&errcsno& VbCRLF
        sqlStr = sqlStr + " where yyyymm='"&mayDay&"'" & VbCRLF
        sqlStr = sqlStr + " and itemgubun='"&itemgubun&"'" & VbCRLF
        sqlStr = sqlStr + " and itemid='"&itemid&"'" & VbCRLF
        sqlStr = sqlStr + " and itemoption='"&itemoption&"'" & VbCRLF

        dbget.Execute sqlStr,AssignedRows

        if (AssignedRows>0) then
            sqlStr = "update [db_summary].[dbo].tbl_monthly_logisstock_summary"& VbCRLF
            sqlStr = sqlStr + " set totsysstock=totipgono+totchulgono-totsellno+errcsno"&VbCRLF
            sqlStr = sqlStr + " ,availsysstock=totipgono+totchulgono-totsellno+errcsno+errbaditemno+erretcno"&VbCRLF
            sqlStr = sqlStr + " ,realstock=totipgono+totchulgono-totsellno+errcsno+errbaditemno+erretcno+errrealcheckno"&VbCRLF
            sqlStr = sqlStr + " where yyyymm='"&mayDay&"'" & VbCRLF
            sqlStr = sqlStr + " and itemgubun='"&itemgubun&"'" & VbCRLF
            sqlStr = sqlStr + " and itemid='"&itemid&"'" & VbCRLF
            sqlStr = sqlStr + " and itemoption='"&itemoption&"'" & VbCRLF

            dbget.Execute sqlStr
        else
            sqlStr = " Insert into [db_summary].[dbo].tbl_monthly_logisstock_summary"& VbCRLF
            sqlStr = sqlStr + " (yyyymm,itemgubun,itemid,itemoption,errcsno,totsysstock,availsysstock,realstock)"& VbCRLF
            sqlStr = sqlStr + " values('"&mayDay&"'"&VbCRLF
            sqlStr = sqlStr + " ,'"&itemgubun&"'"&VbCRLF
            sqlStr = sqlStr + " ,'"&itemid&"'"&VbCRLF
            sqlStr = sqlStr + " ,'"&itemoption&"'"&VbCRLF
            sqlStr = sqlStr + " ,"&errcsno&VbCRLF
            sqlStr = sqlStr + " ,"&errcsno&VbCRLF
            sqlStr = sqlStr + " ,"&errcsno&VbCRLF
            sqlStr = sqlStr + " )"&VbCRLF

            dbget.Execute sqlStr
        end if
    ELSEIF (LEN(mayDay)=10) then
        sqlStr = "update [db_summary].[dbo].tbl_daily_logisstock_summary"& VbCRLF
        sqlStr = sqlStr + " set errcsno="&errcsno& VbCRLF
        sqlStr = sqlStr + " where yyyymmdd='"&mayDay&"'" & VbCRLF
        sqlStr = sqlStr + " and itemgubun='"&itemgubun&"'" & VbCRLF
        sqlStr = sqlStr + " and itemid='"&itemid&"'" & VbCRLF
        sqlStr = sqlStr + " and itemoption='"&itemoption&"'" & VbCRLF

        dbget.Execute sqlStr,AssignedRows

        if (AssignedRows>0) then
            sqlStr = "update [db_summary].[dbo].tbl_daily_logisstock_summary"& VbCRLF
            sqlStr = sqlStr + " set totsysstock=totipgono+totchulgono-totsellno+errcsno"&VbCRLF
            sqlStr = sqlStr + " ,availsysstock=totipgono+totchulgono-totsellno+errcsno+errbaditemno+erretcno"&VbCRLF
            sqlStr = sqlStr + " ,realstock=totipgono+totchulgono-totsellno+errcsno+errbaditemno+erretcno+errrealcheckno"&VbCRLF
            sqlStr = sqlStr + " where yyyymmdd='"&mayDay&"'" & VbCRLF
            sqlStr = sqlStr + " and itemgubun='"&itemgubun&"'" & VbCRLF
            sqlStr = sqlStr + " and itemid='"&itemid&"'" & VbCRLF
            sqlStr = sqlStr + " and itemoption='"&itemoption&"'" & VbCRLF

            dbget.Execute sqlStr
        ELSE
            sqlStr = " Insert into [db_summary].[dbo].tbl_daily_logisstock_summary"& VbCRLF
            sqlStr = sqlStr + " (yyyymmdd,itemgubun,itemid,itemoption,errcsno,totsysstock,availsysstock,realstock)"& VbCRLF
            sqlStr = sqlStr + " values('"&mayDay&"'"&VbCRLF
            sqlStr = sqlStr + " ,'"&itemgubun&"'"&VbCRLF
            sqlStr = sqlStr + " ,'"&itemid&"'"&VbCRLF
            sqlStr = sqlStr + " ,'"&itemoption&"'"&VbCRLF
            sqlStr = sqlStr + " ,"&errcsno&VbCRLF
			sqlStr = sqlStr + " ,"&errcsno&VbCRLF
            sqlStr = sqlStr + " ,"&errcsno&VbCRLF
            sqlStr = sqlStr + " ,"&errcsno&VbCRLF
            sqlStr = sqlStr + " )"&VbCRLF

            dbget.Execute sqlStr
        end if
    ENd IF

    sqlStr = "exec db_summary.dbo.sp_ten_IpChul_Update_All '" & itemgubun & "'," & itemid & ",'" & itemoption & "'"
    dbget.Execute sqlStr

elseif (mode="dummidailyerrlog") then
    sqlStr = " IF not Exists(select * from [db_summary].[dbo].tbl_erritem_daily_summary " & VbCRLF
    sqlStr = sqlStr + " where yyyymmdd='"&yyyymmdd&"'" & VbCRLF
    sqlStr = sqlStr + " and itemgubun='"&itemgubun&"'" & VbCRLF
    sqlStr = sqlStr + " and itemid='"&itemid&"'" & VbCRLF
    sqlStr = sqlStr + " and itemoption='"&itemoption&"'" & VbCRLF
    sqlStr = sqlStr + " )" & VbCRLF
    sqlStr = sqlStr + " BEGIN"
    sqlStr = sqlStr + " insert into [db_summary].[dbo].tbl_erritem_daily_summary"& VbCRLF
    sqlStr = sqlStr + " (yyyymmdd,itemgubun,itemid,itemoption)"
    sqlStr = sqlStr + " values ('"&yyyymmdd&"','"&itemgubun&"','"&itemid&"','"&itemoption&"')" & VbCRLF
    sqlStr = sqlStr + " END"

    dbget.Execute sqlStr

    sqlStr = "exec db_summary.dbo.sp_ten_IpChul_Update_All '" & itemgubun & "'," & itemid & ",'" & itemoption & "'"
    dbget.Execute sqlStr

elseif (mode="errcheckupdate") then

    ''sqlStr = "exec [db_summary].[dbo].[sp_ten_recentOnlineSell_Update] @makerid,@itemgubun,@itemid,@itemoption"
    '' 최근 재고 업데이트
    ''sqlStr = "exec [db_summary].[dbo].[sp_ten_recentOnlineSell_Update] '','" & itemgubun & "'," & itemid & ",'" & itemoption & "'"
    ''dbget.Execute sqlStr


    '' 금일 실사 값으로 오차 입력
    ''sqlStr = "exec [db_summary].[dbo].[sp_ten_realchekErr_Input_By_CurrentStock] @itemgubun,@itemid,@itemoption,@realCheckStock,@reguser"
    sqlStr = "exec [db_summary].[dbo].[sp_ten_realchekErr_Input_By_CurrentStock] '" & itemgubun & "'," & itemid & ",'" & itemoption & "'," & realstock & ",'" & session("ssBctID") & "'"
    dbget.Execute sqlStr

    ''한정수량 조절
    sqlStr = " exec [db_summary].[dbo].[sp_ten_limitSetByRealStock] '" & itemgubun & "'," & itemid & ",'" & itemoption & "'"
    dbget.Execute sqlStr, AssignedRows

    ''한정 일시품절->판매조절
    if (itemgubun="10") and (AssignedRows>0) then
        sqlStr = " exec [db_summary].[dbo].sp_Ten_SellYnSetByLimitNo " & itemid
        dbget.Execute sqlStr
    end if

''    오차 값으로 입력 할 경우..
''    sqlStr = "select (realstock + ipkumdiv5 + offconfirmno) as realCheckStock from [db_summary].[dbo].tbl_current_logisstock_summary"
''    sqlStr = sqlStr + " where itemid=" + itemid
''    sqlStr = sqlStr + " and itemoption='" + itemoption + "'"
''    sqlStr = sqlStr + " and itemgubun='" + itemgubun + "'"
''
''    dim diffNo
''    rsget.Open sqlStr,dbget,1
''        diffNo = realstock - rsget("realCheckStock")
''    rsget.Close
''
''    sqlStr = "exec [db_summary].[dbo].[sp_ten_realchekErr_Input] '" & Left(now(),10) & "','" & itemgubun & "'," & itemid & ",'" & itemoption & "'," & diffNo & ",'" & session("ssBctID") & "'"
''    dbget.Execute sqlStr

elseif (mode="OFFitemAllRefresh") then
    ''-1 월말 업데이트
    sqlStr = "exec [db_summary].[dbo].[sp_Ten_Shop_Stock_UpdateALL] '" & shopid & "','" & itemgubun & "'," & itemid & ",'" & itemoption & "'"
    dbget.Execute sqlStr

    ''-1 일별 업데이트
    sqlStr = "exec [db_summary].[dbo].sp_Ten_Shop_Stock_RecentUpdateByItem '" & shopid & "','" & itemgubun & "'," & itemid & ",'" & itemoption & "'"
    dbget.Execute sqlStr
elseif (mode="Offerrcheckupdate") then
    ''오프 실사 재고 저장.
    sqlStr = "exec [db_summary].[dbo].sp_Ten_Shop_realchekErr_Input_By_CurrentStock '" & shopid & "','" & itemgubun & "'," & itemid & ",'" & itemoption & "'," & realstock & ",'" & session("ssBctID") & "'"
    dbget.Execute sqlStr
elseif (mode="OffErrDelete") then

    sqlStr = "delete from [db_summary].[dbo].tbl_erritem_shop_summary" + VbCrlf
    sqlStr = sqlStr + " where yyyymmdd='" + yyyymmdd + "'" + VbCrlf
    sqlStr = sqlStr + " and itemgubun='" + itemgubun + "'" + VbCrlf
    sqlStr = sqlStr + " and shopitemid=" + CStr(itemid) + "" + VbCrlf
    sqlStr = sqlStr + " and itemoption='" + itemoption + "'" + VbCrlf
    sqlStr = sqlStr + " and shopid='" + shopid + "'"

    dbget.Execute sqlStr

    if (CDate(BasicMonth+"-01")>CDate(yyyymmdd)) then
        ''-1 월말 업데이트
        sqlStr = "exec [db_summary].[dbo].[sp_Ten_Shop_Stock_UpdateALL] '" & shopid & "','" & itemgubun & "'," & itemid & ",'" & itemoption & "'"
        dbget.Execute sqlStr

        response.write "."
    end if
    ''-1 일별 업데이트
    sqlStr = "exec [db_summary].[dbo].sp_Ten_Shop_Stock_RecentUpdateByItem '" & shopid & "','" & itemgubun & "'," & itemid & ",'" & itemoption & "'"
    dbget.Execute sqlStr
elseif (mode="OffErrEdit") then
    sqlStr = "update [db_summary].[dbo].tbl_erritem_shop_summary" + VbCrlf
    sqlStr = sqlStr + " set errrealcheckno=" + errrealcheckno + ", errsampleitemno=" + errsampleitemno + ", modiuser = '" & session("ssBctID") & "', lastupdate = getdate()" + VbCrlf
    sqlStr = sqlStr + " where yyyymmdd='" + yyyymmdd + "'" + VbCrlf
    sqlStr = sqlStr + " and itemgubun='" + itemgubun + "'" + VbCrlf
    sqlStr = sqlStr + " and shopitemid=" + CStr(itemid) + "" + VbCrlf
    sqlStr = sqlStr + " and itemoption='" + itemoption + "'" + VbCrlf
    sqlStr = sqlStr + " and shopid='" + shopid + "'"
'response.write sqlStr
'dbget.close()	:	response.End
    dbget.Execute sqlStr

    if (CDate(BasicMonth+"-01")>CDate(yyyymmdd)) then
        ''-1 월말 업데이트
        sqlStr = "exec [db_summary].[dbo].[sp_Ten_Shop_Stock_UpdateALL] '" & shopid & "','" & itemgubun & "'," & itemid & ",'" & itemoption & "'"
        dbget.Execute sqlStr

        response.write "."
    end if
    ''-1 일별 업데이트
    sqlStr = "exec [db_summary].[dbo].sp_Ten_Shop_Stock_RecentUpdateByItem '" & shopid & "','" & itemgubun & "'," & itemid & ",'" & itemoption & "'"
    dbget.Execute sqlStr

elseif (mode="dummidailyerrlogOFF") then
    sqlStr = " IF not Exists(select * from [db_summary].[dbo].tbl_erritem_shop_summary " & VbCRLF
    sqlStr = sqlStr + " where yyyymmdd='"&yyyymmdd&"'" & VbCRLF
    sqlStr = sqlStr + " and itemgubun='"&itemgubun&"'" & VbCRLF
    sqlStr = sqlStr + " and shopitemid='"&itemid&"'" & VbCRLF
    sqlStr = sqlStr + " and itemoption='"&itemoption&"'" & VbCRLF
	sqlStr = sqlStr + " and shopid='"&shopid&"'" & VbCRLF
    sqlStr = sqlStr + " )" & VbCRLF
    sqlStr = sqlStr + " BEGIN"
    sqlStr = sqlStr + " insert into [db_summary].[dbo].tbl_erritem_shop_summary"& VbCRLF
    sqlStr = sqlStr + " (yyyymmdd,shopid,itemgubun,shopitemid,itemoption,reguser)"
    sqlStr = sqlStr + " values ('"&yyyymmdd&"','" & shopid & "','"&itemgubun&"','"&itemid&"','"&itemoption&"','" & session("ssBctID") & "')" & VbCRLF
    sqlStr = sqlStr + " END"
    dbget.Execute sqlStr
elseif (mode="dummidailyerrlogCHGOFF") then
    sqlStr = " IF not Exists(select * from [db_summary].[dbo].tbl_erritem_shop_summary " & VbCRLF
    sqlStr = sqlStr + " where yyyymmdd='"&yyyymmdd&"'" & VbCRLF
    sqlStr = sqlStr + " and itemgubun='"&itemgubun&"'" & VbCRLF
    sqlStr = sqlStr + " and shopitemid='"&itemid&"'" & VbCRLF
    sqlStr = sqlStr + " and itemoption='"&itemoption&"'" & VbCRLF
	sqlStr = sqlStr + " and shopid='"&shopid&"'" & VbCRLF
    sqlStr = sqlStr + " )" & VbCRLF
    sqlStr = sqlStr + " BEGIN"
    sqlStr = sqlStr + " insert into [db_summary].[dbo].tbl_erritem_shop_summary"& VbCRLF
    sqlStr = sqlStr + " (yyyymmdd,shopid,itemgubun,shopitemid,itemoption,reguser, errrealcheckno)"
    sqlStr = sqlStr + " values ('"&yyyymmdd&"','" & shopid & "','"&itemgubun&"','"&itemid&"','"&itemoption&"','" & session("ssBctID") & "',"&errrealcheckno&")" & VbCRLF
    sqlStr = sqlStr + " END"
    dbget.Execute sqlStr, AssignedRows

    if (AssignedRows<1) then
        sqlStr = " IF Exists(select * from [db_summary].[dbo].tbl_erritem_shop_summary " & VbCRLF
        sqlStr = sqlStr + " where yyyymmdd='"&yyyymmdd&"'" & VbCRLF
        sqlStr = sqlStr + " and itemgubun='"&itemgubun&"'" & VbCRLF
        sqlStr = sqlStr + " and shopitemid='"&itemid&"'" & VbCRLF
        sqlStr = sqlStr + " and itemoption='"&itemoption&"'" & VbCRLF
    	sqlStr = sqlStr + " and shopid='"&shopid&"'" & VbCRLF
    	sqlStr = sqlStr + " and errrealcheckno=0" & VbCRLF
        sqlStr = sqlStr + " )" & VbCRLF
        sqlStr = sqlStr + " BEGIN"
        sqlStr = sqlStr + " update [db_summary].[dbo].tbl_erritem_shop_summary" + VbCrlf
        sqlStr = sqlStr + " set errrealcheckno="&errrealcheckno&VbCrlf
        sqlStr = sqlStr + " , modiuser = '" & session("ssBctID") & "', lastupdate = getdate()" + VbCrlf
        sqlStr = sqlStr + " where yyyymmdd='" + yyyymmdd + "'" + VbCrlf
        sqlStr = sqlStr + " and itemgubun='" + itemgubun + "'" + VbCrlf
        sqlStr = sqlStr + " and shopitemid=" + CStr(itemid) + "" + VbCrlf
        sqlStr = sqlStr + " and itemoption='" + itemoption + "'" + VbCrlf
        sqlStr = sqlStr + " and shopid='" + shopid + "'"
        sqlStr = sqlStr + " END"
    '' rw sqlStr
        dbget.Execute sqlStr, AssignedRows
    end if

    if (AssignedRows<1) then
        response.write yyyymmdd&" 날짜에 이미 에러가 존재합니다."
        dbget.close() : response.end
    end if

    sqlStr = "update [db_summary].[dbo].tbl_erritem_shop_summary" + VbCrlf
    sqlStr = sqlStr + " set errrealcheckno=errrealcheckno+"&errrealcheckno&"*-1"&VbCrlf
    sqlStr = sqlStr + " , modiuser = '" & session("ssBctID") & "', lastupdate = getdate()" + VbCrlf
    sqlStr = sqlStr + " where yyyymmdd='" + preyyyymmdd + "'" + VbCrlf
    sqlStr = sqlStr + " and itemgubun='" + itemgubun + "'" + VbCrlf
    sqlStr = sqlStr + " and shopitemid=" + CStr(itemid) + "" + VbCrlf
    sqlStr = sqlStr + " and itemoption='" + itemoption + "'" + VbCrlf
    sqlStr = sqlStr + " and shopid='" + shopid + "'"
    dbget.Execute sqlStr


    ''-1 월말 업데이트
    sqlStr = "exec [db_summary].[dbo].[sp_Ten_Shop_Stock_UpdateALL] '" & shopid & "','" & itemgubun & "'," & itemid & ",'" & itemoption & "'"
    dbget.Execute sqlStr

    ''-1 일별 업데이트
    sqlStr = "exec [db_summary].[dbo].sp_Ten_Shop_Stock_RecentUpdateByItem '" & shopid & "','" & itemgubun & "'," & itemid & ",'" & itemoption & "'"
    dbget.Execute sqlStr
else
    response.write "<script>alert('정의 되지 않았습니다. - " & mode & "');</script>"
end if

%>
<script language="javascript">
alert('저장 되었습니다.');
location.replace('<%= refer %>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
