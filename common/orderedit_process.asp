<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 고객센터
' History : 이상구 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp"-->
<%
dim orderserial, detailidx, mode
dim buycash, isupchebeasong, songjangdiv, songjangno
dim beasongdate, currstate, upcheconfirmdate
dim requiredetail, itemno, omwdiv, odlvType, applyallitem
dim vatinclude
dim makerid, userid, ipkumdate
dim presongjangno, presongjangdiv, reducedPrice, itemcost
dim prevBonusCouponDiscountPrice, currBonusCouponDiscountPrice, tencardspend, itemcostCouponNotApplied
dim orgorderserial, orgdetailidx
dim detailidxArr, songjangdivArr, songjangnoArr
dim i


orderserial     = request("orderserial")
detailidx       = request("detailidx")
mode            = request("mode")
buycash         = request("buycash")
reducedPrice    = request("reducedPrice")
itemcost    	= request("itemcost")
itemcostCouponNotApplied    	= request("itemcostCouponNotApplied")
isupchebeasong  = request("isupchebeasong")
songjangdiv     = request("songjangdiv")
songjangno      = request("songjangno")

currstate       = request("currstate")
upcheconfirmdate = request("upcheconfirmdate")
beasongdate     = request("beasongdate")
requiredetail   = html2db(request("requiredetail"))
itemno          = request("itemno")
omwdiv          = request("omwdiv")
odlvType        = request("odlvType")
applyallitem    = request("applyallitem")
vatinclude    	= request("vatinclude")
makerid    		= request("makerid")
ipkumdate    	= request("ipkumdate")

presongjangno	= requestCheckvar(request("presongjangno"),32)
presongjangdiv	= requestCheckvar(request("presongjangdiv"),10)

dim tmp
On Error resume Next
if (upcheconfirmdate<>"") then tmp = CDate(upcheconfirmdate)
if Err then
    response.write "<script language='javascript'>alert('업체 확인일이 올바르지 않습니다.');history.back();</script>"
    dbget.close()	:	response.End
end if

if (beasongdate<>"") then tmp = CDate(beasongdate)
if Err then
    response.write "<script language='javascript'>alert('업체 배송일이  올바르지 않습니다.');history.back();</script>"
    dbget.close()	:	response.End
end if

On Error Goto 0


function DoRecalcOrderMaster(orderserial)
    dim sqlStr,nrowCount

    sqlStr = " update m "
    sqlStr = sqlStr & " set "
    sqlStr = sqlStr & "     m.totalsum = T.dtotalsum "
    sqlStr = sqlStr & "     , m.tencardspend = T.dtotalTencardspend "
    sqlStr = sqlStr & "     , m.subtotalprice = T.dtotalsubtotalprice - IsNull(m.miletotalprice, 0) - T.dtotalAllatdiscount "
    sqlStr = sqlStr & "     , m.totalmileage = T.dtotalmileage "
    sqlStr = sqlStr & "     , m.subtotalpriceCouponNotApplied = T.dtotalSubtotalPriceCouponNotApplied "
    sqlStr = sqlStr & "     , m.allatdiscountprice = T.dtotalAllatdiscount "
    sqlStr = sqlStr & " from "
    sqlStr = sqlStr & "     [db_order].[dbo].tbl_order_master m "
    sqlStr = sqlStr & "     join ( "
    sqlStr = sqlStr & "         select "
    sqlStr = sqlStr & "             d.orderserial "
    sqlStr = sqlStr & "             , sum(itemcost*itemno) as dtotalsum "
    sqlStr = sqlStr & "             , sum((itemcost - reducedPrice - IsNull(etcDiscount, 0))*itemno) as dtotalTencardspend "
    sqlStr = sqlStr & "             , sum(reducedPrice*itemno) as dtotalsubtotalprice "
    sqlStr = sqlStr & "             , sum(IsNull(etcDiscount, 0)*itemno) as dtotalAllatdiscount "
    sqlStr = sqlStr & "             , sum(mileage*itemno) as dtotalmileage "
    sqlStr = sqlStr & "             , sum(IsNull(itemcostCouponNotApplied,0)*itemno) as dtotalSubtotalPriceCouponNotApplied "
    sqlStr = sqlStr & "             , sum(IsNull(orgitemcost,0)*itemno) as dtotalorgitemcost "
    sqlStr = sqlStr & "         from "
    sqlStr = sqlStr & "             [db_order].[dbo].tbl_order_detail d "
    sqlStr = sqlStr & "         where "
    sqlStr = sqlStr & "             1 = 1 "
    sqlStr = sqlStr & "             and d.orderserial = '" & orderserial & "' "
    sqlStr = sqlStr & "             and d.cancelyn <> 'Y' "
    sqlStr = sqlStr & "         group by "
    sqlStr = sqlStr & "             d.orderserial "
    sqlStr = sqlStr & "     ) T on m.orderserial = T.orderserial "
    dbget.Execute sqlStr,nrowCount

    sqlStr = " update e set e.realPayedSum = (T.realpayedsum - T.realpayedsum120), e.acctamount = (case when T.ipkumdiv < '4' then (T.realpayedsum - T.realpayedsum120) else e.acctamount end) "
    sqlStr = sqlStr & " from [db_order].[dbo].tbl_order_PaymentEtc e "
    sqlStr = sqlStr & " join ( "
    sqlStr = sqlStr & "     select m.orderserial, m.accountdiv, (m.subtotalprice - m.sumpaymentetc) as realpayedsum, IsNull(sum(Case when e.acctdiv = '120' then e.realpayedsum else 0 end),0) as realpayedsum120, m.ipkumdiv  "
    sqlStr = sqlStr & "     from [db_order].[dbo].tbl_order_master m "
    sqlStr = sqlStr & "     join [db_order].[dbo].tbl_order_PaymentEtc e "
    sqlStr = sqlStr & "         on m.orderserial = e.orderserial "
    sqlStr = sqlStr & "         and e.acctdiv in (m.accountdiv, '120') "
    sqlStr = sqlStr & "     where m.orderserial = '" & orderserial & "' "
    sqlStr = sqlStr & "     group by m.orderserial, m.accountdiv, (m.subtotalprice - m.sumpaymentetc), m.ipkumdiv "
    sqlStr = sqlStr & " ) T "
    sqlStr = sqlStr & "     on e.orderserial = T.orderserial "
    sqlStr = sqlStr & "     and e.acctdiv = T.accountdiv "
    dbget.Execute sqlStr,nrowCount
end function


dim sqlStr, dataExists
dim nrowCount
if (mode="buycash") or (mode="reducedPrice") or (mode="itemcost") or (mode="currstate") or (mode="vatinclude") then
	''정산내역에 반영 되있는지 체크
	'' TODO : 매입상품 체크안됨 => 전월출고내역인 경우 수정불가
	sqlStr = "select count(*) as cnt from [db_jungsan].[dbo].tbl_designer_jungsan_detail"
	sqlStr = sqlStr + " where detailidx=" + CStr(detailidx)
	sqlStr = sqlStr + " and ("
	sqlStr = sqlStr + "     gubuncd='upche' or gubuncd='witaksell' or gubuncd='lecture'"
 	sqlStr = sqlStr + " )"

 	rsget.Open sqlStr,dbget,1
 		dataExists = rsget("cnt")>0
 	rsget.close

 	if (dataExists) and (orderserial <> "18031527847") then
 		response.write "<script>alert('정산 데이타가 존재합니다. 수정하실 수 없습니다.');</script>"
 		response.write "<script>history.back();</script>"
 		dbget.close()	:	response.End
 	end if

    dataExists = ""
	sqlStr = " select top 1 (case "
    if C_ADMIN_AUTH then
        '// 관리자 허용
        sqlStr = sqlStr + " 				when 1=0 then 'J' "
    else
 	    sqlStr = sqlStr + " 				when d.jungsanfixdate is not NULL then 'J' "
 	    sqlStr = sqlStr + " 				when d.currstate = '7' and DateDiff(month, d.beasongdate, getdate()) > 0 then 'P' "
    end if

 	sqlStr = sqlStr + " 				else '' end) errCode "
 	sqlStr = sqlStr + " from "
 	sqlStr = sqlStr + " 	[db_order].[dbo].[tbl_order_master] m "
 	sqlStr = sqlStr + " 	join [db_order].[dbo].[tbl_order_detail] d on m.orderserial = d.orderserial "
 	sqlStr = sqlStr + " where  "
 	sqlStr = sqlStr + " 	1 = 1 "
 	sqlStr = sqlStr + " 	and m.orderserial = '" & orderserial & "' "
 	sqlStr = sqlStr + " 	and (d.idx = " & detailidx & ") "
 	sqlStr = sqlStr + " 	and ( "
 	sqlStr = sqlStr + " 		(d.jungsanfixdate is not NULL) "
 	sqlStr = sqlStr + " 		or "
 	sqlStr = sqlStr + " 		(d.currstate = '7' and DateDiff(month, d.beasongdate, getdate()) > 0) "
 	sqlStr = sqlStr + " 	) "
 	rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
 		dataExists = rsget("errCode")
    end if
 	rsget.close

 	if (dataExists <> "") then
        if dataExists = "J" then
            response.write "정산 데이타가 존재합니다. 수정하실 수 없습니다."
        elseif dataExists = "P" then
            response.write "전월 출고내역입니다. 수정하실 수 없습니다."
        else
            response.write "알 수 없는 오류입니다."
        end if
 		dbget.close()	:	response.End
 	end if
end if

if (mode="buycash") then
    '매입가 변경
		sqlStr = "update [db_order].[dbo].tbl_order_detail" + VbCrlf
		sqlStr = sqlStr + " set buycash=" + CStr(buycash)  + VbCrlf
		sqlStr = sqlStr + " where idx=" + CStr(detailidx)  + VbCrlf

		dbget.Execute sqlStr,nrowCount

        sqlStr = " if not exists(select top 1 1 from [db_datamart].[dbo].[tbl_order_log_remakeQue] where orderserial = '" & Trim(orderserial) & "') "
        sqlStr = sqlStr & " begin "
        sqlStr = sqlStr & "     insert into [db_datamart].[dbo].[tbl_order_log_remakeQue](orderserial, chktype) values('" & Trim(orderserial) & "', 999) "
        sqlStr = sqlStr & "     insert into [tendb].db_temp.dbo.tbl_orderSerial_change(orderserial,lastupdate,gubun) values('" & Trim(orderserial) & "',getdate(),'MAKEQUE') "
        sqlStr = sqlStr & " end "
        db3_dbget.Execute sqlStr

		call AddCsMemo(orderserial,"1",userid,session("ssBctId"), "매입가 수기변경")

		response.write "<script>alert('" + CStr(nrowCount) + "개 수정 되었습니다.');</script>"
 		response.write "<script>location.replace('/common/orderdetailedit.asp?idx=" + detailidx + "');</script>"
 		dbget.close()	:	response.End
elseif (mode="reducedPrice") then
    '쿠폰가 변경
		sqlStr = " select top 1 d.itemcost, d.reducedPrice, d.itemno from [db_order].[dbo].[tbl_order_detail] d where d.idx = " & detailidx
 	    rsget.Open sqlStr,dbget,1
        itemcost = 0
        if Not rsget.Eof then
 		    itemcost = rsget("itemcost")
            prevBonusCouponDiscountPrice = (rsget("itemcost") - rsget("reducedPrice")) * rsget("itemno")
            currBonusCouponDiscountPrice = (rsget("itemcost") - reducedPrice) * rsget("itemno")
        end if
 	    rsget.close

        if (CLng(itemcost) < CLng(reducedPrice)) then
            response.write "판매가보다 쿠폰가가 더 클 수 없습니다."
            dbget.close()	:	response.End
        end if

        sqlStr = "update [db_order].[dbo].tbl_order_detail" + VbCrlf
		sqlStr = sqlStr + " set reducedPrice=" + CStr(reducedPrice)  + VbCrlf
        if (currBonusCouponDiscountPrice > 0) then
            sqlStr = sqlStr + " ,bonuscouponidx=-1"+ VbCrlf
        elseif (currBonusCouponDiscountPrice = 0) then
            sqlStr = sqlStr + " ,bonuscouponidx=NULL"+ VbCrlf
        end if
		sqlStr = sqlStr + " where idx=" + CStr(detailidx)  + VbCrlf

        dbget.Execute sqlStr,nrowCount

        sqlStr = "update [db_order].[dbo].tbl_order_master" + VbCrlf
		sqlStr = sqlStr + " set tencardspend=tencardspend + " + CStr(currBonusCouponDiscountPrice-prevBonusCouponDiscountPrice)  + VbCrlf
        sqlStr = sqlStr + " , subtotalprice=subtotalprice - " + CStr(currBonusCouponDiscountPrice-prevBonusCouponDiscountPrice)  + VbCrlf
		sqlStr = sqlStr + " where orderserial='" & orderserial & "'" + VbCrlf

        dbget.Execute sqlStr,nrowCount

        Call DoRecalcOrderMaster(orderserial)

		call AddCsMemo(orderserial,"1",userid,session("ssBctId"), "쿠폰가 수기변경")

		response.write "<script>alert('" + CStr(nrowCount) + "개 수정 되었습니다.');</script>"
 		response.write "<script>location.replace('/common/orderdetailedit.asp?idx=" + detailidx + "');</script>"
 		dbget.close()	:	response.End
elseif (mode="itemcostCouponNotApplied") then
    '판매가 변경

		sqlStr = " select top 1 d.itemcost, d.reducedPrice, d.itemno from [db_order].[dbo].[tbl_order_detail] d where d.idx = " & detailidx
 	    rsget.Open sqlStr,dbget,1
        reducedPrice = 0
        if Not rsget.Eof then
 		    reducedPrice = rsget("reducedPrice")
            prevBonusCouponDiscountPrice = (rsget("itemcost") - rsget("reducedPrice")) * rsget("itemno")
            currBonusCouponDiscountPrice = (itemcostCouponNotApplied - rsget("reducedPrice")) * rsget("itemno")
        end if
 	    rsget.close

        if (CLng(itemcostCouponNotApplied) < CLng(reducedPrice)) then
            response.write "판매가보다 쿠폰가가 더 클 수 없습니다."
            dbget.close()	:	response.End
        end if

        sqlStr = "update [db_order].[dbo].tbl_order_detail" + VbCrlf
		sqlStr = sqlStr + " set itemcost=" + CStr(itemcostCouponNotApplied) + ", itemcostCouponNotApplied = " + CStr(itemcostCouponNotApplied) + VbCrlf
        if (currBonusCouponDiscountPrice > 0) then
            sqlStr = sqlStr + " ,bonuscouponidx=-1"+ VbCrlf
        elseif (currBonusCouponDiscountPrice = 0) then
            sqlStr = sqlStr + " ,bonuscouponidx=NULL"+ VbCrlf
        end if
		sqlStr = sqlStr + " where idx=" + CStr(detailidx)  + VbCrlf

		dbget.Execute sqlStr,nrowCount

        sqlStr = "update [db_order].[dbo].tbl_order_master" + VbCrlf
		sqlStr = sqlStr + " set tencardspend=tencardspend + " + CStr(currBonusCouponDiscountPrice-prevBonusCouponDiscountPrice)  + VbCrlf
        sqlStr = sqlStr + " , subtotalpriceCouponNotApplied=subtotalpriceCouponNotApplied + " + CStr(currBonusCouponDiscountPrice-prevBonusCouponDiscountPrice)  + VbCrlf
        sqlStr = sqlStr + " , totalsum=totalsum + " + CStr(currBonusCouponDiscountPrice-prevBonusCouponDiscountPrice)  + VbCrlf
		sqlStr = sqlStr + " where orderserial='" & orderserial & "'" + VbCrlf

		dbget.Execute sqlStr,nrowCount

        Call DoRecalcOrderMaster(orderserial)

		call AddCsMemo(orderserial,"1",userid,session("ssBctId"), "판매가 수기변경")

		response.write "<script>alert('" + CStr(nrowCount) + "개 수정 되었습니다.');</script>"
 		response.write "<script>location.replace('/common/orderdetailedit.asp?idx=" + detailidx + "');</script>"
 		dbget.close()	:	response.End
elseif (mode="itemcost") then

        sqlStr = "update [db_order].[dbo].tbl_order_detail" + VbCrlf
		sqlStr = sqlStr + " set itemcost=" + CStr(itemcost) + ", reducedPrice = " + CStr(itemcost) + VbCrlf
        sqlStr = sqlStr + " , itemcouponidx = (case when itemcostCouponNotApplied <> " & itemcost & " then -1 else NULL end)"
		sqlStr = sqlStr + " where idx=" + CStr(detailidx)  + VbCrlf

		dbget.Execute sqlStr,nrowCount

        Call DoRecalcOrderMaster(orderserial)

		call AddCsMemo(orderserial,"1",userid,session("ssBctId"), "상품쿠폰가 수기변경")

		response.write "<script>alert('" + CStr(nrowCount) + "개 수정 되었습니다.');</script>"
 		response.write "<script>location.replace('/common/orderdetailedit.asp?idx=" + detailidx + "');</script>"
 		dbget.close()	:	response.End

elseif mode="currstate" then
    ''상태변경
        if (currstate="") then  ''미확인
            sqlStr = "update [db_order].[dbo].tbl_order_detail" + VbCrlf
    		sqlStr = sqlStr + " set currstate=0"  & VbCrlf
    		sqlStr = sqlStr + " ,upcheconfirmdate=NULL" & VbCrlf
    		sqlStr = sqlStr + " ,beasongdate=NULL" & VbCrlf                                             ''//2013/04/01 수정
    		''sqlStr = sqlStr + " ,beasongdate=IsNULL(beasongdate,NULL)" & VbCrlf
    		sqlStr = sqlStr + " ,songjangdiv=NULL"
    		sqlStr = sqlStr + " ,songjangno=NULL"
    		sqlStr = sqlStr + " where idx=" + CStr(detailidx)  + VbCrlf
    		dbget.Execute sqlStr,nrowCount

			call AddCsMemo(orderserial,"1",userid,session("ssBctId"), "상태변경 : 미확인")
    	elseif (currstate="2") then  ''업체통보(물류통보)
    	    sqlStr = "update D" + VbCrlf
    		sqlStr = sqlStr + " set D.currstate=" + CStr(currstate) + ""  & VbCrlf

			'// 2015-10-14, skyer9
			if C_ADMIN_AUTH then
    			sqlStr = sqlStr + " ,beasongdate=NULL" & VbCrlf
    			sqlStr = sqlStr + " ,songjangdiv=NULL"
    			sqlStr = sqlStr + " ,songjangno=NULL"
			end if

    		sqlStr = sqlStr + " From [db_order].[dbo].tbl_order_detail D" & VbCrlf
    		sqlStr = sqlStr + "     Join [db_order].[dbo].tbl_order_master M" & VbCrlf
    		sqlStr = sqlStr + "     on D.orderserial=M.orderserial" & VbCrlf
    		sqlStr = sqlStr + " where D.idx=" + CStr(detailidx)  + VbCrlf
    		sqlStr = sqlStr + " and M.ipkumdiv>3"

			'// 2015-10-14, skyer9
			if not(C_ADMIN_AUTH) then
    			sqlStr = sqlStr + " and D.currstate=0"
			end if

    		dbget.Execute sqlStr,nrowCount

			call AddCsMemo(orderserial,"1",userid,session("ssBctId"), "상태변경 : 업체통보")
        elseif (currstate="3") then  ''업체확인
            sqlStr = "update [db_order].[dbo].tbl_order_detail" + VbCrlf
    		sqlStr = sqlStr + " set currstate=" + CStr(currstate) + ""  & VbCrlf
    		sqlStr = sqlStr + " ,upcheconfirmdate=IsNULL(upcheconfirmdate,getdate())" & VbCrlf
    		sqlStr = sqlStr + " ,beasongdate=NULL" & VbCrlf                                             ''//2013/04/01 수정
    		''sqlStr = sqlStr + " ,beasongdate=IsNULL(beasongdate,NULL)" & VbCrlf
    		sqlStr = sqlStr + " ,songjangdiv=NULL"
    		sqlStr = sqlStr + " ,songjangno=NULL"
    		sqlStr = sqlStr + " where idx=" + CStr(detailidx)  + VbCrlf
    		dbget.Execute sqlStr,nrowCount

            '// 배송비 출고이전 전환
            sqlStr = " update d "
    		sqlStr = sqlStr + " set d.currstate = 0, d.beasongdate = NULL "
    		sqlStr = sqlStr + " from "
    		sqlStr = sqlStr + " 	[db_order].[dbo].[tbl_order_detail] d "
    		sqlStr = sqlStr + " 	join ( "
    		sqlStr = sqlStr + " 		select top 1 (case when isupchebeasong = 'Y' then makerid else '' end) as makerid "
    		sqlStr = sqlStr + " 		from [db_order].[dbo].[tbl_order_detail] "
    		sqlStr = sqlStr + " 		where idx = " & detailidx
    		sqlStr = sqlStr + " 	) T on 1 = 1 "
    		sqlStr = sqlStr + " where  "
    		sqlStr = sqlStr + " 	1 = 1 "
    		sqlStr = sqlStr + " 	and d.orderserial = '" & orderserial & "' "
    		sqlStr = sqlStr + " 	and d.itemid = 0 "
    		sqlStr = sqlStr + " 	and d.makerid = T.makerid "
    		sqlStr = sqlStr + " 	and d.currstate = '7' "
            sqlStr = sqlStr + " 	and d.jungsanfixdate is NULL "
            dbget.Execute sqlStr,nrowCount

            call AddCsMemo(orderserial,"1",userid,session("ssBctId"), "상태변경 : 업체확인")

        elseif (currstate="7to3") then  ''출고완료(텐배,업배) 출고이전 전환

            '// 회수내역이 있는지 체크
            sqlStr = " select top 1 a.id " + VbCrlf
    		sqlStr = sqlStr + " from " + VbCrlf
    		sqlStr = sqlStr + " 	[db_cs].[dbo].[tbl_new_as_list] a " + VbCrlf
    		sqlStr = sqlStr + " 	join [db_cs].[dbo].[tbl_new_as_detail] d on a.id = d.masterid " + VbCrlf
    		sqlStr = sqlStr + " where " + VbCrlf
    		sqlStr = sqlStr + " 	1 = 1 " + VbCrlf
    		sqlStr = sqlStr + " 	and a.deleteyn = 'N' " + VbCrlf
    		sqlStr = sqlStr + " 	and a.orderserial = '" & orderserial & "' " + VbCrlf
    		sqlStr = sqlStr + " 	and a.divcd in ('A004', 'A010') " + VbCrlf
    		sqlStr = sqlStr + " 	and d.orderdetailidx = " & detailidx & VbCrlf

            dataExists = False
 	        rsget.Open sqlStr,dbget,1
            if Not rsget.Eof then
 		        dataExists = True
            end if
 	        rsget.close

            if dataExists then
                response.write "회수내역이 존재합니다. 회수내역 삭제후 처리가능합니다."
                dbget.close : response.end
            end if

            sqlStr = "update [db_order].[dbo].tbl_order_detail" + VbCrlf
    		sqlStr = sqlStr + " set currstate = (case when itemid = 0 then 0 else 3 end)"  & VbCrlf
    		sqlStr = sqlStr + " ,upcheconfirmdate=IsNULL(upcheconfirmdate,getdate())" & VbCrlf
    		sqlStr = sqlStr + " ,beasongdate=NULL" & VbCrlf                                             ''//2013/04/01 수정
    		sqlStr = sqlStr + " ,songjangdiv=NULL"
    		sqlStr = sqlStr + " ,songjangno=NULL"
    		sqlStr = sqlStr + " where idx=" + CStr(detailidx)  + VbCrlf
            sqlStr = sqlStr + " and currstate = '7' " + VbCrlf
            ''sqlStr = sqlStr + " and isupchebeasong = 'N' " + VbCrlf
    		dbget.Execute sqlStr,nrowCount

            if (nrowCount > 0) then
                '// 배송비 출고이전 전환
                sqlStr = " update d "
    		    sqlStr = sqlStr + " set d.currstate = 0, d.beasongdate = NULL "
    		    sqlStr = sqlStr + " from "
    		    sqlStr = sqlStr + " 	[db_order].[dbo].[tbl_order_detail] d "
    		    sqlStr = sqlStr + " 	join ( "
    			sqlStr = sqlStr + " 		select top 1 (case when isupchebeasong = 'Y' then makerid else '' end) as makerid "
    		    sqlStr = sqlStr + " 		from [db_order].[dbo].[tbl_order_detail] "
    		    sqlStr = sqlStr + " 		where idx = " & detailidx
    		    sqlStr = sqlStr + " 	) T on 1 = 1 "
    		    sqlStr = sqlStr + " where  "
    		    sqlStr = sqlStr + " 	1 = 1 "
    		    sqlStr = sqlStr + " 	and d.orderserial = '" & orderserial & "' "
    		    sqlStr = sqlStr + " 	and d.itemid = 0 "
    		    sqlStr = sqlStr + " 	and d.makerid = T.makerid "
    		    sqlStr = sqlStr + " 	and d.currstate = '7' "
                sqlStr = sqlStr + " 	and d.jungsanfixdate is NULL "
                dbget.Execute sqlStr,nrowCount

                '// 텐배인 경우 재고업데이트
                sqlStr = " exec [db_summary].[dbo].[sp_Ten_RealtimeStock_chulgo_Rollback] " & detailidx
                dbget.Execute sqlStr

                call AddCsMemo(orderserial,"1",userid,session("ssBctId"), "상태변경 : 출고완료 상품준비중 전환")
            else
                response.write "에러 : 출고완료상태가 아닙니다."
                dbget.close : response.end
            end if

        elseif (currstate="7") then  ''출고완료
            sqlStr = "update [db_order].[dbo].tbl_order_detail" + VbCrlf
    		sqlStr = sqlStr + " set currstate=" + CStr(currstate) + ""  & VbCrlf
    		sqlStr = sqlStr + " ,upcheconfirmdate=IsNULL(upcheconfirmdate,getdate())" & VbCrlf
    		sqlStr = sqlStr + " ,beasongdate=IsNULL(beasongdate,getdate())" & VbCrlf
    		sqlStr = sqlStr + " where idx=" + CStr(detailidx)  + VbCrlf
    		dbget.Execute sqlStr,nrowCount

			call AddCsMemo(orderserial,"1",userid,session("ssBctId"), "상태변경 : 출고완료")

			sqlStr = "update [db_order].[dbo].tbl_order_detail" + VbCrlf
			sqlStr = sqlStr + " set songjangdiv='" + CStr(songjangdiv) + "'" + VbCrlf
			sqlStr = sqlStr + " ,songjangno='" + CStr(songjangno) + "'" + VbCrlf
			sqlStr = sqlStr + " where idx=" + CStr(detailidx)  + VbCrlf
			dbget.Execute sqlStr,nrowCount

			call AddCsMemo(orderserial,"1",userid,session("ssBctId"), "택배사수정")

			''로그 / 추적 큐 추가 //2019/06/27 by eastone
			sqlStr = " exec db_order.[dbo].[usp_Ten_Delivery_Trace_ChgOrderSongjang_AddOnlyLog] "&CStr(detailidx)&",'"&orderserial&"','"&presongjangno&"',"&CHKIIF(presongjangdiv="","NULL",presongjangdiv)&",'"&songjangno&"',"&CHKIIF(songjangdiv="","NULL",songjangdiv)&",'"&session("ssBctId")&"'"
			dbget.Execute sqlStr

        end if
        '' MASTER 상태 변경 추가 by eastone (전체 출고 완료인경우 8, 일부출고-7, 확인건 있을경우 상품준비-6, 주문통보-5, 입금완료 4

		''배송비
		sqlStr = "update B "
		sqlStr = sqlStr + " set B.currstate = 0 "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_order].[dbo].[tbl_order_detail] A "
		sqlStr = sqlStr + " 	join [db_order].[dbo].[tbl_order_detail] B "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and A.orderserial = B.orderserial "
		sqlStr = sqlStr + " 		and A.makerid = B.makerid "
		sqlStr = sqlStr + " 		and A.isupchebeasong = 'Y' "
		sqlStr = sqlStr + " 		and A.idx = " + CStr(detailidx) + " "
		sqlStr = sqlStr + " 		and B.itemid = 0 "
		sqlStr = sqlStr + " 		and A.currstate <> 7 "
		sqlStr = sqlStr + " 		and B.currstate = 7 "
		dbget.Execute sqlStr,nrowCount

        sqlStr = "update M" + VbCrlf
        sqlStr = sqlStr + " set ipkumdiv=(CASE WHEN T.TTLCNT=T.chulCNT THEN 8" + VbCrlf
        sqlStr = sqlStr + " 				   WHEN T.chulCNT>0 THEN 7" + VbCrlf
        sqlStr = sqlStr + " 				   WHEN T.confirmCNT>0 THEN 6" + VbCrlf
        sqlStr = sqlStr + " 				   WHEN T.tongCNT>0 THEN 5" + VbCrlf
        sqlStr = sqlStr + " 				   WHEN (M.ipkumdiv>4) and (M.baljudate is Not NULL) and (M.jumundiv<>9) THEN 5" + VbCrlf
        sqlStr = sqlStr + " 				   WHEN (M.ipkumdiv>3) and (M.baljudate is NULL) and (M.jumundiv<>9) THEN 4" + VbCrlf
        sqlStr = sqlStr + " 				   ELSE ipkumdiv END)" + VbCrlf
        sqlStr = sqlStr + " ,beadaldate=(CASE WHEN T.TTLCNT=T.chulCNT THEN getdate()" + VbCrlf
        sqlStr = sqlStr + " 				   ELSE beadaldate END)" + VbCrlf
        sqlStr = sqlStr + " from db_order.dbo.tbl_order_master M" + VbCrlf
        sqlStr = sqlStr + " 	Join (" + VbCrlf
        sqlStr = sqlStr + " 		select orderserial, count(*) as TTLCNT" + VbCrlf
        sqlStr = sqlStr + " 		,SUM(CASE WHEN IsNULL(currstate,0)=0 THEN 1 ELSE 0 END) as nothingCNT" + VbCrlf
        sqlStr = sqlStr + " 		,SUM(CASE WHEN currstate=2 THEN 1 ELSE 0 END) as tongCNT" + VbCrlf
        sqlStr = sqlStr + " 		,SUM(CASE WHEN currstate=3 THEN 1 ELSE 0 END) as confirmCNT" + VbCrlf
        sqlStr = sqlStr + " 		,SUM(CASE WHEN currstate=7 THEN 1 ELSE 0 END) as chulCNT" + VbCrlf
        sqlStr = sqlStr + " 		from db_order.dbo.tbl_order_detail" + VbCrlf
        sqlStr = sqlStr + " 		where orderserial='"&orderserial&"'" + VbCrlf
        sqlStr = sqlStr + " 		and itemid<>0" + VbCrlf
        sqlStr = sqlStr + " 		and cancelyn<>'Y'" + VbCrlf
        sqlStr = sqlStr + " 		group by orderserial" + VbCrlf
        sqlStr = sqlStr + " 	) T on  M.orderserial=T.orderserial" + VbCrlf
        sqlStr = sqlStr + " where M.orderserial='"&orderserial&"'" + VbCrlf

        dbget.Execute sqlStr

		sqlStr = " exec db_order.[dbo].[sp_Ten_recalcuMiChulgoMile_AddQue] '" + CStr(orderserial) + "' "
		dbget.Execute sqlStr

        response.write "<script>alert('" + CStr(nrowCount) + "개 수정 되었습니다.');</script>"
 		response.write "<script>location.replace('/common/orderdetailedit.asp?idx=" + detailidx + "');</script>"
 		dbget.close()	:	response.End

elseif mode="isupchebeasong" then
    '배송구분변경 및 매입구분, 배송방식
        sqlStr = "update [db_order].[dbo].tbl_order_detail" + VbCrlf
		sqlStr = sqlStr + " set isupchebeasong='" + CStr(isupchebeasong) + "'" & VbCrlf
		sqlStr = sqlStr + " ,omwdiv='" + CStr(omwdiv) + "'" & VbCrlf
		sqlStr = sqlStr + " ,odlvType='" + CStr(odlvType) + "'" & VbCrlf
		sqlStr = sqlStr + " where idx=" + CStr(detailidx)  + VbCrlf
		dbget.Execute sqlStr,nrowCount

		call AddCsMemo(orderserial,"1",userid,session("ssBctId"), "배송구분변경")

		response.write "<script>alert('" + CStr(nrowCount) + "개 수정 되었습니다.');</script>"
 		response.write "<script>location.replace('/common/orderdetailedit.asp?idx=" + detailidx + "');</script>"
 		dbget.close()	:	response.End

elseif mode="songjangdiv" then


	'택배사수정
		if (applyallitem = "Y") then

            detailidxArr = ""
            songjangdivArr = ""
            songjangnoArr = ""
            sqlStr = " select b.idx as detailidx, b.songjangdiv, b.songjangno "
            sqlStr = sqlStr + " from "
            sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_detail a "
            sqlStr = sqlStr + "     join [db_order].[dbo].tbl_order_detail b "
            sqlStr = sqlStr + "     on "
            sqlStr = sqlStr + "     	1 = 1 "
            sqlStr = sqlStr + "         and a.orderserial = b.orderserial "
            sqlStr = sqlStr + "         and a.isupchebeasong = b.isupchebeasong "
            sqlStr = sqlStr + "         and b.itemid not in (0, 100) "
            sqlStr = sqlStr + "         and b.cancelyn <> 'Y' "
            sqlStr = sqlStr + "         and ( "
            sqlStr = sqlStr + "         	(a.isupchebeasong = 'Y' and a.makerid = b.makerid) "
            sqlStr = sqlStr + " 			or "
            sqlStr = sqlStr + " 			(a.isupchebeasong = 'N') "
            sqlStr = sqlStr + "         ) "
            sqlStr = sqlStr + "         and IsNull(b.currstate, 0) = 7 "
            sqlStr = sqlStr + "         and a.idx = " & detailidx

            rsget.CursorLocation = adUseClient
	        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

	        if  not rsget.EOF  then
		        do until rsget.eof
                    detailidxArr = detailidxArr & "|" & rsget("detailidx")
	    	        songjangdivArr = songjangdivArr & "|" & rsget("songjangdiv")
                    songjangnoArr = songjangnoArr & "|" & rsget("songjangno")
			        rsget.moveNext
		        loop
	        end if
	        rsget.close

            detailidxArr = Split(detailidxArr, "|")
            songjangdivArr = Split(songjangdivArr, "|")
            songjangnoArr = Split(songjangnoArr, "|")

            for i = 0 to UBound(detailidxArr)
                if Trim(detailidxArr(i)) <> "" then
                    presongjangno = Trim(songjangnoArr(i))
                    presongjangdiv = Trim(songjangdivArr(i))

			        sqlStr = "update [db_order].[dbo].tbl_order_detail" + VbCrlf
			        sqlStr = sqlStr + " set songjangdiv='" + CStr(songjangdiv) + "'" + VbCrlf
			        sqlStr = sqlStr + " ,songjangno='" + CStr(songjangno) + "'" + VbCrlf
			        sqlStr = sqlStr + " where idx=" + CStr(Trim(detailidxArr(i)))  + VbCrlf
			        dbget.Execute sqlStr,nrowCount

			        ''로그 / 추적 큐 추가 //2019/06/27 by eastone
			        sqlStr = " exec db_order.[dbo].[usp_Ten_Delivery_Trace_ChgOrderSongjang_AddOnlyLog] "&CStr(Trim(detailidxArr(i)))&",'"&orderserial&"','"&presongjangno&"',"&CHKIIF(presongjangdiv="","NULL",presongjangdiv)&",'"&songjangno&"',"&CHKIIF(songjangdiv="","NULL",songjangdiv)&",'"&session("ssBctId")&"'"
			        dbget.Execute sqlStr
                end if
            Next
		else
			'출고완료만 입력가능
			if (currstate<>"7") and (currstate <> "7to3") then
				response.write "<script>alert('출고완료로 변경 후, 입력하세요.[" & currstate & "]');</script>"
				response.write "<script>history.back();</script>"
				dbget.close()	:	response.End
			end if

			sqlStr = "update [db_order].[dbo].tbl_order_detail" + VbCrlf
			sqlStr = sqlStr + " set songjangdiv='" + CStr(songjangdiv) + "'" + VbCrlf
			sqlStr = sqlStr + " ,songjangno='" + CStr(songjangno) + "'" + VbCrlf
			sqlStr = sqlStr + " where idx=" + CStr(detailidx)  + VbCrlf
			dbget.Execute sqlStr,nrowCount

			''로그 / 추적 큐 추가 //2019/06/27 by eastone
			sqlStr = " exec db_order.[dbo].[usp_Ten_Delivery_Trace_ChgOrderSongjang_AddOnlyLog] "&CStr(detailidx)&",'"&orderserial&"','"&presongjangno&"',"&CHKIIF(presongjangdiv="","NULL",presongjangdiv)&",'"&songjangno&"',"&CHKIIF(songjangdiv="","NULL",songjangdiv)&",'"&session("ssBctId")&"'"
			dbget.Execute sqlStr
		end if

		call AddCsMemo(orderserial,"1",userid,session("ssBctId"), "택배사수정")

		''dbget.Execute sqlStr,nrowCount

		response.write "<script>alert('" + CStr(nrowCount) + "개 수정 되었습니다.');</script>"
 		response.write "<script>location.replace('/common/orderdetailedit.asp?idx=" + detailidx + "');</script>"
 		dbget.close()	:	response.End

elseif mode="songjangno" then
    '운송장번호수정
        sqlStr = "update [db_order].[dbo].tbl_order_detail" + VbCrlf
		sqlStr = sqlStr + " set songjangno='" + CStr(songjangno) + "'" + VbCrlf
		sqlStr = sqlStr + " where idx=" + CStr(detailidx)  + VbCrlf

		dbget.Execute sqlStr,nrowCount

		call AddCsMemo(orderserial,"1",userid,session("ssBctId"), "운송장번호수정")

		response.write "<script>alert('" + CStr(nrowCount) + "개 수정 되었습니다.');</script>"
 		response.write "<script>location.replace('/common/orderdetailedit.asp?idx=" + detailidx + "');</script>"
 		dbget.close()	:	response.End

elseif mode="vatinclude" then
    '과세구분수정
        sqlStr = "update [db_order].[dbo].tbl_order_detail" + VbCrlf
		sqlStr = sqlStr + " set vatinclude='" + CStr(vatinclude) + "'" + VbCrlf
		sqlStr = sqlStr + " where idx=" + CStr(detailidx)  + VbCrlf

		dbget.Execute sqlStr,nrowCount

		call AddCsMemo(orderserial,"1",userid,session("ssBctId"), "과세구분수정")

		response.write "<script>alert('" + CStr(nrowCount) + "개 수정 되었습니다.');</script>"
 		response.write "<script>location.replace('/common/orderdetailedit.asp?idx=" + detailidx + "');</script>"
 		dbget.close()	:	response.End

elseif mode="requiredetail" then
        response.write "사용중지 메뉴, 관리자 문의 요망."
        response.end

        sqlStr = "update [db_order].[dbo].tbl_order_detail" + VbCrlf
		sqlStr = sqlStr + " set requiredetail='" + CStr(requiredetail) + "'" + VbCrlf
		sqlStr = sqlStr + " where idx=" + CStr(detailidx)  + VbCrlf

		dbget.Execute sqlStr,nrowCount

        response.write "<script>alert('" + CStr(nrowCount) + "개 수정 되었습니다.');</script>"
 		response.write "<script>location.replace('/common/orderdetailedit.asp?idx=" + detailidx + "');</script>"
 		dbget.close()	:	response.End

elseif mode="itemno" then
        '수량수정  -- 오더마스터도 같이 수정해야 함 / 관리자 메뉴
        sqlStr = "update [db_order].[dbo].tbl_order_detail" + VbCrlf
		sqlStr = sqlStr + " set itemno='" + CStr(itemno) + "'" + VbCrlf
		sqlStr = sqlStr + " where idx=" + CStr(detailidx)  + VbCrlf

		dbget.Execute sqlStr,nrowCount

        if (nrowCount>0) then
			call AddCsMemo(orderserial,"1",userid,session("ssBctId"), "수량수정")
            Call recalcuOrderMaster(orderserial)
        end if

        response.write "<script>alert('" + CStr(nrowCount) + "개 수정 되었습니다.');</script>"
 		response.write "<script>location.replace('/common/orderdetailedit.asp?idx=" + detailidx + "');</script>"
 		dbget.close()	:	response.End

elseif mode="recalcmaster" then

    Call DoRecalcOrderMaster(orderserial)

    nrowCount = 1

    response.write "<script>alert('" + CStr(nrowCount) + "개 수정 되었습니다.');</script>"
 	response.write "<script>location.replace('/common/orderdetailedit.asp?idx=" + detailidx + "');</script>"
 	dbget.close()	:	response.End
elseif mode="jungsan" then

    if (CheckJungsanExists(orderserial) = True) then
        response.write "<script>alert('정산내역이 존재합니다.');</script>"
    else
        response.write "<script>alert('정산내역이 없습니다.');</script>"
    end if

 	response.write "<script>location.replace('/common/orderdetailedit.asp?idx=" + detailidx + "');</script>"
 	dbget.close()	:	response.End
elseif mode="10x10logistics" then

	sqlStr = "insert into [db_order].[dbo].tbl_order_detail"
	sqlStr = sqlStr + " (masteridx, orderserial,itemid,itemoption,itemno," & vbCrlf
	sqlStr = sqlStr + " itemcost,itemvat,mileage,reducedPrice,itemname," & vbCrlf
	sqlStr = sqlStr + " itemoptionname,makerid,buycash,vatinclude,isupchebeasong,issailitem,oitemdiv,omwdiv,odlvType,orgitemcost, itemcostCouponNotApplied, buycashCouponNotApplied, plussaleDiscount, specialShopDiscount, currstate,beasongdate,upcheconfirmdate,itemcouponidx, bonuscouponidx)" & vbCrlf
	sqlStr = sqlStr + " select masteridx " & vbCrlf
	sqlStr = sqlStr + " ,orderserial " & vbCrlf
	sqlStr = sqlStr + " ,0" & vbCrlf
	sqlStr = sqlStr + " ,'0101'" & vbCrlf
	sqlStr = sqlStr + " ,1" & vbCrlf
	sqlStr = sqlStr + " , 500 " & vbCrlf
	sqlStr = sqlStr + " , 45 " & vbCrlf
	sqlStr = sqlStr + " , 0 " & vbCrlf
	sqlStr = sqlStr + " , 500 " & vbCrlf
	sqlStr = sqlStr + " , '배송비' " & vbCrlf
	sqlStr = sqlStr + " , '' " & vbCrlf
	sqlStr = sqlStr + " , '10x10logistics' " & vbCrlf
	sqlStr = sqlStr + " , 0 " & vbCrlf
	sqlStr = sqlStr + " , 'Y' " & vbCrlf
	sqlStr = sqlStr + " , NULL " & vbCrlf
	sqlStr = sqlStr + " , 'N' " & vbCrlf
	sqlStr = sqlStr + " , '01' " & vbCrlf
	sqlStr = sqlStr + " , NULL " & vbCrlf
	sqlStr = sqlStr + " , NULL " & vbCrlf
	sqlStr = sqlStr + " , 500 " & vbCrlf
	sqlStr = sqlStr + " , 500 " & vbCrlf
	sqlStr = sqlStr + " , 0 " & vbCrlf
	sqlStr = sqlStr + " , 0 " & vbCrlf
	sqlStr = sqlStr + " , 0 " & vbCrlf
	sqlStr = sqlStr + " ,'0'" & vbCrlf
	sqlStr = sqlStr + " ,NULL" & vbCrlf
	sqlStr = sqlStr + " ,NULL" & vbCrlf
	sqlStr = sqlStr + " ,NULL, NULL " & vbCrlf
	sqlStr = sqlStr + " from " & vbCrlf
	sqlStr = sqlStr + "		[db_order].[dbo].tbl_order_detail d " & vbCrlf
	sqlStr = sqlStr + " where idx = " & detailidx & vbCrlf
    dbget.Execute sqlStr,nrowCount

    Call recalcuOrderMaster(orderserial)

    response.write "<script>alert('" + CStr(nrowCount) + "개 수정 되었습니다.');</script>"
 	response.write "<script>location.replace('/common/orderdetailedit.asp?idx=" + detailidx + "');</script>"
 	dbget.close()	:	response.End

elseif mode="balju" then

	sqlStr = " select top 1 orderserial "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " [db_order].[dbo].[tbl_baljudetail] "
	sqlStr = sqlStr + " where orderserial = '" & orderserial & "' "

    dataExists = False
 	rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
 		dataExists = True
    end if
 	rsget.close

    if dataExists then
        response.write "에러 : 정상 발주된 내역이니다."
        dbget.close()	:	response.End
    end if

	sqlStr = " update "
	sqlStr = sqlStr + " [db_order].[dbo].[tbl_order_master] "
	sqlStr = sqlStr + " set ipkumdiv = '4', baljudate = NULL "
	sqlStr = sqlStr + " where orderserial = '" & orderserial & "' "
    dbget.Execute sqlStr,nrowCount

    call AddCsMemo(orderserial,"1",userid,session("ssBctId"), "발주중 => 결제완료 전환")

    response.write "<script>alert('" + CStr(nrowCount) + "개 수정 되었습니다.');</script>"
 	response.write "<script>location.replace('/common/orderdetailedit.asp?idx=" + detailidx + "');</script>"
 	dbget.close()	:	response.End

elseif mode="updmastercoupon" then

    Call recalcuOrderMaster(orderserial)

	sqlStr = " select top 1 tencardspend "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " [db_order].[dbo].[tbl_order_master] "
	sqlStr = sqlStr + " where orderserial = '" & orderserial & "' "

    dataExists = False
 	rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
 		tencardspend = rsget("tencardspend")
    end if
 	rsget.close

	sqlStr = " select sum((itemcost - reducedPrice - isnull(etcDiscount, 0) )*itemno) as tencardspend "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " [db_order].[dbo].[tbl_order_detail] "
	sqlStr = sqlStr + " where orderserial = '" & orderserial & "' and cancelyn <> 'Y' "

 	rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        if (CLng(tencardspend) <> CLng(rsget("tencardspend"))) then
            dataExists = True
            tencardspend = rsget("tencardspend")
        end if
    end if
 	rsget.close

    nrowCount = 0
    if dataExists = True then
        sqlStr = " update "
	    sqlStr = sqlStr + " [db_order].[dbo].[tbl_order_master] "
	    sqlStr = sqlStr + " set tencardspend = " & tencardspend & ", subtotalprice = totalsum - " & tencardspend
	    sqlStr = sqlStr + " where orderserial = '" & orderserial & "' "
        dbget.Execute sqlStr,nrowCount

        Call recalcuOrderMaster(orderserial)

        call AddCsMemo(orderserial,"1",userid,session("ssBctId"), "쿠폰액 오류 보정")
    end if

    response.write "<script>alert('" + CStr(nrowCount) + "개 수정 되었습니다.');</script>"
 	response.write "<script>location.replace('/common/orderdetailedit.asp?idx=" + detailidx + "');</script>"
 	dbget.close()	:	response.End

elseif mode="ipkumdate" then

	sqlStr = " update "
	sqlStr = sqlStr + " [db_order].[dbo].[tbl_order_master] "
	sqlStr = sqlStr + " set ipkumdate = '" & ipkumdate & "' "
	sqlStr = sqlStr + " where orderserial = '" & orderserial & "' "
	dbget.Execute sqlStr,nrowCount

	call AddCsMemo(orderserial,"1",userid,session("ssBctId"), "입금일 수정")

    response.write "<script>alert('" + CStr(nrowCount) + "개 수정 되었습니다.');</script>"
 	response.write "<script>location.replace('/common/orderdetailedit.asp?idx=" + detailidx + "');</script>"
 	dbget.close()	:	response.End

elseif mode="additemid" then

    orgorderserial     = request("orgorderserial")
    orgdetailidx     = request("orgdetailidx")

	sqlStr = " insert into [db_order].[dbo].[tbl_order_detail]( "
	sqlStr = sqlStr + " 	[orderserial], "
	sqlStr = sqlStr + " 	[itemid], [itemoption], "
	sqlStr = sqlStr + " 	[masteridx], "
	sqlStr = sqlStr + " 	[makerid], [itemno], [itemcost], [mileage], [reducedPrice], [cancelyn], [currstate], [songjangno], [songjangdiv], "
	sqlStr = sqlStr + " 	[itemname], [itemoptionname], [buycash], [itemvat], [vatinclude], [beasongdate], "
	sqlStr = sqlStr + " 	[isupchebeasong], [omwdiv], [odlvType], [issailitem], [upcheconfirmdate], [oitemdiv], [requiredetail], [itemcouponidx], [bonuscouponidx], [canceldate], "
	sqlStr = sqlStr + " 	[passday], [orgitemcost], [itemcostCouponNotApplied], [buycashCouponNotApplied], [odlvfixday], [plusSaleDiscount], [specialshopDiscount], [etcDiscount], "
	sqlStr = sqlStr + " 	[dlvfinishdt], [jungsanfixdate], [plus_sale_item_idx] "
	sqlStr = sqlStr + " ) "
	sqlStr = sqlStr + " select top 1 "
	sqlStr = sqlStr + " 	'" & orderserial & "', "
	sqlStr = sqlStr + " 	[itemid], [itemoption], "
	sqlStr = sqlStr + " 	(select top 1 masteridx from [db_order].[dbo].[tbl_order_detail] where orderserial = '" & orderserial & "'), "
	sqlStr = sqlStr + " 	[makerid], 1, [itemcost], [mileage], [reducedPrice], 'N', '1', NULL, NULL, "
	sqlStr = sqlStr + " 	[itemname], [itemoptionname], [buycash], [itemvat], [vatinclude], NULL, "
	sqlStr = sqlStr + " 	[isupchebeasong], [omwdiv], [odlvType], [issailitem], [upcheconfirmdate], [oitemdiv], [requiredetail], [itemcouponidx], [bonuscouponidx], NULL, "
	sqlStr = sqlStr + " 	[passday], [orgitemcost], [itemcostCouponNotApplied], [buycashCouponNotApplied], [odlvfixday], [plusSaleDiscount], [specialshopDiscount], [etcDiscount], "
	sqlStr = sqlStr + " 	NULL, NULL, [plus_sale_item_idx] "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	[db_order].[dbo].[tbl_order_detail] d "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and d.orderserial = '" & orgorderserial & "' "
	sqlStr = sqlStr + " 	and d.idx = " & orgdetailidx
    dbget.Execute sqlStr,nrowCount

    if (nrowCount = 0) then
	    sqlStr = " insert into [db_order].[dbo].[tbl_order_detail]( "
	    sqlStr = sqlStr + " 	[orderserial], "
	    sqlStr = sqlStr + " 	[itemid], [itemoption], "
	    sqlStr = sqlStr + " 	[masteridx], "
	    sqlStr = sqlStr + " 	[makerid], [itemno], [itemcost], [mileage], [reducedPrice], [cancelyn], [currstate], [songjangno], [songjangdiv], "
	    sqlStr = sqlStr + " 	[itemname], [itemoptionname], [buycash], [itemvat], [vatinclude], [beasongdate], "
	    sqlStr = sqlStr + " 	[isupchebeasong], [omwdiv], [odlvType], [issailitem], [upcheconfirmdate], [oitemdiv], [requiredetail], [itemcouponidx], [bonuscouponidx], [canceldate], "
	    sqlStr = sqlStr + " 	[passday], [orgitemcost], [itemcostCouponNotApplied], [buycashCouponNotApplied], [odlvfixday], [plusSaleDiscount], [specialshopDiscount], [etcDiscount], "
	    sqlStr = sqlStr + " 	[dlvfinishdt], [jungsanfixdate], [plus_sale_item_idx] "
	    sqlStr = sqlStr + " ) "
	    sqlStr = sqlStr + " select top 1 "
	    sqlStr = sqlStr + " 	'" & orderserial & "', "
	    sqlStr = sqlStr + " 	[itemid], [itemoption], "
	    sqlStr = sqlStr + " 	(select top 1 masteridx from [db_order].[dbo].[tbl_order_detail] where orderserial = '" & orderserial & "'), "
	    sqlStr = sqlStr + " 	[makerid], 1, [itemcost], [mileage], [reducedPrice], 'N', '1', NULL, NULL, "
	    sqlStr = sqlStr + " 	[itemname], [itemoptionname], [buycash], [itemvat], [vatinclude], NULL, "
	    sqlStr = sqlStr + " 	[isupchebeasong], [omwdiv], [odlvType], [issailitem], [upcheconfirmdate], [oitemdiv], [requiredetail], [itemcouponidx], [bonuscouponidx], NULL, "
	    sqlStr = sqlStr + " 	[passday], [orgitemcost], [itemcostCouponNotApplied], [buycashCouponNotApplied], [odlvfixday], [plusSaleDiscount], [specialshopDiscount], [etcDiscount], "
	    sqlStr = sqlStr + " 	NULL, NULL, [plus_sale_item_idx] "
	    sqlStr = sqlStr + " from "
	    sqlStr = sqlStr + " 	[db_log].[dbo].[tbl_old_order_detail_2003] d "
	    sqlStr = sqlStr + " where "
	    sqlStr = sqlStr + " 	1 = 1 "
	    sqlStr = sqlStr + " 	and d.orderserial = '" & orgorderserial & "' "
	    sqlStr = sqlStr + " 	and d.idx = " & orgdetailidx
        dbget.Execute sqlStr,nrowCount
    end if

	call AddCsMemo(orderserial,"1",userid,session("ssBctId"), "고객추가결제 : 원주문 상품추가")

    response.write "<script>alert('" + CStr(nrowCount) + "개 수정 되었습니다.');</script>"
 	response.write "<script>location.replace('/common/orderdetailedit.asp?idx=" + detailidx + "');</script>"
 	dbget.close()	:	response.End

end if
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
