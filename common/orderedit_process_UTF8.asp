<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
%>
<%
'###########################################################
' Description : 고객센터
' History : 이상구 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib_UTF8.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/lib/function_UTF8.asp"-->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp"-->
<%
dim orderserial, detailidx, mode
dim buycash, isupchebeasong, songjangdiv, songjangno
dim beasongdate, currstate, upcheconfirmdate
dim requiredetail, itemno, omwdiv, odlvType, applyallitem
dim vatinclude
dim makerid, userid
dim presongjangno, presongjangdiv, reducedPrice, itemcost
dim prevBonusCouponDiscountPrice, currBonusCouponDiscountPrice

orderserial     = request("orderserial")
detailidx       = request("detailidx")
mode            = request("mode")
buycash         = request("buycash")
reducedPrice    = request("reducedPrice")
itemcost    	= request("itemcost")
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

presongjangno	= requestCheckvar(request("presongjangno"),32)
presongjangdiv	= requestCheckvar(request("presongjangdiv"),10)

dim tmp
On Error resume Next
if (upcheconfirmdate<>"") then tmp = CDate(upcheconfirmdate)
if Err then
    response.write "<script language='javascript'>alert('업체 확인일이 올바르지 않습니다.');history.back();</script>"
	session.codePage = 949
    dbget.close()	:	response.End
end if

if (beasongdate<>"") then tmp = CDate(beasongdate)
if Err then
    response.write "<script language='javascript'>alert('업체 배송일이  올바르지 않습니다.');history.back();</script>"
	session.codePage = 949
    dbget.close()	:	response.End
end if

On Error Goto 0

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
		 session.codePage = 949
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
		session.codePage = 949
 		dbget.close()	:	response.End
 	end if
end if

if (mode="buycash") then
    '매입가 변경
		sqlStr = "update [db_order].[dbo].tbl_order_detail" + VbCrlf
		sqlStr = sqlStr + " set buycash=" + CStr(buycash)  + VbCrlf
		sqlStr = sqlStr + " where idx=" + CStr(detailidx)  + VbCrlf

		dbget.Execute sqlStr,nrowCount

		call AddCsMemo(orderserial,"1",userid,session("ssBctId"), "매입가 수기변경")

		response.write "<script>alert('" + CStr(nrowCount) + "개 수정 되었습니다.');</script>"
		session.codePage = 949
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
			session.codePage = 949
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

		call AddCsMemo(orderserial,"1",userid,session("ssBctId"), "쿠폰가 수기변경")

		response.write "<script>alert('" + CStr(nrowCount) + "개 수정 되었습니다.');</script>"
		session.codePage = 949
 		response.write "<script>location.replace('/common/orderdetailedit.asp?idx=" + detailidx + "');</script>"
 		dbget.close()	:	response.End
elseif (mode="itemcost") then
    '판매가 변경
		sqlStr = " select top 1 d.itemcost, d.reducedPrice, d.itemno from [db_order].[dbo].[tbl_order_detail] d where d.idx = " & detailidx
 	    rsget.Open sqlStr,dbget,1
        reducedPrice = 0
        if Not rsget.Eof then
 		    reducedPrice = rsget("reducedPrice")
            prevBonusCouponDiscountPrice = (rsget("itemcost") - rsget("reducedPrice")) * rsget("itemno")
            currBonusCouponDiscountPrice = (itemcost - rsget("reducedPrice")) * rsget("itemno")
        end if
 	    rsget.close

        if (CLng(itemcost) < CLng(reducedPrice)) then
            response.write "판매가보다 쿠폰가가 더 클 수 없습니다."
			session.codePage = 949
            dbget.close()	:	response.End
        end if

        sqlStr = "update [db_order].[dbo].tbl_order_detail" + VbCrlf
		sqlStr = sqlStr + " set itemcost=" + CStr(itemcost) + ", itemcostCouponNotApplied = " + CStr(itemcost) + VbCrlf
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

		call AddCsMemo(orderserial,"1",userid,session("ssBctId"), "판매가 수기변경")

		response.write "<script>alert('" + CStr(nrowCount) + "개 수정 되었습니다.');</script>"
		session.codePage = 949
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
			if C_ADMIN_AUTH then
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

        elseif (currstate="7to3") then  ''출고완료(텐배) 출고이전 전환

            '// 회수내역이 있는지 체크
            sqlStr = " select top 1 a.id " + VbCrlf
    		sqlStr = sqlStr + " from " + VbCrlf
    		sqlStr = sqlStr + " 	[db_cs].[dbo].[tbl_new_as_list] a " + VbCrlf
    		sqlStr = sqlStr + " 	join [db_cs].[dbo].[tbl_new_as_detail] d on a.id = d.masterid " + VbCrlf
    		sqlStr = sqlStr + " where " + VbCrlf
    		sqlStr = sqlStr + " 	1 = 1 " + VbCrlf
    		sqlStr = sqlStr + " 	and a.deleteyn = 'N' " + VbCrlf
    		sqlStr = sqlStr + " 	and a.orderserial = '" & orderserial & "' " + VbCrlf
    		sqlStr = sqlStr + " 	and a.divcd = 'A010' " + VbCrlf
    		sqlStr = sqlStr + " 	and d.orderdetailidx = " & detailidx & VbCrlf

            dataExists = False
 	        rsget.Open sqlStr,dbget,1
            if Not rsget.Eof then
 		        dataExists = True
            end if
 	        rsget.close

            if dataExists then
                response.write "회수내역이 존재합니다. 회수내역 삭제후 처리가능합니다."
				session.codePage = 949
                dbget.close : response.end
            end if

            sqlStr = "update [db_order].[dbo].tbl_order_detail" + VbCrlf
    		sqlStr = sqlStr + " set currstate=3"  & VbCrlf
    		sqlStr = sqlStr + " ,upcheconfirmdate=IsNULL(upcheconfirmdate,getdate())" & VbCrlf
    		sqlStr = sqlStr + " ,beasongdate=NULL" & VbCrlf                                             ''//2013/04/01 수정
    		sqlStr = sqlStr + " ,songjangdiv=NULL"
    		sqlStr = sqlStr + " ,songjangno=NULL"
    		sqlStr = sqlStr + " where idx=" + CStr(detailidx)  + VbCrlf
            sqlStr = sqlStr + " and currstate = '7' " + VbCrlf
            sqlStr = sqlStr + " and isupchebeasong = 'N' " + VbCrlf
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

                sqlStr = " exec [db_summary].[dbo].[sp_Ten_RealtimeStock_chulgo_Rollback] " & detailidx
                dbget.Execute sqlStr

                call AddCsMemo(orderserial,"1",userid,session("ssBctId"), "상태변경 : 출고완료 상품준비중 전환")
            else
                response.write "에러 : 출고완료상태가 아닙니다."
				session.codePage = 949
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
		session.codePage = 949
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
		session.codePage = 949
 		response.write "<script>location.replace('/common/orderdetailedit.asp?idx=" + detailidx + "');</script>"
 		dbget.close()	:	response.End

elseif mode="songjangdiv" then


	'택배사수정
		if (applyallitem = "Y") then
			response.write "사용중지" '' 쿼리가 좀 이상함.
			session.codePage = 949
			dbget.close()	:	response.End


			'업체/텐배 각각 전상품 출고완료만 입력가능
			' sqlStr = "select "
			' sqlStr = sqlStr + " 	count(*) as cnt "
			' sqlStr = sqlStr + " from "
			' sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_detail a "
			' sqlStr = sqlStr + " 	join [db_order].[dbo].tbl_order_detail b "
			' sqlStr = sqlStr + " 	on "
			' sqlStr = sqlStr + " 		1 = 1 "
			' sqlStr = sqlStr + " 		and a.orderserial = b.orderserial "
			' sqlStr = sqlStr + " 		and a.makerid = b.makerid "
			' sqlStr = sqlStr + " 		and b.itemid <> 0 "
			' sqlStr = sqlStr + " 		and a.isupchebeasong = 'Y' "
			' sqlStr = sqlStr + " where "
			' sqlStr = sqlStr + " 	1 = 1 "
			' sqlStr = sqlStr + " 	and a.idx = " + CStr(detailidx) + " "
			' sqlStr = sqlStr + " 	and IsNull(b.currstate, 0) < 7 "

			' rsget.Open sqlStr,dbget,1
			' 	dataExists = rsget("cnt")>0
			' rsget.close

			' sqlStr = "select "
			' sqlStr = sqlStr + " 	count(*) as cnt "
			' sqlStr = sqlStr + " from "
			' sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_detail a "
			' sqlStr = sqlStr + " 	join [db_order].[dbo].tbl_order_detail b "
			' sqlStr = sqlStr + " 	on "
			' sqlStr = sqlStr + " 		1 = 1 "
			' sqlStr = sqlStr + " 		and a.orderserial = b.orderserial "
			' sqlStr = sqlStr + " 		and b.itemid not in (0, 100) "
			' sqlStr = sqlStr + " 		and a.isupchebeasong = 'N' "
			' sqlStr = sqlStr + " where "
			' sqlStr = sqlStr + " 	1 = 1 "
			' sqlStr = sqlStr + " 	and a.idx = " + CStr(detailidx) + " "
			' sqlStr = sqlStr + " 	and IsNull(b.currstate, 0) < 7 "

			' rsget.Open sqlStr,dbget,1
			' 	dataExists = dataExists + rsget("cnt")>0
			' rsget.close

			' if (dataExists) then
			' 	response.write "<script>alert('출고완료 상태가 아닌 상품이 있습니다.\n\n전체 상품이 출고완료 상태여야 합니다.');</script>"
			' 	response.write "<script>history.back();</script>"
			' 	dbget.close()	:	response.End
			' end if

			' sqlStr = "update b " + vbCrLf
			' sqlStr = sqlStr + " 	set b.songjangdiv='" + CStr(songjangdiv) + "' ,songjangno='" + CStr(songjangno) + "'  " + vbCrLf
			' sqlStr = sqlStr + " from " + vbCrLf
			' sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_detail a " + vbCrLf
			' sqlStr = sqlStr + " 	join [db_order].[dbo].tbl_order_detail b " + vbCrLf
			' sqlStr = sqlStr + " 	on " + vbCrLf
			' sqlStr = sqlStr + " 		1 = 1 " + vbCrLf
			' sqlStr = sqlStr + " 		and a.orderserial = b.orderserial " + vbCrLf
			' sqlStr = sqlStr + " 		and a.makerid = b.makerid " + vbCrLf
			' sqlStr = sqlStr + " 		and b.itemid <> 0 " + vbCrLf
			' sqlStr = sqlStr + " 		and a.isupchebeasong = 'Y' " + vbCrLf
			' sqlStr = sqlStr + " where " + vbCrLf
			' sqlStr = sqlStr + " 	1 = 1 " + vbCrLf
			' sqlStr = sqlStr + " 	and a.idx = " + CStr(detailidx) + " " + vbCrLf
			' sqlStr = sqlStr + " 	and IsNull(b.currstate, 0) = 7 " + vbCrLf
			' dbget.Execute sqlStr,dataExists

			' sqlStr = "update b " + vbCrLf
			' sqlStr = sqlStr + " 	set b.songjangdiv='" + CStr(songjangdiv) + "' ,songjangno='" + CStr(songjangno) + "'  " + vbCrLf
			' sqlStr = sqlStr + " from " + vbCrLf
			' sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_detail a " + vbCrLf
			' sqlStr = sqlStr + " 	join [db_order].[dbo].tbl_order_detail b " + vbCrLf
			' sqlStr = sqlStr + " 	on " + vbCrLf
			' sqlStr = sqlStr + " 		1 = 1 " + vbCrLf
			' sqlStr = sqlStr + " 		and a.orderserial = b.orderserial " + vbCrLf
			' sqlStr = sqlStr + " 		and b.itemid not in (0, 100) "
			' sqlStr = sqlStr + " 		and a.isupchebeasong = 'N' "
			' sqlStr = sqlStr + " where " + vbCrLf
			' sqlStr = sqlStr + " 	1 = 1 " + vbCrLf
			' sqlStr = sqlStr + " 	and a.idx = " + CStr(detailidx) + " " + vbCrLf
			' sqlStr = sqlStr + " 	and IsNull(b.currstate, 0) = 7 " + vbCrLf
			' dbget.Execute sqlStr,nrowCount

			' nrowCount = nrowCount + dataExists

		else
			'출고완료만 입력가능
			if (currstate<>"7") then
				response.write "<script>alert('출고완료로 변경 후, 입력하세요.[" & currstate & "]');</script>"
				session.codePage = 949
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
		session.codePage = 949
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
		session.codePage = 949
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
		session.codePage = 949
 		response.write "<script>location.replace('/common/orderdetailedit.asp?idx=" + detailidx + "');</script>"
 		dbget.close()	:	response.End

elseif mode="requiredetail" then
        response.write "사용중지 메뉴, 관리자 문의 요망."
		session.codePage = 949
        response.end

        sqlStr = "update [db_order].[dbo].tbl_order_detail" + VbCrlf
		sqlStr = sqlStr + " set requiredetail='" + CStr(requiredetail) + "'" + VbCrlf
		sqlStr = sqlStr + " where idx=" + CStr(detailidx)  + VbCrlf

		dbget.Execute sqlStr,nrowCount

        response.write "<script>alert('" + CStr(nrowCount) + "개 수정 되었습니다.');</script>"
		session.codePage = 949
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
		session.codePage = 949
 		response.write "<script>location.replace('/common/orderdetailedit.asp?idx=" + detailidx + "');</script>"
 		dbget.close()	:	response.End
end if

session.codePage = 949
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
