<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/requestlecturecls.asp"-->
<!-- #include virtual="/academy/lib/classes/fingers_lecturecls.asp"-->
<%

dim oordermaster, oorderdetail

dim id, divcd, gubun01, gubun02, orderserial, customername, userid, writeuser, finishuser, title, contents_jupsu, contents_finish
dim currstate, regdate, finishdate, refundrequire, refundresult, songjangno, beasongdate, cause, causedetail, requireupche, makerid, deleteyn
dim extsitename, rebankname, rebankaccount, rebankownername, refundbeasongpay, refunditemcostsum, refunddeliverypay, refundadjustpay, returnmethod

dim detailitemlist, detailitemnolist
dim did, dmasterid, dorderserial, ditemid, ditemoption, dmakerid, ditemname, ditemoptionname
dim dregitemno, dconfirmitemno, ditemcost, dbuycash, disupchebeasong, dregdetailstate, dcausediv, dcausedetail, dcausecontent, dcurrstate

dim canceldetailno, canceldetailnamelist

dim sitename
dim canceldetailall
dim refundcstitle


dim mode, i, j, k, tmp
dim sqlStr

mode = html2db(request("mode"))

id              = html2db(RequestCheckvar(request("id"),10))
divcd           = html2db(RequestCheckvar(request("divcd"),10))
gubun01         = html2db(RequestCheckvar(request("gubun01"),10))
gubun02         = html2db(RequestCheckvar(request("gubun02"),10))
orderserial     = html2db(RequestCheckvar(request("orderserial"),16))
customername    = html2db(RequestCheckvar(request("customername"),32))
userid          = html2db(RequestCheckvar(request("userid"),32))
writeuser       = html2db(RequestCheckvar(request("writeuser"),32))
finishuser      = html2db(RequestCheckvar(request("finishuser"),32))
title           = html2db(RequestCheckvar(request("title"),128))
contents_jupsu  = html2db(request("contents_jupsu"))
contents_finish = html2db(request("contents_finish"))
currstate       = html2db(RequestCheckvar(request("currstate"),10))
regdate         = html2db(RequestCheckvar(request("regdate"),32))
finishdate      = html2db(RequestCheckvar(request("finishdate"),32))
refundrequire   = html2db(RequestCheckvar(request("refundrequire"),10))
refundresult    = html2db(RequestCheckvar(request("refundresult"),10))
songjangno      = html2db(RequestCheckvar(request("songjangno"),16))
beasongdate     = html2db(RequestCheckvar(request("beasongdate"),32))
cause           = html2db(RequestCheckvar(request("causecd"),10))
causedetail     = html2db(request("causedetail"))
requireupche    = html2db(RequestCheckvar(request("requireupche"),2))
makerid         = html2db(RequestCheckvar(request("makerid"),32))
deleteyn        = html2db(RequestCheckvar(request("deleteyn"),2))
extsitename     = html2db(RequestCheckvar(request("extsitename"),32))
rebankname              = html2db(RequestCheckvar(request("rebankname"),32))
rebankaccount           = html2db(RequestCheckvar(request("rebankaccount"),32))
rebankownername         = html2db(RequestCheckvar(request("rebankownername"),32))
refundbeasongpay        = html2db(RequestCheckvar(request("refundbeasongpay"),10))
refunditemcostsum       = html2db(RequestCheckvar(request("refunditemcostsum"),10))
refunddeliverypay       = html2db(RequestCheckvar(request("refunddeliverypay"),10))
refundadjustpay         = html2db(RequestCheckvar(request("refundadjustpay"),10))
returnmethod            = html2db(RequestCheckvar(request("returnmethod"),32))
sitename            	= html2db(RequestCheckvar(request("sitename"),32))


detailitemlist 			= html2db(request("detailitemlist"))
detailitemnolist 		= html2db(request("detailitemnolist"))

response.write "divcd=" & divcd & "<br>"
response.write "gubun01=" & gubun01 & "<br>"
response.write "gubun02=" & gubun02 & "<br>"
'dbget.close()	:	response.End
if contents_jupsu <> "" then
	if checkNotValidHTML(contents_jupsu) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
	response.write "</script>"
	response.End
	end if
end If
if contents_finish <> "" then
	if checkNotValidHTML(contents_finish) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
	response.write "</script>"
	response.End
	end if
end If
if causedetail <> "" then
	if checkNotValidHTML(causedetail) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
	response.write "</script>"
	response.End
	end if
end if


if (Len(divcd)=1) then divcd="A00" & divcd
if (Len(divcd)=2) then divcd="A0" & divcd

'==============================================================================
set oordermaster = new CRequestLecture
oordermaster.FRectOrderSerial = orderserial
oordermaster.GetRequestLectureMasterOne

set oorderdetail = new CRequestLecture
oorderdetail.FRectOrderSerial = orderserial
oorderdetail.CRequestLectureDetailList

'==============================================================================
dim olecture
set olecture = new CLecture
olecture.FRectIdx = oordermaster.FOneItem.Fitemid

if (olecture.FRectIdx = "") then
    olecture.FRectIdx = "0"
end if
olecture.GetOneLecture

'==============================================================================
'등록체크
'취소된 강좌에 대해 강좌취소 불가능
'취소불가일(강좌시작일 -3일)을 지났는지 체크한다. >> 가능하게 변경
'TODO : 전체취소중 이전 취소환불내역이 있다면 다음 무통장환불페이지에서 나머지금액만 환불할수 있도록 변경
'전체취소된 신청은 부분취소가 불가능합니다.
'카드/실시간이체 취소는 강좌신청이 취소되어야만 가응합니다.
'
if ((mode = "cancelorder") and (oordermaster.FOneItem.Fcancelyn = "Y")) then
    response.write "<script>alert('이미 취소된 강좌입니다.'); history.back();</script>"
    dbget.close()	:	response.End
end if

if ((mode = "cancelitem") and (oordermaster.FOneItem.Fcancelyn = "Y")) then
    response.write "<script>alert('이미 취소된 강좌입니다.'); history.back();</script>"
    dbget.close()	:	response.End
end if

'if ((mode = "cancelorder") and (Left(DateAdd("d",3,now), 10)  > Left(olecture.FOneItem.Flec_startday1, 10))) then
'    response.write "<script>alert('강좌취소는 강좌시작 3일전까지 가능합니다.'); history.back();</script>"
'    dbget.close()	:	response.End
'end if

if ((mode = "cancelcard") and (oordermaster.FOneItem.Fcancelyn <> "Y")) then
    response.write "<script>alert('강좌신청이 취소되지 않았습니다.'); history.back();</script>"
    dbget.close()	:	response.End
end if

if ((mode = "revalidate") and ((oordermaster.FOneItem.Fipkumdiv <> "2") and (oordermaster.FOneItem.Faccountdiv <> "7") and (oordermaster.FOneItem.Fcancelyn <> "Y"))) then
    response.write "<script>alert('취소된 주문이 아니거나, 무통장주문이 아니거나, 주문접수상태가 아닙니다.'); history.back();</script>"
    dbget.close()	:	response.End
end if

if (mode = "revalidate") then
    '비회원이 아닌경우(쿠폰 확인)
    if (oordermaster.FOneItem.FUserID <> "") then
    	'==============================================================
    	if (oordermaster.FOneItem.Ftencardspend <> 0) then
    	    '사용쿠폰 확인
    	    sqlStr = " select top 1 orderserial "
    	    sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_coupon "
    	    sqlStr = sqlStr + " where orderserial = '" + CStr(orderserial) + "' and deleteyn = 'N' and isusing = 'N' "
    	    rsget.Open sqlStr,dbget,1
    	    i = "N"
    	    if  not rsget.EOF  then
    	        i = "Y"
    	    end if
    	    rsget.close

    	    if (i = "N") then
                response.write "<script>alert('주문시 사용되었던 쿠폰이 존재하지 않습니다. 재주문 하세요.'); history.back();</script>"
                dbget.close()	:	response.End
    	    end if
        end if
    end if
end if

if (sitename = "diyitem") then
	refundcstitle = "환불요청(DIY상품)"
else
	refundcstitle = "환불요청(강좌)"
end if

'==============================================================================
if (mode = "cancelorder") then
    '전체취소
    ' - 입금전 취소일 경우, 주문취소AS등록및종료처리, 주문건취소, 마일리지재부여, 사용쿠폰사용안함표시, 강좌 한정수량 재증가
    ' - 입금후 취소일 경우, 주문취소AS등록및종료처리, 주문건취소, 마일리지재부여, 사용쿠폰사용안함표시, 강좌 한정수량 재증가, 환불금액이 있을경우(신용카드취소 또는 무통장입금AS등록페이지로 이동)

    '======================================================================
    '주문취소AS등록및종료처리(마스타)
	sqlStr = " select * from [db_cs].[dbo].tbl_new_as_list where 1=0 "
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("divcd")          = divcd
	rsget("orderserial")    = orderserial
	rsget("customername")   = html2db(oordermaster.FOneItem.FBuyName)
	rsget("userid")         = html2db(oordermaster.FOneItem.FUserID)
	rsget("writeuser")      = session("ssBctId")
	rsget("title")          = title
	rsget("contents_jupsu") = contents_jupsu
	'rsget("refundrequire")  = 0
	'rsget("cause")          = cause
	'rsget("causedetail")    = causedetail
	rsget("requireupche")   = "N"
	rsget("makerid")        = ""
	rsget("deleteyn")       = "N"
	rsget("songjangno")     = songjangno
	'rsget("rebankname")     = rebankname
	'rsget("rebankaccount")  = rebankaccount
	'rsget("rebankownername")        = rebankownername
	'rsget("refundbeasongpay")       = 0
	'rsget("refunditemcostsum")      = 0
	'rsget("refunddeliverypay")      = 0
	'rsget("refundadjustpay")        = 0
	rsget("sitegubun")      = "FI"

	rsget.update
	id = rsget("id")
	rsget.close

	sqlStr = " update [db_cs].[dbo].tbl_new_as_list "
	sqlStr = sqlStr + " set finishdate=getdate() "
	sqlStr = sqlStr + " ,finishuser = '" + session("ssBctId") + "' "
	sqlStr = sqlStr + " ,contents_finish = '" + html2db("강좌취소") + "' "
	sqlStr = sqlStr + " ,currstate = 'B007' "
	'sqlStr = sqlStr + " ,refundresult = 0 "
	sqlStr = sqlStr + " where id=" + CStr(id) + " "
	rsget.Open sqlStr,dbget,1

	'======================================================================
	'주문건취소
	sqlStr = " update [db_academy].[dbo].tbl_academy_order_master "
	sqlStr = sqlStr + " set cancelyn='Y' "
	sqlStr = sqlStr + " where orderserial = '" + CStr(orderserial) + "' "
	rsAcademyget.Open sqlStr,dbAcademyget,1

	'======================================================================
	''한정수량 증가.'강좌신청가능인원 증가(대기자가 없을 경우만)

	dim WaitExist : WaitExist = false
	sqlStr = " select count(*) as cnt from " + vbCrlf
    sqlStr = sqlStr + " db_academy.dbo.tbl_academy_order_master m " + vbCrlf
    sqlStr = sqlStr + " 	Join db_academy.dbo.tbl_academy_order_detail d" + vbCrlf
    sqlStr = sqlStr + " 	on m.orderserial=d.orderserial" + vbCrlf
    sqlStr = sqlStr + " 	Join db_academy.dbo.tbl_lec_waiting_user w" + vbCrlf
    sqlStr = sqlStr + " 	on d.itemid=w.lec_idx" + vbCrlf
    sqlStr = sqlStr + " 	and d.itemoption=w.lecOption" + vbCrlf
    sqlStr = sqlStr + " 	and w.isusing='Y'" + vbCrlf
    sqlStr = sqlStr + " 	and w.currstate<7" + vbCrlf
    sqlStr = sqlStr + " 	and IsNULL(w.regendday,'9999-12-12')>getdate()" + vbCrlf
    sqlStr = sqlStr + " where m.orderserial='" + CStr(orderserial) + "'" + vbCrlf
    rsAcademyget.Open sqlStr,dbAcademyget,1
    if Not rsAcademyget.Eof then
    	WaitExist = (rsAcademyget("cnt")>0)
    end if
    rsAcademyget.Close


	if (Not WaitExist) then
    	sqlStr = "update [db_academy].[dbo].tbl_lec_item_option " + vbCrlf
    	sqlStr = sqlStr + " set limit_sold=limit_sold - T.cnt" + vbCrlf
    	sqlStr = sqlStr + " from " + vbCrlf
    	sqlStr = sqlStr + " (select d.itemid, d.itemoption, sum(d.itemno) as cnt" + vbCrlf
    	sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_detail d" + vbCrlf
    	sqlStr = sqlStr + " where d.orderserial='" + CStr(orderserial) + "'" + vbCrlf
    	sqlStr = sqlStr + " and d.itemid<>0" + vbCrlf
    	sqlStr = sqlStr + " group by d.itemid, d.itemoption ) as T" + vbCrlf
    	sqlStr = sqlStr + " where [db_academy].[dbo].tbl_lec_item_option.lecidx=T.Itemid"
    	sqlStr = sqlStr + " and [db_academy].[dbo].tbl_lec_item_option.lecoption=T.itemoption"
    	rsAcademyget.Open sqlStr,dbAcademyget,1

    	sqlStr = "update [db_academy].[dbo].tbl_lec_item" + vbCrlf
        sqlStr = sqlStr + " set limit_count=T.limit_count" + vbCrlf
        sqlStr = sqlStr + " ,limit_sold=T.limit_sold" + vbCrlf
        sqlStr = sqlStr + " ,wait_count=T.wait_count" + vbCrlf
        sqlStr = sqlStr + " from (" + vbCrlf
        sqlStr = sqlStr + " 	select o.lecidx, sum(limit_count) as limit_count, sum(limit_sold) as limit_sold" + vbCrlf
        sqlStr = sqlStr + " 	,sum(wait_count) as wait_count" + vbCrlf
        sqlStr = sqlStr + " 	from [db_academy].[dbo].tbl_lec_item_option o" + vbCrlf
        sqlStr = sqlStr + " 		Join (select distinct itemid from [db_academy].[dbo].tbl_academy_order_detail where orderserial='" + CStr(orderserial) + "') A" + vbCrlf
        sqlStr = sqlStr + " 		on o.lecidx=A.itemid" + vbCrlf
        sqlStr = sqlStr + " 	group by o.lecidx" + vbCrlf
        sqlStr = sqlStr + " ) T" + vbCrlf
        sqlStr = sqlStr + " where [db_academy].[dbo].tbl_lec_item.idx=T.lecidx" + vbCrlf

    	rsAcademyget.Open sqlStr,dbAcademyget,1
	end if
	'강좌신청가능인원 증가(대기자가 없을 경우만)
'	if (olecture.FOneItem.FWaitCount = 0) then
'    	sqlStr = " update [db_academy].[dbo].tbl_lec_item "
'    	sqlStr = sqlStr + " set limit_sold = limit_sold - " + CStr(oordermaster.FOneItem.Ftotalitemno) + " "
'    	sqlStr = sqlStr + " where idx = " + CStr(oordermaster.FOneItem.Fitemid) + " "
'    	rsAcademyget.Open sqlStr,dbAcademyget,1
'    end if

    '비회원이 아닌경우(마일리지/쿠폰 처리)
    if (oordermaster.FOneItem.FUserID <> "") then
    	'==============================================================
    	if (oordermaster.FOneItem.Fmiletotalprice <> 0) then
        	'사용마일리지 취소
        	sqlStr = " update [db_user].[dbo].tbl_mileagelog "
        	sqlStr = sqlStr + " set deleteyn='Y' "
        	sqlStr = sqlStr + " where orderserial = '" + CStr(orderserial) + "' "
        	rsget.Open sqlStr,dbget,1
        	response.write "<script>alert('마일리지 사용이 취소되었습니다.');</script>"
        end if

    	'==============================================================
    	if (oordermaster.FOneItem.Ftencardspend <> 0) then
    	    '사용쿠폰 재사용가능하게 전환
        	sqlStr = " update [db_user].[dbo].tbl_user_coupon "
        	sqlStr = sqlStr + " set isusing='N' "
        	sqlStr = sqlStr + " where orderserial = '" + CStr(orderserial) + "' and deleteyn = 'N' and isusing = 'Y' "
        	rsget.Open sqlStr,dbget,1
        	response.write "<script>alert('쿠폰 사용이 취소되었습니다.');</script>"
        end if

        '==============================================================
        '고객 사용마일리지/획득마일리지 재계산
        updateUserMileage oordermaster.FOneItem.FUserID
    end if

    response.write "<script>alert('강좌신청이 취소되었습니다.');</script>"
    if (returnmethod = "bank") then
        insertRepayBank orderserial, id, refundrequire, rebankname, rebankaccount, rebankownername, refundcstitle
        response.write "<script>alert('무통장환불 CS 가 등록되었습니다.'); opener.location.reload(); opener.parent.topframe.location.reload(); opener.focus(); window.close();</script>"
        dbget.close()	:	response.End
    elseif (returnmethod = "creditcard") then
        insertCancelCardRequest orderserial, id, refundrequire, refundcstitle
        'cancelInicisCardPay orderserial
        response.write "<script>alert('카드취소 CS 가 등록되었습니다.'); opener.location.reload(); opener.parent.topframe.location.reload(); opener.focus(); window.close();</script>"
        dbget.close()	:	response.End
    elseif (returnmethod = "realtimetransfer") then
        insertCancelRealTimeTransferRequest orderserial, id, refundrequire, refundcstitle
        response.write "<script>alert('실시간이체취소 CS 가 등록되었습니다.'); opener.location.reload(); opener.parent.topframe.location.reload(); opener.focus(); window.close();</script>"
        dbget.close()	:	response.End
    elseif (returnmethod = "point") then
        insertCancelPointRequest orderserial, id, refundrequire
        response.write "<script>alert('포인트결제 취소요청 CS 가 등록되었습니다.'); opener.location.reload(); opener.parent.topframe.location.reload(); opener.focus(); window.close();</script>"
        dbget.close()	:	response.End
    elseif (returnmethod = "mall") then
        insertCancelMallRequest orderserial, id, refundrequire
        response.write "<script>alert('외부몰결제 취소요청 CS 가 등록되었습니다.'); opener.location.reload(); opener.parent.topframe.location.reload(); opener.focus(); window.close();</script>"
        dbget.close()	:	response.End
    elseif (returnmethod = "allatcard") then
        insertCancelAllAtCardRequest orderserial, id, refundrequire
        response.write "<script>alert('올앳카드결제 취소요청 CS 가 등록되었습니다.'); opener.location.reload(); opener.parent.topframe.location.reload(); opener.focus(); window.close();</script>"
        dbget.close()	:	response.End
    elseif (returnmethod = "ticket") then
        insertCancelTicketRequest orderserial, id, refundrequire
        response.write "<script>alert('상품권결제 취소요청 CS 가 등록되었습니다.'); opener.location.reload(); opener.parent.topframe.location.reload(); opener.focus(); window.close();</script>"
        dbget.close()	:	response.End
    elseif (returnmethod = "mileage") then
        response.write "<script>alert('고객 결재금액이 마일리지 전환되었습니다. 현재 작업중.'); opener.location.reload(); opener.parent.topframe.location.reload(); opener.focus(); window.close();</script>"
        dbget.close()	:	response.End
    else
        if ((oordermaster.FOneItem.Fsubtotalprice > 0) and (oordermaster.FOneItem.Fipkumdiv >= 4)) then
            response.write "<script>alert('환불방식이 선택되지 않았습니다.'); opener.focus(); window.close();</script>"
            dbget.close()	:	response.End
        else
            response.write "<script>opener.location.reload(); opener.parent.topframe.location.reload(); opener.focus(); window.close();</script>"
            dbget.close()	:	response.End
        end if
    end if
end if

if (mode = "cancelitem") then
    '부분취소
    ' - 입금전 취소일 경우, 부분취소AS등록및종료처리, 주문테이블에서 상품수량변경, 주문테이블에서 획득마일리지 재계산, 사용자마일리지 요약 재계산, 사용쿠폰마일리지재계산, 강좌 한정수량 재증가
    ' - 입금후 취소일 경우, 부분취소AS등록및종료처리, 주문테이블에서 상품수량변경, 주문테이블에서 획득마일리지 재계산, 사용자마일리지 요약 재계산, 사용쿠폰마일리지재계산, 강좌 한정수량 재증가, 신용카드취소 또는 무통장입금AS등록페이지로 이동

    '======================================================================
    '부분취소AS등록및종료처리(마스타)
	sqlStr = " select * from [db_cs].[dbo].tbl_new_as_list where 1=0 "
	'response.write sqlStr
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("divcd")          = divcd
	rsget("orderserial")    = orderserial
	rsget("customername")   = html2db(oordermaster.FOneItem.FBuyName)
	rsget("userid")         = html2db(oordermaster.FOneItem.FUserID)
	rsget("writeuser")      = session("ssBctId")
	rsget("title")          = title
	rsget("contents_jupsu") = contents_jupsu
	''rsget("refundrequire")  = 0
	''rsget("cause")          = cause
	''rsget("causedetail")    = causedetail
	rsget("requireupche")   = "N"
	rsget("makerid")        = ""
	rsget("deleteyn")       = "N"
	rsget("songjangno")     = songjangno
	'rsget("rebankname")     = rebankname
	'rsget("rebankaccount")  = rebankaccount
	'rsget("rebankownername")        = rebankownername
	'rsget("refundbeasongpay")       = 0
	'rsget("refunditemcostsum")      = 0
	'rsget("refunddeliverypay")      = 0
	'rsget("refundadjustpay")        = 0
	rsget("sitegubun")      = "FI"

	rsget.update
	id = rsget("id")
	rsget.close




    '======================================================================
    '부분취소AS등록및종료처리(상품목록)
    '신청내역중 해당 수강생취소
    'TODO : 디테일에는 하나의 강좌만 있다고 가정한다.
    dmasterid = id
    dorderserial = orderserial

    contents_finish = ""
    canceldetailno = 0
    canceldetailnamelist = ""
    detailitemlist = split(detailitemlist, "|")
    detailitemnolist = split(detailitemnolist, "|")
	for i = 0 to UBound(detailitemlist)
		if (trim(detailitemlist(i)) <> "") then
            for j = 0 to oorderdetail.FResultCount - 1
                if (CLng(oorderdetail.FItemList(j).Fdetailidx) = CLng(trim(detailitemlist(i)))) then
                    canceldetailno = canceldetailno + 1

                    canceldetailall = True
					sqlStr = " select itemno from " + vbCrlf
				    sqlStr = sqlStr + " db_academy.dbo.tbl_academy_order_master m " + vbCrlf
				    sqlStr = sqlStr + " 	Join db_academy.dbo.tbl_academy_order_detail d" + vbCrlf
				    sqlStr = sqlStr + " 	on m.orderserial=d.orderserial" + vbCrlf
				    sqlStr = sqlStr + " where d.detailidx='" + CStr(detailitemlist(i)) + "'" + vbCrlf
				    rsAcademyget.Open sqlStr,dbAcademyget,1
				    if Not rsAcademyget.Eof then
				    	canceldetailall = (rsAcademyget("itemno") <= CLng(detailitemnolist(i)))
				    end if
				    rsAcademyget.Close


					if (sitename = "academy") then
	                    '강좌
	                    if (canceldetailnamelist = "") then
	                        canceldetailnamelist = oorderdetail.FItemList(j).Fentryname
	                    else
	                        canceldetailnamelist = canceldetailnamelist + "," + oorderdetail.FItemList(j).Fentryname
	                    end if
					else
	                    if (canceldetailnamelist = "") then
	                        canceldetailnamelist = oorderdetail.FItemList(j).FItemName & "[" & CStr(oorderdetail.FItemList(j).Fitemoptionname) & "] " & detailitemnolist(i) & " 개"
	                    else
	                        canceldetailnamelist = canceldetailnamelist + "<br>" + oorderdetail.FItemList(j).FItemName & "[" & CStr(oorderdetail.FItemList(j).Fitemoptionname) & "] " & detailitemnolist(i) & " 개"
	                    end if
					end if

					if (canceldetailall = true) then
	                    sqlStr = " update [db_academy].[dbo].tbl_academy_order_detail set cancelyn = 'Y' "
	                    sqlStr = sqlStr + " where orderserial = '" + CStr(orderserial) + "' and detailidx = " + CStr(trim(detailitemlist(i))) + " "
	                    rsAcademyget.Open sqlStr,dbAcademyget,1
	                    'response.write sqlStr
					else
	                    sqlStr = " update [db_academy].[dbo].tbl_academy_order_detail set itemno = itemno - " & detailitemnolist(i) & " "
	                    sqlStr = sqlStr + " where orderserial = '" + CStr(orderserial) + "' and detailidx = " + CStr(trim(detailitemlist(i))) + " "
	                    rsAcademyget.Open sqlStr,dbAcademyget,1
					end if

                    if (sitename = "academy") and (canceldetailno = 1) then
    			        sqlStr = " insert into [db_cs].[dbo].tbl_new_as_detail(masterid,orderserial,itemid,itemoption,makerid,itemname,itemoptionname,regitemno,confirmitemno,itemcost,isupchebeasong,regdetailstate,gubun01,gubun02) "
    			        sqlStr = sqlStr + " values(" + CStr(dmasterid) + ",'" + CStr(dorderserial) + "'," + CStr(oordermaster.FOneItem.Fitemid) + ",'" + CStr(oordermaster.FOneItem.Fitemoption) + "','" + CStr(oordermaster.FOneItem.Fmakerid) + "','" + html2db(oordermaster.FOneItem.FItemName) + "','" + html2db(oordermaster.FOneItem.FItemoptionName) + "'," + CStr(oordermaster.FOneItem.Ftotalitemno) + ",1," + CStr(oordermaster.FOneItem.Fitemcost) + ",'N','" + CStr(oordermaster.FOneItem.Fipkumdiv) + "','','') "
    			        rsget.Open sqlStr,dbget,1
    			        'response.write sqlStr
    			    end if

					if (sitename <> "academy") then
				        sqlStr = " insert into [db_cs].[dbo].tbl_new_as_detail(masterid,orderserial,itemid,itemoption,makerid,itemname,itemoptionname,regitemno,confirmitemno,itemcost,isupchebeasong,regdetailstate,gubun01,gubun02) "
				        sqlStr = sqlStr + " values(" + CStr(dmasterid) + ",'" + CStr(dorderserial) + "'," + CStr(oorderdetail.FItemList(j).Fitemid) + ",'" + CStr(oorderdetail.FItemList(j).Fitemoption) + "','" + CStr(oorderdetail.FItemList(j).Fmakerid) + "','" + html2db(oorderdetail.FItemList(j).FItemName) + "','" + html2db(oorderdetail.FItemList(j).FItemoptionName) + "'," + CStr(detailitemnolist(i)) + "," & detailitemnolist(i) & "," + CStr(oorderdetail.FItemList(j).Freducedprice) + ",'Y','" + CStr(oorderdetail.FItemList(j).Fcurrstate) + "','','') "
				        rsget.Open sqlStr,dbget,1
				        'response.write sqlStr
			    	end if
                end if
            next
		end if
	next

    if (sitename = "academy") and (canceldetailno > 1) then
        sqlStr = " update [db_cs].[dbo].tbl_new_as_detail set confirmitemno = " + CStr(canceldetailno) + " "
        sqlStr = sqlStr + " where masterid = " + CStr(dmasterid) + " "
        rsget.Open sqlStr,dbget,1
        'response.write sqlStr
    end if

	if (sitename = "academy") then
		sqlStr = " update [db_cs].[dbo].tbl_new_as_list "
		sqlStr = sqlStr + " set finishdate=getdate() "
		sqlStr = sqlStr + " ,finishuser = '" + session("ssBctId") + "' "
		sqlStr = sqlStr + " ,contents_finish = '" + html2db("부분취소(" + CStr(canceldetailno) + " 명[" + CStr(html2db(canceldetailnamelist)) + "]" + ")") + "' "
		sqlStr = sqlStr + " ,currstate = 'B007' "
		'sqlStr = sqlStr + " ,refundresult = 0 "
		sqlStr = sqlStr + " where id=" + CStr(dmasterid) + " "
		rsget.Open sqlStr,dbget,1
		'response.write sqlStr
	else
		sqlStr = " update [db_cs].[dbo].tbl_new_as_list "
		sqlStr = sqlStr + " set finishdate=getdate() "
		sqlStr = sqlStr + " ,finishuser = '" + session("ssBctId") + "' "
		sqlStr = sqlStr + " ,contents_finish = '" + html2db("부분취소(DIY상품)") + "' "
		sqlStr = sqlStr + " ,currstate = 'B007' "
		'sqlStr = sqlStr + " ,refundresult = 0 "
		sqlStr = sqlStr + " where id=" + CStr(dmasterid) + " "
		rsget.Open sqlStr,dbget,1
		'response.write sqlStr
	end if

	'======================================================================
	'강좌신청가능인원 증가
	if (sitename = "academy") and (olecture.FOneItem.FWaitCount = 0) then
    	sqlStr = " update [db_academy].[dbo].tbl_lec_item "
    	sqlStr = sqlStr + " set limit_sold = limit_sold - " + CStr(canceldetailno) + " "
    	sqlStr = sqlStr + " where idx = " + CStr(oordermaster.FOneItem.Fitemid) + " "
    	rsAcademyget.Open sqlStr,dbAcademyget,1
    end if

	'======================================================================
	'오더마스터테이블 정보 업데이트
    recalculateOrderMaster orderserial

	'======================================================================
	'고객 사용마일리지/획득마일리지 재계산
    if (oordermaster.FOneItem.FUserID <> "") then
        updateUserMileage oordermaster.FOneItem.FUserID
    end if


    response.write "<script>alert('부분취소신청이 등록되었습니다.');</script>"
    if (returnmethod = "bank") then
        insertRepayBank orderserial, dmasterid, refundrequire, rebankname, rebankaccount, rebankownername, refundcstitle
        response.write "<script>alert('무통장환불요청CS 가 등록되었습니다.'); opener.location.reload(); opener.parent.topframe.location.reload(); opener.focus(); window.close();</script>"
        dbget.close()	:	response.End
    elseif (returnmethod = "mileage") then
        response.write "<script>alert('고객 결재금액이 마일리지 전환되었습니다. 현재 작업중.'); opener.focus(); window.close();</script>"
        dbget.close()	:	response.End
    else
        if ((oordermaster.FOneItem.Fsubtotalprice > 0) and (oordermaster.FOneItem.Fipkumdiv >= 4)) then
            response.write "<script>alert('환불방식이 선택되지 않았습니다.'); opener.location.reload(); opener.parent.topframe.location.reload(); opener.focus(); window.close();</script>"
            dbget.close()	:	response.End
        else
            response.write "<script>opener.location.reload(); opener.parent.topframe.location.reload(); opener.focus(); window.close();</script>"
            dbget.close()	:	response.End
        end if
    end if
end if

if (mode = "revalidate") then
    '정상전환
    ' - 무통장주문중 주문접수상태이며, 취소된 주문을 정상전환 합니다.

    '======================================================================
    '정상전환AS등록및종료처리(마스타)
	sqlStr = " select * from [db_cs].[dbo].tbl_new_as_list where 1=0 "
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("divcd")          = divcd
	rsget("orderserial")    = orderserial
	rsget("customername")   = html2db(oordermaster.FOneItem.FBuyName)
	rsget("userid")         = html2db(oordermaster.FOneItem.FUserID)
	rsget("writeuser")      = session("ssBctId")
	rsget("title")          = title
	rsget("contents_jupsu") = contents_jupsu
'	rsget("refundrequire")  = 0
	'rsget("cause")          = cause
	'rsget("causedetail")    = causedetail
	rsget("requireupche")   = "N"
	rsget("makerid")        = ""
	rsget("deleteyn")       = "N"
	rsget("songjangno")     = songjangno
'	rsget("rebankname")     = rebankname
'	rsget("rebankaccount")  = rebankaccount
'	rsget("rebankownername")        = rebankownername
'	rsget("refundbeasongpay")       = 0
'	rsget("refunditemcostsum")      = 0
'	rsget("refunddeliverypay")      = 0
'	rsget("refundadjustpay")        = 0
	rsget("sitegubun")      = "FI"

	rsget.update
	id = rsget("id")
	rsget.close

	sqlStr = " update [db_cs].[dbo].tbl_new_as_list "
	sqlStr = sqlStr + " set finishdate=getdate() "
	sqlStr = sqlStr + " ,finishuser = '" + session("ssBctId") + "' "
	sqlStr = sqlStr + " ,contents_finish = '" + html2db("강좌신청정상전환") + "' "
	sqlStr = sqlStr + " ,currstate = 'B007' "
	''sqlStr = sqlStr + " ,refundresult = 0 "
	sqlStr = sqlStr + " where id=" + CStr(id) + " "
	rsget.Open sqlStr,dbget,1

	'======================================================================
	'주문건정상화
	sqlStr = " update [db_academy].[dbo].tbl_academy_order_master "
	sqlStr = sqlStr + " set cancelyn='N' "
	sqlStr = sqlStr + " where orderserial = '" + CStr(orderserial) + "' "
	rsAcademyget.Open sqlStr,dbAcademyget,1

	'======================================================================
	'강좌신청가능인원 감소
	sqlStr = " update [db_academy].[dbo].tbl_lec_item "
	sqlStr = sqlStr + " set limit_sold = limit_sold + " + CStr(oordermaster.FOneItem.Ftotalitemno) + " "
	sqlStr = sqlStr + " where idx = " + CStr(oordermaster.FOneItem.Fitemid) + " "
	rsAcademyget.Open sqlStr,dbAcademyget,1

    '비회원이 아닌경우(마일리지/쿠폰 처리)
    if (oordermaster.FOneItem.FUserID <> "") then
    	'==============================================================
    	if (oordermaster.FOneItem.Fmiletotalprice <> 0) then
        	'사용마일리지 정상화
        	sqlStr = " update [db_user].[dbo].tbl_mileagelog "
        	sqlStr = sqlStr + " set deleteyn='N' "
        	sqlStr = sqlStr + " where orderserial = '" + CStr(orderserial) + "' "
        	rsget.Open sqlStr,dbget,1
        	response.write "<script>alert('마일리지 사용이 정상화되었습니다.');</script>"
        end if

    	'==============================================================
    	if (oordermaster.FOneItem.Ftencardspend <> 0) then
    	    '사용쿠폰 정상화
        	sqlStr = " update [db_user].[dbo].tbl_user_coupon "
        	sqlStr = sqlStr + " set isusing='Y' "
        	sqlStr = sqlStr + " where orderserial = '" + CStr(orderserial) + "' and deleteyn = 'N' and isusing = 'N' "
        	rsget.Open sqlStr,dbget,1
        	response.write "<script>alert('쿠폰 사용이 정상화되었습니다.');</script>"
        end if

        '==============================================================
        '고객 사용마일리지/획득마일리지 재계산
        updateUserMileage oordermaster.FOneItem.FUserID
    end if

    response.write "<script>alert('강좌신청이 정상화되었습니다.');</script>"
    response.write "<script>opener.location.reload(); opener.parent.topframe.location.reload(); opener.focus(); window.close();</script>"
    dbget.close()	:	response.End
end if


if (mode = "receveupche") then
    response.write "관리자 문의 요망"
    dbget.close()	:	response.End

        '업체반품
        '선택된 상품중, 출고완료가 아닌 상품이 있을 경우, 진행정지(부분취소불가)
        '선택된 상품중, 출고완료가 아닌 상품이 없을 경우,
        ' - 선택된 상품목록을 저장하고, 반품/회수 CS 를 접수상태로 저장한다.
        ' - 이후, 반품이 들어왔을때, 접수상태로 저장된 CS 를 완료처리하면서, 마이너스 주문과 무통장환불 CS 등록을 한다. 부여된 마일리지 회수도 마이너스 주문에서 처리한다.(다른 루틴에서 처리한다.)

        '======================================================================
        '업체반품AS등록(마스타)
	sqlStr = " select * from [db_cs].[dbo].tbl_as_list where 1=0 "
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("divcd")          = divcd
	rsget("orderserial")    = orderserial
	rsget("customername")   = html2db(oordermaster.FOneItem.FBuyName)
	rsget("userid")         = oordermaster.FOneItem.FUserID
	rsget("writeuser")      = session("ssBctId")
	rsget("title")          = html2db(title)
	rsget("contents_jupsu") = html2db(contents_jupsu)
	rsget("refundrequire")  = refundrequire
	rsget("cause")          = cause
	rsget("causedetail")    = html2db(causedetail)
	rsget("requireupche")   = "Y"
	rsget("makerid")        = makerid
	rsget("deleteyn")       = "N"
	rsget("songjangno")     = songjangno
	rsget("rebankname")     = rebankname
	rsget("rebankaccount")  = html2db(rebankaccount)
	rsget("rebankownername")        = html2db(rebankownername)
	rsget("refundbeasongpay")       = refundbeasongpay
	rsget("refunditemcostsum")      = refunditemcostsum
	rsget("refunddeliverypay")      = refunddeliverypay
	rsget("refundadjustpay")        = refundadjustpay

	rsget.update
	id = rsget("id")
	rsget.close

	'sqlStr = " update [db_cs].[dbo].tbl_as_list "
	'sqlStr = sqlStr + " set finishdate=getdate() "
	'sqlStr = sqlStr + " ,finishuser = '" + session("ssBctId") + "' "
	'sqlStr = sqlStr + " ,contents_finish = '" + html2db("부분취소(" + oordermaster.FOneItem.JumunMethodName + ")") + "' "
	'sqlStr = sqlStr + " ,currstate = '7' "
	'sqlStr = sqlStr + " ,refundresult = 0 "
	'sqlStr = sqlStr + " where id=" + CStr(id) + " "
	'rsget.Open sqlStr,dbget,1

        '======================================================================
        '업체반품AS등록(상품목록)
        dmasterid = id
        dorderserial = orderserial

        detailitemlist = split(detailitemlist, "|")
	for i = 0 to UBound(detailitemlist)
		if (trim(detailitemlist(i)) <> "") then
			tmp = split(detailitemlist(i), Chr(9))

			did             = tmp(0)
			dcausediv       = tmp(1)
			dcausedetail    = html2db(tmp(2))
			dconfirmitemno  = tmp(3)
			dcausecontent   = html2db(tmp(4))

			j = -1
                        for j = 0 to oorderdetail.FResultCount - 1
                                if (CLng(oorderdetail.FItemList(j).Fidx) = CLng(did)) then
                                        exit for
                                end if
                        next

			if (j <> -1) then
			        if isnull(oorderdetail.FItemList(j).Fcurrstate) then
			                oorderdetail.FItemList(j).Fcurrstate = ""
			        end if

                                '선택된 상품중 출고완료가 아닌 상품이 있는지 체크
                                'if (oorderdetail.FItemList(i).GetStateName <> "출고완료") then
                                '        sqlStr = " update [db_cs].[dbo].tbl_as_list set deleteyn = 'Y' where id = " + CStr(dmasterid) + " "
                                '        rsget.Open sqlStr,dbget,1
                                '
                                '        sqlStr = " delete from [db_cs].[dbo].tbl_as_detail where masterid = " + CStr(dmasterid) + " "
                                '        rsget.Open sqlStr,dbget,1
                                '
                                '        response.write "<script>alert('상품중 출고되지 않은 상품이 있습니다. 등록이 취소됩니다.'); history.back();</script>"
                                '        dbget.close()	:	response.End
                                'end if

			        sqlStr = " insert into [db_cs].[dbo].tbl_as_detail(masterid,orderserial,itemid,itemoption,makerid,itemname,itemoptionname,regitemno,confirmitemno,itemcost,isupchebeasong,regdetailstate,causediv,causedetail,causecontent) "
			        sqlStr = sqlStr + " values(" + CStr(dmasterid) + ",'" + CStr(dorderserial) + "'," + CStr(oorderdetail.FItemList(j).Fitemid) + ",'" + CStr(oorderdetail.FItemList(j).Fitemoption) + "','" + CStr(oorderdetail.FItemList(j).Fmakerid) + "','" + html2db(oorderdetail.FItemList(j).FItemName) + "','" + html2db(oorderdetail.FItemList(j).FItemoptionName) + "'," + CStr(oorderdetail.FItemList(j).Fitemno) + "," + CStr(dconfirmitemno) + "," + CStr(oorderdetail.FItemList(j).Fitemcost) + ",'" + CStr(oorderdetail.FItemList(j).Fisupchebeasong) + "','" + CStr(oorderdetail.FItemList(j).Fcurrstate) + "','" + CStr(dcausediv) + "','" + CStr(dcausedetail) + "','" + CStr(dcausecontent) + "') "
			        rsget.Open sqlStr,dbget,1
			end if
		end if
	next

        response.write "<script>alert('반품접수가 등록되었습니다.'); opener.focus(); window.close();</script>"
        dbget.close()	:	response.End
end if

if (mode = "recevetenten") then
    response.write "관리자 문의 요망"
    dbget.close()	:	response.End
        '회수요청
        '선택된 상품중, 출고완료가 아닌 상품이 있을 경우, 진행정지(부분취소불가)
        '선택된 상품중, 출고완료가 아닌 상품이 없을 경우,
        ' - 선택된 상품목록을 저장하고, 반품/회수 CS 를 접수상태로 저장한다.
        ' - 이후, 반품이 들어왔을때, 접수상태로 저장된 CS 를 완료처리하면서, 마이너스 주문과 무통장환불 CS 등록을 한다. 부여된 마일리지 회수도 마이너스 주문에서 처리한다.(다른 루틴에서 처리한다.)

        '======================================================================
        '회수요청AS등록(마스타)
	sqlStr = " select * from [db_cs].[dbo].tbl_as_list where 1=0 "
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("divcd")          = divcd
	rsget("orderserial")    = orderserial
	rsget("customername")   = html2db(oordermaster.FOneItem.FBuyName)
	rsget("userid")         = oordermaster.FOneItem.FUserID
	rsget("writeuser")      = session("ssBctId")
	rsget("title")          = html2db(title)
	rsget("contents_jupsu") = html2db(contents_jupsu)
	rsget("refundrequire")  = refundrequire
	rsget("cause")          = cause
	rsget("causedetail")    = html2db(causedetail)
	rsget("requireupche")   = "N"
	rsget("makerid")        = ""
	rsget("deleteyn")       = "N"
	rsget("songjangno")     = songjangno
	rsget("rebankname")     = rebankname
	rsget("rebankaccount")  = html2db(rebankaccount)
	rsget("rebankownername")        = html2db(rebankownername)
	rsget("refundbeasongpay")       = refundbeasongpay
	rsget("refunditemcostsum")      = refunditemcostsum
	rsget("refunddeliverypay")      = refunddeliverypay
	rsget("refundadjustpay")        = refundadjustpay

	rsget.update
	id = rsget("id")
	rsget.close

	'sqlStr = " update [db_cs].[dbo].tbl_as_list "
	'sqlStr = sqlStr + " set finishdate=getdate() "
	'sqlStr = sqlStr + " ,finishuser = '" + session("ssBctId") + "' "
	'sqlStr = sqlStr + " ,contents_finish = '" + html2db("부분취소(" + oordermaster.FOneItem.JumunMethodName + ")") + "' "
	'sqlStr = sqlStr + " ,currstate = '7' "
	'sqlStr = sqlStr + " ,refundresult = 0 "
	'sqlStr = sqlStr + " where id=" + CStr(id) + " "
	'rsget.Open sqlStr,dbget,1

        '======================================================================
        '회수요청AS등록(상품목록)
        dmasterid = id
        dorderserial = orderserial

        detailitemlist = split(detailitemlist, "|")
	for i = 0 to UBound(detailitemlist)
		if (trim(detailitemlist(i)) <> "") then
			tmp = split(detailitemlist(i), Chr(9))

			did             = tmp(0)
			dcausediv       = tmp(1)
			dcausedetail    = html2db(tmp(2))
			dconfirmitemno  = tmp(3)
			dcausecontent   = html2db(tmp(4))

			j = -1
                        for j = 0 to oorderdetail.FResultCount - 1
                                if (CLng(oorderdetail.FItemList(j).Fidx) = CLng(did)) then
                                        exit for
                                end if
                        next

			if (j <> -1) then
			        if isnull(oorderdetail.FItemList(j).Fcurrstate) then
			                oorderdetail.FItemList(j).Fcurrstate = ""
			        end if

                                '선택된 상품중 출고완료가 아닌 상품이 있는지 체크
                                'if (oorderdetail.FItemList(i).GetStateName <> "출고완료") then
                                '        sqlStr = " update [db_cs].[dbo].tbl_as_list set deleteyn = 'Y' where id = " + CStr(dmasterid) + " "
                                '        rsget.Open sqlStr,dbget,1
                                '
                                '        sqlStr = " delete from [db_cs].[dbo].tbl_as_detail where masterid = " + CStr(dmasterid) + " "
                                '        rsget.Open sqlStr,dbget,1
                                '
                                '        response.write "<script>alert('상품중 출고되지 않은 상품이 있습니다. 등록이 취소됩니다.'); history.back();</script>"
                                '        dbget.close()	:	response.End
                                'end if

			        sqlStr = " insert into [db_cs].[dbo].tbl_as_detail(masterid,orderserial,itemid,itemoption,makerid,itemname,itemoptionname,regitemno,confirmitemno,itemcost,isupchebeasong,regdetailstate,causediv,causedetail,causecontent) "
			        sqlStr = sqlStr + " values(" + CStr(dmasterid) + ",'" + CStr(dorderserial) + "'," + CStr(oorderdetail.FItemList(j).Fitemid) + ",'" + CStr(oorderdetail.FItemList(j).Fitemoption) + "','" + CStr(oorderdetail.FItemList(j).Fmakerid) + "','" + html2db(oorderdetail.FItemList(j).FItemName) + "','" + html2db(oorderdetail.FItemList(j).FItemoptionName) + "'," + CStr(oorderdetail.FItemList(j).Fitemno) + "," + CStr(dconfirmitemno) + "," + CStr(oorderdetail.FItemList(j).Fitemcost) + ",'" + CStr(oorderdetail.FItemList(j).Fisupchebeasong) + "','" + CStr(oorderdetail.FItemList(j).Fcurrstate) + "','" + CStr(dcausediv) + "','" + CStr(dcausedetail) + "','" + CStr(dcausecontent) + "') "
			        rsget.Open sqlStr,dbget,1
			end if
		end if
	next

        response.write "<script>alert('회수요청이 등록되었습니다.'); opener.focus(); window.close();</script>"
        dbget.close()	:	response.End
end if

if (mode = "change") then
    response.write "관리자 문의 요망"
    dbget.close()	:	response.End
        '맞교환
        '선택된 상품중, 출고완료가 아닌 상품이 있을 경우, 진행정지(부분취소불가)
        '선택된 상품중, 출고완료가 아닌 상품이 없을 경우,
        ' - 선택된 상품목록을 저장하고, 맞교환 CS 를 접수상태로 저장한다.
        ' - 이후, 실제 물건을 보내면서, 송장번호를 입력하고 종료처리한다.

        '======================================================================
        if (makerid = "-") then
                makerid = ""
        end if


        '맞교환AS등록(마스타)
	sqlStr = " select * from [db_cs].[dbo].tbl_as_list where 1=0 "
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("divcd")          = divcd
	rsget("orderserial")    = orderserial
	rsget("customername")   = html2db(oordermaster.FOneItem.FBuyName)
	rsget("userid")         = oordermaster.FOneItem.FUserID
	rsget("writeuser")      = session("ssBctId")
	rsget("title")          = html2db(title)
	rsget("contents_jupsu") = html2db(contents_jupsu)
	rsget("refundrequire")  = 0
	rsget("cause")          = cause
	rsget("causedetail")    = html2db(causedetail)
	if (makerid = "") then
	        rsget("requireupche")   = "N"
	        rsget("makerid")        = ""
	else
	        rsget("requireupche")   = "Y"
	        rsget("makerid")        = makerid
	end if
	rsget("deleteyn")       = "N"
	rsget("songjangno")     = songjangno
	rsget("rebankname")     = ""
	rsget("rebankaccount")  = ""
	rsget("rebankownername")        = ""
	rsget("refundbeasongpay")       = 0
	rsget("refunditemcostsum")      = 0
	rsget("refunddeliverypay")      = 0
	rsget("refundadjustpay")        = 0

	rsget.update
	id = rsget("id")
	rsget.close

	'sqlStr = " update [db_cs].[dbo].tbl_as_list "
	'sqlStr = sqlStr + " set finishdate=getdate() "
	'sqlStr = sqlStr + " ,finishuser = '" + session("ssBctId") + "' "
	'sqlStr = sqlStr + " ,contents_finish = '" + html2db("부분취소(" + oordermaster.FOneItem.JumunMethodName + ")") + "' "
	'sqlStr = sqlStr + " ,currstate = '7' "
	'sqlStr = sqlStr + " ,refundresult = 0 "
	'sqlStr = sqlStr + " where id=" + CStr(id) + " "
	'rsget.Open sqlStr,dbget,1

        '======================================================================
        '맞교환AS등록(상품목록)
        dmasterid = id
        dorderserial = orderserial

        detailitemlist = split(detailitemlist, "|")
	for i = 0 to UBound(detailitemlist)
		if (trim(detailitemlist(i)) <> "") then
			tmp = split(detailitemlist(i), Chr(9))

			did             = tmp(0)
			dcausediv       = tmp(1)
			dcausedetail    = html2db(tmp(2))
			dconfirmitemno  = tmp(3)
			dcausecontent   = html2db(tmp(4))

			j = -1
                        for j = 0 to oorderdetail.FResultCount - 1
                                if (CLng(oorderdetail.FItemList(j).Fidx) = CLng(did)) then
                                        exit for
                                end if
                        next

			if (j <> -1) then
			        if isnull(oorderdetail.FItemList(j).Fcurrstate) then
			                oorderdetail.FItemList(j).Fcurrstate = ""
			        end if

                                '선택된 상품중 출고완료가 아닌 상품이 있는지 체크
                                'if (oorderdetail.FItemList(i).GetStateName <> "출고완료") then
                                '        sqlStr = " update [db_cs].[dbo].tbl_as_list set deleteyn = 'Y' where id = " + CStr(dmasterid) + " "
                                '        rsget.Open sqlStr,dbget,1
                                '
                                '        sqlStr = " delete from [db_cs].[dbo].tbl_as_detail where masterid = " + CStr(dmasterid) + " "
                                '        rsget.Open sqlStr,dbget,1
                                '
                                '        response.write "<script>alert('상품중 출고되지 않은 상품이 있습니다. 등록이 취소됩니다.'); history.back();</script>"
                                '        dbget.close()	:	response.End
                                'end if

			        sqlStr = " insert into [db_cs].[dbo].tbl_as_detail(masterid,orderserial,itemid,itemoption,makerid,itemname,itemoptionname,regitemno,confirmitemno,itemcost,isupchebeasong,regdetailstate,causediv,causedetail,causecontent) "
			        sqlStr = sqlStr + " values(" + CStr(dmasterid) + ",'" + CStr(dorderserial) + "'," + CStr(oorderdetail.FItemList(j).Fitemid) + ",'" + CStr(oorderdetail.FItemList(j).Fitemoption) + "','" + CStr(oorderdetail.FItemList(j).Fmakerid) + "','" + html2db(oorderdetail.FItemList(j).FItemName) + "','" + html2db(oorderdetail.FItemList(j).FItemoptionName) + "'," + CStr(oorderdetail.FItemList(j).Fitemno) + "," + CStr(dconfirmitemno) + "," + CStr(oorderdetail.FItemList(j).Fitemcost) + ",'" + CStr(oorderdetail.FItemList(j).Fisupchebeasong) + "','" + CStr(oorderdetail.FItemList(j).Fcurrstate) + "','" + CStr(dcausediv) + "','" + CStr(dcausedetail) + "','" + CStr(dcausecontent) + "') "
			        rsget.Open sqlStr,dbget,1
			end if
		end if
	next

        response.write "<script>alert('맞교환이 등록되었습니다.'); opener.focus(); window.close();</script>"
        dbget.close()	:	response.End
end if

if (mode = "omit") then
    response.write "관리자 문의 요망"
    dbget.close()	:	response.End
        '누락재발송
        '선택된 상품중, 출고완료가 아닌 상품이 있을 경우, 진행정지(부분취소불가)
        '선택된 상품중, 출고완료가 아닌 상품이 없을 경우,
        ' - 선택된 상품목록을 저장하고, 누락재발송 CS 를 접수상태로 저장한다.
        ' - 이후, 실제 물건을 보내면서, 송장번호를 입력하고 종료처리한다.

        '======================================================================
        if (makerid = "-") then
                makerid = ""
        end if


        '누락재발송AS등록(마스타)
	sqlStr = " select * from [db_cs].[dbo].tbl_as_list where 1=0 "
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("divcd")          = divcd
	rsget("orderserial")    = orderserial
	rsget("customername")   = html2db(oordermaster.FOneItem.FBuyName)
	rsget("userid")         = oordermaster.FOneItem.FUserID
	rsget("writeuser")      = session("ssBctId")
	rsget("title")          = html2db(title)
	rsget("contents_jupsu") = html2db(contents_jupsu)
	rsget("refundrequire")  = 0
	rsget("cause")          = cause
	rsget("causedetail")    = html2db(causedetail)
	if (makerid = "") then
	        rsget("requireupche")   = "N"
	        rsget("makerid")        = ""
	else
	        rsget("requireupche")   = "Y"
	        rsget("makerid")        = makerid
	end if
	rsget("deleteyn")       = "N"
	rsget("songjangno")     = songjangno
	rsget("rebankname")     = ""
	rsget("rebankaccount")  = ""
	rsget("rebankownername")        = ""
	rsget("refundbeasongpay")       = 0
	rsget("refunditemcostsum")      = 0
	rsget("refunddeliverypay")      = 0
	rsget("refundadjustpay")        = 0

	rsget.update
	id = rsget("id")
	rsget.close

	'sqlStr = " update [db_cs].[dbo].tbl_as_list "
	'sqlStr = sqlStr + " set finishdate=getdate() "
	'sqlStr = sqlStr + " ,finishuser = '" + session("ssBctId") + "' "
	'sqlStr = sqlStr + " ,contents_finish = '" + html2db("부분취소(" + oordermaster.FOneItem.JumunMethodName + ")") + "' "
	'sqlStr = sqlStr + " ,currstate = '7' "
	'sqlStr = sqlStr + " ,refundresult = 0 "
	'sqlStr = sqlStr + " where id=" + CStr(id) + " "
	'rsget.Open sqlStr,dbget,1

        '======================================================================
        '누락재발송AS등록(상품목록)
        dmasterid = id
        dorderserial = orderserial

        detailitemlist = split(detailitemlist, "|")
	for i = 0 to UBound(detailitemlist)
		if (trim(detailitemlist(i)) <> "") then
			tmp = split(detailitemlist(i), Chr(9))

			did             = tmp(0)
			dcausediv       = tmp(1)
			dcausedetail    = html2db(tmp(2))
			dconfirmitemno  = tmp(3)
			dcausecontent   = html2db(tmp(4))

			j = -1
                        for j = 0 to oorderdetail.FResultCount - 1
                                if (CLng(oorderdetail.FItemList(j).Fidx) = CLng(did)) then
                                        exit for
                                end if
                        next

			if (j <> -1) then
			        if isnull(oorderdetail.FItemList(j).Fcurrstate) then
			                oorderdetail.FItemList(j).Fcurrstate = ""
			        end if

                                '선택된 상품중 출고완료가 아닌 상품이 있는지 체크
                                'if (oorderdetail.FItemList(i).GetStateName <> "출고완료") then
                                '        sqlStr = " update [db_cs].[dbo].tbl_as_list set deleteyn = 'Y' where id = " + CStr(dmasterid) + " "
                                '        rsget.Open sqlStr,dbget,1
                                '
                                '        sqlStr = " delete from [db_cs].[dbo].tbl_as_detail where masterid = " + CStr(dmasterid) + " "
                                '        rsget.Open sqlStr,dbget,1
                                '
                                '        response.write "<script>alert('상품중 출고되지 않은 상품이 있습니다. 등록이 취소됩니다.'); history.back();</script>"
                                '        dbget.close()	:	response.End
                                'end if

			        sqlStr = " insert into [db_cs].[dbo].tbl_as_detail(masterid,orderserial,itemid,itemoption,makerid,itemname,itemoptionname,regitemno,confirmitemno,itemcost,isupchebeasong,regdetailstate,causediv,causedetail,causecontent) "
			        sqlStr = sqlStr + " values(" + CStr(dmasterid) + ",'" + CStr(dorderserial) + "'," + CStr(oorderdetail.FItemList(j).Fitemid) + ",'" + CStr(oorderdetail.FItemList(j).Fitemoption) + "','" + CStr(oorderdetail.FItemList(j).Fmakerid) + "','" + html2db(oorderdetail.FItemList(j).FItemName) + "','" + html2db(oorderdetail.FItemList(j).FItemoptionName) + "'," + CStr(oorderdetail.FItemList(j).Fitemno) + "," + CStr(dconfirmitemno) + "," + CStr(oorderdetail.FItemList(j).Fitemcost) + ",'" + CStr(oorderdetail.FItemList(j).Fisupchebeasong) + "','" + CStr(oorderdetail.FItemList(j).Fcurrstate) + "','" + CStr(dcausediv) + "','" + CStr(dcausedetail) + "','" + CStr(dcausecontent) + "') "
			        rsget.Open sqlStr,dbget,1
			end if
		end if
	next

        response.write "<script>alert('누락재발송이 등록되었습니다.'); opener.focus(); window.close();</script>"
        dbget.close()	:	response.End
end if

if (mode = "more") then
    response.write "관리자 문의 요망"
    dbget.close()	:	response.End
        '서비스발송
        '선택된 상품중, 출고완료가 아닌 상품이 있을 경우, 진행정지(부분취소불가)
        '선택된 상품중, 출고완료가 아닌 상품이 없을 경우,
        ' - 선택된 상품목록을 저장하고, 서비스발송 CS 를 접수상태로 저장한다.
        ' - 이후, 실제 물건을 보내면서, 송장번호를 입력하고 종료처리한다.

        '======================================================================
        if (makerid = "-") then
                makerid = ""
        end if


        '서비스발송AS등록(마스타)
	sqlStr = " select * from [db_cs].[dbo].tbl_as_list where 1=0 "
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("divcd")          = divcd
	rsget("orderserial")    = orderserial
	rsget("customername")   = html2db(oordermaster.FOneItem.FBuyName)
	rsget("userid")         = oordermaster.FOneItem.FUserID
	rsget("writeuser")      = session("ssBctId")
	rsget("title")          = html2db(title)
	rsget("contents_jupsu") = html2db(contents_jupsu)
	rsget("refundrequire")  = 0
	rsget("cause")          = cause
	rsget("causedetail")    = html2db(causedetail)
	if (makerid = "") then
	        rsget("requireupche")   = "N"
	        rsget("makerid")        = ""
	else
	        rsget("requireupche")   = "Y"
	        rsget("makerid")        = makerid
	end if
	rsget("deleteyn")       = "N"
	rsget("songjangno")     = songjangno
	rsget("rebankname")     = ""
	rsget("rebankaccount")  = ""
	rsget("rebankownername")        = ""
	rsget("refundbeasongpay")       = 0
	rsget("refunditemcostsum")      = 0
	rsget("refunddeliverypay")      = 0
	rsget("refundadjustpay")        = 0

	rsget.update
	id = rsget("id")
	rsget.close

	'sqlStr = " update [db_cs].[dbo].tbl_as_list "
	'sqlStr = sqlStr + " set finishdate=getdate() "
	'sqlStr = sqlStr + " ,finishuser = '" + session("ssBctId") + "' "
	'sqlStr = sqlStr + " ,contents_finish = '" + html2db("부분취소(" + oordermaster.FOneItem.JumunMethodName + ")") + "' "
	'sqlStr = sqlStr + " ,currstate = '7' "
	'sqlStr = sqlStr + " ,refundresult = 0 "
	'sqlStr = sqlStr + " where id=" + CStr(id) + " "
	'rsget.Open sqlStr,dbget,1

        '======================================================================
        '서비스발송AS등록(상품목록)
        dmasterid = id
        dorderserial = orderserial

        detailitemlist = split(detailitemlist, "|")
	for i = 0 to UBound(detailitemlist)
		if (trim(detailitemlist(i)) <> "") then
			tmp = split(detailitemlist(i), Chr(9))

			did             = tmp(0)
			dcausediv       = tmp(1)
			dcausedetail    = html2db(tmp(2))
			dconfirmitemno  = tmp(3)
			dcausecontent   = html2db(tmp(4))

			j = -1
                        for j = 0 to oorderdetail.FResultCount - 1
                                if (CLng(oorderdetail.FItemList(j).Fidx) = CLng(did)) then
                                        exit for
                                end if
                        next

			if (j <> -1) then
			        if isnull(oorderdetail.FItemList(j).Fcurrstate) then
			                oorderdetail.FItemList(j).Fcurrstate = ""
			        end if

                                '선택된 상품중 출고완료가 아닌 상품이 있는지 체크
                                'if (oorderdetail.FItemList(i).GetStateName <> "출고완료") then
                                '        sqlStr = " update [db_cs].[dbo].tbl_as_list set deleteyn = 'Y' where id = " + CStr(dmasterid) + " "
                                '        rsget.Open sqlStr,dbget,1
                                '
                                '        sqlStr = " delete from [db_cs].[dbo].tbl_as_detail where masterid = " + CStr(dmasterid) + " "
                                '        rsget.Open sqlStr,dbget,1
                                '
                                '        response.write "<script>alert('상품중 출고되지 않은 상품이 있습니다. 등록이 취소됩니다.'); history.back();</script>"
                                '        dbget.close()	:	response.End
                                'end if

			        sqlStr = " insert into [db_cs].[dbo].tbl_as_detail(masterid,orderserial,itemid,itemoption,makerid,itemname,itemoptionname,regitemno,confirmitemno,itemcost,isupchebeasong,regdetailstate,causediv,causedetail,causecontent) "
			        sqlStr = sqlStr + " values(" + CStr(dmasterid) + ",'" + CStr(dorderserial) + "'," + CStr(oorderdetail.FItemList(j).Fitemid) + ",'" + CStr(oorderdetail.FItemList(j).Fitemoption) + "','" + CStr(oorderdetail.FItemList(j).Fmakerid) + "','" + html2db(oorderdetail.FItemList(j).FItemName) + "','" + html2db(oorderdetail.FItemList(j).FItemoptionName) + "'," + CStr(oorderdetail.FItemList(j).Fitemno) + "," + CStr(dconfirmitemno) + "," + CStr(oorderdetail.FItemList(j).Fitemcost) + ",'" + CStr(oorderdetail.FItemList(j).Fisupchebeasong) + "','" + CStr(oorderdetail.FItemList(j).Fcurrstate) + "','" + CStr(dcausediv) + "','" + CStr(dcausedetail) + "','" + CStr(dcausecontent) + "') "
			        rsget.Open sqlStr,dbget,1
			end if
		end if
	next

        response.write "<script>alert('서비스발송이 등록되었습니다.'); opener.focus(); window.close();</script>"
        dbget.close()	:	response.End
end if

if (mode = "cancelcard") then
        '신용카드/상품권/실시간이체취소요청

        '======================================================================
        '서비스발송AS등록(마스타)
	sqlStr = " select * from [db_cs].[dbo].tbl_new_as_list where 1=0 "
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("divcd")          = divcd
	rsget("orderserial")    = orderserial
	rsget("customername")   = html2db(oordermaster.FOneItem.FBuyName)
	rsget("userid")         = oordermaster.FOneItem.FUserID
	rsget("writeuser")      = session("ssBctId")
	rsget("title")          = html2db(title)
	rsget("contents_jupsu") = html2db(contents_jupsu)
	'rsget("refundrequire")  = refundrequire
	'rsget("cause")          = cause
	'rsget("causedetail")    = html2db(causedetail)
        rsget("requireupche")   = "N"
        rsget("makerid")        = ""
	rsget("deleteyn")       = "N"
	rsget("songjangno")     = songjangno
	'rsget("rebankname")     = ""
	'rsget("rebankaccount")  = ""
	'rsget("rebankownername")        = ""
	'rsget("refundbeasongpay")       = 0
	'rsget("refunditemcostsum")      = 0
	'rsget("refunddeliverypay")      = 0
	'rsget("refundadjustpay")        = 0

	rsget.update
	id = rsget("id")
	rsget.close

        response.write "<script>alert('신용카드/상품권/실시간이체취소 요청이 등록되었습니다.'); opener.focus(); window.close();</script>"
        dbget.close()	:	response.End
end if

if (mode = "cancelbank") then
        '환불요청

        '======================================================================
        '서비스발송AS등록(마스타)
	sqlStr = " select * from [db_cs].[dbo].tbl_new_as_list where 1=0 "
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("divcd")          = divcd
	rsget("orderserial")    = orderserial
	rsget("customername")   = html2db(oordermaster.FOneItem.FBuyName)
	rsget("userid")         = oordermaster.FOneItem.FUserID
	rsget("writeuser")      = session("ssBctId")
	rsget("title")          = html2db(title)
	rsget("contents_jupsu") = html2db(contents_jupsu)
	'rsget("refundrequire")  = refundrequire
	'rsget("cause")          = cause
	'rsget("causedetail")    = html2db(causedetail)
        rsget("requireupche")   = "N"
        rsget("makerid")        = ""
	rsget("deleteyn")       = "N"
	rsget("songjangno")     = songjangno
	'rsget("rebankname")     = ""
	'rsget("rebankaccount")  = ""
	'rsget("rebankownername")        = ""
	'rsget("refundbeasongpay")       = 0
	'rsget("refunditemcostsum")      = 0
	'rsget("refunddeliverypay")      = 0
	'rsget("refundadjustpay")        = 0

	rsget.update
	id = rsget("id")
	rsget.close

        response.write "<script>alert('환불요청이 등록되었습니다.'); opener.focus(); window.close();</script>"
        dbget.close()	:	response.End
end if

if (mode = "cancelothersite") then
    response.write "관리자 문의 요망"
    dbget.close()	:	response.End
        '외부몰취소요청

        '======================================================================
        '서비스발송AS등록(마스타)
	sqlStr = " select * from [db_cs].[dbo].tbl_as_list where 1=0 "
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("divcd")          = divcd
	rsget("orderserial")    = orderserial
	rsget("customername")   = html2db(oordermaster.FOneItem.FBuyName)
	rsget("userid")         = oordermaster.FOneItem.FUserID
	rsget("writeuser")      = session("ssBctId")
	rsget("title")          = html2db(title)
	rsget("contents_jupsu") = html2db(contents_jupsu)
	rsget("refundrequire")  = refundrequire
	rsget("cause")          = cause
	rsget("causedetail")    = html2db(causedetail)
        rsget("requireupche")   = "N"
        rsget("makerid")        = ""
	rsget("deleteyn")       = "N"
	rsget("songjangno")     = songjangno
	rsget("rebankname")     = ""
	rsget("rebankaccount")  = ""
	rsget("rebankownername")        = ""
	rsget("refundbeasongpay")       = 0
	rsget("refunditemcostsum")      = 0
	rsget("refunddeliverypay")      = 0
	rsget("refundadjustpay")        = 0

	rsget.update
	id = rsget("id")
	rsget.close

        response.write "<script>alert('외부몰취소요청이 등록되었습니다.'); opener.focus(); window.close();</script>"
        dbget.close()	:	response.End
end if

if (mode = "writereadme") then
    response.write "관리자 문의 요망"
    dbget.close()	:	response.End
        '배송유의사항

        '======================================================================
        '서비스발송AS등록(마스타)
	sqlStr = " select * from [db_cs].[dbo].tbl_as_list where 1=0 "
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("divcd")          = divcd
	rsget("orderserial")    = orderserial
	rsget("customername")   = html2db(oordermaster.FOneItem.FBuyName)
	rsget("userid")         = oordermaster.FOneItem.FUserID
	rsget("writeuser")      = session("ssBctId")
	rsget("title")          = html2db(title)
	rsget("contents_jupsu") = html2db(contents_jupsu)
	rsget("refundrequire")  = refundrequire
	rsget("cause")          = cause
	rsget("causedetail")    = html2db(causedetail)
	if (makerid = "") then
	        rsget("requireupche")   = "N"
	        rsget("makerid")        = ""
	else
	        rsget("requireupche")   = "Y"
	        rsget("makerid")        = makerid
	end if
	rsget("deleteyn")       = "N"
	rsget("songjangno")     = songjangno
	rsget("rebankname")     = ""
	rsget("rebankaccount")  = ""
	rsget("rebankownername")        = ""
	rsget("refundbeasongpay")       = 0
	rsget("refunditemcostsum")      = 0
	rsget("refunddeliverypay")      = 0
	rsget("refundadjustpay")        = 0

	rsget.update
	id = rsget("id")
	rsget.close

        response.write "<script>alert('배송유의사항이 등록되었습니다.'); opener.focus(); window.close();</script>"
        dbget.close()	:	response.End
end if

if (mode = "writeetcnote") then
    response.write "관리자 문의 요망"
    dbget.close()	:	response.End
        '기타내역

        '======================================================================
        '서비스발송AS등록(마스타)
	sqlStr = " select * from [db_cs].[dbo].tbl_as_list where 1=0 "
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("divcd")          = divcd
	rsget("orderserial")    = orderserial
	rsget("customername")   = html2db(oordermaster.FOneItem.FBuyName)
	rsget("userid")         = oordermaster.FOneItem.FUserID
	rsget("writeuser")      = session("ssBctId")
	rsget("title")          = html2db(title)
	rsget("contents_jupsu") = html2db(contents_jupsu)
	rsget("refundrequire")  = refundrequire
	rsget("cause")          = cause
	rsget("causedetail")    = html2db(causedetail)
	if (makerid = "") then
	        rsget("requireupche")   = "N"
	        rsget("makerid")        = ""
	else
	        rsget("requireupche")   = "Y"
	        rsget("makerid")        = makerid
	end if
	rsget("deleteyn")       = "N"
	rsget("songjangno")     = songjangno
	rsget("rebankname")     = ""
	rsget("rebankaccount")  = ""
	rsget("rebankownername")        = ""
	rsget("refundbeasongpay")       = 0
	rsget("refunditemcostsum")      = 0
	rsget("refunddeliverypay")      = 0
	rsget("refundadjustpay")        = 0

	rsget.update
	id = rsget("id")
	rsget.close

        response.write "<script>alert('기타내역이 등록되었습니다.'); opener.focus(); window.close();</script>"
        dbget.close()	:	response.End
end if



























sub recalculateOrderMaster(byVal orderserial)
	dim sqlStr
	dim jumundiv, discountrate, linkorderserial, miletotalprice, tencardspend, spendmembership, userid, sitename, ipkumdiv
	dim itemcostsum, itemvatsum, itemmileagesum, deliverpay, minusitemcostsum
	dim discountitemcostsum, discountitemvatsum, discountminusitemcostsum
	dim isallreturn, hasreturn, notreturnsubtotal
	dim subtotal, totalsum, totalitemno, cancelitemno, cancelprice

	sqlStr = " select * from [db_academy].[dbo].tbl_academy_order_master where orderserial = '" + CStr(orderserial) + "' "
    rsAcademyget.Open sqlStr,dbAcademyget,1

    if Not rsAcademyget.Eof then
            jumundiv = rsAcademyget("jumundiv")
            discountrate = rsAcademyget("discountrate")
            linkorderserial = rsAcademyget("linkorderserial")
            miletotalprice = rsAcademyget("miletotalprice")
            tencardspend = rsAcademyget("tencardspend")
            spendmembership = rsAcademyget("spendmembership")

            userid = rsAcademyget("userid")
            sitename = rsAcademyget("sitename")
            ipkumdiv = rsAcademyget("ipkumdiv")
    else
            jumundiv = "0"
            discountrate = 1.0
            linkorderserial = ""
            miletotalprice = 0
            tencardspend = 0
            spendmembership = 0

            userid = ""
            sitename = ""
            ipkumdiv = "0"
    end if
    rsAcademyget.close

	'서브합계 구하기
    sqlStr = "          select   sum((case when cancelyn = 'Y' then 0 else itemcost end) * itemno) as itemcostsum "
    sqlStr = sqlStr + "         ,sum((case when cancelyn = 'Y' then 0 else mileage end) * itemno) as itemmileagesum "
    sqlStr = sqlStr + "         ,sum((case when ((cancelyn <> 'Y') and (itemno < 0)) then itemcost else 0 end) * itemno) as minusitemcostsum "
    sqlStr = sqlStr + "         ,sum((case when cancelyn = 'Y' then 0 else round((" + CStr(discountrate) + " * itemcost), 2) end) * itemno) as discountitemcostsum "
    sqlStr = sqlStr + "         ,sum((case when ((cancelyn <> 'Y') and (itemno < 0)) then round((" + CStr(discountrate) + " * itemcost), 2) else 0 end) * itemno) as discountminusitemcostsum "
    sqlStr = sqlStr + "         ,sum(case when cancelyn <> 'Y' then itemno else 0 end) as totalitemno "
    sqlStr = sqlStr + "         ,sum(case when cancelyn = 'Y' then itemno else 0 end) as cancelitemno "
    sqlStr = sqlStr + "         ,sum((case when cancelyn <> 'Y' then 0 else itemcost end) * itemno) as cancelprice "
    sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_detail "
    sqlStr = sqlStr + " where orderserial = '" + CStr(orderserial) + "' "
    rsAcademyget.Open sqlStr,dbAcademyget,1
    'response.write sqlStr

    if Not rsAcademyget.Eof then
        itemcostsum = rsAcademyget("itemcostsum")
        itemmileagesum = rsAcademyget("itemmileagesum")
        deliverpay = 0
        minusitemcostsum = rsAcademyget("minusitemcostsum")

        discountitemcostsum = rsAcademyget("discountitemcostsum")
        discountitemvatsum = 0
        discountminusitemcostsum = rsAcademyget("discountminusitemcostsum")

        totalitemno = rsAcademyget("totalitemno")
        cancelitemno = rsAcademyget("cancelitemno")
        cancelprice = rsAcademyget("cancelprice")
    else
        itemcostsum = 0
        itemmileagesum = 0
        deliverpay = 0
        minusitemcostsum = 0

        discountitemcostsum = 0
        discountitemvatsum = 0
        discountminusitemcostsum = 0

        totalitemno = 0
        cancelitemno = 0
        cancelprice = 0
    end if
    rsAcademyget.close

    '전체반품/부분반품 확인
    if (linkorderserial<>"") and (jumundiv="9") then
        if (discountminusitemcostsum < 0) then
                hasreturn = "Y"
        end if

        if (discountitemcostsum = discountminusitemcostsum) then
                isallreturn = "Y"
        end if

        notreturnsubtotal = discountitemcostsum - discountminusitemcostsum
    end if

    subtotal = discountitemcostsum + deliverpay

	if (jumundiv<>"9") then
		subtotal = subtotal - miletotalprice - tencardspend - spendmembership
	else
		if (isallreturn = "Y") or (Abs(miletotalprice + tencardspend + spendmembership) > Abs(notreturnsubtotal)) then
			'전체반품인경우
			'부분반품인경우 : (원구매금액-반품금액)이 (쿠폰+마일리지사용)금액보다 작은경우
			subtotal = subtotal + miletotalprice + tencardspend + spendmembership
		end if
	end if

    totalsum = itemcostsum + deliverpay

    sqlStr = "update [db_academy].[dbo].tbl_academy_order_master set " + vbCrlf
	'sqlStr = sqlStr & " totalvat=" & itemvatsum & "," + vbCrlf
	sqlStr = sqlStr & " totalitemno=" & totalitemno & "," + vbCrlf
	sqlStr = sqlStr & " cancelitemno=" & cancelitemno & "," + vbCrlf
	sqlStr = sqlStr & " cancelprice=" & cancelprice & "," + vbCrlf
	'sqlStr = sqlStr & " totalcost=" & totalsum & "," + vbCrlf
	sqlStr = sqlStr & " totalsum=" & totalsum & "," + vbCrlf
	sqlStr = sqlStr & " totalmileage=" & itemmileagesum & "," + vbCrlf
	sqlStr = sqlStr & " subtotalprice=" & subtotal & vbCrlf
	sqlStr = sqlStr & " where orderserial='" + CStr(orderserial) + "' "
	'response.write sqlStr
	rsAcademyget.Open sqlStr,dbAcademyget,1


	'if (userid<>"") and ((sitename="10x10") or (sitename="way2way")) and (CInt(ipkumdiv)>3) then
	'	sqlStr = "update [db_user].[dbo].tbl_user_current_mileage" + vbCrlf
	'	sqlStr = sqlStr + " set [db_user].[dbo].tbl_user_current_mileage.jumunmileage=T.totmile" + vbCrlf
	'	sqlStr = sqlStr + " from " + vbCrlf
	'	sqlStr = sqlStr + " (select sum(totalmileage) as totmile" + vbCrlf
	'	sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_master" + vbCrlf
	'	sqlStr = sqlStr + " where userid='" + userid + "'" + vbCrlf
	'	sqlStr = sqlStr + " and sitename in ('10x10','way2way')" + vbCrlf
	'	sqlStr = sqlStr + " and cancelyn='N'" + vbCrlf
	'	sqlStr = sqlStr + " and ipkumdiv>3" + vbCrlf
	'	sqlStr = sqlStr + " ) as T" + vbCrlf
	'	sqlStr = sqlStr + " where [db_user].[dbo].tbl_user_current_mileage.userid='" + userid + "'"
	'	rsget.Open sqlStr,dbget,1
	'end if
end sub

sub updateUserMileage(byVal userid)
	dim sqlStr

	'// 보너스/사용마일리지 요약 재계산(신규Proc)
	sqlStr = " exec [db_user].[dbo].sp_Ten_ReCalcu_His_BonusMileage '"&userid&"'"
	dbget.Execute sqlStr

	dim totmile
	sqlStr = " select IsNULL(sum(totalmileage),0) as totmile"
    sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_master"
    sqlStr = sqlStr + " where userid='" + CStr(userid) + "' "
    sqlStr = sqlStr + " and cancelyn='N'"
    sqlStr = sqlStr + " and ipkumdiv>3"
    rsAcademyget.Open sqlStr,dbAcademyget,1
    if Not rsAcademyget.Eof then
    	totmile = rsAcademyget("totmile")
    else
    	totmile = 0
    end if
    rsAcademyget.Close


	'==============================================================
	'주문마일리지 요약 재계산([db_academy].[dbo].tbl_academy_order_master)
    sqlStr = "update [db_user].[dbo].tbl_user_current_mileage"
    sqlStr = sqlStr + " set academymileage=" + CStr(totmile) + ""
    sqlStr = sqlStr + " where userid='" + CStr(userid) + "' "
    rsget.Open sqlStr,dbget,1
end sub

sub insertRepayBank(byVal orderserial, byVal basecsid, byVal refundrequire, rebankname, rebankaccount, rebankownername, refundcstitle)
    dim sqlStr
    dim cause, causedetail
    dim buyname, userid
    dim id
    dim orgsubtotalprice

'    sqlStr = " select top 1 * from [db_cs].[dbo].tbl_new_as_list where id = " + CStr(basecsid) + " "
'    rsget.Open sqlStr,dbget,1
'
'    if Not rsget.Eof then
'            rebankname = db2html(rsget("rebankname"))
'            rebankaccount = db2html(rsget("rebankaccount"))
'            rebankownername = db2html(rsget("rebankownername"))
'
'            cause = ""
'            causedetail = ""
'    else
'            rebankname = ""
'            rebankaccount = ""
'            rebankownername = ""
'            refundrequire = 0
'
'            cause = ""
'            causedetail = ""
'    end if
'    rsget.close

    sqlStr = " select top 1 * from [db_academy].[dbo].tbl_academy_order_master where orderserial = '" + CStr(orderserial) + "' "
    rsAcademyget.Open sqlStr,dbAcademyget,1

    if Not rsAcademyget.Eof then
            buyname = db2html(rsAcademyget("buyname"))
            userid = rsAcademyget("userid")
            orgsubtotalprice = rsAcademyget("subtotalprice")
    else
            buyname = ""
            userid = ""
            orgsubtotalprice = "0"
    end if
    rsAcademyget.close


	sqlStr = " select * from [db_cs].[dbo].tbl_new_as_list where 1=0 "
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("divcd")          = "A003"
	rsget("orderserial")    = orderserial
	rsget("customername")   = html2db(buyname)
	rsget("userid")         = userid
	rsget("writeuser")      = session("ssBctId")
	rsget("title")          = html2db(refundcstitle)
	rsget("contents_jupsu") = html2db("계좌번호 : " + rebankaccount + " / 은행 : " + rebankname + " / 예금주 : " + rebankownername + " ")
	'rsget("refundrequire")  = refundrequire
	'rsget("cause")          = cause
	'rsget("causedetail")    = html2db(causedetail)
	'rsget("requireupche")   = "N"
	'rsget("makerid")        = ""
	rsget("deleteyn")       = "N"
	rsget("songjangno")     = ""
	'rsget("rebankname")     = html2db(rebankname)
	'rsget("rebankaccount")  = html2db(rebankaccount)
	'rsget("rebankownername")        = html2db(rebankownername)

	rsget.update
	id = rsget("id")
	rsget.close


	sqlStr = "insert into [db_cs].[dbo].tbl_as_refund_info"
	sqlStr = sqlStr + " (asid,returnmethod,refundrequire, refundresult, orgsubtotalprice"
	'sqlStr = sqlStr + " ,orgitemcostsum, orgbeasongpay, orgmileagesum, orgcouponsum, orgallatdiscountsum"
    'sqlStr = sqlStr + " ,canceltotal, refunditemcostsum, refundmileagesum, refundcouponsum, allatsubtractsum"
    'sqlStr = sqlStr + " ,refundbeasongpay, refunddeliverypay, refundadjustpay,"
    sqlStr = sqlStr + " ,rebankname, rebankaccount, rebankownername"
    sqlStr = sqlStr + " )"
    sqlStr = sqlStr + " values(" & id
    sqlStr = sqlStr + " ,'" & "R007" & "'"
    sqlStr = sqlStr + " ," & refundrequire
    sqlStr = sqlStr + " ," & "0"
    sqlStr = sqlStr + " ," & orgsubtotalprice
    sqlStr = sqlStr + " ,'" & rebankname &"'"
    sqlStr = sqlStr + " ,'" & rebankaccount &"'"
    sqlStr = sqlStr + " ,'" & rebankownername &"'"
    sqlStr = sqlStr + " )"

    dbget.execute sqlStr
end sub

'신용카드취소
sub insertCancelCardRequest(byVal orderserial, byVal basecsid, byVal refundrequire, refundcstitle)
    dim sqlStr
    dim cause, causedetail
    dim buyname, userid, paygatetid
    dim id
    dim orgsubtotalprice

    cause = ""
    causedetail = ""

    sqlStr = " select top 1 * from [db_academy].[dbo].tbl_academy_order_master where orderserial = '" + CStr(orderserial) + "' "
    rsAcademyget.Open sqlStr,dbAcademyget,1

    if Not rsAcademyget.Eof then
            buyname = db2html(rsAcademyget("buyname"))
            userid = rsAcademyget("userid")
            paygatetid = db2html(rsAcademyget("paygatetid"))
            orgsubtotalprice = rsAcademyget("subtotalprice")
    else
            buyname = ""
            userid = ""
            paygatetid = ""
            orgsubtotalprice = 0
    end if
    rsAcademyget.close


	sqlStr = " select * from [db_cs].[dbo].tbl_new_as_list where 1=0 "
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("divcd")          = "A007"
	rsget("orderserial")    = orderserial
	rsget("customername")   = html2db(buyname)
	rsget("userid")         = userid
	rsget("writeuser")      = session("ssBctId")
	rsget("title")          = html2db(refundcstitle + " - 카드취소")
	rsget("contents_jupsu") = html2db("TID[ " + paygatetid + " ]")
	'rsget("refundrequire")  = refundrequire
	'rsget("cause")          = cause
	'rsget("causedetail")    = html2db(causedetail)
	rsget("requireupche")   = "N"
	rsget("makerid")        = ""
	rsget("deleteyn")       = "N"
	rsget("songjangno")     = ""
	'rsget("rebankname")     = ""
	'rsget("rebankaccount")  = ""
	'rsget("rebankownername")        = ""

	rsget.update
	id = rsget("id")
	rsget.close


	sqlStr = "insert into [db_cs].[dbo].tbl_as_refund_info"
	sqlStr = sqlStr + " (asid,returnmethod,refundrequire, refundresult, orgsubtotalprice"
	'sqlStr = sqlStr + " ,orgitemcostsum, orgbeasongpay, orgmileagesum, orgcouponsum, orgallatdiscountsum"
    'sqlStr = sqlStr + " ,canceltotal, refunditemcostsum, refundmileagesum, refundcouponsum, allatsubtractsum"
    'sqlStr = sqlStr + " ,refundbeasongpay, refunddeliverypay, refundadjustpay,"
    sqlStr = sqlStr + " ,rebankname, rebankaccount, rebankownername, paygateTid"
    sqlStr = sqlStr + " )"
    sqlStr = sqlStr + " values(" & id
    sqlStr = sqlStr + " ,'" & "R100" & "'"
    sqlStr = sqlStr + " ," & refundrequire
    sqlStr = sqlStr + " ," & "0"
    sqlStr = sqlStr + " ," & orgsubtotalprice
    sqlStr = sqlStr + " ,''"
    sqlStr = sqlStr + " ,''"
    sqlStr = sqlStr + " ,''"
    sqlStr = sqlStr + " ,'" & paygatetid & "'"
    sqlStr = sqlStr + " )"

    dbget.execute sqlStr
end sub

'실시간이체 취소
sub insertCancelRealTimeTransferRequest(byVal orderserial, byVal basecsid, byVal refundrequire, refundcstitle)
    dim sqlStr
    dim cause, causedetail
    dim buyname, userid, paygatetid
    dim id
    dim orgsubtotalprice

    cause = ""
    causedetail = ""

    sqlStr = " select top 1 * from [db_academy].[dbo].tbl_academy_order_master where orderserial = '" + CStr(orderserial) + "' "
    rsAcademyget.Open sqlStr,dbAcademyget,1

    if Not rsAcademyget.Eof then
            buyname = db2html(rsAcademyget("buyname"))
            userid = rsAcademyget("userid")
            paygatetid = db2html(rsAcademyget("paygatetid"))
            orgsubtotalprice = rsAcademyget("subtotalprice")
    else
            buyname = ""
            userid = ""
            paygatetid = ""
            orgsubtotalprice = 0
    end if
    rsAcademyget.close


	sqlStr = " select * from [db_cs].[dbo].tbl_new_as_list where 1=0 "
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("divcd")          = "A007"
	rsget("orderserial")    = orderserial
	rsget("customername")   = html2db(buyname)
	rsget("userid")         = userid
	rsget("writeuser")      = session("ssBctId")
	rsget("title")          = html2db(refundcstitle + " - 실시간이체취소")
	rsget("contents_jupsu") = html2db("실시간이체[ " + paygatetid + " ]")
	'rsget("refundrequire")  = refundrequire
	'rsget("cause")          = cause
	'rsget("causedetail")    = html2db(causedetail)
	rsget("requireupche")   = "N"
	rsget("makerid")        = ""
	rsget("deleteyn")       = "N"
	rsget("songjangno")     = ""
	'rsget("rebankname")     = ""
	'rsget("rebankaccount")  = ""
	'rsget("rebankownername")        = ""

	rsget.update
	id = rsget("id")
	rsget.close



	sqlStr = "insert into [db_cs].[dbo].tbl_as_refund_info"
	sqlStr = sqlStr + " (asid,returnmethod,refundrequire, refundresult, orgsubtotalprice"
	'sqlStr = sqlStr + " ,orgitemcostsum, orgbeasongpay, orgmileagesum, orgcouponsum, orgallatdiscountsum"
    'sqlStr = sqlStr + " ,canceltotal, refunditemcostsum, refundmileagesum, refundcouponsum, allatsubtractsum"
    'sqlStr = sqlStr + " ,refundbeasongpay, refunddeliverypay, refundadjustpay,"
    sqlStr = sqlStr + " ,rebankname, rebankaccount, rebankownername, paygateTid"
    sqlStr = sqlStr + " )"
    sqlStr = sqlStr + " values(" & id
    sqlStr = sqlStr + " ,'" & "R020" & "'"
    sqlStr = sqlStr + " ," & refundrequire
    sqlStr = sqlStr + " ," & "0"
    sqlStr = sqlStr + " ," & orgsubtotalprice
    sqlStr = sqlStr + " ,''"
    sqlStr = sqlStr + " ,''"
    sqlStr = sqlStr + " ,''"
    sqlStr = sqlStr + " ,'" & paygatetid & "'"
    sqlStr = sqlStr + " )"

    dbget.execute sqlStr
end sub

'포인트 취소
sub insertCancelPointRequest(byVal orderserial, byVal basecsid)
        dim sqlStr
        dim cause, causedetail
        dim buyname, userid, paygatetid
        dim id

        cause = ""
        causedetail = ""

        sqlStr = " select top 1 * from [db_academy].[dbo].tbl_academy_order_master where orderserial = '" + CStr(orderserial) + "' "
        rsAcademyget.Open sqlStr,dbAcademyget,1

        if Not rsAcademyget.Eof then
                buyname = db2html(rsAcademyget("buyname"))
                userid = rsAcademyget("userid")
                paygatetid = db2html(rsAcademyget("paygatetid"))
        else
                buyname = ""
                userid = ""
                paygatetid = ""
        end if
        rsAcademyget.close


	sqlStr = " select * from [db_cs].[dbo].tbl_as_list where 1=0 "
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("divcd")          = "7"
	rsget("orderserial")    = orderserial
	rsget("customername")   = html2db(buyname)
	rsget("userid")         = userid
	rsget("writeuser")      = session("ssBctId")
	rsget("title")          = html2db("강좌취소(포인트취소)")
	rsget("contents_jupsu") = html2db("포인트[ " + paygatetid + " ]")
	rsget("refundrequire")  = 0
	rsget("cause")          = cause
	rsget("causedetail")    = html2db(causedetail)
	rsget("requireupche")   = "N"
	rsget("makerid")        = ""
	rsget("deleteyn")       = "N"
	rsget("songjangno")     = ""
	rsget("rebankname")     = ""
	rsget("rebankaccount")  = ""
	rsget("rebankownername")        = ""

	rsget.update
	id = rsget("id")
	rsget.close
end sub

'입점몰 취소
sub insertCancelMallRequest(byVal orderserial, byVal basecsid)
        dim sqlStr
        dim cause, causedetail
        dim buyname, userid, paygatetid
        dim id

        cause = ""
        causedetail = ""

        sqlStr = " select top 1 * from [db_academy].[dbo].tbl_academy_order_master where orderserial = '" + CStr(orderserial) + "' "
        rsAcademyget.Open sqlStr,dbAcademyget,1

        if Not rsAcademyget.Eof then
                buyname = db2html(rsAcademyget("buyname"))
                userid = rsAcademyget("userid")
                paygatetid = db2html(rsAcademyget("paygatetid"))
        else
                buyname = ""
                userid = ""
                paygatetid = ""
        end if
        rsAcademyget.close


	sqlStr = " select * from [db_cs].[dbo].tbl_as_list where 1=0 "
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("divcd")          = "7"
	rsget("orderserial")    = orderserial
	rsget("customername")   = html2db(buyname)
	rsget("userid")         = userid
	rsget("writeuser")      = session("ssBctId")
	rsget("title")          = html2db("강좌취소(입점몰취소)")
	rsget("contents_jupsu") = html2db("입점몰[ " + paygatetid + " ]")
	rsget("refundrequire")  = 0
	rsget("cause")          = cause
	rsget("causedetail")    = html2db(causedetail)
	rsget("requireupche")   = "N"
	rsget("makerid")        = ""
	rsget("deleteyn")       = "N"
	rsget("songjangno")     = ""
	rsget("rebankname")     = ""
	rsget("rebankaccount")  = ""
	rsget("rebankownername")        = ""

	rsget.update
	id = rsget("id")
	rsget.close
end sub

'올앳카드 취소
sub insertCancelAllAtCardRequest(byVal orderserial, byVal basecsid, byVal refundrequire)
        dim sqlStr
        dim cause, causedetail
        dim buyname, userid, paygatetid
        dim id

        cause = ""
        causedetail = ""

        sqlStr = " select top 1 * from [db_academy].[dbo].tbl_academy_order_master where orderserial = '" + CStr(orderserial) + "' "
        rsAcademyget.Open sqlStr,dbAcademyget,1

        if Not rsAcademyget.Eof then
                buyname = db2html(rsAcademyget("buyname"))
                userid = rsAcademyget("userid")
                paygatetid = db2html(rsAcademyget("paygatetid"))
        else
                buyname = ""
                userid = ""
                paygatetid = ""
        end if
        rsAcademyget.close


	sqlStr = " select * from [db_cs].[dbo].tbl_as_list where 1=0 "
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("divcd")          = "7"
	rsget("orderserial")    = orderserial
	rsget("customername")   = html2db(buyname)
	rsget("userid")         = userid
	rsget("writeuser")      = session("ssBctId")
	rsget("title")          = html2db("강좌취소(올앳카드취소)")
	rsget("contents_jupsu") = html2db("올앳카드[ " + paygatetid + " ]")
	rsget("refundrequire")  = refundrequire
	rsget("cause")          = cause
	rsget("causedetail")    = html2db(causedetail)
	rsget("requireupche")   = "N"
	rsget("makerid")        = ""
	rsget("deleteyn")       = "N"
	rsget("songjangno")     = ""
	rsget("rebankname")     = ""
	rsget("rebankaccount")  = ""
	rsget("rebankownername")        = ""

	rsget.update
	id = rsget("id")
	rsget.close
end sub

'상품권 취소
sub insertCancelTicketRequest(byVal orderserial, byVal basecsid)
        dim sqlStr
        dim cause, causedetail
        dim buyname, userid, paygatetid
        dim id

        cause = ""
        causedetail = ""

        sqlStr = " select top 1 * from [db_academy].[dbo].tbl_academy_order_master where orderserial = '" + CStr(orderserial) + "' "
        rsAcademyget.Open sqlStr,dbAcademyget,1

        if Not rsAcademyget.Eof then
                buyname = db2html(rsAcademyget("buyname"))
                userid = rsAcademyget("userid")
                paygatetid = db2html(rsAcademyget("paygatetid"))
        else
                buyname = ""
                userid = ""
                paygatetid = ""
        end if
        rsAcademyget.close


	sqlStr = " select * from [db_cs].[dbo].tbl_as_list where 1=0 "
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("divcd")          = "7"
	rsget("orderserial")    = orderserial
	rsget("customername")   = html2db(buyname)
	rsget("userid")         = userid
	rsget("writeuser")      = session("ssBctId")
	rsget("title")          = html2db("강좌취소(상품권취소)")
	rsget("contents_jupsu") = html2db("상품권[ " + paygatetid + " ]")
	rsget("refundrequire")  = 0
	rsget("cause")          = cause
	rsget("causedetail")    = html2db(causedetail)
	rsget("requireupche")   = "N"
	rsget("makerid")        = ""
	rsget("deleteyn")       = "N"
	rsget("songjangno")     = ""
	rsget("rebankname")     = ""
	rsget("rebankaccount")  = ""
	rsget("rebankownername")        = ""

	rsget.update
	id = rsget("id")
	rsget.close
end sub

sub cancelInicisCardPay(byVal orderserial)
    dim sqlStr
    dim refundrequire, cause, causedetail
    dim buyname, userid, paygatetid, accountdiv
    dim id

    sqlStr = " select top 1 * from [db_academy].[dbo].tbl_academy_order_master where orderserial = '" + CStr(orderserial) + "' "
    rsAcademyget.Open sqlStr,dbAcademyget,1
    'response.write sqlStr

    if Not rsAcademyget.Eof then
            buyname = db2html(rsAcademyget("buyname"))
            userid = rsAcademyget("userid")
            paygatetid = rsAcademyget("paygatetid")
            accountdiv = rsAcademyget("accountdiv")
            if ((accountdiv <> "20") and (accountdiv <> "90") and (accountdiv <> "100")) then
                    paygatetid = ""
            end if
    else
            buyname = ""
            userid = ""
            paygatetid = ""
            accountdiv = ""
    end if
    rsAcademyget.close

    '이니시스 모듈(카드취소)
    dim INIpay, PInst, ResultCode, ResultMsg

    ResultCode = "--"
    ResultMsg = "TID 없음"
    if (paygatetid <> "") then
        Set INIpay = Server.CreateObject("INItx41.INItx41.1")
        PInst = INIpay.Initialize("")
        INIpay.SetActionType CLng(PInst), "CANCEL"

        INIpay.SetField CLng(PInst), "pgid", "IniTechPG_"       'PG ID (고정)
        INIpay.SetField CLng(PInst), "spgip", "203.238.3.10"    '예비 PG IP (고정)
        INIpay.SetField CLng(PInst), "mid", "teenxteen3"        '상점아이디
        INIpay.SetField CLng(PInst), "admin", "1111"            '키패스워드(상점아이디에 따라 변경)
        INIpay.SetField CLng(PInst), "tid", paygatetid          '취소할 거래번호(TID)
        INIpay.SetField CLng(PInst), "msg", "CS카드취소"        '취소 사유
        INIpay.SetField CLng(PInst), "uip", Request.ServerVariables("REMOTE_ADDR") 'IP
        INIpay.SetField CLng(PInst), "debug", "true"            '로그모드("true"로 설정하면 상세한 로그를 남김)
        INIpay.SetField CLng(PInst), "merchantreserved", "예비" '예비

        INIpay.StartAction(CLng(PInst))

        ResultCode = INIpay.GetResult(CLng(PInst), "resultcode") '결과코드 ("00"이면 취소성공)
        ResultMsg = INIpay.GetResult(CLng(PInst), "resultmsg") '결과내용
        'CancelDate = INIpay.GetResult(CLng(PInst), "pgcanceldate") '이니시스 취소날짜
        'CancelTime = INIpay.GetResult(CLng(PInst), "pgcanceltime") '이니시스 취소시각
        'Rcash_cancel_noappl = INIpay.GetResult(CLng(PInst), "rcash_cancel_noappl") '현금영수증 취소 승인번호

        INIpay.Destroy CLng(PInst)
    end if

    'response.write ResultMsg
end sub

%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
