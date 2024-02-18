<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/classes/requestlecturecls.asp"-->
<!-- #include virtual="/academy/lib/classes/fingers_lecturecls.asp"-->
<!-- #include virtual="/academy/lib/email/maillib.asp"-->
<%

dim oordermaster, oorderdetail

dim orderserial
dim mode, i, j, k, tmp
dim sqlStr

mode = RequestCheckvar((request("mode"),16)

orderserial     = RequestCheckvar((request("orderserial"),16)


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
'
if ((oordermaster.FOneItem.Fipkumdiv <> "2") or (oordermaster.FOneItem.Faccountdiv <> "7") or (oordermaster.FOneItem.Fcancelyn <> "N")) then
    response.write "<script>alert('정상주문중 무통장에 대한 결재완료 전환만 가능합니다.'); opener.focus(); window.close();</script>"
    dbget.close()	:	response.End
else
    mode = "4"
end if


'==============================================================================
if (mode = "4") then
	'결재완료처리
	sqlStr = " update [db_academy].[dbo].tbl_academy_order_master "
	sqlStr = sqlStr + " set ipkumdiv='4', ipkumdate = getdate() "
	sqlStr = sqlStr + " where orderserial = '" + CStr(orderserial) + "' "
	rsAcademyget.Open sqlStr,dbAcademyget,1

    '==========================================================================
    'TODO : 메일발송
    call sendmailbankok(oordermaster.FOneItem.FBuyEmail, oordermaster.FOneItem.FBuyName, orderserial)


    '==========================================================================
    '비회원이 아닌경우
    if (oordermaster.FOneItem.FUserID <> "") then
        '==============================================================
        '고객 사용마일리지/획득마일리지 재계산
        updateUserMileage oordermaster.FOneItem.FUserID
    end if

    response.write "<script>alert('결재완료 처리되었습니다.');</script>"
    response.write "<script>opener.location.reload(); opener.parent.topframe.location.reload(); opener.focus(); window.close();</script>"
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
    rsget.close

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
	rsAcademyget.Open sqlStr,dbAcademyget,1
	'response.write sqlStr

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
	sqlStr = " select sum(totalmileage) as totmile"
    sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_master"
    sqlStr = sqlStr + " where userid='" + CStr(userid) + "' "
    sqlStr = sqlStr + " and cancelyn='N'"
    sqlStr = sqlStr + " and ipkumdiv>3"

    rsAcademyget.Open sqlStr,dbAcademyget,1
    if Not rsAcademyget.Eof then
    	totmile = rsAcademyget("totmile")
    end if
    rsAcademyget.close

	'==============================================================
	'주문마일리지 요약 재계산([db_academy].[dbo].tbl_academy_order_master)
    sqlStr = "update [db_user].[dbo].tbl_user_current_mileage"
    sqlStr = sqlStr + " set academymileage=" + CStr(totmile)
    sqlStr = sqlStr + " where userid='" + CStr(userid) + "' "
    rsget.Open sqlStr,dbget,1
end sub

sub insertRepayBank(byVal orderserial, byVal basecsid)
        dim sqlStr
        dim rebankname, rebankaccount, rebankownername, refundrequire, cause, causedetail
        dim buyname, userid
        dim id

        sqlStr = " select top 1 * from [db_cs].[dbo].tbl_as_list where id = " + CStr(basecsid) + " "
        rsget.Open sqlStr,dbget,1

        if Not rsget.Eof then
                rebankname = db2html(rsget("rebankname"))
                rebankaccount = db2html(rsget("rebankaccount"))
                rebankownername = db2html(rsget("rebankownername"))
                refundrequire = rsget("refundrequire")

                cause = rsget("cause")
                causedetail = db2html(rsget("causedetail"))
        else
                rebankname = ""
                rebankaccount = ""
                rebankownername = ""
                refundrequire = 0

                cause = ""
                causedetail = ""
        end if
        rsget.close

        sqlStr = " select top 1 * from [db_academy].[dbo].tbl_academy_order_master where orderserial = " + CStr(orderserial) + " "
        rsAcademyget.Open sqlStr,dbAcademyget,1

        if Not rsAcademyget.Eof then
                buyname = db2html(rsAcademyget("buyname"))
                userid = rsAcademyget("userid")
        else
                buyname = ""
                userid = ""
        end if
        rsAcademyget.close


	sqlStr = " select * from [db_cs].[dbo].tbl_as_list where 1=0 "
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("divcd")          = "3"
	rsget("orderserial")    = orderserial
	rsget("customername")   = html2db(buyname)
	rsget("userid")         = userid
	rsget("writeuser")      = session("ssBctId")
	rsget("title")          = html2db("환불요청")
	rsget("contents_jupsu") = html2db("")
	rsget("refundrequire")  = refundrequire
	rsget("cause")          = cause
	rsget("causedetail")    = html2db(causedetail)
	rsget("requireupche")   = "N"
	rsget("makerid")        = ""
	rsget("deleteyn")       = "N"
	rsget("songjangno")     = ""
	rsget("rebankname")     = rebankname
	rsget("rebankaccount")  = html2db(rebankaccount)
	rsget("rebankownername")        = html2db(rebankownername)

	rsget.update
	id = rsget("id")
	rsget.close
end sub

'신용카드취소
sub insertCancelCardRequest(byVal orderserial, byVal basecsid)
        dim sqlStr
        dim cause, causedetail
        dim buyname, userid, paygatetid
        dim id

        sqlStr = " select top 1 * from [db_cs].[dbo].tbl_as_list where id = " + CStr(basecsid) + " "
        rsget.Open sqlStr,dbget,1

        if Not rsget.Eof then
                cause = rsget("cause")
                causedetail = db2html(rsget("causedetail"))
        else
                cause = ""
                causedetail = ""
        end if
        rsget.close

        sqlStr = " select top 1 * from [db_academy].[dbo].tbl_academy_order_master where orderserial = " + CStr(orderserial) + " "
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
	rsget("title")          = html2db("카드취소")
	rsget("contents_jupsu") = html2db("신용카드[" + paygatetid + "]")
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

'실시간이체 취소
sub insertCancelRealTimeTransferRequest(byVal orderserial, byVal basecsid)
        dim sqlStr
        dim cause, causedetail
        dim buyname, userid, paygatetid
        dim id

        sqlStr = " select top 1 * from [db_cs].[dbo].tbl_as_list where id = " + CStr(basecsid) + " "
        rsget.Open sqlStr,dbget,1

        if Not rsget.Eof then
                cause = rsget("cause")
                causedetail = db2html(rsget("causedetail"))
        else
                cause = ""
                causedetail = ""
        end if
        rsget.close

        sqlStr = " select top 1 * from [db_academy].[dbo].tbl_academy_order_master where orderserial = " + CStr(orderserial) + " "
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
	rsget("title")          = html2db("실시간이체취소")
	rsget("contents_jupsu") = html2db("실시간이체[" + paygatetid + "]")
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

'포인트 취소
sub insertCancelPointRequest(byVal orderserial, byVal basecsid)
        dim sqlStr
        dim cause, causedetail
        dim buyname, userid, paygatetid
        dim id

        sqlStr = " select top 1 * from [db_cs].[dbo].tbl_as_list where id = " + CStr(basecsid) + " "
        rsget.Open sqlStr,dbget,1

        if Not rsget.Eof then
                cause = rsget("cause")
                causedetail = db2html(rsget("causedetail"))
        else
                cause = ""
                causedetail = ""
        end if
        rsget.close

        sqlStr = " select top 1 * from [db_academy].[dbo].tbl_academy_order_master where orderserial = " + CStr(orderserial) + " "
        rsAcademyget.Open sqlStr,dbAcademyget,1

        if Not rsget.Eof then
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
	rsget("title")          = html2db("포인트취소")
	rsget("contents_jupsu") = html2db("포인트[" + paygatetid + "]")
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

        sqlStr = " select top 1 * from [db_cs].[dbo].tbl_as_list where id = " + CStr(basecsid) + " "
        rsget.Open sqlStr,dbget,1

        if Not rsget.Eof then
                cause = rsget("cause")
                causedetail = db2html(rsget("causedetail"))
        else
                cause = ""
                causedetail = ""
        end if
        rsget.close

        sqlStr = " select top 1 * from [db_academy].[dbo].tbl_academy_order_master where orderserial = " + CStr(orderserial) + " "
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
	rsget("title")          = html2db("입점몰취소")
	rsget("contents_jupsu") = html2db("입점몰[" + paygatetid + "]")
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
sub insertCancelAllAtCardRequest(byVal orderserial, byVal basecsid)
        dim sqlStr
        dim cause, causedetail
        dim buyname, userid, paygatetid
        dim id

        sqlStr = " select top 1 * from [db_cs].[dbo].tbl_as_list where id = " + CStr(basecsid) + " "
        rsget.Open sqlStr,dbget,1

        if Not rsget.Eof then
                cause = rsget("cause")
                causedetail = db2html(rsget("causedetail"))
        else
                cause = ""
                causedetail = ""
        end if
        rsget.close

        sqlStr = " select top 1 * from [db_academy].[dbo].tbl_academy_order_master where orderserial = " + CStr(orderserial) + " "
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
	rsget("title")          = html2db("올앳카드취소")
	rsget("contents_jupsu") = html2db("올앳카드[" + paygatetid + "]")
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

'상품권 취소
sub insertCancelTicketRequest(byVal orderserial, byVal basecsid)
        dim sqlStr
        dim cause, causedetail
        dim buyname, userid, paygatetid
        dim id

        sqlStr = " select top 1 * from [db_cs].[dbo].tbl_as_list where id = " + CStr(basecsid) + " "
        rsget.Open sqlStr,dbget,1

        if Not rsget.Eof then
                cause = rsget("cause")
                causedetail = db2html(rsget("causedetail"))
        else
                cause = ""
                causedetail = ""
        end if
        rsget.close

        sqlStr = " select top 1 * from [db_academy].[dbo].tbl_academy_order_master where orderserial = " + CStr(orderserial) + " "
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
	rsget("title")          = html2db("상품권취소")
	rsget("contents_jupsu") = html2db("상품권[" + paygatetid + "]")
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

sub cancelInicisCardPay(byVal orderserial, byVal basecsid)
        dim sqlStr
        dim refundrequire, cause, causedetail
        dim buyname, userid, paygatetid, accountdiv
        dim id

        '카드취소접수CS 작성
        sqlStr = " select top 1 * from [db_cs].[dbo].tbl_as_list where id = " + CStr(basecsid) + " "
        rsget.Open sqlStr,dbget,1

        if Not rsget.Eof then
                refundrequire = rsget("refundrequire")

                cause = rsget("cause")
                causedetail = db2html(rsget("causedetail"))
        else
                refundrequire = 0

                cause = ""
                causedetail = ""
        end if
        rsget.close

        sqlStr = " select top 1 * from [db_academy].[dbo].tbl_academy_order_master where orderserial = " + CStr(orderserial) + " "
        rsAcademyget.Open sqlStr,dbAcademyget,1

        if Not rsget.Eof then
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


	sqlStr = " select * from [db_cs].[dbo].tbl_as_list where 1=0 "
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("divcd")          = "7"
	rsget("orderserial")    = orderserial
	rsget("customername")   = html2db(buyname)
	rsget("userid")         = userid
	rsget("writeuser")      = session("ssBctId")
	rsget("title")          = html2db("카드취소")
	rsget("contents_jupsu") = html2db("")
	rsget("refundrequire")  = refundrequire
	rsget("cause")          = cause
	rsget("causedetail")    = html2db(causedetail)
	rsget("requireupche")   = "N"
	rsget("makerid")        = ""
	rsget("deleteyn")       = "N"
	rsget("songjangno")     = ""

	rsget.update
	id = rsget("id")
	rsget.close


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

	sqlStr = " update [db_cs].[dbo].tbl_as_list "
	sqlStr = sqlStr + " set finishdate=getdate() "
	sqlStr = sqlStr + " ,finishuser = '" + session("ssBctId") + "' "
	sqlStr = sqlStr + " ,contents_finish = '[" + CStr(ResultCode) + "]" + CStr(ResultMsg) + "' "
	sqlStr = sqlStr + " ,currstate = '7' "
	if (ResultCode = "00") then
	        sqlStr = sqlStr + " ,refundresult = " + CStr(refundrequire) + " "
	else
	        sqlStr = sqlStr + " ,refundresult = 0 "
	end if
	sqlStr = sqlStr + " where id=" + CStr(id) + " "
	rsget.Open sqlStr,dbget,1
end sub

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
