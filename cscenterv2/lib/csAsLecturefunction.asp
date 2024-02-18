<%
dim GC_IsOldOrder
GC_IsOldOrder = false

function CheckIsOldOrder(orderserial)
    ''과거 주문인지 Check
    dim sqlStr

	sqlStr = " select orderserial from " & TABLE_ORDERMASTER & " where orderserial='" & orderserial & "'"
	rsget.Open sqlStr,dbget,1
	if Not rsget.Eof then
	    CheckIsOldOrder = False
	else
        CheckIsOldOrder = True
    end if
    rsget.close

    if (CheckIsOldOrder) then
        sqlStr = " select orderserial from db_log.dbo.tbl_old_order_master_2003 where orderserial='" & orderserial & "'"
	    rsget.Open sqlStr,dbget,1
	    if rsget.Eof then
	        CheckIsOldOrder = False
	    end if
	    rsget.close
    end if
end function

function getCardRibonName(cardribbon)
    if IsNULL(cardribbon) then Exit Function

    if (cardribbon="1") then
        getCardRibonName  = "카드"
    elseif (cardribbon="2") then
        getCardRibonName  = "리본"
    elseif (cardribbon="3") then
        getCardRibonName  = "없음"
    end if
end function

function FinishCSMaster(iAsid, finishuser, contents_finish)
    dim sqlStr
    dim IsCsErrStockUpdateRequire
    IsCsErrStockUpdateRequire = False

    sqlStr = "select divcd, finishdate, currstate"
    sqlStr = sqlStr + " from " & TABLE_CSMASTER & ""
    sqlStr = sqlStr + " where id=" + CStr(iAsid)
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        IsCsErrStockUpdateRequire = (rsget("divcd")="A011") and (IsNULL(rsget("finishdate"))) and (rsget("currstate")<>"B007")
    end if
    rsget.close

    sqlStr = " update " & TABLE_CSMASTER & ""                      + VbCrlf
    sqlStr = sqlStr + " set finishuser='" + finishuser + "'"            + VbCrlf
    sqlStr = sqlStr + " , contents_finish='" + contents_finish + "'"    + VbCrlf
    sqlStr = sqlStr + " , finishdate=getdate()"                         + VbCrlf
    sqlStr = sqlStr + " , currstate='B007'"                             + VbCrlf
    sqlStr = sqlStr + " where id=" + CStr(iAsid)

    dbget.Execute sqlStr

    ''맞교환회수 완료일경우 재고없데이트. 2007.11.16
    if (IsCsErrStockUpdateRequire) then
        sqlStr = " exec db_summary.dbo.ten_RealTimeStock_CsErr " & iAsid & ",'','" & finishuser & "'"
        'dbget.Execute sqlStr
    end if
end function

function GetDefaultTitle(divcd, id, orderserial)
    dim opentitle, opencontents
    dim ipkumdiv, accountdiv, cancelyn, comm_name, ipkumdivName, accountdivName
    dim sqlStr

    sqlStr = " select m.ipkumdiv, m.accountdiv, m.cancelyn, C.comm_name"
    if (GC_IsOLDOrder) then
        sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_master_2003 m"
    else
        sqlStr = sqlStr + " from " & TABLE_ORDERMASTER & " m"
    end if
    sqlStr = sqlStr + " left join " & TABLE_CSMASTER & " A"
    sqlStr = sqlStr + "     on A.orderserial='" + orderserial + "'"
    if (id<>"") then
        sqlStr = sqlStr + " and A.id=" + CStr(id)
    end if
    sqlStr = sqlStr + " left join [db_academy].[dbo].tbl_academy_cs_lecutre_comm_code C"
    sqlStr = sqlStr + " on C.comm_cd='" + divcd + "'"

    sqlStr = sqlStr + " where m.orderserial='" + orderserial + "'"

    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        ipkumdiv    = rsget("ipkumdiv")
        cancelyn    = rsget("cancelyn")
        comm_name   = rsget("comm_name")
        accountdiv  = Trim(rsget("accountdiv"))
    end if
    rsget.close


    if (ipkumdiv="2") then
        ipkumdivName = "입금 대기"
    elseif (ipkumdiv="4") then
        ipkumdivName = "결제 완료"
    elseif (ipkumdiv="5") then
        ipkumdivName = "상품 준비"
    elseif (ipkumdiv="6") then
        ipkumdivName = "출고 준비"
    elseif (ipkumdiv="7") then
        ipkumdivName = "일부 출고"
    elseif (ipkumdiv="8") then
        ipkumdivName = "출고 완료"
    end if

    if (accountdiv="7") then
        accountdivName = "무통장"
    elseif (accountdiv="100") then
        accountdivName = "신용카드"
    elseif (accountdiv="80") then
        accountdivName = "올엣카드"
    elseif (accountdiv="50") then
        accountdivName = "제휴몰결제"
    elseif (accountdiv="20") then
        accountdivName = "실시간이체"

    end if

    ''취소만..
    if (divcd="A007") or (divcd="A008") then
        GetDefaultTitle = accountdivName + " " + ipkumdivName + " 상태 중 " + comm_name
    else
        GetDefaultTitle = comm_name + " 접수"
    end if
end function

function AddCsMemoWithMemoGubun(orderserial,divcd,userid,writeuser,contents_jupsu,mmgubun)
	dim sqlStr

	if divcd="1" then
        ''일반메모
        sqlStr = "insert into " & TABLE_CS_MEMO & ""
        sqlStr = sqlStr + "(orderserial,divcd,userid,mmgubun,writeuser,finishuser,contents_jupsu,finishyn,finishdate)"
        sqlStr = sqlStr + " values('" + orderserial + "','" + divcd + "','" + userid + "','" + mmgubun + "','" + writeuser + "','" + writeuser + "','" + html2db(contents_jupsu) + "','Y',getdate())"

        dbget.Execute sqlStr
    else
        ''처리요청메모
        sqlStr = "insert into " & TABLE_CS_MEMO & ""
        sqlStr = sqlStr + "(orderserial,divcd,userid,mmgubun,writeuser,contents_jupsu,finishyn)"
        sqlStr = sqlStr + " values('" + orderserial + "','" + divcd + "','" + userid + "','" + mmgubun + "','" + writeuser + "','" + html2db(contents_jupsu) + "','N')"

        dbget.Execute sqlStr
    end if
end function

function AddCsMemo(orderserial,divcd,userid,writeuser,contents_jupsu)
    dim sqlStr
    dim mmgubun ''메모구분
    if (LCase(LEFT(contents_jupsu,4))="[sms") then
    	mmgubun = "4"
    elseif (LCase(LEFT(contents_jupsu,5))="[mail") then
    	mmgubun = "5"
    else
    	mmgubun = "0"
    end if

    if divcd="1" then
        ''일반메모
        sqlStr = "insert into " & TABLE_CS_MEMO & ""
        sqlStr = sqlStr + "(orderserial,divcd,userid,mmgubun,writeuser,finishuser,contents_jupsu,finishyn,finishdate)"
        sqlStr = sqlStr + " values('" + orderserial + "','" + divcd + "','" + userid + "','" + mmgubun + "','" + writeuser + "','" + writeuser + "','" + html2db(contents_jupsu) + "','Y',getdate())"

        dbget.Execute sqlStr
    else
        ''처리요청메모
        sqlStr = "insert into " & TABLE_CS_MEMO & ""
        sqlStr = sqlStr + "(orderserial,divcd,userid,mmgubun,writeuser,contents_jupsu,finishyn)"
        sqlStr = sqlStr + " values('" + orderserial + "','" + divcd + "','" + userid + "','" + mmgubun + "','" + writeuser + "','" + html2db(contents_jupsu) + "','N')"

        dbget.Execute sqlStr
    end if

end function

function SetCustomerOpenMsg(id, opentitle, opencontents)
    dim sqlStr

    sqlStr = " update " & TABLE_CSMASTER & ""        + VbCrlf
    sqlStr = sqlStr + " set opentitle='" + opentitle + "'"  + VbCrlf
    sqlStr = sqlStr + " , opencontents='" + opencontents + "'" + VbCrlf
    sqlStr = sqlStr + " where id=" + CStr(id)

    dbget.Execute sqlStr

end function

'function AddCustomerOpenMsg(id, orderserial, addcontents)
'    dim sqlStr
'
'    sqlStr = " update " & TABLE_CSMASTER & ""        + VbCrlf
'    sqlStr = sqlStr + " set opentitle=opentitle + '" + VbCrlf + addcontents + "'" + VbCrlf
'    sqlStr = sqlStr + " where id=" + CStr(id)
'
'    dbget.Execute sqlStr
'
'end function


function AddCustomerOpenContents(id, addcontents)
    dim sqlStr

    if ((addcontents="") or (id="")) then Exit Function

    sqlStr = " update " & TABLE_CSMASTER & ""        + VbCrlf
    sqlStr = sqlStr + " set opencontents=IsNULL(opencontents,'') + (Case When (IsNULL(opencontents,'')='') then '" & addcontents & "' else '" & VbCrlf & addcontents + "' End )" + VbCrlf
    sqlStr = sqlStr + " where id=" + CStr(id)

    dbget.Execute sqlStr

end function


function RegCSMasterAddUpche(id, imakerid)
    dim sqlStr
    sqlStr = " update " & TABLE_CSMASTER & ""    + VbCrlf
    sqlStr = sqlStr + " set makerid='" + imakerid + "'"   + VbCrlf
    sqlStr = sqlStr + " , requireupche='Y'"               + VbCrlf
    sqlStr = sqlStr + " where id=" + CStr(id)

    dbget.Execute sqlStr
end function


function RegCSMaster(divcd, orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
    '' CS Master 저장
    dim sqlStr, InsertedId

    sqlStr = " select * from " & TABLE_CSMASTER & " where 1=0 "
    rsget.Open sqlStr,dbget,1,3
    rsget.AddNew
        rsget("divcd")          = divcd
    	rsget("orderserial")    = orderserial
    	rsget("customername")   = ""
    	rsget("userid")         = ""
    	rsget("writeuser")      = reguserid
    	rsget("title")          = title
    	rsget("contents_jupsu") = contents_jupsu
    	rsget("gubun01")        = gubun01
    	rsget("gubun02")        = gubun02

    	rsget("currstate")      = "B001"
    	rsget("deleteyn")       = "N"


        ''''''''''''''''''''''''''''''''''
    	''rsget("requireupche")   = "N"
    	''rsget("makerid")        = ""
    	''''''''''''''''''''''''''''''''''

    rsget.update
	    InsertedId = rsget("id")
	rsget.close

	dim opentitle, opencontents

	opentitle = GetDefaultTitle(divcd, InsertedId, orderserial)

	sqlStr = " update " & TABLE_CSMASTER & ""  + VbCrlf
	sqlStr = sqlStr + " set userid=T.userid"        + VbCrlf
	sqlStr = sqlStr + " , customername=T.buyname"   + VbCrlf
	sqlStr = sqlStr + " , opentitle='" + html2db(opentitle) + "'" + VbCrlf
	sqlStr = sqlStr + " , opencontents='" + html2db(opencontents) + "'" + VbCrlf
	if (GC_IsOLDOrder) then
	    sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_master_2003 T" + VbCrlf
	else
    	sqlStr = sqlStr + " from " & TABLE_ORDERMASTER & " T" + VbCrlf
    end if

	sqlStr = sqlStr + " where T.orderserial='" + orderserial + "'"  + VbCrlf
	sqlStr = sqlStr + " and " & TABLE_CSMASTER & ".id=" + CStr(InsertedId)

	dbget.Execute sqlStr


	''회수신청 접수인경우 - 기본 회수 배송지 저장
	''맞교환, 서비스 발송, 누락발송
	if (divcd="A010") or (divcd="A010") or (divcd="A000") or (divcd="A001") or (divcd="A002") then
	    Call RegDefaultDEliverInfo(InsertedId, orderserial)
    end if

	RegCSMaster = InsertedId
end function

''기본 회수/맞교환/서비스발송 주소지 입력 - 접수시 주문번호 기본 주소지로 저장됨. - 저장후 수정하는 Procsess
function RegDefaultDEliverInfo(AsID, orderserial)
    dim sqlStr
    sqlStr = "insert into " & TABLE_CS_DELIVERY & ""
    sqlStr = sqlStr + " (asid, reqname, reqphone, reqhp, reqzipcode, reqzipaddr, reqetcaddr)"
    ''sqlStr = sqlStr + " select " + CStr(AsID) + ",reqname, reqphone, reqhp, reqzipcode, reqaddress, reqzipaddr" ''바꼈음.
    sqlStr = sqlStr + " select " + CStr(AsID) + ",reqname, reqphone, reqhp, reqzipcode, reqzipaddr, reqaddress"
    if (GC_IsOLDOrder) then
        sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_master_2003 T" + VbCrlf
    else
        sqlStr = sqlStr + " from " & TABLE_ORDERMASTER & ""
    end if
    sqlStr = sqlStr + " where orderserial='" + orderserial + "'"

    dbget.Execute sqlStr
end function

function EditCSMaster(divcd, orderserial, modiuserid, title, contents_jupsu, gubun01, gubun02)
    '' CS Master 수정
    dim sqlStr

    sqlStr = " update " & TABLE_CSMASTER & ""
    sqlStr = sqlStr + " set writeuser='" + modiuserid + "'"
    sqlStr = sqlStr + " ,title='" + title + "'"
    sqlStr = sqlStr + " ,contents_jupsu='" + contents_jupsu + "'"
    sqlStr = sqlStr + " ,gubun01='" + gubun01 + "'"
    sqlStr = sqlStr + " ,gubun02='" + gubun02 + "'"
    sqlStr = sqlStr + " where id=" + CStr(id)

    dbget.Execute sqlStr

end function

function EditCSMasterFinished(divcd, orderserial, modiuserid, title, contents_jupsu, gubun01, gubun02, finishuserid, contents_finish)
    '' CS Master 완료된 내역 수정
    dim sqlStr

    sqlStr = " update " & TABLE_CSMASTER & ""
    sqlStr = sqlStr + " set finishuser='" + finishuserid + "'"
    sqlStr = sqlStr + " ,title='" + title + "'"
    sqlStr = sqlStr + " ,contents_jupsu='" + contents_jupsu + "'"
    sqlStr = sqlStr + " ,contents_finish='" + contents_finish + "'"
    sqlStr = sqlStr + " ,gubun01='" + gubun01 + "'"
    sqlStr = sqlStr + " ,gubun02='" + gubun02 + "'"
    sqlStr = sqlStr + " where id=" + CStr(id)

    dbget.Execute sqlStr
end function


function RegCSMasterRefundInfo(asid, returnmethod, refundrequire , orgsubtotalprice, orgitemcostsum, orgbeasongpay , orgmileagesum, orgcouponsum, orgallatdiscountsum , canceltotal, refunditemcostsum, refundmileagesum, refundcouponsum, allatsubtractsum , refundbeasongpay, refunddeliverypay, refundadjustpay  , rebankname, rebankaccount, rebankownername, paygateTid)

    dim sqlStr

    sqlStr = " insert into " & TABLE_CS_REFUND & ""
    sqlStr = sqlStr + " (asid"
    sqlStr = sqlStr + " ,returnmethod"
    sqlStr = sqlStr + " ,refundrequire"
    sqlStr = sqlStr + " ,orgsubtotalprice"
    sqlStr = sqlStr + " ,orgitemcostsum"
    sqlStr = sqlStr + " ,orgbeasongpay"
    sqlStr = sqlStr + " ,orgmileagesum"
    sqlStr = sqlStr + " ,orgcouponsum"
    sqlStr = sqlStr + " ,orgallatdiscountsum"

    sqlStr = sqlStr + " ,canceltotal"
    sqlStr = sqlStr + " ,refunditemcostsum"
    sqlStr = sqlStr + " ,refundmileagesum"
    sqlStr = sqlStr + " ,refundcouponsum"
    sqlStr = sqlStr + " ,allatsubtractsum"
    sqlStr = sqlStr + " ,refundbeasongpay"
    sqlStr = sqlStr + " ,refunddeliverypay"
    sqlStr = sqlStr + " ,refundadjustpay"
    sqlStr = sqlStr + " ,rebankname"
    sqlStr = sqlStr + " ,rebankaccount"
    sqlStr = sqlStr + " ,rebankownername"
    sqlStr = sqlStr + " ,paygateTid"
    sqlStr = sqlStr + " )"

	'response.write "aaaaaaaaaaa" & sqlStr

    sqlStr = sqlStr + " values("
    sqlStr = sqlStr + " " + CStr(asid)
    sqlStr = sqlStr + " ,'" + returnmethod + "'"
    sqlStr = sqlStr + " ," + CStr(refundrequire)
    sqlStr = sqlStr + " ," + CStr(orgsubtotalprice)
    sqlStr = sqlStr + " ," + CStr(orgitemcostsum)
    sqlStr = sqlStr + " ," + CStr(orgbeasongpay)
    sqlStr = sqlStr + " ," + CStr(orgmileagesum)
    sqlStr = sqlStr + " ," + CStr(orgcouponsum)
    sqlStr = sqlStr + " ," + CStr(orgallatdiscountsum)

	'response.write "aaaaaaaaaaa" & sqlStr

    sqlStr = sqlStr + " ," + CStr(canceltotal)
    sqlStr = sqlStr + " ," + CStr(refunditemcostsum)
    sqlStr = sqlStr + " ," + CStr(refundmileagesum)
    sqlStr = sqlStr + " ," + CStr(refundcouponsum)
    sqlStr = sqlStr + " ," + CStr(allatsubtractsum)
    sqlStr = sqlStr + " ," + CStr(refundbeasongpay)
    sqlStr = sqlStr + " ," + CStr(refunddeliverypay)
    sqlStr = sqlStr + " ," + CStr(refundadjustpay)
    sqlStr = sqlStr + " ,'" + rebankname + "'"
    sqlStr = sqlStr + " ,'" + rebankaccount + "'"
    sqlStr = sqlStr + " ,'" + rebankownername + "'"
    sqlStr = sqlStr + " ,'" + paygateTid + "'"
    sqlStr = sqlStr + " )"

	'response.write "aaaaaaaaaaa" & sqlStr
    dbget.Execute sqlStr

end function

function RegCSMasterRefundInfoLecture(asid, returnmethod, refundrequire , orgsubtotalprice, orgitemcostsum, orgbeasongpay , orgmileagesum, orgcouponsum, orgallatdiscountsum , canceltotal, refunditemcostsum, refundmileagesum, refundcouponsum, allatsubtractsum , refundbeasongpay, refunddeliverypay, refundadjustpay  , rebankname, rebankaccount, rebankownername, paygateTid, refundmatdiv)

    dim sqlStr

    sqlStr = " insert into " & TABLE_CS_REFUND & ""
    sqlStr = sqlStr + " (asid"
    sqlStr = sqlStr + " ,returnmethod"
    sqlStr = sqlStr + " ,refundrequire"
    sqlStr = sqlStr + " ,orgsubtotalprice"
    sqlStr = sqlStr + " ,orgitemcostsum"
    sqlStr = sqlStr + " ,orgbeasongpay"
    sqlStr = sqlStr + " ,orgmileagesum"
    sqlStr = sqlStr + " ,orgcouponsum"
    sqlStr = sqlStr + " ,orgallatdiscountsum"

    sqlStr = sqlStr + " ,canceltotal"
    sqlStr = sqlStr + " ,refunditemcostsum"
    sqlStr = sqlStr + " ,refundmileagesum"
    sqlStr = sqlStr + " ,refundcouponsum"
    sqlStr = sqlStr + " ,allatsubtractsum"
    sqlStr = sqlStr + " ,refundbeasongpay"
    sqlStr = sqlStr + " ,refunddeliverypay"
    sqlStr = sqlStr + " ,refundadjustpay"
    sqlStr = sqlStr + " ,rebankname"
    sqlStr = sqlStr + " ,rebankaccount"
    sqlStr = sqlStr + " ,rebankownername"
    sqlStr = sqlStr + " ,paygateTid"
    sqlStr = sqlStr + " ,refundmatdiv"
    sqlStr = sqlStr + " )"

	'response.write "aaaaaaaaaaa" & sqlStr

    sqlStr = sqlStr + " values("
    sqlStr = sqlStr + " " + CStr(asid)
    sqlStr = sqlStr + " ,'" + returnmethod + "'"
    sqlStr = sqlStr + " ," + CStr(refundrequire)
    sqlStr = sqlStr + " ," + CStr(orgsubtotalprice)
    sqlStr = sqlStr + " ," + CStr(orgitemcostsum)
    sqlStr = sqlStr + " ," + CStr(orgbeasongpay)
    sqlStr = sqlStr + " ," + CStr(orgmileagesum)
    sqlStr = sqlStr + " ," + CStr(orgcouponsum)
    sqlStr = sqlStr + " ," + CStr(orgallatdiscountsum)

	'response.write "aaaaaaaaaaa" & sqlStr

    sqlStr = sqlStr + " ," + CStr(canceltotal)
    sqlStr = sqlStr + " ," + CStr(refunditemcostsum)
    sqlStr = sqlStr + " ," + CStr(refundmileagesum)
    sqlStr = sqlStr + " ," + CStr(refundcouponsum)
    sqlStr = sqlStr + " ," + CStr(allatsubtractsum)
    sqlStr = sqlStr + " ," + CStr(refundbeasongpay)
    sqlStr = sqlStr + " ," + CStr(refunddeliverypay)
    sqlStr = sqlStr + " ," + CStr(refundadjustpay)
    sqlStr = sqlStr + " ,'" + rebankname + "'"
    sqlStr = sqlStr + " ,'" + rebankaccount + "'"
    sqlStr = sqlStr + " ,'" + rebankownername + "'"
    sqlStr = sqlStr + " ,'" + paygateTid + "'"
    sqlStr = sqlStr + " ,'" + refundmatdiv + "'"
    sqlStr = sqlStr + " )"

	'response.write "aaaaaaaaaaa" & sqlStr
    dbget.Execute sqlStr

end function

function RegCSUpcheAddJungsanPay(iasid, iadd_upchejungsandeliverypay, iadd_upchejungsancause, buf_requiremakerid)
    dim sqlStr

    sqlStr = " insert into [db_cs].[dbo].tbl_as_upcheAddjungsan"
    sqlStr = sqlStr + " (asid, add_upchejungsandeliverypay, add_upchejungsancause)"
    sqlStr = sqlStr + " values(" &iasid
    sqlStr = sqlStr + " ," & iadd_upchejungsandeliverypay
    sqlStr = sqlStr + " ,'" & iadd_upchejungsancause & "')"

    dbget.Execute sqlStr

    ''기타 정산 추가인경우만 makerid 지정 : 강좌확정후 취소 접수(업체배송) / 맞교환(업체)인 경우는 기 지정됨.
    sqlStr = " update " & TABLE_CSMASTER & "" & VbCrlf
    sqlStr = sqlStr + " set makerid='" & buf_requiremakerid & "'" & VbCrlf
    sqlStr = sqlStr + " where id=" & iasid & "" & VbCrlf
    sqlStr = sqlStr + " and divcd='A700'" & VbCrlf

    dbget.Execute sqlStr

end function

function EditCSUpcheAddJungsanPay(iasid, iadd_upchejungsandeliverypay, iadd_upchejungsancause, buf_requiremakerid)
    dim sqlStr

    sqlStr = " IF EXISTS( select * from [db_cs].[dbo].tbl_as_upcheAddjungsan where asid=" & iasid & ")" & VbCrlf
    sqlStr = sqlStr + " BEGIN" & VbCrlf
    sqlStr = sqlStr + "     update [db_cs].[dbo].tbl_as_upcheAddjungsan" & VbCrlf
    sqlStr = sqlStr + "     set add_upchejungsandeliverypay=" & add_upchejungsandeliverypay & VbCrlf
    sqlStr = sqlStr + "     , add_upchejungsancause='" & iadd_upchejungsancause & "'" & VbCrlf
    sqlStr = sqlStr + "     where asid = " & iasid & VbCrlf
    sqlStr = sqlStr + " END" & VbCrlf
    sqlStr = sqlStr + " ELSE " & VbCrlf
    sqlStr = sqlStr + " BEGIN" & VbCrlf
    sqlStr = sqlStr + "     IF (0<>" & iadd_upchejungsandeliverypay & ")" & VbCrlf
    sqlStr = sqlStr + "     BEGIN" & VbCrlf
    sqlStr = sqlStr + "         insert into [db_cs].[dbo].tbl_as_upcheAddjungsan" & VbCrlf
    sqlStr = sqlStr + "         (asid, add_upchejungsandeliverypay, add_upchejungsancause)" & VbCrlf
    sqlStr = sqlStr + "         values(" &iasid & VbCrlf
    sqlStr = sqlStr + "         ," & iadd_upchejungsandeliverypay & VbCrlf
    sqlStr = sqlStr + "         ,'" & iadd_upchejungsancause & "')" & VbCrlf
    sqlStr = sqlStr + "     END" & VbCrlf
    sqlStr = sqlStr + " END" & VbCrlf

    dbget.Execute sqlStr


    ''기타 정산 추가인경우만 makerid 지정 : 강좌확정후 취소 접수(업체배송) / 맞교환(업체)인 경우는 기 지정됨.
    sqlStr = " update " & TABLE_CSMASTER & "" & VbCrlf
    sqlStr = sqlStr + " set makerid='" & buf_requiremakerid & "'" & VbCrlf
    sqlStr = sqlStr + " where id=" & iasid & "" & VbCrlf
    sqlStr = sqlStr + " and divcd='A700'" & VbCrlf
    sqlStr = sqlStr + " and IsNULL(makerid,'')<>'" & buf_requiremakerid & "'" & VbCrlf

    dbget.Execute sqlStr
end function


function AddCSDetailByArrStr(byval detailitemlist, id, orderserial)
    dim sqlStr, tmp, buf, i
    dim dorderdetailidx, dgubun01, dgubun02, dregitemno

    buf = split(detailitemlist, "|")

    for i = 0 to UBound(buf)
		if (TRIM(buf(i)) <> "") then
			tmp = split(buf(i), Chr(9))

			dorderdetailidx = tmp(0)
			dgubun01        = tmp(1)
			dgubun02        = tmp(2)
			dregitemno      = tmp(3)

	        call AddOneCSDetail(id, dorderdetailidx, dgubun01, dgubun02, orderserial, dregitemno)
		end if
	next

	sqlStr = " update " & TABLE_CSDETAIL & ""
	sqlStr = sqlStr + " set itemid=T.itemid"
	sqlStr = sqlStr + " , itemoption=T.itemoption"
	sqlStr = sqlStr + " , makerid=T.makerid"
	sqlStr = sqlStr + " , itemname=T.itemname"
	sqlStr = sqlStr + " , itemoptionname=T.itemoptionname"
	sqlStr = sqlStr + " , itemcost=T.itemcost"
	sqlStr = sqlStr + " , orderitemno=T.itemno"
	sqlStr = sqlStr + " , isupchebeasong=T.isupchebeasong"
	sqlStr = sqlStr + " , regdetailstate=T.currstate"
	if (GC_IsOldOrder) then
	    sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_detail_2003 T"
	else
	    sqlStr = sqlStr + " from " & TABLE_ORDERDETAIL & " T"
	end if
	sqlStr = sqlStr + " where T.orderserial='" + orderserial + "'"
	sqlStr = sqlStr + " and " & TABLE_CSDETAIL & ".masterid=" + CStr(id)
	sqlStr = sqlStr + " and " & TABLE_CSDETAIL & ".orderdetailidx=T." & FIELD_DETAILIDX & ""

	dbget.Execute sqlStr

end function

function EditCSDetailByArrStr(byval detailitemlist, id, orderserial)
    dim sqlStr, tmp, buf, i
    dim dorderdetailidx, dgubun01, dgubun02, dregitemno, dcausecontent

    buf = split(detailitemlist, "|")

    for i = 0 to UBound(buf)
		if (TRIM(buf(i)) <> "") then
			tmp = split(buf(i), Chr(9))

			dorderdetailidx = tmp(0)
			dgubun01        = tmp(1)
			dgubun02        = tmp(2)
			dregitemno      = tmp(3)
			dcausecontent   = tmp(4)

	        call EditOneCSDetail(id, dorderdetailidx, dgubun01, dgubun02, orderserial, dregitemno, dcausecontent)
		end if
	next

end function


function AddOneCSDetail(id, dorderdetailidx, dgubun01, dgubun02, orderserial, dregitemno)
    dim sqlStr

    sqlStr = " insert into " & TABLE_CSDETAIL & ""
    sqlStr = sqlStr + " (masterid, orderdetailidx, gubun01,gubun02"
    sqlStr = sqlStr + " ,orderserial, itemid, itemoption, makerid, itemname, itemoptionname, regitemno, confirmitemno,orderitemno) "
    sqlStr = sqlStr + " values(" + CStr(id) + ""
    sqlStr = sqlStr + " ," + CStr(dorderdetailidx) + ""
    sqlStr = sqlStr + " ,'" + CStr(dgubun01) + "'"
    sqlStr = sqlStr + " ,'" + CStr(dgubun02) + "'"
    sqlStr = sqlStr + " ,'" + CStr(orderserial) + "'"
    sqlStr = sqlStr + " ,0"
    sqlStr = sqlStr + " ,''"
    sqlStr = sqlStr + " ,''"
    sqlStr = sqlStr + " ,''"
    sqlStr = sqlStr + " ,''"
    sqlStr = sqlStr + " ," + CStr(dregitemno) + ""
    sqlStr = sqlStr + " ," + CStr(dregitemno) + ""
    sqlStr = sqlStr + " ,0"
    sqlStr = sqlStr + " )"

    dbget.Execute sqlStr
end function


function EditOneCSDetail(id, dorderdetailidx, dgubun01, dgubun02, orderserial, dregitemno, dcausecontent)
    dim sqlStr

    sqlStr = " update " & TABLE_CSDETAIL & ""
    sqlStr = sqlStr + " set gubun01='" + dgubun01 + "'"
    sqlStr = sqlStr + " , gubun02='" + dgubun02 + "'"
    sqlStr = sqlStr + " , regitemno=" + dregitemno + ""
    sqlStr = sqlStr + " , confirmitemno=" + dregitemno + ""
    sqlStr = sqlStr + " where masterid=" + CStr(id)
    sqlStr = sqlStr + " and orderdetailidx=" + CStr(dorderdetailidx)

    dbget.Execute sqlStr
end function

function AddOneDeliveryInfoCSDetail(id, gubun01, gubun02, orderserial)
    dim sqlStr

    sqlStr = " insert into " & TABLE_CSDETAIL & ""
    sqlStr = sqlStr + " (masterid, orderdetailidx, gubun01, gubun02,"
    sqlStr = sqlStr + " orderserial, itemid, itemoption, makerid,itemname, itemoptionname,"
    sqlStr = sqlStr + " regitemno, confirmitemno, orderitemno, itemcost, buycash, isupchebeasong,regdetailstate) "
    sqlStr = sqlStr + " select top 1 "
    sqlStr = sqlStr + " " + CStr(id)
    sqlStr = sqlStr + " ,d." & FIELD_DETAILIDX & ""
    sqlStr = sqlStr + " ,'" + CStr(gubun01) + "'"
    sqlStr = sqlStr + " ,'" + CStr(gubun02) + "'"
    sqlStr = sqlStr + " ,d.orderserial"
    sqlStr = sqlStr + " ,d.itemid"
    sqlStr = sqlStr + " ,d.itemoption"
    sqlStr = sqlStr + " ,IsNULL(d.makerid,'')"
    sqlStr = sqlStr + " ,IsNULL(d.itemname,'배송료')"
    sqlStr = sqlStr + " ,IsNULL(d.itemoptionname,(case when d.itemcost=0 then '무료배송' else '일반택배' end))"
    sqlStr = sqlStr + " ,d.itemno"
    sqlStr = sqlStr + " ,d.itemno"
    sqlStr = sqlStr + " ,d.itemno"
    sqlStr = sqlStr + " ,d.itemcost"
    sqlStr = sqlStr + " ,d.buycash"
    sqlStr = sqlStr + " ,d.isupchebeasong"
    sqlStr = sqlStr + " ,d.currstate"
    if (GC_IsOLDOrder) then
        sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_detail_2003 d"
    else
        sqlStr = sqlStr + " from " & TABLE_ORDERDETAIL & " d"
    end if
    sqlStr = sqlStr + " where d.orderserial='" + orderserial + "'"
    sqlStr = sqlStr + " and d.itemid=0"
    sqlStr = sqlStr + " and d.cancelyn<>'Y'"

    dbget.Execute sqlStr

end function



''바로 완료 처리로 진행 할지 여부.
function IsDirectProceedFinish(divcd, Asid, orderserial, byRef EtcStr)
    dim sqlStr
    dim cancelyn, ipkumdiv
    IsDirectProceedFinish = false

    '' currstate:2 업체(물류) 통보
    if (divcd="A008") then
        ''' 취소 Case
        '' 등록된 상품중 업체 확인중 상태가 있으면 접수상태로 진행
        sqlStr = " select count(d." & FIELD_DETAILIDX & ") as invalidcount"
        sqlStr = sqlStr + " from "
        if (GC_IsOLDOrder) then
            sqlStr = sqlStr + " [db_log].[dbo].tbl_old_order_master_2003 m,"
            sqlStr = sqlStr + " [db_log].[dbo].tbl_old_order_detail_2003 d,"
        else
            sqlStr = sqlStr + " " & TABLE_ORDERMASTER & " m,"
            sqlStr = sqlStr + " " & TABLE_ORDERDETAIL & " d,"
        end if
        sqlStr = sqlStr + " " & TABLE_CSDETAIL & " c"
        sqlStr = sqlStr + " where m.orderserial='" + orderserial + "'"
        sqlStr = sqlStr + " and m.orderserial=d.orderserial"
        sqlStr = sqlStr + " and d.itemid<>0"
        sqlStr = sqlStr + " and c.masterid=" + CStr(Asid)
        sqlStr = sqlStr + " and d." & FIELD_DETAILIDX & "=c.orderdetailidx"
        sqlStr = sqlStr + " and d.currstate>=3"
        sqlStr = sqlStr + " and d.cancelyn<>'Y'"

        rsget.Open sqlStr,dbget,1
            IsDirectProceedFinish = (rsget("invalidcount")=0)
        rsget.close

    else

    end if

end function

''검증. 전체 취소 맞는지.
function IsAllCancelRegValid(Asid, orderserial)
    dim sqlStr
    IsAllCancelRegValid = false

    sqlStr = "select count(d." & FIELD_DETAILIDX & ") as cnt"
    if (GC_IsOLDOrder) then
        sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_detail_2003 d"
    else
        sqlStr = sqlStr + " from " & TABLE_ORDERDETAIL & " d"
    end if
    sqlStr = sqlStr + " left join " & TABLE_CSDETAIL & " c"
    sqlStr = sqlStr + "     on c.masterid=" + CStr(Asid)
    sqlStr = sqlStr + "     and c.orderdetailidx=d." & FIELD_DETAILIDX & ""
    sqlStr = sqlStr + " where d.orderserial='" + orderserial + "'"
    sqlStr = sqlStr + " and d.itemid<>0"
    sqlStr = sqlStr + " and d.cancelyn<>'Y'"
    sqlStr = sqlStr + " and d.itemno<>IsNULL(c.regitemno,0)"
''rw sqlStr
    rsget.Open sqlStr,dbget,1
        IsAllCancelRegValid = (rsget("cnt")=0)
    rsget.close

end function

''검증. 부분 취소 맞는지.
function IsPartialCancelRegValid(Asid, orderserial)
    dim sqlStr
    IsPartialCancelRegValid = false

    sqlStr = "select count(d." & FIELD_DETAILIDX & ") as cnt, sum(case when d.itemno=IsNULL(c.regitemno,0) then 1 else 0 end) as Matchcount"
    if (GC_IsOLDOrder) then
        sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_detail_2003 d"
    else
        sqlStr = sqlStr + " from " & TABLE_ORDERDETAIL & " d"
    end if
    sqlStr = sqlStr + " left join " & TABLE_CSDETAIL & " c"
    sqlStr = sqlStr + "     on c.masterid=" + CStr(Asid)
    sqlStr = sqlStr + "     and c.orderdetailidx=d." & FIELD_DETAILIDX & ""
    sqlStr = sqlStr + " where d.orderserial='" + orderserial + "'"
    sqlStr = sqlStr + " and d.itemid<>0"
    sqlStr = sqlStr + " and d.cancelyn<>'Y'"

    rsget.Open sqlStr,dbget,1
        IsPartialCancelRegValid = Not (rsget("cnt")=rsget("Matchcount"))
    rsget.close
end function


''주문 상세 내역이 취소 가능한지 체크 - 출고 완료된 내역이 있는지, 주문건이 취소된내역이 있는지
function IsCancelValidState(Asid, orderserial)
    dim sqlStr

    IsCancelValidState = false

    sqlStr = " select m.cancelyn, m.ipkumdiv, count(d." & FIELD_DETAILIDX & ") as invalidcount, sum(case when d.cancelyn='Y' then 1 else 0 end) as detailcancelcount "
    sqlStr = sqlStr + " from "
    if (GC_IsOLDOrder) then
        sqlStr = sqlStr + " [db_log].[dbo].tbl_old_order_master_2003 m,"
        sqlStr = sqlStr + " [db_log].[dbo].tbl_old_order_detail_2003 d,"
    else
        sqlStr = sqlStr + " " & TABLE_ORDERMASTER & " m,"
        sqlStr = sqlStr + " " & TABLE_ORDERDETAIL & " d,"
    end if
    sqlStr = sqlStr + " " & TABLE_CSDETAIL & " c"
    sqlStr = sqlStr + " where m.orderserial='" + orderserial + "'"
    sqlStr = sqlStr + " and m.orderserial=d.orderserial"
    sqlStr = sqlStr + " and c.masterid=" + CStr(Asid)
    sqlStr = sqlStr + " and d." & FIELD_DETAILIDX & "=c.orderdetailidx"
    sqlStr = sqlStr + " and d.currstate>=7"
    sqlStr = sqlStr + " group by m.cancelyn, m.ipkumdiv"

    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        IsCancelValidState = (rsget("cancelyn")="N") and (rsget("ipkumdiv")<7) and (rsget("invalidcount")<1) and (rsget("detailcancelcount")<1)
    else
        IsCancelValidState = true
    end if
    rsget.close

end function

''강좌확정후 취소/ 회수 접수내역 체크
function IsReturnRegValid(Asid, orderserial,byref ScanErr, upcheMakerid)
    ''  업체배송과 자체배송을 같이 접수하지 못함.
    ''  업체배송이 존재할 경우 MakerID가 1개만 존재 해야함.

    dim sqlStr
    sqlStr = " select count(d." & FIELD_DETAILIDX & ") as cnt, d.isupchebeasong "
    sqlStr = sqlStr + " from "
     if (GC_IsOLDOrder) then
        sqlStr = sqlStr + " [db_log].[dbo].tbl_old_order_master_2003 m,"
        sqlStr = sqlStr + " [db_log].[dbo].tbl_old_order_detail_2003 d,"
    else
        sqlStr = sqlStr + " " & TABLE_ORDERMASTER & " m,"
        sqlStr = sqlStr + " " & TABLE_ORDERDETAIL & " d,"
    end if
    sqlStr = sqlStr + " " & TABLE_CSDETAIL & " c"
    sqlStr = sqlStr + " where m.orderserial='" + orderserial + "'"
    sqlStr = sqlStr + " and m.orderserial=d.orderserial"
    sqlStr = sqlStr + " and d.itemid<>0"
    sqlStr = sqlStr + " and c.masterid=" + CStr(Asid)
    sqlStr = sqlStr + " and d." & FIELD_DETAILIDX & "=c.orderdetailidx"
    sqlStr = sqlStr + " group by d.isupchebeasong"

    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        if (rsget.RecordCount>1) then
            ScanErr = "텐바이텐 배송과 업체배송을 동시에 접수하실 수 없습니다."
        end if
    end if
    rsget.Close

    if ScanErr<>"" then
        IsReturnRegValid = false
        exit function
    end if


    sqlStr = " select count(d." & FIELD_DETAILIDX & ") as cnt, d.makerid "
    sqlStr = sqlStr + " from "
    if (GC_IsOLDOrder) then
        sqlStr = sqlStr + " [db_log].[dbo].tbl_old_order_master_2003 m,"
        sqlStr = sqlStr + " [db_log].[dbo].tbl_old_order_detail_2003 d,"
    else
        sqlStr = sqlStr + " " & TABLE_ORDERMASTER & " m,"
        sqlStr = sqlStr + " " & TABLE_ORDERDETAIL & " d,"
    end if
    sqlStr = sqlStr + " " & TABLE_CSDETAIL & " c"
    sqlStr = sqlStr + " where m.orderserial='" + orderserial + "'"
    sqlStr = sqlStr + " and m.orderserial=d.orderserial"
    sqlStr = sqlStr + " and d.isupchebeasong='Y'"
    sqlStr = sqlStr + " and d.itemid<>0"
    sqlStr = sqlStr + " and c.masterid=" + CStr(Asid)
    sqlStr = sqlStr + " and d." & FIELD_DETAILIDX & "=c.orderdetailidx"
    sqlStr = sqlStr + " group by d.makerid"

    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        if (rsget.RecordCount>1) then
            ScanErr = "업체배송의 경우 각 브랜드 별로 접수하셔야 합니다."
        else
            upcheMakerid = rsget("makerid")
        end if
    end if
    rsget.Close

    if ScanErr<>"" then
        IsReturnRegValid = false
        exit function
    end if

    IsReturnRegValid = true
end function

function IsReturnValidState(Asid, orderserial, byref iScanErr)
    dim sqlStr
    IsReturnValidState = false

    sqlStr = " select ipkumdiv, cancelyn "
    if (GC_IsOLDOrder) then
        sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_master_2003"
    else
        sqlStr = sqlStr + " from " & TABLE_ORDERMASTER & ""
    end if
    sqlStr = sqlStr + " where orderserial='" + orderserial + "'"

    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        cancelyn    = rsget("cancelyn")
        ipkumdiv    = rsget("ipkumdiv")
    end if
    esget.Close

    if (cancelyn<>"N") then Exit function

    IsReturnValidState = true
end function

function setCancelMaster(Asid, orderserial)
    dim sqlStr

    sqlStr = "update " & TABLE_ORDERMASTER & "" + VbCrlf
    sqlStr = sqlStr + " set cancelyn='Y'" + VbCrlf
    '' 취소일 추가
    sqlStr = sqlStr + " ,canceldate=IsNULL(canceldate,getdate())" + VbCrlf
    sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf

    dbget.Execute sqlStr
end function



'' 수량이 같으면 취소 Flag 다르면 수량변경
function setCancelDetail(Asid, orderserial)
    dim sqlStr
    ''취소일 추가
    sqlStr = "update " & TABLE_ORDERDETAIL & "" + VbCrlf
    sqlStr = sqlStr + " set cancelyn='Y'" + VbCrlf
    sqlStr = sqlStr + " ,canceldate=IsNULL(canceldate,getdate())" + VbCrlf
    sqlStr = sqlStr + " from " & TABLE_CSDETAIL & " c" + VbCrlf
    sqlStr = sqlStr + " where " & TABLE_ORDERDETAIL & ".orderserial='" + orderserial + "'" + VbCrlf
    sqlStr = sqlStr + " and c.masterid=" + CStr(Asid)
    sqlStr = sqlStr + " and " & TABLE_ORDERDETAIL & "." & FIELD_DETAILIDX & "=c.orderdetailidx" + VbCrlf
    sqlStr = sqlStr + " and " & TABLE_ORDERDETAIL & ".itemno=c.regitemno" + VbCrlf
    '''sqlStr = sqlStr + " and " & TABLE_ORDERDETAIL & ".itemid<>0"
    '''배송비도 취소?

    dbget.Execute sqlStr

    '' 수량조정 ::: (몇개 만 취소인경우)
    sqlStr = "update " & TABLE_ORDERDETAIL & "" + VbCrlf
    sqlStr = sqlStr + " set itemno=itemno-c.regitemno" + VbCrlf
    'sqlStr = sqlStr + " ,orderdate=getdate()" + VbCrlf
    sqlStr = sqlStr + " from " & TABLE_CSDETAIL & " c" + VbCrlf
    sqlStr = sqlStr + " where " & TABLE_ORDERDETAIL & ".orderserial='" + orderserial + "'" + VbCrlf
    sqlStr = sqlStr + " and c.masterid=" + CStr(Asid)
    sqlStr = sqlStr + " and " & TABLE_ORDERDETAIL & "." & FIELD_DETAILIDX & "=c.orderdetailidx" + VbCrlf
    sqlStr = sqlStr + " and " & TABLE_ORDERDETAIL & ".itemno>c.regitemno" + VbCrlf
    sqlStr = sqlStr + " and " & TABLE_ORDERDETAIL & ".itemid<>0"

    dbget.Execute sqlStr


end function



''주문 마스타 재계산
function recalcuOrderMaster(byVal orderserial)
	dim sqlStr

	sqlStr = "update " & TABLE_ORDERMASTER & "" + VbCrlf
	sqlStr = sqlStr + " set totalsum=IsNULL(T.dtotalsum,0)" + VbCrlf
	''sqlStr = sqlStr + " , totalcost=IsNULL(T.dtotalsum,0)"  + VbCrlf
	sqlStr = sqlStr + " , totalmileage=IsNULL(T.dtotalmileage,0)" + VbCrlf
	sqlStr = sqlStr + " from (" + VbCrlf
	sqlStr = sqlStr + "     select sum(itemcost*itemno) as dtotalsum, sum(mileage*itemno) as dtotalmileage" + VbCrlf
	sqlStr = sqlStr + "     from " & TABLE_ORDERDETAIL & "" + VbCrlf
	sqlStr = sqlStr + "     where orderserial='" + orderserial + "'" + VbCrlf
	sqlStr = sqlStr + "     and cancelyn<>'Y'" + VbCrlf
	sqlStr = sqlStr + " ) T" + VbCrlf
	sqlStr = sqlStr + " where " & TABLE_ORDERMASTER & ".orderserial='" + orderserial + "'" + VbCrlf

	dbget.Execute sqlStr


	sqlStr = "update " & TABLE_ORDERMASTER & "" + VbCrlf
	sqlStr = sqlStr + " set subtotalprice=totalsum-(IsNULL(tencardspend,0) + IsNULL(miletotalprice,0) + IsNULL(spendmembership,0) + IsNULL(allatdiscountprice,0)) "+ VbCrlf
    sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf

    dbget.Execute sqlStr



end function



function updateUserMileage(byVal userid)
	dim sqlStr
	dim totmile, michulgoMile

	'==============================================================
	'// 보너스/사용마일리지 요약 재계산(신규Proc)
	sqlStr = " exec [db_user].[dbo].sp_Ten_ReCalcu_His_BonusMileage '"&userid&"'"
	'dbget.Execute sqlStr
	if (CS_COMPANYID = "thefingers") then
		dbget_CS.Execute  sqlStr
	else
		dbget.Execute  sqlStr
	end if

	'==============================================================
	'주문마일리지 요약 재계산
    sqlStr = " select IsNull(sum(totalmileage), 0) as totmile, IsNull(sum(case when IsNull(sitename, '') <> 'academy' and ipkumdiv < 8 then totalmileage when IsNull(sitename, '') = 'academy' and ipkumdiv < 7 then totalmileage else 0 end),0) as michulgoMile " + VbCrlf
    sqlStr = sqlStr + "     from " & TABLE_ORDERMASTER & "" + VbCrlf
    sqlStr = sqlStr + "     where userid='" + CStr(userid) + "' " + VbCrlf
    sqlStr = sqlStr + "     and sitename in ('" & MAIN_SITENAME1 & "', '" & MAIN_SITENAME2 & "')" + VbCrlf
    sqlStr = sqlStr + "     and cancelyn='N'" + VbCrlf
    sqlStr = sqlStr + "     and ipkumdiv>3" + VbCrlf

    rsget.Open sqlStr,dbget,1
		totmile = rsget("totmile")
		michulgoMile = rsget("michulgoMile")
    rsget.Close

    sqlStr = "update " & TABLE_USER_CURRENT_MILEAGE& "" + VbCrlf
    sqlStr = sqlStr + " set " & FIELD_CURRENT_MILEAGE & "=" & totmile & ", " & FIELD_MICHULGO_MILEAGE & "=" & michulgoMile & " " + VbCrlf
    sqlStr = sqlStr + " where userid='" + CStr(userid) + "' " + VbCrlf

	if (CS_COMPANYID = "thefingers") then
		dbget_CS.Execute  sqlStr
	else
		dbget.Execute  sqlStr
	end if

end function


function ValidDeleteCS(id)
    dim sqlStr
    dim currstate

    ValidDeleteCS = false

    sqlStr = "select * from " & TABLE_CSMASTER & ""
    sqlStr = sqlStr + " where id=" + CStr(id)

    rsget.Open sqlStr,dbget,1
        currstate = rsget("currstate")
    rsget.Close

    If (currstate>="B006") then Exit function

    ValidDeleteCS = true
end function

function DeleteCSProcess(id, finishuserid)
    dim sqlStr, resultCount

    sqlStr = " update " & TABLE_CSMASTER & "" + VbCrlf
    sqlStr = sqlStr + " set deleteyn='Y'" + VbCrlf
    sqlStr = sqlStr + " , finishuser = '" + finishuserid+ "'" + VbCrlf
    sqlStr = sqlStr + " , finishdate = getdate()" + VbCrlf
    sqlStr = sqlStr + " where id=" + CStr(id)
    sqlStr = sqlStr + " and currstate<'B006'"

    dbget.Execute sqlStr, resultCount

    DeleteCSProcess = (resultCount>0)
end function


function CancelProcess(id, orderserial)
    dim IsAllCancel, IsUpdatedMile

    dim sqlStr, userid, ipkumdiv, miletotalprice, tencardspend, allatdiscountprice

    dim refundmileagesum, refundcouponsum, allatsubtractsum
    dim refundbeasongpay, refunditemcostsum, refunddeliverypay
    dim refundadjustpay, canceltotal

    dim detailidx, orgbeasongpay, deliveritemoption, deliverbeasongpay
    dim InsureCd
    dim openMessage

    dim regDetailRows, i
    dim remaintencardspend, gubun01, gubun02

    dim itemid, itemoption, cancelno

    IsAllCancel = IsAllCancelRegValid(id, orderserial)

    sqlStr = " select userid, ipkumdiv, IsNULL(miletotalprice,0) as miletotalprice "
    sqlStr = sqlStr + " ,IsNULL(tencardspend,0) as tencardspend, IsNULL(allatdiscountprice,0) as allatdiscountprice" + VbCrlf
    sqlStr = sqlStr + " ,IsNULL(InsureCd,'') as InsureCd" + VbCrlf
    sqlStr = sqlStr + " from " & TABLE_ORDERMASTER & "" + VbCrlf
    sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf

    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        userid              = rsget("userid")
        miletotalprice      = rsget("miletotalprice")
        tencardspend        = rsget("tencardspend")
        allatdiscountprice  = rsget("allatdiscountprice")
        InsureCd            = rsget("InsureCd")
        ipkumdiv            = rsget("ipkumdiv")
    end if
    rsget.close

    sqlStr = " select r.*, a.gubun01, a.gubun02 from " & TABLE_CSMASTER & " a"
    sqlStr = sqlStr + " , " & TABLE_CS_REFUND & " r"
    sqlStr = sqlStr + " where a.id=" + CStr(id)
    sqlStr = sqlStr + " and a.id=r.asid"
    sqlStr = sqlStr + " and a.deleteyn='N'"
    sqlStr = sqlStr + " and a.currstate<>'B007'"


    rsget.Open sqlStr,dbget,1

    if Not rsget.Eof then
        refundmileagesum    = rsget("refundmileagesum")
        refundcouponsum     = rsget("refundcouponsum")
        allatsubtractsum    = rsget("allatsubtractsum")

        refunditemcostsum   = rsget("refunditemcostsum")

        refundbeasongpay    = rsget("refundbeasongpay")
        refunddeliverypay   = rsget("refunddeliverypay")
        refundadjustpay     = rsget("refundadjustpay")
        canceltotal         = rsget("canceltotal")
        gubun01             = rsget("gubun01")
        gubun02             = rsget("gubun02")

    else
        refundmileagesum    = 0
        refundcouponsum     = 0
        allatsubtractsum    = 0

        refunditemcostsum   = 0

        refundbeasongpay    = 0
        refunddeliverypay   = 0
        refundadjustpay     = 0
        canceltotal         = 0
    end if
    rsget.close

'' 마일리지 환급

    IsUpdatedMile = false

    if (userid<>"") and (IsAllCancel) and (miletotalprice<>0) then
        '' 전체 취소인경우 주문건 취소로 jukyocd : 2 상품구매, 3 : 부분취소시 환급마일리지
        sqlStr = " update " & TABLE_MILEAGELOG & " " + VbCrlf
        sqlStr = sqlStr + " set deleteyn='Y' " + VbCrlf
        sqlStr = sqlStr + " where userid='" + userid + "'" + VbCrlf
        sqlStr = sqlStr + " and orderserial='" + orderserial + "'" + VbCrlf
        sqlStr = sqlStr + " and deleteyn='N'" + VbCrlf
        sqlStr = sqlStr + " and jukyocd in ('2','3')" + VbCrlf

		if (CS_COMPANYID = "thefingers") then
			dbget_CS.Execute  sqlStr
		else
			dbget.Execute  sqlStr
		end if

        IsUpdatedMile = true

        if openMessage="" then
            openMessage = openMessage + "사용 마일리지 환급 : " & miletotalprice
        else
            openMessage = openMessage + VbCrlf + "사용 마일리지 환급 : " & miletotalprice
        end if

    end if

    if (userid<>"") and (Not IsAllCancel) and (refundmileagesum<>0) then
        '' 부분 취소인데 마일리지 환급할 경우.
        sqlStr = " update " & TABLE_ORDERMASTER & "" + VbCrlf
        sqlStr = sqlStr + " set miletotalprice=miletotalprice + " + CStr(refundmileagesum) + VbCrlf
        sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf

        dbget.Execute sqlStr


        sqlStr = " insert into " & TABLE_MILEAGELOG & " " + VbCrlf
        sqlStr = sqlStr + " (userid, mileage, jukyocd, jukyo, orderserial, deleteyn) " + VbCrlf
        sqlStr = sqlStr + " values ("
        sqlStr = sqlStr + " '" + userid + "'"
        sqlStr = sqlStr + " ," + CStr(refundmileagesum*-1) + ""
        sqlStr = sqlStr + " ,'3'"
        sqlStr = sqlStr + " ,'상품구매 취소 환급'"
        sqlStr = sqlStr + " ,'" + orderserial + "'"
        sqlStr = sqlStr + " ,'N'"
        sqlStr = sqlStr + " )"

		if (CS_COMPANYID = "thefingers") then
			dbget_CS.Execute  sqlStr
		else
			dbget.Execute  sqlStr
		end if

        IsUpdatedMile = true

        if openMessage="" then
            openMessage = openMessage + "사용 마일리지 환급 : " & refundmileagesum
        else
            openMessage = openMessage + VbCrlf + "사용 마일리지 환급 : " & refundmileagesum
        end if
    end if


'' 할인권 환급
    if (IsAllCancel) and (tencardspend<>0) then
        sqlStr = " update [db_user].[dbo].tbl_user_coupon "   + VbCrlf
	    sqlStr = sqlStr + " set isusing='N' "                   + VbCrlf
	    sqlStr = sqlStr + " where orderserial = '" + CStr(orderserial) + "' and deleteyn = 'N' and isusing = 'Y' "

	    dbget_CS.Execute sqlStr

	    if openMessage="" then
            openMessage = openMessage + "사용 보너스쿠폰 환급"
        else
            openMessage = openMessage + VbCrlf + "사용 보너스쿠폰 환급"
        end if
    end if

    if (Not IsAllCancel) and (refundcouponsum<>0) then
         '' 부분 취소인경우 - 환급한 만큼 깜..
        sqlStr = " update " & TABLE_ORDERMASTER & "" + VbCrlf
        sqlStr = sqlStr + " set tencardspend=tencardspend + " + CStr(refundcouponsum) + VbCrlf
        sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf

        dbget.Execute sqlStr

        ''전체 환급인 경우만 쿠폰을 돌려줌
        sqlStr = "select IsNULL(tencardspend,0) as tencardspend from " & TABLE_ORDERMASTER & "" + VbCrlf
        sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf

        rsget.Open sqlStr,dbget,1
            remaintencardspend = rsget("tencardspend")
        rsget.close

        ''원래 할인권 사용액이 있고, 남은 쿠폰사용액이 없을경우 전체  환급
        if (tencardspend>0) then
            if (remaintencardspend=0)   then
                sqlStr = " update [db_user].[dbo].tbl_user_coupon "   + VbCrlf
            	sqlStr = sqlStr + " set isusing='N' "                   + VbCrlf
            	sqlStr = sqlStr + " where orderserial = '" + CStr(orderserial) + "' and deleteyn = 'N' and isusing = 'Y' "

            	dbget_CS.Execute sqlStr

            	if openMessage="" then
                    openMessage = openMessage + "사용 할인권  환급"
                else
                    openMessage = openMessage + VbCrlf + "사용 할인권  환급"
                end if
            else
                ''(또는, %쿠폰인 경우 공통,단순변심인 경우 제외하고 환급해줌./ 부분취소 ) C004 CD01
                if (ipkumdiv>3) and (Not ((gubun01="C004") and (gubun02="CD01"))) then
                    sqlStr = " update [db_user].[dbo].tbl_user_coupon "   + VbCrlf
                	sqlStr = sqlStr + " set isusing='N' "                   + VbCrlf
                	sqlStr = sqlStr + " where orderserial = '" + CStr(orderserial) + "' and deleteyn = 'N' and isusing = 'Y' "
                	sqlStr = sqlStr + " and coupontype=1"

                	dbget_CS.Execute sqlStr

                	if openMessage="" then
                        openMessage = openMessage + "사용 할인권  환급."
                    else
                        openMessage = openMessage + VbCrlf + "사용 할인권  환급."
                    end if
                end if
            end if
        end if



    end if



    '' 올엣카드 할인 차감
    if (IsAllCancel) and (allatdiscountprice<>0) then

    end if

    if (Not IsAllCancel) and (allatsubtractsum<>0) then
        sqlStr = " update " & TABLE_ORDERMASTER & "" + VbCrlf
        sqlStr = sqlStr + " set allatdiscountprice=allatdiscountprice + " + CStr(allatsubtractsum) + VbCrlf
        sqlStr = sqlStr + " where orderserial='" + orderserial + "'" + VbCrlf

        dbget.Execute sqlStr

        if openMessage="" then
            openMessage = openMessage + "올엣카드 할인 차감 : " & allatsubtractsum
        else
            openMessage = openMessage + VbCrlf + "올엣카드 할인 차감 : " & allatsubtractsum
        end if
    end if


'' 배송비 재계산. : 현재 배송비와 다를경우만. 부분 취소인 경우만. :: 업체 개별 배송비로 수정
    dim detailRefundBeasongPay
    detailRefundBeasongPay = 0
    sqlStr = " select IsNULL(sum(itemcost),0) as detailRefundBeasongPay from " & TABLE_CSDETAIL & ""
    sqlStr = sqlStr + " where masterid=" + CStr(id)
    sqlStr = sqlStr + " and itemid=0"
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        detailRefundBeasongPay = rsget("detailRefundBeasongPay")
    end if
    rsget.Close

    if (Not IsAllCancel) and (refundbeasongpay<>0) then
        orgbeasongpay =0

        ''기본배송비.
        sqlStr = " select * from " & TABLE_ORDERDETAIL & " "
        sqlStr = sqlStr + " where orderserial='" + orderserial + "'"
        sqlStr = sqlStr + " and itemid=0"
        sqlStr = sqlStr + " and IsNULL(makerid,'')=''"
        sqlStr = sqlStr + " and cancelyn<>'Y'"

        rsget.Open sqlStr,dbget,1
            detailidx     = rsget("idx")
            orgbeasongpay = rsget("itemcost")
        rsget.Close

        ''원래 텐배송 비가 >0 이고, 환불배송비가=텐배송비고,
'response.write "orgbeasongpay=" & orgbeasongpay & "<br>"
'response.write "refundbeasongpay=" & refundbeasongpay & "<br>"
'response.write "detailRefundBeasongPay=" & detailRefundBeasongPay & "<br>"

        if (orgbeasongpay>0) and (orgbeasongpay-refundbeasongpay=0) and (refundbeasongpay-detailRefundBeasongPay>0) then
             sqlStr = " update " & TABLE_ORDERDETAIL & " "
             sqlStr = sqlStr + " set itemoption='0000'"
             sqlStr = sqlStr + " ,itemcost=0"
             sqlStr = sqlStr + " where idx=" + CStr(detailidx)

             dbget.Execute sqlStr
             response.write   "원 기본 배송비(" & orgbeasongpay & ") 0 원 처리"
        else

        end if
    end if

    if (IsAllCancel) then
	    ''전체 취소인경우
	    '' 주문  master 취소 변경
	    call setCancelMaster(id, orderserial)

	    if openMessage="" then
            openMessage = openMessage + "주문취소 완료"
        else
            openMessage = openMessage + VbCrlf + "주문취소 완료"
        end if
	else
	    ''부분 취소인경우
	    '' 주문  detail 취소 변경
	    call setCancelDetail(id, orderserial)

	    call reCalcuOrderMaster(orderserial)

	    if openMessage="" then
            openMessage = openMessage + "주문부분취소 완료"
        else
            openMessage = openMessage + VbCrlf + "주문부분취소 완료"
        end if
	end if

    ''마일리지는 주문건 취소 후 재계산해야함.
    if (userid<>"") then
        Call updateUserMileage(userid)
    end if



    ''전자보증서 발급된 경우 취소
    if (InsureCd="0") then
        Call UsafeCancel(orderserial)
    end if

    ''재고 및 한정수량 조절(2007-09-01 서동석 추가)
    ''Call LimitItemRecover(orderserial) : 기존
    if (IsAllCancel) then
	    ''전체 취소인경우
	    'sqlStr = " exec [db_summary].[dbo].sp_Ten_RealtimeStock_cancelOrderAll '" & orderserial & "'"
	    'dbget.Execute sqlStr

	    Call LimitItemRecoverOnOrderCancel(orderserial)

	else

	    ''부분 취소인경우
	    sqlStr = " select itemid,itemoption,regitemno "
        sqlStr = sqlStr & " from " & TABLE_CSDETAIL & " "
        sqlStr = sqlStr & " where masterid=" & id
        sqlStr = sqlStr & " and orderserial='" & orderserial & "'"

        rsget.Open sqlStr,dbget,1
        if Not rsget.Eof then
            regDetailRows = rsget.getRows()
        end if
        rsget.Close

        if IsArray(regDetailRows) then
            for i=0 to UBound(regDetailRows,2)
    	        'sqlStr = " exec [db_summary].[dbo].sp_Ten_RealtimeStock_cancelOrderPartial '" & orderserial & "'," & regDetailRows(0,i) & ",'" & regDetailRows(1,i) & "'," & regDetailRows(2,i)

                Call LimitItemRecoverOnItemCancel(orderserial, regDetailRows(0,i), regDetailRows(1,i), regDetailRows(2,i))
                'dbget.Execute sqlStr


            Next
        end if
	end if

    ''전자보증서 발급된 경우 취소
    if (InsureCd="0") then
        Call UsafeCancel(orderserial)
    end if

    if (openMessage<>"") then
        call AddCustomerOpenContents(id, openMessage)
    end if
end function

function CheckRefundFinish(id, orderserial,byRef RefreturnMethod,byRef Refrealrefund)
    dim sqlStr
    dim returnmethod, refundrequire, refundresult
    dim realrefund ,userid

    realrefund = 0

    sqlStr = "select r.*, a.userid from "
    sqlStr = sqlStr + " " & TABLE_CS_REFUND & " r,"
    sqlStr = sqlStr + " " & TABLE_CSMASTER & " a"
    sqlStr = sqlStr + " where r.asid=" + CStr(id)
    sqlStr = sqlStr + " and r.asid=a.id"

    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        returnmethod    = rsget("returnmethod")
        refundrequire   = rsget("refundrequire")
        refundresult    = rsget("refundresult")
        userid          = rsget("userid")
        realrefund      = refundrequire-refundresult

        RefreturnMethod = returnmethod
        Refrealrefund   = realrefund
    end if
    rsget.Close

    ''마일리지 환급
    if (returnmethod="R900") then
        sqlStr = "insert into " & TABLE_MILEAGELOG & ""
        sqlStr = sqlStr + " (userid, mileage, jukyocd, jukyo, orderserial, deleteyn)"
        sqlStr = sqlStr + " values('" + userid + "'," + CStr(realrefund) + ",'999','구매환불','" + orderserial + "','N')"

		if (CS_COMPANYID = "thefingers") then
			dbget_CS.Execute  sqlStr
		else
			dbget.Execute  sqlStr
		end if

        sqlStr = " update " & TABLE_CS_REFUND & ""
        sqlStr = sqlStr + " set refundresult=" + CStr(realrefund)
        sqlStr = sqlStr + " where asid=" + CStr(id)
        dbget.Execute sqlStr

        call updateUserMileage(userid)

        call AddCustomerOpenContents(id, "마일리지 환급 완료: " & CStr(realrefund))
    elseif (returnmethod<>"R000") then
        sqlStr = " update " & TABLE_CS_REFUND & ""
        sqlStr = sqlStr + " set refundresult=" + CStr(realrefund)
        sqlStr = sqlStr + " where asid=" + CStr(id)
        dbget.Execute sqlStr

        call AddCustomerOpenContents(id, "환불(취소) 완료: " & CStr(realrefund))
    end if

end function

function CheckNRegRefund(id, orderserial, reguserid)
    '' A003 환불요청 , A005 외부몰환불요청 , A007 신용카드/실시간이체취소요청
    '' Result -1, or newAsID
    dim divcd
    dim returnmethod, gubun01, gubun02

    dim sqlStr, RegDivCd
    dim title, contents_jupsu
    dim NewRegedID

    CheckNRegRefund = -1

    sqlStr = " select a.divcd, a.gubun01, a.gubun02"
    sqlStr = sqlStr + " , r.returnmethod "
    sqlStr = sqlStr + " from " & TABLE_CSMASTER & " a"
    sqlStr = sqlStr + " left join " & TABLE_CS_REFUND & " r"
    sqlStr = sqlStr + "     on a.id=r.asid"
    sqlStr = sqlStr + " where a.id=" + CStr(id)
    sqlStr = sqlStr + " and a.deleteyn='N'"
    sqlStr = sqlStr + " and a.currstate<>'B007'"

    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        divcd                = rsget("divcd")
        returnmethod         = rsget("returnmethod")
        gubun01              = rsget("gubun01")
        gubun02              = rsget("gubun02")

        if IsNULL(returnmethod) then returnmethod=""
    end if
    rsget.close

    'R007 무통장환불
    'R020 실시간이체취소
	'R022 실시간이체부분취소
    'R050 입점몰결제 취소
    'R080 올엣카드취소
    'R100 신용카드취소
	'R120 신용카드부분취소
    'R400 휴대폰취소
    'R900 마일리지로환불

    if (returnmethod="R000") or (returnmethod="") or (Trim(returnmethod)="") then
        Exit function
	elseif (returnmethod="R020") or (returnmethod="R022") or (returnmethod="R080") or (returnmethod="R100") or (returnmethod="R120") or (returnmethod="R400") then
        RegDivCd = "A007"

        if (returnmethod="R020") then
            title = "실시간이체취소"
        elseif (returnmethod="R022") then
            title = "실시간이체부분취소"
        elseif (returnmethod="R080") then
            title = "올엣카드취소"
        elseif (returnmethod="R100") then
            title = "신용카드취소"
        elseif (returnmethod="R120") then
            title = "신용카드부분취소"
		elseif (returnmethod="R400") then
            title = "휴대폰취소"
        end if

        contents_jupsu = paygateTid
    elseif (returnmethod="R050") then
        RegDivCd = "A005"
        title = "입점몰결제 취소"

        ''contents_jupsu = The Ext site name
    elseif (returnmethod="R900") then
        RegDivCd = "A003"
        title = "마일리지 환불"

    elseif (returnmethod<>"") then
        RegDivCd = "A003"
        title = "무통장환불"
        contents_jupsu = ""
    end if

    if (divcd="A008") then
        title = "주문 취소 후 " + title + " 요청"
    elseif (divcd="A004") then
        title = "강좌확정 후 일부취소 처리 후 " + title + " 요청"
    elseif (divcd="A010") then
        title = "회수 처리 후 " + title + " 요청"
    end if

    if (RegDivCd<>"") then
        NewRegedID =  RegCSMaster(RegDivCd, orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)

        sqlStr = " insert into " & TABLE_CS_REFUND & ""
        sqlStr = sqlStr + " (asid, returnmethod, refundrequire, rebankname, rebankaccount, "
        sqlStr = sqlStr + " rebankownername, paygateTid, paygateresultTid, paygateresultMsg, encmethod, encaccount) "
        sqlStr = sqlStr + " select " + CStr(NewRegedID)
        sqlStr = sqlStr + " ,returnmethod, refundrequire, rebankname, rebankaccount, "
        sqlStr = sqlStr + " rebankownername, paygateTid, paygateresultTid, paygateresultMsg, encmethod, encaccount "
        sqlStr = sqlStr + " from " & TABLE_CS_REFUND & ""
        sqlStr = sqlStr + " where asid=" + CStr(id)

''        sqlStr = " update " & TABLE_CS_REFUND & ""
''        sqlStr = sqlStr + " set asid=" + CStr(NewRegedID)
''        sqlStr = sqlStr + " where asid=" + CStr(id)

        dbget.Execute sqlStr

        CheckNRegRefund = NewRegedID
    end if
end function


function CheckNAddMinusOrder(id, orderserial, reguserid,byref MinusOrderserial, byref ErrStr)
    dim sqlStr
    dim orgsubtotalprice, currjupsusum, orgidx
    dim userid, sitename
    dim AsDetailExists
    dim MinusMiletotalprice
    dim InvalidCnt
    ''dim totalpreminussum
    ''totalpreminussum = 0

    orgidx           = 0
    orgsubtotalprice = 0
    currjupsusum = 0
    AsDetailExists = false
    MinusMiletotalprice = 0

''   총금액으로 체크 사용안함.
'    sqlStr = "select sum(subtotalprice*-1) as totalpreminussum from " & TABLE_ORDERMASTER & ""
'    sqlStr = sqlStr  + " where linkorderserial='" + orderserial + "'"
'    sqlStr = sqlStr  + " and jumundiv='9'"
'    sqlStr = sqlStr  + " and cancelyn='N'"
'
'    rsget.Open sqlStr,dbget,1
'    if Not rsget.Eof then
'        totalpreminussum    = rsget("totalpreminussum")
'    end if
'    rsget.Close

	'TODO : 재료비환불 중복체크 필요
    ''접수되는 내역보다 기존 마이너스+ 추가 마이너스  합계가 큰지 체크 (중복접수)
    if (GC_IsOLDOrder) then
        '' 과거 주문인 경우.. Skip
        InvalidCnt = 0
    else
        sqlStr = " exec " & PROC_MINUS_ORDER_INVALID_CNT & " " & CStr(id) & ",'" & orderserial & "'"

        rsget.Open sqlStr, dbget, 1
        if Not (rsget.Eof) then
            InvalidCnt = rsget("InvalidCnt")
        end if
        rsget.Close

    end if

    if (InvalidCnt>0) then
        CheckNAddMinusOrder = false
        ErrStr = "마이너스 주문 상품 합계가 원 상품보타 클 수 있습니다.\n(중복 접수되었을 수 있습니다. 명령이 취소 됩니다.)"
        exit function
    end if

    ''원주문건 조회
    sqlStr = " select idx, subtotalprice, userid, sitename "
    if (GC_IsOLDOrder) then
        sqlStr = sqlStr  + " from [db_log].[dbo].tbl_old_order_master_2003"
    else
        sqlStr = sqlStr  + " from " & TABLE_ORDERMASTER & ""
    end if
    sqlStr = sqlStr  + " where orderserial='" + orderserial + "'"
    sqlStr = sqlStr  + " and cancelyn='N'"

    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        orgidx              = rsget("idx")
        orgsubtotalprice    = rsget("subtotalprice")
        userid              = rsget("userid")
        sitename            = rsget("sitename")
    end if
    rsget.Close

    if (orgidx=0) then
        CheckNAddMinusOrder = false
        ErrStr = "원 주문건이 존재하지 않습니다."
        exit function
    end if


    ''as_detail에 상품내역 있는지 체크
    sqlStr = " select count(*) as cnt from" & Vbcrlf
    sqlStr = sqlStr  + " " & TABLE_CSMASTER & " a," & Vbcrlf
    sqlStr = sqlStr  + " " & TABLE_CSDETAIL & " d" & Vbcrlf
    sqlStr = sqlStr  + " where a.id=" & CStr(id) & Vbcrlf
    sqlStr = sqlStr  + " and a.id=d.masterid" & Vbcrlf
    sqlStr = sqlStr  + " and a.orderserial='" + orderserial + "'"

    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        AsDetailExists    = rsget("cnt")>0
    end if
    rsget.Close

    if (Not AsDetailExists) then
        CheckNAddMinusOrder = false
        ErrStr = "강좌확정 후 일부취소 주문건 상세내역이 없습니다. - 관리자 문의요망"
        exit function
    end if

    MinusOrderSerial =  AddMinusOrder(id, orderserial)

    if (MinusOrderSerial="") then
        CheckNAddMinusOrder = false
        ErrStr = "강좌확정 후 일부취소 주문건 생성 실패 - 반드시! 관리자 문의요망."
        exit function
    end if



    sqlStr = " select IsNULL(subtotalprice*-1,0) as subtotalprice, IsNULL(miletotalprice,0) as miletotalprice from " & TABLE_ORDERMASTER & ""
    sqlStr = sqlStr + " where orderserial='" + MinusOrderSerial + "'"

    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
       ''totalpreminussum = totalpreminussum + rsget("subtotalprice")

       ''강좌확정후 취소 환급 마일리지
       MinusMiletotalprice = rsget("miletotalprice")
    end if
    rsget.Close

'    if (totalpreminussum>orgsubtotalprice) then
'        CheckNAddMinusOrder = false
'        ErrStr = "기존 마이너스 주문 합보다 원주문 금액이 작습니다.(중복 접수되었을 수 있습니다.)"
'        exit function
'    end if

    ''마일리지 재계산
    if (userid<>"") and ((sitename=MAIN_SITENAME1) or (sitename=MAIN_SITENAME2)) then

        ''강좌확정후 취소 환급 마일리지 추가
        if (MinusMiletotalprice<>0) then
            sqlStr = "insert into " & TABLE_MILEAGELOG & "(userid,mileage,jukyocd,jukyo,orderserial)" + vbCrlf
			sqlStr = sqlStr + " values('" + CStr(userid) + "'," + CStr(-1*CLng(MinusMiletotalprice)) + ",'02','강좌확정 후 일부취소 환급','" + MinusOrderSerial + "')"

			if (CS_COMPANYID = "thefingers") then
				dbget_CS.Execute  sqlStr
			else
				dbget.Execute  sqlStr
			end if

        end if

        Call updateUserMileage(userid)
    end if

    CheckNAddMinusOrder = true
end function

function AddMinusOrder(id, orderserial)
    dim sqlStr
    dim iid
    dim rndjumunno
    dim neworderserial

    dim subtotalprice, miletotalprice, tencardspend, spendmembership, allatdiscountprice
    sqlStr = " select subtotalprice, IsNULL(miletotalprice,0) as miletotalprice,"
    sqlStr = sqlStr + " IsNULL(tencardspend,0) as tencardspend, IsNULL(spendmembership,0) as spendmembership,"
    sqlStr = sqlStr + " IsNULL(allatdiscountprice,0) as allatdiscountprice "
    if (GC_IsOLDOrder) then
        sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_master_2003 m"
    else
        sqlStr = sqlStr + " from " & TABLE_ORDERMASTER & " m"
    end if
    sqlStr = sqlStr + " where orderserial='" + orderserial + "'"

    rsget.Open sqlStr,dbget,1
        subtotalprice       = rsget("subtotalprice")
        miletotalprice      = rsget("miletotalprice")
        tencardspend        = rsget("tencardspend")
        spendmembership     = rsget("spendmembership")
        allatdiscountprice  = rsget("allatdiscountprice")
    rsget.close



    dim refundmileagesum, refundcouponsum, allatsubtractsum, refunditemcostsum
    dim refundbeasongpay, refunddeliverypay, refundadjustpay, canceltotal
    dim refundmatdiv

    ''쿠폰 마일리지 환급 계산
    'refundmatdiv - 재료비 환불방식(1:10%, 2:20% 차감)
    sqlStr = " select r.* from " & TABLE_CSMASTER & " a"
    sqlStr = sqlStr + " , " & TABLE_CS_REFUND & " r"
    sqlStr = sqlStr + " where a.id=" + CStr(id)
    sqlStr = sqlStr + " and a.id=r.asid"
    sqlStr = sqlStr + " and a.deleteyn='N'"
    sqlStr = sqlStr + " and a.currstate<>'B007'"
    sqlStr = sqlStr + " and r.returnmethod<>'R000'"

    rsget.Open sqlStr,dbget,1

    if Not rsget.Eof then
        refundrequire       = rsget("refundrequire")
        refundmileagesum    = rsget("refundmileagesum")
        refundcouponsum     = rsget("refundcouponsum")
        allatsubtractsum    = rsget("allatsubtractsum")

        refunditemcostsum   = rsget("refunditemcostsum")

        refundbeasongpay    = rsget("refundbeasongpay")
        refunddeliverypay   = rsget("refunddeliverypay")
        refundadjustpay     = rsget("refundadjustpay")
        canceltotal         = rsget("canceltotal")

        refundmatdiv        = rsget("refundmatdiv")


    else
        refundrequire       = 0
        refundmileagesum    = 0
        refundcouponsum     = 0
        allatsubtractsum    = 0

        refunditemcostsum   = 0

        refundbeasongpay    = 0
        refunddeliverypay   = 0
        refundadjustpay     = 0
        canceltotal         = 0

        refundmatdiv        = ""
    end if
    rsget.Close

    ''환불 상세 내역이 없을 수 있음
    if (subtotalprice=refundrequire) then
        refundmileagesum    = miletotalprice * -1
        refundcouponsum     = tencardspend * -1
        allatsubtractsum    = allatdiscountprice * -1
    end if


	Randomize
	rndjumunno = CLng(Rnd * 100000) + 1
	rndjumunno = CStr(rndjumunno)

	sqlStr = "select * from " & TABLE_ORDERMASTER & " where 1=0"
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("orderserial") = rndjumunno
	rsget("jumundiv") = "9"
	rsget("userid") = ""
	rsget("accountname") = ""
	rsget("accountdiv") = "7"
	rsget("sitename") = ""

	if (CStr(refundmatdiv) = "1") then
		rsget("goodsnames") = "강좌 마이너스 주문(10% 차감)"
	elseif (CStr(refundmatdiv) = "2") then
		rsget("goodsnames") = "강좌 마이너스 주문(20% 차감)"
	else
		rsget("goodsnames") = "강좌 마이너스 주문(강좌료)"
	end if

	rsget.update
	    iid = rsget("idx")
	rsget.close

	neworderserial = Mid(replace(CStr(DateSerial(Year(now),month(now),Day(now))),"-",""),3,256)
	neworderserial = neworderserial & Format00(5,Right(CStr(iid),5))
	neworderserial = "B" & Right(neworderserial, (Len(neworderserial) - 1))

    sqlStr = "update " & TABLE_ORDERMASTER & "" & vbCrlf
    sqlStr = sqlStr + " set orderserial='" + neworderserial + "'" & vbCrlf
    sqlStr = sqlStr + " where idx=" + CStr(iid)

    dbget.Execute sqlStr


    sqlStr = "update " & TABLE_ORDERMASTER & "" & vbCrlf
	sqlStr = sqlStr + " set userid=O.userid" & vbCrlf
	sqlStr = sqlStr + " ,accountname=O.accountname" & vbCrlf
	sqlStr = sqlStr + " ,accountdiv=O.accountdiv" & vbCrlf
	sqlStr = sqlStr + " ,ipkumdiv='8'" & vbCrlf
	sqlStr = sqlStr + " ,ipkumdate=getdate()" & vbCrlf
	sqlStr = sqlStr + " ,regdate=getdate()" & vbCrlf
	'사용안함
	'sqlStr = sqlStr + " ,beadaldiv=O.beadaldiv" & vbCrlf
	sqlStr = sqlStr + " ,beadaldate=getdate()" & vbCrlf
	sqlStr = sqlStr + " ,buyname=O.buyname" & vbCrlf
	sqlStr = sqlStr + " ,buyphone=O.buyphone" & vbCrlf
	sqlStr = sqlStr + " ,buyhp=O.buyhp" & vbCrlf
	sqlStr = sqlStr + " ,buyemail=O.buyemail" & vbCrlf
	sqlStr = sqlStr + " ,reqname=O.reqname" & vbCrlf
	sqlStr = sqlStr + " ,reqzipcode=O.reqzipcode" & vbCrlf
	sqlStr = sqlStr + " ,reqaddress=O.reqaddress" & vbCrlf
	sqlStr = sqlStr + " ,reqphone=O.reqphone" & vbCrlf
	sqlStr = sqlStr + " ,reqhp=O.reqhp" & vbCrlf
	sqlStr = sqlStr + " ,comment='원주문번호:" + orderserial +"'" & vbCrlf
	sqlStr = sqlStr + " ,linkorderserial=O.orderserial" & vbCrlf
	sqlStr = sqlStr + " ,deliverno=''" & vbCrlf
	sqlStr = sqlStr + " ,sitename=O.sitename" & vbCrlf
	sqlStr = sqlStr + " ,discountrate=O.discountrate" & vbCrlf
	sqlStr = sqlStr + " ,subtotalprice=O.subtotalprice" & vbCrlf
	sqlStr = sqlStr + " ,miletotalprice=" & CStr(refundmileagesum) & vbCrlf
	sqlStr = sqlStr + " ,tencardspend=" & CStr(refundcouponsum) & vbCrlf
	sqlStr = sqlStr + " ,spendmembership=0" & vbCrlf
	sqlStr = sqlStr + " ,allatdiscountprice=" & CStr(allatsubtractsum) & vbCrlf
	sqlStr = sqlStr + " ,rduserid=O.rduserid" & vbCrlf
	sqlStr = sqlStr + " ,sentenceidx=O.sentenceidx" & vbCrlf
	sqlStr = sqlStr + " ,reqzipaddr=O.reqzipaddr" & vbCrlf
	sqlStr = sqlStr + " ,rdsite=O.rdsite" & vbCrlf
	sqlStr = sqlStr + " ,pggubun=O.pggubun" & vbCrlf
	sqlStr = sqlStr + " ,bcpnidx=O.bcpnidx" & vbCrlf

	if (GC_IsOLDOrder) then
	    sqlStr = sqlStr + " from (select top 1 * from [db_log].[dbo].tbl_old_order_master_2003 where orderserial='" + orderserial + "') O" & vbCrlf
	else
	    sqlStr = sqlStr + " from (select top 1 * from " & TABLE_ORDERMASTER & " where orderserial='" + orderserial + "') O" & vbCrlf
	end if
	sqlStr = sqlStr + " where " & TABLE_ORDERMASTER & ".idx=" + CStr(iid)

	dbget.Execute sqlStr

	''원배송비 환급 있을경우
'	if (refundbeasongpay<>0) then
'	    sqlStr = "insert into " & TABLE_ORDERDETAIL & ""
'	    sqlStr = sqlStr + " (masteridx, orderserial,itemid,itemoption,itemno," & vbCrlf
'	    sqlStr = sqlStr + " itemcost," & FIELD_ITEMVAT & ",mileage,reducedPrice,itemname," & vbCrlf
'        sqlStr = sqlStr + " itemoptionname,makerid,buycash,vatinclude,isupchebeasong,issailitem,oitemdiv,currstate,beasongdate,upcheconfirmdate)" & vbCrlf
'        sqlStr = sqlStr + " select " & CStr(iid) & vbCrlf
'        sqlStr = sqlStr + " ,'" & neworderserial & "'" & vbCrlf
'        sqlStr = sqlStr + " ,d.itemid" & vbCrlf
'        sqlStr = sqlStr + " ,d.itemoption" & vbCrlf
'        sqlStr = sqlStr + " ,d.itemno*-1" & vbCrlf
'        sqlStr = sqlStr + " ,d.itemcost" & vbCrlf
'        sqlStr = sqlStr + " ,d." & FIELD_ITEMVAT & "" & vbCrlf
'        sqlStr = sqlStr + " ,d.mileage" & vbCrlf
'        sqlStr = sqlStr + " ,d.reducedPrice" & vbCrlf
'        sqlStr = sqlStr + " ,d.itemname" & vbCrlf
'        sqlStr = sqlStr + " ,d.itemoptionname" & vbCrlf
'        sqlStr = sqlStr + " ,d.makerid" & vbCrlf
'        sqlStr = sqlStr + " ,d.buycash" & vbCrlf
'        sqlStr = sqlStr + " ,d.vatinclude" & vbCrlf
'        sqlStr = sqlStr + " ,d.isupchebeasong" & vbCrlf
'        sqlStr = sqlStr + " ,d.issailitem" & vbCrlf
'        sqlStr = sqlStr + " ,d.oitemdiv" & vbCrlf
'        sqlStr = sqlStr + " ,'7'" & vbCrlf
'        sqlStr = sqlStr + " ,getdate()" & vbCrlf
'        sqlStr = sqlStr + " ,getdate()" & vbCrlf
'        sqlStr = sqlStr + " from " & TABLE_ORDERDETAIL & " d" & vbCrlf
'        sqlStr = sqlStr + " where d.orderserial='" & orderserial & "'"  & vbCrlf
'        sqlStr = sqlStr + " and d.itemid=0" & vbCrlf
'        sqlStr = sqlStr + " and d.cancelyn<>'Y'"
'
'        dbget.Execute sqlStr
'	end if

	''강좌확정후 취소 상세내역
	sqlStr = "insert into " & TABLE_ORDERDETAIL & ""
	sqlStr = sqlStr + " (masteridx, orderserial,itemid,itemoption,itemno " & vbCrlf
    sqlStr = sqlStr + " ,itemname," & vbCrlf
    sqlStr = sqlStr + " itemoptionname,makerid,vatinclude,isupchebeasong,issailitem,oitemdiv,omwdiv,odlvType,currstate,beasongdate,upcheconfirmdate, matinclude_yn, itemcost, mileage, reducedPrice, buycash, matcostadded, matbuycashadded, couponnotasigncost)" & vbCrlf
    sqlStr = sqlStr + " select " & CStr(iid) & vbCrlf
    sqlStr = sqlStr + " ,'" & neworderserial & "'" & vbCrlf
    sqlStr = sqlStr + " ,d.itemid" & vbCrlf
    sqlStr = sqlStr + " ,d.itemoption" & vbCrlf
    sqlStr = sqlStr + " ,J.confirmitemno*-1" & vbCrlf
    sqlStr = sqlStr + " ,d.itemname" & vbCrlf
   sqlStr = sqlStr + " ,d.itemoptionname" & vbCrlf
    sqlStr = sqlStr + " ,d.makerid" & vbCrlf
    sqlStr = sqlStr + " ,d.vatinclude" & vbCrlf
    sqlStr = sqlStr + " ,d.isupchebeasong" & vbCrlf
    sqlStr = sqlStr + " ,d.issailitem" & vbCrlf
    sqlStr = sqlStr + " ,d.oitemdiv" & vbCrlf
    sqlStr = sqlStr + " ,d.omwdiv" & vbCrlf
    sqlStr = sqlStr + " ,d.odlvType" & vbCrlf
    sqlStr = sqlStr + " ,'7'" & vbCrlf
    sqlStr = sqlStr + " ,getdate()" & vbCrlf
    sqlStr = sqlStr + " ,getdate()" & vbCrlf

	if (CStr(refundmatdiv) = "1") then
		'재료비 10% 차감
		sqlStr = sqlStr + " ,d.matinclude_yn" & vbCrlf
	    sqlStr = sqlStr + " , (CASE WHEN d.matinclude_yn = 'C' THEN ROUND(d.matcostadded * 0.9       , 0) ELSE 0 END) as itemcost " & vbCrlf
	    sqlStr = sqlStr + " , (CASE WHEN d.matinclude_yn = 'C' THEN ROUND(d.matcostadded * 0.9 * 0.01, 0) ELSE 0 END) as mileage " & vbCrlf
	    sqlStr = sqlStr + " , (CASE WHEN d.matinclude_yn = 'C' THEN ROUND(d.matcostadded * 0.9       , 0) ELSE 0 END) as reducedPrice " & vbCrlf
	    sqlStr = sqlStr + " , (CASE WHEN d.matinclude_yn = 'C' THEN ROUND(d.matbuycashadded * 0.9    , 0) ELSE 0 END) as buycash " & vbCrlf
	    sqlStr = sqlStr + " , (CASE WHEN d.matinclude_yn = 'C' THEN ROUND(d.matcostadded * 0.9       , 0) ELSE 0 END) as matcostadded " & vbCrlf
	    sqlStr = sqlStr + " , (CASE WHEN d.matinclude_yn = 'C' THEN ROUND(d.matbuycashadded * 0.9    , 0) ELSE 0 END) as matbuycashadded " & vbCrlf
	    sqlStr = sqlStr + " , (CASE WHEN d.matinclude_yn = 'C' THEN ROUND(d.matcostadded * 0.9       , 0) ELSE 0 END) as couponnotasigncost " & vbCrlf
	elseif (CStr(refundmatdiv) = "2") then
		'20% 차감
	    sqlStr = sqlStr + " ,d.matinclude_yn" & vbCrlf
	    sqlStr = sqlStr + " , (CASE WHEN d.matinclude_yn = 'C' THEN ROUND(d.matcostadded * 0.8       , 0) ELSE 0 END) as itemcost " & vbCrlf
	    sqlStr = sqlStr + " , (CASE WHEN d.matinclude_yn = 'C' THEN ROUND(d.matcostadded * 0.8 * 0.01, 0) ELSE 0 END) as mileage " & vbCrlf
	    sqlStr = sqlStr + " , (CASE WHEN d.matinclude_yn = 'C' THEN ROUND(d.matcostadded * 0.8       , 0) ELSE 0 END) as reducedPrice " & vbCrlf
	    sqlStr = sqlStr + " , (CASE WHEN d.matinclude_yn = 'C' THEN ROUND(d.matbuycashadded * 0.8    , 0) ELSE 0 END) as buycash " & vbCrlf
	    sqlStr = sqlStr + " , (CASE WHEN d.matinclude_yn = 'C' THEN ROUND(d.matcostadded * 0.8       , 0) ELSE 0 END) as matcostadded " & vbCrlf
	    sqlStr = sqlStr + " , (CASE WHEN d.matinclude_yn = 'C' THEN ROUND(d.matbuycashadded * 0.8    , 0) ELSE 0 END) as matbuycashadded " & vbCrlf
	    sqlStr = sqlStr + " , (CASE WHEN d.matinclude_yn = 'C' THEN ROUND(d.matcostadded * 0.8       , 0) ELSE 0 END) as couponnotasigncost " & vbCrlf
	else
		'수강료환불
		''재료비 포함 환불로 변경(skyer9, 2014-09-02)
	    sqlStr = sqlStr + " ,'N' " & vbCrlf
	    sqlStr = sqlStr + " , d.itemcost as itemcost " & vbCrlf
	    sqlStr = sqlStr + " , d.mileage as mileage " & vbCrlf
	    sqlStr = sqlStr + " , d.reducedPrice as reducedPrice " & vbCrlf
	    sqlStr = sqlStr + " , d.buycash as buycash " & vbCrlf
	    sqlStr = sqlStr + " , d.matcostadded as matcostadded " & vbCrlf
	    sqlStr = sqlStr + " , d.matbuycashadded as matbuycashadded " & vbCrlf
	    sqlStr = sqlStr + " , d.couponnotasigncost as couponnotasigncost " & vbCrlf
	end if

    sqlStr = sqlStr + " from " & TABLE_CSDETAIL & " J" & vbCrlf
    if (GC_IsOLDOrder) then
        sqlStr = sqlStr + " ,[db_log].[dbo].tbl_old_order_detail_2003 d" & vbCrlf
    else
        sqlStr = sqlStr + " ," & TABLE_ORDERDETAIL & " d" & vbCrlf
    end if
    sqlStr = sqlStr + " where J.masterid=" & CStr(id)
    sqlStr = sqlStr + " and d.orderserial='" & orderserial & "'"  & vbCrlf
    sqlStr = sqlStr + " and J.orderdetailidx=d." & FIELD_DETAILIDX & ""  & vbCrlf
    sqlStr = sqlStr + " and J.confirmitemno<>0"
    sqlStr = sqlStr + " and d.cancelyn<>'Y'"
	'response.write sqlStr
    dbget.Execute sqlStr

    ''주문금액 재계산
    call recalcuOrderMaster(neworderserial)

    ''재고수량조정 - 한정수량은 조정 안됨
    '강좌확정후 취소 (재료비환불)은 재고하고 무관하다.
    'sqlStr = " exec [db_summary].[dbo].sp_Ten_RealtimeStock_minusOrder '" & neworderserial & "'"
    'dbget.Execute sqlStr

    AddMinusOrder    = neworderserial
end function

function CheckNEditRefundInfo(asid, returnmethod, rebankaccount, rebankownername, rebankname, paygateTid , refundrequire , orgsubtotalprice, orgitemcostsum, orgbeasongpay , orgmileagesum, orgcouponsum, orgallatdiscountsum , canceltotal, refunditemcostsum, refundmileagesum, refundcouponsum, allatsubtractsum , refundbeasongpay, refunddeliverypay, refundadjustpay, refundmatdiv)
    dim sqlStr
    dim refundInfoExists, oldrefundrequire
    refundInfoExists     = false
    CheckNEditRefundInfo = false

    if ((returnmethod="") ) then Exit function
    if ((Not IsNumeric(refundrequire)) or (refundrequire="")) then Exit function


    sqlStr = " select * from " & TABLE_CS_REFUND & ""
    sqlStr = sqlStr + " where asid=" + CStr(asid)

    rsget.Open sqlStr,dbget,1
    if (Not rsget.Eof) then
        refundInfoExists = True
        oldrefundrequire = rsget("refundrequire")
    end if
    rsget.Close

    if (Not refundInfoExists) then Exit function


    sqlStr = " update " & TABLE_CS_REFUND & ""                             + VbCrlf
    sqlStr = sqlStr + " set returnmethod='" + returnmethod + "'"                    + VbCrlf
    sqlStr = sqlStr + " , rebankaccount='" + rebankaccount + "'"                    + VbCrlf
    sqlStr = sqlStr + " , rebankownername='" + rebankownername + "'"                    + VbCrlf
    sqlStr = sqlStr + " , rebankname='" + rebankname + "'"                          + VbCrlf
    sqlStr = sqlStr + " , paygateTid='" + paygateTid + "'"                          + VbCrlf

    sqlStr = sqlStr + " , orgsubtotalprice=" & orgsubtotalprice & VbCrlf
    sqlStr = sqlStr + " , orgitemcostsum=" & orgitemcostsum & VbCrlf
    sqlStr = sqlStr + " , orgbeasongpay=" & orgbeasongpay & VbCrlf
    sqlStr = sqlStr + " , orgmileagesum=" & orgmileagesum & VbCrlf
    sqlStr = sqlStr + " , orgcouponsum=" & orgcouponsum & VbCrlf
    sqlStr = sqlStr + " , orgallatdiscountsum=" & orgallatdiscountsum & VbCrlf
    sqlStr = sqlStr + " , canceltotal=" & canceltotal & VbCrlf
    sqlStr = sqlStr + " , refunditemcostsum=" & refunditemcostsum & VbCrlf
    sqlStr = sqlStr + " , refundmileagesum=" & refundmileagesum & VbCrlf
    sqlStr = sqlStr + " , refundcouponsum=" & refundcouponsum & VbCrlf
    sqlStr = sqlStr + " , allatsubtractsum=" & allatsubtractsum & VbCrlf
    sqlStr = sqlStr + " , refundbeasongpay=" & refundbeasongpay & VbCrlf
    sqlStr = sqlStr + " , refunddeliverypay=" & refunddeliverypay & VbCrlf
    sqlStr = sqlStr + " , refundadjustpay=" & refundadjustpay & VbCrlf
    sqlStr = sqlStr + " , refundmatdiv='" & refundmatdiv & "' " & VbCrlf



    ''무통장이나 마일리지 환불인 경우만 수기 수정 가능
    ''if ((returnmethod="R007") or (returnmethod="R900") or (returnmethod="R000")) and (refundrequire<>oldrefundrequire) then
    if (refundrequire<>oldrefundrequire) then
        sqlStr = sqlStr + " , refundrequire=" + CStr(refundrequire)                     + VbCrlf
        '''sqlStr = sqlStr + " , refundadjustpay=" + CStr(refundrequire) + "-canceltotal"  + VbCrlf
    end if
    sqlStr = sqlStr + " where asid=" + CStr(asid)

'response.write   sqlStr
    dbget.Execute sqlStr

    CheckNEditRefundInfo = true
end Function

function EditCSMasterRefundEncInfo(asid, encmethod, bnkaccount)
    dim sqlStr

    IF (encmethod="PH1") then
        IF (bnkaccount="") then
            sqlStr = " update " & TABLE_CS_REFUND & " " & VbCRLF
            sqlStr = sqlStr + " set encmethod = '' " & VbCRLF
            sqlStr = sqlStr + " 	, encaccount = NULL" & VbCRLF
            sqlStr = sqlStr + " 	, rebankaccount=''" & VbCRLF
            sqlStr = sqlStr + " where asid = " & CStr(asid) & " " & VbCRLF

            dbget.Execute sqlStr
        ELSE
            sqlStr = " update " & TABLE_CS_REFUND & " " & VbCRLF
            sqlStr = sqlStr + " set encmethod = '" & Left(CStr(encmethod), 8) & "' " & VbCRLF
            sqlStr = sqlStr + " 	, encaccount = db_academy.dbo.uf_EncAcctPH1('"&bnkaccount&"')" & VbCRLF
            sqlStr = sqlStr + " 	, rebankaccount=''" & VbCRLF
            sqlStr = sqlStr + " where asid = " & CStr(asid) & " " & VbCRLF

            dbget.Execute sqlStr
        END IF
    end IF

end function

function CheckNEditRefundInfo_OLD(id,returnmethod,rebankaccount,rebankownername,rebankname,paygateTid,refundrequire)
    dim sqlStr
    dim refundInfoExists, oldrefundrequire
    refundInfoExists     = false
    CheckNEditRefundInfo_OLD = false

    if ((returnmethod="") or (returnmethod="R000")) then Exit function
    if ((Not IsNumeric(refundrequire)) or (refundrequire="")) then Exit function


    sqlStr = " select * from " & TABLE_CS_REFUND & ""
    sqlStr = sqlStr + " where asid=" + CStr(id)

    rsget.Open sqlStr,dbget,1
    if (Not rsget.Eof) then
        refundInfoExists = True
        oldrefundrequire = rsget("refundrequire")
    end if
    rsget.Close

    if (Not refundInfoExists) then Exit function


    sqlStr = " update " & TABLE_CS_REFUND & ""                             + VbCrlf
    sqlStr = sqlStr + " set returnmethod='" + returnmethod + "'"                    + VbCrlf
    sqlStr = sqlStr + " , rebankaccount='" + rebankaccount + "'"                    + VbCrlf
    sqlStr = sqlStr + " , rebankownername='" + rebankownername + "'"                    + VbCrlf
    sqlStr = sqlStr + " , rebankname='" + rebankname + "'"                          + VbCrlf
    sqlStr = sqlStr + " , paygateTid='" + paygateTid + "'"                          + VbCrlf

    ''무통장이나 마일리지 환불인 경우만 수기 수정 가능
    if ((returnmethod="R007") or (returnmethod="R900")) and (refundrequire<>oldrefundrequire) then
        sqlStr = sqlStr + " , refundrequire=" + CStr(refundrequire)                     + VbCrlf
        '''sqlStr = sqlStr + " , refundadjustpay=" + CStr(refundrequire) + "-canceltotal"  + VbCrlf
    end if
    sqlStr = sqlStr + " where asid=" + CStr(id)

'response.write   sqlStr
    dbget.Execute sqlStr

    CheckNEditRefundInfo_OLD = true
end function


function LimitItemRecoverOnOrderCancel(byval orderserial)
    dim sqlStr
    dim sitename

    On Error Resume Next

	if (CS_COMPANYID = "thefingers") then

	    ''강좌신청인지 DIY 주문인지 확인
	    sqlStr = " select top 1 sitename "
        sqlStr = sqlStr & " from " & TABLE_ORDERMASTER & " "
        sqlStr = sqlStr & " where orderserial='" & orderserial & "'"

        rsget.Open sqlStr,dbget,1
        if Not rsget.Eof then
            sitename      = rsget("sitename")
        end if
        rsget.Close

		if (sitename = "diyitem") then
			'DIY 주문인 경우

		    ''한정수량 조정 -
	        sqlStr = "update [db_academy].[dbo].tbl_diy_item" + vbCrlf
	        sqlStr = sqlStr + " set limitsold=(case when 0>limitsold - T.itemno then 0 else limitsold - T.itemno end)" + vbCrlf
	        sqlStr = sqlStr + " from " + vbCrlf
	        sqlStr = sqlStr + " ("
	        sqlStr = sqlStr + " 	select d.itemid, d.itemno" + vbCrlf
	        sqlStr = sqlStr + " 	from " & TABLE_ORDERDETAIL & " d" + vbCrlf
	        sqlStr = sqlStr + " 	where d.orderserial='" + CStr(orderserial) + "'" + vbCrlf
	        sqlStr = sqlStr + " 	and d.itemid<>0 "
	        sqlStr = sqlStr + " 	and d.cancelyn<>'Y'"
	        sqlStr = sqlStr + " ) as T" + vbCrlf
	        sqlStr = sqlStr + " where [db_academy].[dbo].tbl_diy_item.itemid=T.Itemid"
	        sqlStr = sqlStr + " and [db_academy].[dbo].tbl_diy_item.limityn='Y'"

	        dbget.Execute(sqlStr)

	        ''옵션있는상품
	        sqlStr = "update [db_academy].[dbo].tbl_diy_item_option" + vbCrlf
	        sqlStr = sqlStr + " set optlimitsold=(case when 0 >optlimitsold - T.itemno then 0 else optlimitsold - T.itemno end)" + vbCrlf
	        sqlStr = sqlStr + " from " + vbCrlf
	        sqlStr = sqlStr + " ("
	        sqlStr = sqlStr + " 	select d.itemid, d.itemoption, d.itemno" + vbCrlf
	        sqlStr = sqlStr + " 	from " & TABLE_ORDERDETAIL & " d " + vbCrlf
	        sqlStr = sqlStr + " 	where d.orderserial='" + CStr(orderserial) + "'" + vbCrlf
	        sqlStr = sqlStr + " 	and d.itemid<>0" + vbCrlf
	        sqlStr = sqlStr + " 	and d.itemoption<>'0000'" + vbCrlf
	        sqlStr = sqlStr + " 	and d.cancelyn<>'Y'"
	        sqlStr = sqlStr + " ) as T" + vbCrlf
	        sqlStr = sqlStr + " where [db_academy].[dbo].tbl_diy_item_option.itemid=T.Itemid"
	        sqlStr = sqlStr + " and [db_academy].[dbo].tbl_diy_item_option.itemoption=T.itemoption"
	        sqlStr = sqlStr + " and [db_academy].[dbo].tbl_diy_item_option.optlimityn='Y'"

	        dbget.Execute(sqlStr)


		elseif (sitename = "academy") then
			'강좌인경우

			'======================================================================
			''한정수량 증가. 대기자가 없는경우만..
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
		    rsget.Open sqlStr,dbget,1
		    if Not rsget.Eof then
		    	WaitExist = (rsget("cnt")>0)
	    	end if
		    rsget.Close

		    if (Not WaitExist) then
		    	sqlStr = "update [db_academy].[dbo].tbl_lec_item_option " + vbCrlf
		    	sqlStr = sqlStr + " set limit_sold=limit_sold - T.cnt" + vbCrlf
		    	sqlStr = sqlStr + " from " + vbCrlf
		    	sqlStr = sqlStr + " (select d.itemid, d.itemoption, sum(d.itemno) as cnt" + vbCrlf
		    	sqlStr = sqlStr + " from " & TABLE_ORDERDETAIL & " d" + vbCrlf
		    	sqlStr = sqlStr + " where d.orderserial='" + CStr(orderserial) + "'" + vbCrlf
		    	sqlStr = sqlStr + " and d.itemid<>0" + vbCrlf
		    	sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		    	sqlStr = sqlStr + " group by d.itemid, d.itemoption ) as T" + vbCrlf
		    	sqlStr = sqlStr + " where [db_academy].[dbo].tbl_lec_item_option.lecidx=T.Itemid"
		    	sqlStr = sqlStr + " and [db_academy].[dbo].tbl_lec_item_option.lecoption=T.itemoption"
		    	dbget.Execute sqlStr

		    	sqlStr = "update [db_academy].[dbo].tbl_lec_item" + vbCrlf
		        sqlStr = sqlStr + " set limit_count=T.limit_count" + vbCrlf
		        sqlStr = sqlStr + " ,limit_sold=T.limit_sold" + vbCrlf
		        sqlStr = sqlStr + " ,wait_count=T.wait_count" + vbCrlf
		        sqlStr = sqlStr + " from (" + vbCrlf
		        sqlStr = sqlStr + " 	select o.lecidx, sum(limit_count) as limit_count, sum(limit_sold) as limit_sold" + vbCrlf
		        sqlStr = sqlStr + " 	,sum(wait_count) as wait_count" + vbCrlf
		        sqlStr = sqlStr + " 	from [db_academy].[dbo].tbl_lec_item_option o" + vbCrlf
		        sqlStr = sqlStr + " 		Join (select distinct itemid from " & TABLE_ORDERDETAIL & " where orderserial='" + CStr(orderserial) + "') A" + vbCrlf
		        sqlStr = sqlStr + " 		on o.lecidx=A.itemid" + vbCrlf
		        sqlStr = sqlStr + " 	where o.isusing <> 'N' " + vbCrlf
		        sqlStr = sqlStr + " 	group by o.lecidx" + vbCrlf
		        sqlStr = sqlStr + " ) T" + vbCrlf
		        sqlStr = sqlStr + " where [db_academy].[dbo].tbl_lec_item.idx=T.lecidx" + vbCrlf

		    	dbget.Execute sqlStr
			end if

		else
			'에러
		end if

	else
		'dbget.Execute  sqlStr
	end if

    On Error Goto 0
end function

function LimitItemRecoverOnItemCancel(byval orderserial, byval itemid, byval itemoption, byval cancelno)
    dim sqlStr
    dim sitename

    On Error Resume Next

	if (CS_COMPANYID = "thefingers") then

	    ''강좌신청인지 DIY 주문인지 확인
	    sqlStr = " select top 1 sitename "
        sqlStr = sqlStr & " from " & TABLE_ORDERMASTER & " "
        sqlStr = sqlStr & " where orderserial='" & orderserial & "'"

        rsget.Open sqlStr,dbget,1
        if Not rsget.Eof then
            sitename      = rsget("sitename")
        end if
        rsget.Close

		if (sitename = "diyitem") then
			'DIY 주문인 경우

		    ''한정수량 조정 -
	        sqlStr = " update " + vbCrlf
	        sqlStr = sqlStr + " 	[db_academy].[dbo].tbl_diy_item " + vbCrlf
	        sqlStr = sqlStr + " set " + vbCrlf
	        sqlStr = sqlStr + " 	limitsold=(case when 0>limitsold - " & CStr(cancelno) & " then 0 else limitsold - " & CStr(cancelno) & " end) " + vbCrlf
	        sqlStr = sqlStr + " where " + vbCrlf
	        sqlStr = sqlStr + " 	1 = 1 " + vbCrlf
	        sqlStr = sqlStr + " 	and limityn='Y' " + vbCrlf
	        sqlStr = sqlStr + " 	and itemid = " & CStr(itemid) & " " + vbCrlf
	        sqlStr = sqlStr + " 	and itemid <> 0  " + vbCrlf
	        dbget.Execute(sqlStr)

	        ''옵션있는상품

	        sqlStr = " update " + vbCrlf
	        sqlStr = sqlStr + " 	[db_academy].[dbo].tbl_diy_item_option " + vbCrlf
	        sqlStr = sqlStr + " set " + vbCrlf
	        sqlStr = sqlStr + " 	optlimitsold=(case when 0 >optlimitsold - " & CStr(cancelno) & " then 0 else optlimitsold - " & CStr(cancelno) & " end) " + vbCrlf
	        sqlStr = sqlStr + " where " + vbCrlf
	        sqlStr = sqlStr + " 	1 = 1 " + vbCrlf
	        sqlStr = sqlStr + " 	and optlimityn='Y' " + vbCrlf
	        sqlStr = sqlStr + " 	and itemid = " & CStr(itemid) & " " + vbCrlf
	        sqlStr = sqlStr + " 	and itemoption = '" & CStr(itemoption) & "' " + vbCrlf
	        sqlStr = sqlStr + " 	and itemoption <> '0000' " + vbCrlf
	        sqlStr = sqlStr + " 	and itemid <> 0  " + vbCrlf
	        dbget.Execute(sqlStr)

		elseif (sitename = "academy") then
			'강좌인경우

			'======================================================================
			''한정수량 증가. 대기자가 없는경우만..
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
		    rsget.Open sqlStr,dbget,1
		    if Not rsget.Eof then
		    	WaitExist = (rsget("cnt")>0)
	    	end if
		    rsget.Close

		    if (Not WaitExist) then
		        sqlStr = " update " + vbCrlf
		        sqlStr = sqlStr + " 	[db_academy].[dbo].tbl_lec_item_option " + vbCrlf
		        sqlStr = sqlStr + " set " + vbCrlf
		        sqlStr = sqlStr + " 	limit_sold=(case when 0>limit_sold - " & CStr(cancelno) & " then 0 else limit_sold - " & CStr(cancelno) & " end) " + vbCrlf
		        sqlStr = sqlStr + " where " + vbCrlf
		        sqlStr = sqlStr + " 	1 = 1 " + vbCrlf
		        sqlStr = sqlStr + " 	and lecidx = " & CStr(itemid) & " " + vbCrlf
		        sqlStr = sqlStr + " 	and lecoption = '" & CStr(itemoption) & "' " + vbCrlf
		        sqlStr = sqlStr + " 	and lecidx <> 0  " + vbCrlf
		        dbget.Execute sqlStr

		    	sqlStr = "update [db_academy].[dbo].tbl_lec_item" + vbCrlf
		        sqlStr = sqlStr + " set limit_count=T.limit_count" + vbCrlf
		        sqlStr = sqlStr + " ,limit_sold=T.limit_sold" + vbCrlf
		        sqlStr = sqlStr + " ,wait_count=T.wait_count" + vbCrlf
		        sqlStr = sqlStr + " from (" + vbCrlf
		        sqlStr = sqlStr + " 	select o.lecidx, sum(limit_count) as limit_count, sum(limit_sold) as limit_sold" + vbCrlf
		        sqlStr = sqlStr + " 	,sum(wait_count) as wait_count" + vbCrlf
		        sqlStr = sqlStr + " 	from [db_academy].[dbo].tbl_lec_item_option o" + vbCrlf
		        sqlStr = sqlStr + " 		Join (select distinct itemid from " & TABLE_ORDERDETAIL & " where orderserial='" + CStr(orderserial) + "') A" + vbCrlf
		        sqlStr = sqlStr + " 		on o.lecidx=A.itemid" + vbCrlf
		        sqlStr = sqlStr + " 	group by o.lecidx" + vbCrlf
		        sqlStr = sqlStr + " ) T" + vbCrlf
		        sqlStr = sqlStr + " where [db_academy].[dbo].tbl_lec_item.idx=T.lecidx" + vbCrlf
		    	dbget.Execute sqlStr
			end if

		else
			'에러
		end if

	else
		'dbget.Execute  sqlStr
	end if

    On Error Goto 0
end function

function LimitItemRecover(byval orderserial)
    dim sqlStr

        ''한정수량 조정 -
        sqlStr = "update [db_item].[dbo].tbl_item" + vbCrlf
        sqlStr = sqlStr + " set limitsold=(case when 0>limitsold - T.itemno then 0 else limitsold - T.itemno end)" + vbCrlf
        sqlStr = sqlStr + " from " + vbCrlf
        sqlStr = sqlStr + " ("
        sqlStr = sqlStr + " 	select d.itemid, d.itemno" + vbCrlf
        sqlStr = sqlStr + " 	from " & TABLE_ORDERDETAIL & " d" + vbCrlf
        sqlStr = sqlStr + " 	where d.orderserial='" + CStr(orderserial) + "'" + vbCrlf
        sqlStr = sqlStr + " 	and d.itemid<>0 "
        sqlStr = sqlStr + " 	and d.cancelyn<>'Y'"
        sqlStr = sqlStr + " ) as T" + vbCrlf
        sqlStr = sqlStr + " where [db_item].[dbo].tbl_item.itemid=T.Itemid"
        sqlStr = sqlStr + " and [db_item].[dbo].tbl_item.limityn='Y'"

        dbget.Execute(sqlStr)

        ''옵션있는상품
        sqlStr = "update [db_item].[dbo].tbl_item_option" + vbCrlf
        sqlStr = sqlStr + " set optlimitsold=(case when 0 >optlimitsold - T.itemno then 0 else optlimitsold - T.itemno end)" + vbCrlf
        sqlStr = sqlStr + " from " + vbCrlf
        sqlStr = sqlStr + " ("
        sqlStr = sqlStr + " 	select d.itemid, d.itemoption, d.itemno" + vbCrlf
        sqlStr = sqlStr + " 	from " & TABLE_ORDERDETAIL & " d " + vbCrlf
        sqlStr = sqlStr + " 	where d.orderserial='" + CStr(orderserial) + "'" + vbCrlf
        sqlStr = sqlStr + " 	and d.itemid<>0" + vbCrlf
        sqlStr = sqlStr + " 	and d.itemoption<>'0000'" + vbCrlf
        sqlStr = sqlStr + " 	and d.cancelyn<>'Y'"
        sqlStr = sqlStr + " ) as T" + vbCrlf
        sqlStr = sqlStr + " where [db_item].[dbo].tbl_item_option.itemid=T.Itemid"
        sqlStr = sqlStr + " and [db_item].[dbo].tbl_item_option.itemoption=T.itemoption"
        sqlStr = sqlStr + " and [db_item].[dbo].tbl_item_option.optlimityn='Y'"

        dbget.Execute(sqlStr)
    On Error Goto 0
end function

function IsExtSiteOrder(orderserial)
    dim sqlStr

    sqlStr = " select IsNULL(sitename,'') as sitename from " & TABLE_ORDERMASTER & ""
    sqlStr = sqlStr + " where orderserial='" + orderserial + "'"
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        IsExtSiteOrder = (rsget("sitename")<>MAIN_SITENAME1 and rsget("sitename")<>MAIN_SITENAME2)
    else
        IsExtSiteOrder = False
    end if
    rsget.close

end function

sub UsafeCancel(byval orderserial)
    '// 전자보증서가 있으면 보증서 취소 요청 (2006.06.15; 운영관리팀 허진원)
    dim objUsafe, result, result_code, result_msg
    On Error Resume Next
    	Set objUsafe = CreateObject( "USafeCom.guarantee.1"  )

    '	Test일 때
    '	objUsafe.Port = 80
    '	objUsafe.Url = "gateway2.usafe.co.kr"
    '	objUsafe.CallForm = "/esafe/guartrn.asp"

        ' Real일 때
        objUsafe.Port = 80
        objUsafe.Url = "gateway.usafe.co.kr"
        objUsafe.CallForm = "/esafe/guartrn.asp"

    	objUsafe.gubun	= "B0"				'// 전문구분 (A0:신규발급, B0:보증서취소, C0:입금확인)
    	objUsafe.EncKey	= ""			'널값인 경우 암호화 안됨
    	objUsafe.mallId	= "ZZcube1010"		'// 쇼핑몰ID
    	objUsafe.oId	= CStr(orderserial)	'// 주문번호

    	'처리 실행!
    	result = objUsafe.cancelInsurance

    	result_code	= Left( result , 1 )
    	result_msg	= Mid( result , 3 )

    	Set objUsafe = Nothing
    On Error Goto 0
end Sub

%>
