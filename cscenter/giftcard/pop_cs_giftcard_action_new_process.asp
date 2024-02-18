<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<!-- #include virtual="/lib/util/base64unicode.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_giftcard_ordercls.asp" -->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp"-->
<!-- #include virtual="/cscenter/lib/cs_action_mail_Function.asp"-->
<!-- #include virtual="/lib/email/MailLib2.asp"-->
<!-- #include virtual="/lib/util/DcCyberAcctUtil.asp"-->

<%

'[기본구성]
'
'if (mode="regcsas") then
'	'==========================================================================
'	'CS 접수
'
'elseif (mode="deletecsas") then
'	'==========================================================================
'	'CS 접수 삭제
'
'elseif (mode="editcsas") then
'	'==========================================================================
'	'CS 접수 내역 수정
'
'elseif (mode="finishcsas") then
'	'==========================================================================
'	'CS 접수 내역 완료처리
'
'else
'	'==========================================================================
'    '에러
'
'end if



'[코드정리]
'------------------------------------------------------------------------------
'A008			주문취소
'


dim mode, modeflag2, divcd, id, reguserid, ipkumdiv
dim title, giftorderserial, gubun01, gubun02, contents_jupsu
dim finishuser, contents_finish

dim requireupche, requiremakerid, ForceReturnByTen
dim detailitemlist

''취소 관련
dim refundmileagesum, refundcouponsum, allatsubtractsum
dim refunditemcostsum, canceltotal, nextsubtotal
dim refundbeasongpay, remainbeasongpay, refunddeliverypay, refundadjustpay
dim remainitemcostsum
dim refundgiftcardsum, refunddepositsum

''환불 관련 maybe (refundrequire==canceltotal)
dim refundrequire, returnmethod
dim rebankname, rebankaccount, rebankownername, paygateTid, encmethod

''업체 추가 정산비
dim add_upchejungsandeliverypay, add_upchejungsancause, add_upchejungsancauseText

''원주문 금액
dim orgitemcostsum, orgbeasongpay, orgmileagesum, orgcouponsum, orgallatdiscountsum, orgsubtotalprice, orggiftcardsum, orgdepositsum

''고객 Open msg
dim opentitle, opencontents

''추가정산ID
dim buf_requiremakerid

''추가로 등록된 CSID
dim newasid

''CS메일발송할지
dim isCsMailSend

newasid = -1

mode        = request.Form("mode")
modeflag2   = request.Form("modeflag2")
divcd       = request.Form("divcd")
id          = request.Form("id")
ipkumdiv    = request.Form("ipkumdiv")
reguserid   = session("ssbctid")
finishuser  = reguserid
title       = html2DB(request.Form("title"))
giftorderserial = request.Form("giftorderserial")
gubun01     = request.Form("gubun01")
gubun02     = request.Form("gubun02")
contents_jupsu  = html2DB(request.Form("contents_jupsu"))
detailitemlist  = html2db(request.Form("detailitemlist"))
contents_finish = html2DB(request.Form("contents_finish"))

''업체 처리 요청
requireupche = request.Form("requireupche")
requiremakerid = request.Form("requiremakerid")
ForceReturnByTen = request.Form("ForceReturnByTen")

orgitemcostsum      = request.Form("orgitemcostsum")
orgbeasongpay       = 0
orgmileagesum       = 0
orgcouponsum        = 0
orgallatdiscountsum = 0
orgsubtotalprice    = request.Form("subtotalprice")

orggiftcardsum    	= 0
refundgiftcardsum   = 0
orgdepositsum    	= 0
refunddepositsum    = 0

refunditemcostsum   = request.Form("refunditemcostsum")
nextsubtotal        = request.Form("nextsubtotal")
canceltotal         = request.Form("canceltotal")

refundbeasongpay    = 0
remainbeasongpay    = 0
refunddeliverypay   = 0

refundmileagesum    = 0
refundcouponsum     = 0
allatsubtractsum    = 0
refundadjustpay     = 0
remainitemcostsum   = 0



''환불요청액
refundrequire       = request.Form("refundrequire")
returnmethod        = request.Form("returnmethod")

rebankname          = request.Form("rebankname")
rebankaccount       = request.Form("rebankaccount")
rebankownername     = request.Form("rebankownername")

encmethod 			= "AE2"

paygateTid          = request.Form("paygateTid")

add_upchejungsandeliverypay = 0
add_upchejungsancause       = ""
add_upchejungsancauseText   = ""

buf_requiremakerid  = ""


isCsMailSend = (request.Form("csmailsend")="on")

if (Not IsNumeric(orgitemcostsum)) or (orgitemcostsum="") then orgitemcostsum     = 0
if (Not IsNumeric(orgbeasongpay)) or (orgbeasongpay="") then orgbeasongpay      = 0
if (Not IsNumeric(orgmileagesum)) or (orgmileagesum="") then orgmileagesum   = 0
if (Not IsNumeric(orgcouponsum)) or (orgcouponsum="") then orgcouponsum    = 0
if (Not IsNumeric(orgallatdiscountsum)) or (orgallatdiscountsum="") then orgallatdiscountsum= 0
if (Not IsNumeric(orgsubtotalprice)) or (orgsubtotalprice="") then orgsubtotalprice   = 0

if (Not IsNumeric(refunditemcostsum)) or (refunditemcostsum="") then refunditemcostsum  = 0
if (Not IsNumeric(refundmileagesum)) or (refundmileagesum="") then refundmileagesum = 0
if (Not IsNumeric(refundcouponsum)) or (refundcouponsum="") then refundcouponsum = 0
if (Not IsNumeric(allatsubtractsum)) or (allatsubtractsum="") then allatsubtractsum = 0

if (Not IsNumeric(refundbeasongpay)) or (refundbeasongpay="") then refundbeasongpay = 0
if (Not IsNumeric(remainbeasongpay)) or (remainbeasongpay="") then remainbeasongpay = 0
if (Not IsNumeric(refunddeliverypay)) or (refunddeliverypay="") then refunddeliverypay = 0

if (Not IsNumeric(refundadjustpay)) or (refundadjustpay="") then refundadjustpay = 0
if (Not IsNumeric(remainitemcostsum)) or (remainitemcostsum="") then remainitemcostsum = 0
if (Not IsNumeric(canceltotal)) or (canceltotal="") then canceltotal = 0
if (Not IsNumeric(refundrequire)) or (refundrequire="") then refundrequire = 0

if (Not IsNumeric(orggiftcardsum)) or (orggiftcardsum="") then orggiftcardsum = 0
if (Not IsNumeric(refundgiftcardsum)) or (refundgiftcardsum="") then refundgiftcardsum = 0
if (Not IsNumeric(orgdepositsum)) or (orgdepositsum="") then orgdepositsum = 0
if (Not IsNumeric(refunddepositsum)) or (refunddepositsum="") then refunddepositsum = 0

if (returnmethod="") then returnmethod="R000"

''올엣카드할인.. -상품별로 차감.

dim sqlStr, errcode, i
dim ScanErr
dim ResultMsg, ReturnUrl, EtcStr
dim ProceedFinish

ScanErr = ""
ProceedFinish = False

dim IsAllCancel
dim CancelValidResultMessage



'==============================================================================
''주문 마스타
dim ogiftcardordermaster

set ogiftcardordermaster = new cGiftCardOrder

ogiftcardordermaster.FRectgiftorderserial = giftorderserial

ogiftcardordermaster.getCSGiftcardOrderDetail



if (mode="regcsas") then
    '==========================================================================
	'CS 접수
    if (divcd="A008") then

		'----------------------------------------------------------------------
        'CS 접수 - 주문취소
        'On Error Resume Next
        dbget.beginTrans

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "001"

            'CS Master 접수
            id = RegCSMaster(divcd, giftorderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
        end if

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "002"

            'CS Master 환불 관련정보 저장
	        Call RegCSMasterRefundInfo(id, returnmethod, refundrequire , orgsubtotalprice, orgitemcostsum, orgbeasongpay , orgmileagesum, orgcouponsum, orgallatdiscountsum  , canceltotal, refunditemcostsum, refundmileagesum, refundcouponsum, allatsubtractsum , refundbeasongpay, refunddeliverypay, refundadjustpay , rebankname, rebankaccount, rebankownername, paygateTid)
	        Call AddCSMasterRefundInfo(id, orggiftcardsum, orgdepositsum, refundgiftcardsum, refunddepositsum)

            '''계좌 암호화 추가.
	        Call EditCSMasterRefundEncInfo(id, encmethod, rebankaccount)
	    End if

	    If (Err.Number = 0) and (ScanErr="") Then
            errcode = "005"

        	IsAllCancel = true
        	CancelValidResultMessage = ""
        	if (ogiftcardordermaster.FOneItem.FCancelyn <> "N") then
        		CancelValidResultMessage = "취소된 주문입니다."
        	end if

			if (ogiftcardordermaster.FOneItem.Fjumundiv = "7") then
				CancelValidResultMessage = "등록된 Gift카드주문은 취소할 수 없습니다. 등록이전 상태로 전환하세요."
			end if

			if (CancelValidResultMessage <> "") then
				ScanErr = CancelValidResultMessage
			end if
        End If

        ResultMsg = ResultMsg + "->. [주문 취소 CS] 접수\n\n"

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "008"

		    sqlStr = "update [db_order].[dbo].tbl_giftcard_order " + VbCrlf
		    sqlStr = sqlStr + " set cancelyn='Y'" + VbCrlf
		    sqlStr = sqlStr + " ,canceldate=IsNULL(canceldate,getdate())" + VbCrlf
		    sqlStr = sqlStr + " where giftorderserial='" + giftorderserial + "'" + VbCrlf
		    dbget.Execute sqlStr

		    ''전자보증서 발급된 경우 취소
		    if (ogiftcardordermaster.FOneItem.FInsureCd="0") then
		        Call UsafeCancel(giftorderserial)
		    end if

            ResultMsg = ResultMsg + "->. 주문건 취소 완료\n\n"
        End IF

        ''순서?. 위로?
        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "009"

            '환불 정보가 있는지 체크 후 무통장환불/마일리지환불/신용카드취소 CS 접수 등록
            newasid = CheckNRegRefund(id, giftorderserial,reguserid)

            If (newasid>0) then
                ResultMsg = ResultMsg + "->. 환불 접수 완료\n\n"
            end if
        End If

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "010"

            Call FinishCSMaster(id, reguserid, contents_finish)

            ResultMsg = ResultMsg + "->. [주문 취소 CS] 완료 처리\n\n"
        End If

	    If (Err.Number = 0) and (ScanErr="") Then
            dbget.CommitTrans
        Else
            dbget.RollBackTrans
            response.write "<script>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(에러코드 : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
            response.write "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(에러코드 : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")"
            ''response.write "<script>history.back()</script>"
            dbget.close()	:	response.End
        End If

        ''가상계좌 금액/마감일 수정
        ''''Call CheckNChangeCyberAcct(giftorderserial)
        response.write "<script>alert('TODO : 가상계좌 금액/마감일 수정')</script>"


        ''이메일 발송. 바로 완료인경우만.
        If (isCsMailSend) then
            If (ProceedFinish) then
                ''주문취소 완료 메일
                '''Call SendCsActionMail(id)
                response.write "<script>alert('TODO : SendCsActionMail')</script>"

                ''환불 접수 메일
                if (newasid>0) then
                    '''''Call SendCsActionMail(newasid)
                    response.write "<script>alert('TODO : SendCsActionMail')</script>"
                end if
            End If
        End IF
        'on error Goto 0

        ReturnUrl = "/cscenter/giftcard/pop_cs_giftcard_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd + "&mode=editreginfo"

    else
        ResultMsg = "정의되지 않았습니다. : mode=" + mode + " , divcd=" + divcd
        response.write "<script>alert('" + ResultMsg + "');</script>"
        response.write "<script>history.back();</script>"
        dbget.close()	:	response.End
    end if



elseif (mode="deletecsas") then
	'==========================================================================
	'CS 접수 삭제

    On Error Resume Next
        dbget.beginTrans

        ''Check Valid Delete - 현재는 B006 업체처리완료 , B007 완료 내역은 취소(삭제) 불가
        if (NOT ValidDeleteCS(id)) then
            response.write "<script>alert(" + Chr(34) + "현재 취소 가능 상태가 아닙니다. 관리자 문의 요망." + Chr(34) + ")</script>"
            response.write "<script>history.back()</script>"
            dbget.close()	:	response.End
        end if

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "001"

            If Not DeleteCSProcess(id, reguserid) then
                ScanErr = "데이터 삭제시 오류"
            else
                ResultMsg = ResultMsg + "->. [CS 처리건 삭제] 처리\n\n"
            End if
        end if

        ResultMsg = "OK"
        ReturnUrl = "/cscenter/giftcard/pop_cs_giftcard_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd + "&mode=editreginfo"

        If (Err.Number = 0) and (ScanErr="") Then
            dbget.CommitTrans
        Else
            dbget.RollBackTrans
            response.write "<script>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(에러코드 : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
            'response.write "<script>history.back()</script>"
            dbget.close()	:	response.End
        End If
    On error Goto 0

elseif (mode="editcsas") then
	'==========================================================================
	'CS 접수 내역 수정

    On Error Resume Next
        dbget.beginTrans

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "001"

            '' CS Master 수정
            if (divcd = "A003") or (divcd = "A007")  then
            	title = GetCSRefundTitle(id, divcd, giftorderserial, returnmethod, title)
            end if

            Call EditCSMaster(id, reguserid, title, contents_jupsu, gubun01, gubun02)

            ''환불방식이 바뀐 경우.. 2011-07-20 서동석추가
            if (divcd="A007") and Not ((returnmethod="R020") or (returnmethod="R080") or (returnmethod="R100") or (returnmethod="R120") or (returnmethod="R400")) then
                sqlStr = " update [db_cs].[dbo].tbl_new_as_list"
                sqlStr = sqlStr + " set divcd='A003'"
                sqlStr = sqlStr + " where id=" + CStr(id)

                dbget.Execute sqlStr
            end if

            if (divcd="A003") and ((returnmethod="R020") or (returnmethod="R080") or (returnmethod="R100") or (returnmethod="R120") or (returnmethod="R400")) then
                sqlStr = " update [db_cs].[dbo].tbl_new_as_list"
                sqlStr = sqlStr + " set divcd='A007'"
                sqlStr = sqlStr + " where id=" + CStr(id)

                dbget.Execute sqlStr
            end if
        end if

        ResultMsg = ResultMsg + "->. [CS 처리건 수정] 처리\n\n"

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "003"

            '' 환불 정보 수정
            if (CheckNEditRefundInfo(id,returnmethod,rebankaccount,rebankownername,rebankname,paygateTid,refundrequire, orgsubtotalprice, orgitemcostsum, orgbeasongpay , orgmileagesum, orgcouponsum, orgallatdiscountsum , canceltotal, refunditemcostsum, refundmileagesum, refundcouponsum, allatsubtractsum , refundbeasongpay, refunddeliverypay, refundadjustpay)) then
            	Call AddCSMasterRefundInfo(id, orggiftcardsum, orgdepositsum, refundgiftcardsum, refunddepositsum)

            	'''계좌 암호화 추가.
	            Call EditCSMasterRefundEncInfo(id, encmethod, rebankaccount)
                ResultMsg = ResultMsg + "->. [환불정보 수정] 처리\n\n"
            end if
        end If

        ReturnUrl = "/cscenter/giftcard/pop_cs_giftcard_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd + "&mode=editreginfo"

        If (Err.Number = 0) and (ScanErr="") Then
            dbget.CommitTrans
        Else
            dbget.RollBackTrans
            response.write "<script>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(에러코드 : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
            'response.write "<script>history.back()</script>"
            dbget.close()	:	response.End
        End If
    On error Goto 0

elseif (mode="finishcsas") then
	'==========================================================================
	'CS 접수 내역 완료처리

	if (divcd="A003") or (divcd="A007") then
    	'----------------------------------------------------------------------
        'CS 접수 내역 완료처리 - 환불요청 / 카드,이체,휴대폰취소요청
        dim RefreturnMethod, Refrealrefund

        On Error Resume Next
	        dbget.beginTrans

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "001"

	            Call FinishCSMaster(id, reguserid, contents_finish)
	        End If

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "002"

				'마일리지 환불 및 예치금전환은 실제로 환불이 처리되지만, 그 밖에 신용카드/무통장 등의 환불은 별도 환불 프로세스에서 처리된다.
				'따라서 완료처리한다고 해서 실제로 환불이 일어나지 않는다.
	            Call CheckRefundFinish(id, giftorderserial, RefreturnMethod, Refrealrefund)
	        End If

	        ResultMsg = "처리 완료"
	        if (RefreturnMethod="R007") and (Refrealrefund>0) then
	            ReturnUrl = "/cscenter/giftcard/pop_cs_giftcard_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd + "&finishtype=1"
	        else
	            ReturnUrl = "/cscenter/giftcard/pop_cs_giftcard_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd
	        end if

	        If (Err.Number = 0) and (ScanErr="") Then
	            dbget.CommitTrans
	        Else
	            dbget.RollBackTrans
	            response.write "<script>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(에러코드 : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
	            response.write "<script>history.back()</script>"
	            dbget.close()	:	response.End
	        End If

	        ''환불 완료 메일
	        If (isCsMailSend) then
	            Call SendCsActionMail(id)
	        End IF
        On error Goto 0

    elseif (divcd="A004") or (divcd="A010") then
		'----------------------------------------------------------------------
        'CS 접수 내역 완료처리 - 반품접수(업체배송)  // 회수신청(텐바이텐배송)
        dim MinusOrderserial

        On Error Resume Next
	        dbget.beginTrans

	        If (modeflag2<>"norefund") and (Err.Number = 0) and (ScanErr="") Then
	            errcode = "001"

	            ''마이너스 주문 등록
	            if (CheckNAddMinusOrder(id, orderserial, reguserid, MinusOrderserial, ScanErr)) then
	                ResultMsg = ResultMsg + "->. [반품 주문] 등록\n\n"
	            end if
	        End If

	        If (modeflag2<>"norefund") and (Err.Number = 0) and (ScanErr="") Then
	            errcode = "002"

	            '환불 정보가 있는지 체크 후 무통장환불/마일리지환불/신용카드취소 CS 접수 등록
	            'newasid = CheckNRegRefund(id, MinusOrderserial, reguserid)

	            '원주문에 대해 CS등록한다. 마이너스 주문에는 CS를 등록할 수 없다.
	            newasid = CheckNRegRefund(id, orderserial, reguserid)
	            call AddminusOrderLink(newasid,MinusOrderserial)

	            if (newasid>0) then
	                ResultMsg = ResultMsg + "->. [취소(환불)접수] 처리\n\n"
	            end if
	        End If


	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "009"

	            Call FinishCSMaster(id, reguserid, contents_finish)

	            if (divcd="A004") then
	                ResultMsg = ResultMsg + "->. 반품 처리 완료\n\n"
	            elseif (divcd="A010") then
	                ResultMsg = ResultMsg + "->. 회수 처리 완료\n\n"
	            end if
	        End If

	        ResultMsg = ResultMsg
	        ReturnUrl = "/cscenter/action/pop_cs_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd
	        If (Err.Number = 0) and (ScanErr="") Then
	            dbget.CommitTrans
	        Else
	            dbget.RollBackTrans
	            response.write "<script>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(에러코드 : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
	            response.write "<script>history.back()</script>"
	            dbget.close()	:	response.End
	        End If

	        ''회수 완료 메일
	        If (isCsMailSend) then
	            Call SendCsActionMail(id)

	            ''환불 접수 메일
	            if (newasid>0) then
	                Call SendCsActionMail(newasid)
	            end if
	        End If
        On error Goto 0
    elseif  (divcd="A011") or (divcd="A012") then
    	'----------------------------------------------------------------------
    	'CS 접수 내역 완료처리 - 맞교환회수(텐바이텐배송)
        On Error Resume Next
	        dbget.beginTrans

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "009"

	            Call FinishCSMaster(id, reguserid, contents_finish)
	        End If

	        ResultMsg = "처리 완료"
	        ReturnUrl = "/cscenter/action/pop_cs_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd

	        If (Err.Number = 0) and (ScanErr="") Then
	            dbget.CommitTrans
	        Else
	            dbget.RollBackTrans
	            response.write "<script>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(에러코드 : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
	           ' response.write "<script>history.back()</script>"
	            dbget.close()	:	response.End
	        End If

	        ''맞교환 완료 메일
	        If (isCsMailSend) then
	            if (divcd="A011") then
    	            Call SendCsActionMail(id)
    	        end if
	        End If
        On error Goto 0
    elseif  (divcd="A000") or (divcd="A001") or (divcd="A002") or (divcd="A009") or (divcd="A006") or (divcd="A005") or (divcd="A700") then
    	'----------------------------------------------------------------------
        'CS 접수 내역 완료처리 - 맞교환 출고 / 누락 / 서비스 발송 / 기타 /  출고시 유의사항
        On Error Resume Next
	        dbget.beginTrans

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "009"

	            Call FinishCSMaster(id, reguserid, contents_finish)
	        End If


	        ResultMsg = "처리 완료"
	        ReturnUrl = "/cscenter/action/pop_cs_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd

	        If (Err.Number = 0) and (ScanErr="") Then
	            dbget.CommitTrans
	        Else
	            dbget.RollBackTrans
	            response.write "<script>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(에러코드 : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
	           ' response.write "<script>history.back()</script>"
	            dbget.close()	:	response.End
	        End If

	        If (isCsMailSend) then
	            if ((divcd="A000") or (divcd="A001") or (divcd="A002")) then
	                ''맞교환/누락/서비스 완료 메일
	                Call SendCsActionMail(id)
	            end if
	        End If
        On error Goto 0
    else
        ResultMsg = "정의되지 않았습니다. : mode=" + mode + " , divcd=" + divcd
        response.write "<script>alert('" + ResultMsg + "');</script>"
        response.write "<script>history.back();</script>"
        dbget.close()	:	response.End
    end if

else
	'==========================================================================
    ResultMsg = "정의되지 않았습니다. : mode=" + mode + " , divcd=" + divcd
    response.write "<script>alert('" + ResultMsg + "');</script>"
    response.write "<script>history.back();</script>"
    dbget.close()	:	response.End
end if

%>

<!-- #include virtual="/lib/db/dbclose.asp" -->

<%
response.write "<script>alert('" + ResultMsg + "');</script>"
response.write "<script>location.replace('" + ReturnUrl + "');</script>"
%>
