<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/cscenterv2/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/lecture/lecturecls.asp"-->
<!-- #include virtual="/cscenterv2/lib/csAsLecturefunction.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/cs_lecture/lec_cs_aslistcls.asp"-->
<!-- #include virtual="/cscenterv2/lib/cs_action_mail_Function.asp"-->
<!-- #include virtual="/lib/email/MailLib2.asp"-->
<%

dim mode, modeflag2, divcd, id, reguserid, ipkumdiv
dim title, orderserial, gubun01, gubun02, contents_jupsu
dim finishuser, contents_finish

dim requireupche, requiremakerid, ForceReturnByTen
dim detailitemlist

''취소 관련
dim refundmileagesum, refundcouponsum, allatsubtractsum
dim refunditemcostsum, canceltotal, nextsubtotal
dim refundbeasongpay, recalcubeasongpay, refunddeliverypay, refundadjustpay
dim refundmatdiv

''환불 관련 maybe (refundrequire==canceltotal)
dim refundrequire, returnmethod
dim rebankname, rebankaccount, rebankownername, paygateTid, encmethod

''업체 추가 정산비
dim add_upchejungsandeliverypay, add_upchejungsancause, add_upchejungsancauseText

''원주문 금액
dim orgitemcostsum, orgbeasongpay, orgmileagesum, orgcouponsum, orgallatdiscountsum, orgsubtotalprice

''고객 Open msg
dim opentitle, opencontents

''추가정산ID
dim buf_requiremakerid

''추가로 등록된 CSID
dim newasid

''CS메일발송할지
dim isCsMailSend

newasid = -1

mode        = RequestCheckvar(request.Form("mode"),16)
modeflag2   = RequestCheckvar(request.Form("modeflag2"),16)
divcd       = RequestCheckvar(request.Form("divcd"),4)
id          = RequestCheckvar(request.Form("id"),10)
ipkumdiv    = RequestCheckvar(request.Form("ipkumdiv"),2)
reguserid   = session("ssbctid")
finishuser  = reguserid
title       = html2DB(request.Form("title"))
if title <> "" then
	if checkNotValidHTML(title) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
orderserial = RequestCheckvar(request.Form("orderserial"),16)
gubun01     = RequestCheckvar(request.Form("gubun01"),4)
gubun02     = RequestCheckvar(request.Form("gubun02"),4)
contents_jupsu  = html2DB(request.Form("contents_jupsu"))
detailitemlist  = html2db(request.Form("detailitemlist"))
contents_finish = html2DB(request.Form("contents_finish"))
if contents_jupsu <> "" then
	if checkNotValidHTML(contents_jupsu) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end If
if detailitemlist <> "" then
	if checkNotValidHTML(detailitemlist) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end If
if contents_finish <> "" then
	if checkNotValidHTML(contents_finish) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
''업체 처리 요청
requireupche = RequestCheckvar(request.Form("requireupche"),1)
requiremakerid = RequestCheckvar(request.Form("requiremakerid"),32)
ForceReturnByTen = RequestCheckvar(request.Form("ForceReturnByTen"),32)

orgitemcostsum      = RequestCheckvar(request.Form("orgitemcostsum"),10)
orgbeasongpay       = RequestCheckvar(request.Form("orgbeasongpay"),10)
orgmileagesum       = RequestCheckvar(request.Form("miletotalprice"),10)
orgcouponsum        = RequestCheckvar(request.Form("tencardspend"),10)
orgallatdiscountsum = RequestCheckvar(request.Form("allatdiscountprice"),10)
orgsubtotalprice    = RequestCheckvar(request.Form("subtotalprice"),10)


refunditemcostsum   = RequestCheckvar(request.Form("refunditemcostsum"),10)
nextsubtotal        = RequestCheckvar(request.Form("nextsubtotal"),10)
canceltotal         = RequestCheckvar(request.Form("canceltotal"),10)

refundbeasongpay    = RequestCheckvar(request.Form("refundbeasongpay"),10)
recalcubeasongpay   = RequestCheckvar(request.Form("recalcubeasongpay"),10)
refunddeliverypay   = RequestCheckvar(request.Form("refunddeliverypay"),10)

refundmileagesum    = RequestCheckvar(request.Form("refundmileagesum"),10)
refundcouponsum     = RequestCheckvar(request.Form("refundcouponsum"),10)
allatsubtractsum    = RequestCheckvar(request.Form("allatsubtractsum"),10)
refundadjustpay     = RequestCheckvar(request.Form("refundadjustpay"),10)

'재료비환불방식
refundmatdiv     	= RequestCheckvar(request.Form("ckmaterialpay20"),1)
if (refundmatdiv = "") then
	refundmatdiv = RequestCheckvar(request.Form("ckmaterialpay10"),1)
end if
if (refundmatdiv = "") then
	refundmatdiv = RequestCheckvar(request.Form("cklecturepay0"),1)
end if



''환불요청액
refundrequire       = RequestCheckvar(request.Form("refundrequire"),10)
returnmethod        = RequestCheckvar(request.Form("returnmethod"),4)

rebankname          = RequestCheckvar(request.Form("rebankname"),32)
rebankaccount       = RequestCheckvar(request.Form("rebankaccount"),32)
rebankownername     = RequestCheckvar(request.Form("rebankownername"),64)
if rebankownername <> "" then
	if checkNotValidHTML(rebankownername) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
encmethod 			= "PH1"

paygateTid          = RequestCheckvar(request.Form("paygateTid"),64)
if paygateTid <> "" then
	if checkNotValidHTML(paygateTid) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
add_upchejungsandeliverypay = RequestCheckvar(request.Form("add_upchejungsandeliverypay"),10)
add_upchejungsancause       = RequestCheckvar(request.Form("add_upchejungsancause"),32)
add_upchejungsancauseText   = RequestCheckvar(request.Form("add_upchejungsancauseText"),32)

buf_requiremakerid  = RequestCheckvar(request.Form("buf_requiremakerid"),32)


isCsMailSend = (RequestCheckvar(request.Form("csmailsend"),10)="on")

if (add_upchejungsancause="직접입력") then add_upchejungsancause = add_upchejungsancauseText


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
if (Not IsNumeric(recalcubeasongpay)) or (recalcubeasongpay="") then recalcubeasongpay = 0
if (Not IsNumeric(refunddeliverypay)) or (refunddeliverypay="") then refunddeliverypay = 0

if (Not IsNumeric(refundadjustpay)) or (refundadjustpay="") then refundadjustpay = 0
if (Not IsNumeric(canceltotal)) or (canceltotal="") then canceltotal = 0
if (Not IsNumeric(refundrequire)) or (refundrequire="") then refundrequire = 0

if (returnmethod="") then returnmethod="R000"

''올엣카드할인.. -상품별로 차감.

dim sqlStr, errcode, i
dim ScanErr
dim ResultMsg, ReturnUrl, EtcStr
dim ProceedFinish

ScanErr = ""
ProceedFinish = False

dim IsAllCancel

''과거 주문 내역인지 Check
GC_IsOLDOrder = CheckIsOldOrder(orderserial)

''response.write mode & "_" & divcd
''response.end

if (mode="regcsas") then
    '' 취소
    if (divcd="A008") then
        ''On Error Resume Next
        dbget.beginTrans

        if (GC_IsOLDOrder) then ScanErr="6개월 이전 과거 내역 취소 불가 - 관리자 문의 요망"

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "001"

            '' CS Master 접수
            id = RegCSMaster(divcd, orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
        end if

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "002"

            '' CS Master 취소/ 환불 관련정보
	        Call RegCSMasterRefundInfoLecture(id, returnmethod, refundrequire , orgsubtotalprice, orgitemcostsum, orgbeasongpay , orgmileagesum, orgcouponsum, orgallatdiscountsum  , canceltotal, refunditemcostsum, refundmileagesum, refundcouponsum, allatsubtractsum , refundbeasongpay, refunddeliverypay, refundadjustpay , rebankname, rebankaccount, rebankownername, paygateTid, refundmatdiv)

			'''계좌 암호화 추가.
			Call EditCSMasterRefundEncInfo(id, encmethod, rebankaccount)
	    End if

	    If (Err.Number = 0) and (ScanErr="") Then
            errcode = "003"

            '' CS Detail 접수
	        Call AddCSDetailByArrStr(detailitemlist, id, orderserial)

        end if

'''검토..
	    If (Err.Number = 0) and (ScanErr="") Then
            errcode = "005"
            ''전체 취소인지 여부 확인 - AsDetail 입력후 검사 해야 함.
        	IsAllCancel     = IsAllCancelRegValid(id, orderserial)

        	if (IsAllCancel) And (orgsubtotalprice<>canceltotal) then
        	    ScanErr = "취소 금액 불일치 - 전체취소시 취소금액과 결제금액이 일치해야함"
        	end if
        End If


        '출고완료인 상품이 있을 경우, 진행정지(주문취소 불가)

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "006"

            ''출고 완료 또는 취소된 내역이 있는지 확인
            if Not (IsCancelValidState(id, orderserial)) then
                ScanErr = "취소 검증 오류. - 출고된 내역이 있거나 취소된 내역이 있습니다."
            end if
        end if


        '' 완료처리 바로 진행할지 검토
        '' 업체 확인중 상태가 있는경우 - > 접수로만 진행
        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "007"
        	''바로 완료처리로 진행 할지 여부 - AsDetail 입력후 검사
            ProceedFinish   = IsDirectProceedFinish(divcd, id, orderserial, EtcStr)
            contents_finish = ""
        End If

        ResultMsg = ResultMsg + "->. [주문 취소 CS] 접수\n\n"



        '' 완료처리 프로세스
        'TODO : 한정수량 재증가 시켜주는 기능 :

        If (ProceedFinish) then
            If (Err.Number = 0) and (ScanErr="") Then
                errcode = "008"

                Call CancelProcess(id, orderserial)

                ResultMsg = ResultMsg + "->. 주문건 취소 완료\n\n"
            End IF


            ''순서?. 위로?
            If (Err.Number = 0) and (ScanErr="") Then
                errcode = "009"
                ''환불 등록건이 있는지 체크 후 환불요청/신용카드 취소요청 등록
                newasid = CheckNRegRefund(id, orderserial,reguserid)
                If (newasid>0) then
                    ResultMsg = ResultMsg + "->. 환불 접수 완료\n\n"
                end if
            End If



            If (Err.Number = 0) and (ScanErr="") Then
                errcode = "010"
                Call FinishCSMaster(id, reguserid, contents_finish)

                ResultMsg = ResultMsg + "->. [주문 취소 CS] 완료 처리\n\n"

            End If
        ELSE
            ResultMsg = ResultMsg + "->. 상품 준비중 상태인 상품이 존재하므로\n\n 주문 취소 접수만 진행 되었습니다.\n\n 확인후 완료 처리하셔야 합니다."
        End If

	    If (Err.Number = 0) and (ScanErr="") Then
            dbget.CommitTrans
        Else
            dbget.RollBackTrans
            response.write "<script>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(에러코드 : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
            ''response.write "<script>history.back()</script>"
            dbget.close()	:	response.End
        End If

        ''가상계좌 금액/마감일 수정
        'TODO : 가상계좌 일단 스킵
        'Call CheckNChangeCyberAcct(orderserial)

        ''이메일 발송. 바로 완료인경우만.
        If (isCsMailSend) then
            If (ProceedFinish) then
                ''주문취소 완료 메일
                Call SendCsActionMail(id)

                ''환불 접수 메일
                if (newasid>0) then
                    Call SendCsActionMail(newasid)
                end if
            End If
        End IF
        on error Goto 0


        ReturnUrl = "pop_lec_cs_register.asp?id="  + CStr(id) + "&divcd=" + divcd + "&mode=editreginfo"
    elseif (divcd="A004") or (divcd="A010") then
        ''강좌확정후 취소 접수 또는 회수신청.
        On Error Resume Next
        dbget.beginTrans

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "001"

            '' CS Master 접수
            id = RegCSMaster(divcd, orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
        end if

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "002"

            '' CS Master 취소/ 환불 관련정보
	        Call RegCSMasterRefundInfoLecture(id, returnmethod, refundrequire , orgsubtotalprice, orgitemcostsum, orgbeasongpay , orgmileagesum, orgcouponsum, orgallatdiscountsum  , canceltotal, refunditemcostsum, refundmileagesum, refundcouponsum, allatsubtractsum , refundbeasongpay, refunddeliverypay, refundadjustpay , rebankname, rebankaccount, rebankownername, paygateTid, refundmatdiv)

			'''계좌 암호화 추가.
			Call EditCSMasterRefundEncInfo(id, encmethod, rebankaccount)
	    End if


	    If (Err.Number = 0) and (ScanErr="") Then
            errcode = "003"

            '' CS Detail 접수
	        Call AddCSDetailByArrStr(detailitemlist, id, orderserial)

        end if

        '' To Do : 반품, 회수 접수시 : 업체배송과 자체 배송이 섞여 있을경우 -> 업체배송은 브랜드별로 입력 자체배송은 통합 입력
        '' Check - 업체배송과 자체배송을 같이 접수하지 못함.
        ''       - 업체배송이 존재할 경우 MakerID가 1개만 존재 해야함.

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "004"

            if (IsReturnRegValid(id, orderserial, ScanErr, requiremakerid)) then
                ''업체 요청인 경우 업체 require, 등록
                if ((requiremakerid<>"") and (ForceReturnByTen="")) then
                    call RegCSMasterAddUpche(id, requiremakerid)
                end if

                ResultMsg = ResultMsg + "->. 강좌확정 후 일부취소 접수\n\n"
            end if
        End if

        ''업체 추가 정산비 2008.11.10
        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "005"

            if (add_upchejungsandeliverypay<>"0") and (add_upchejungsandeliverypay<>"")  then
                call RegCSUpcheAddJungsanPay(id, add_upchejungsandeliverypay, add_upchejungsancause, buf_requiremakerid)
            end if
        end if

        ReturnUrl = "pop_lec_cs_register.asp?id="  + CStr(id) + "&divcd=" + divcd


        If (Err.Number = 0) and (ScanErr="") Then
            dbget.CommitTrans
        Else
            dbget.RollBackTrans
            response.write "<script>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(에러코드 : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
            'response.write "<script>history.back()</script>"
            dbget.close()	:	response.End
        End If

        ''이메일 발송. 강좌확정후취소 회수 접수
        If (isCsMailSend) then
            '이메일발송 안함(환불접수만 이메일 발송)
            'Call SendCsActionMail(id)
        end if
        on error Goto 0
    elseif (divcd="A001") or (divcd="A002") then
        ''누락재발송, 서비스발송
        On Error Resume Next
        dbget.beginTrans

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "001"
            '' CS Master 접수
            id = RegCSMaster(divcd, orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
        end if

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "002"
            '' CS Detail 접수
	        Call AddCSDetailByArrStr(detailitemlist, id, orderserial)

        end if

        if (requiremakerid<>"") then
            call RegCSMasterAddUpche(id, requiremakerid)
        end if

        ResultMsg = "접수완료"
        ReturnUrl = "pop_lec_cs_register.asp?id="  + CStr(id) + "&divcd=" + divcd


        If (Err.Number = 0) and (ScanErr="") Then
            dbget.CommitTrans
        Else
            dbget.RollBackTrans
            response.write "<script>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(에러코드 : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
            'response.write "<script>history.back()</script>"
            dbget.close()	:	response.End
        End If

        ''이메일 발송 누락 서비스 접수
        If (isCsMailSend) then
            Call SendCsActionMail(id)
        End If
        on error Goto 0
    elseif (divcd="A000") then
        ''맞교환출고
        On Error Resume Next
        dbget.beginTrans

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "001"
            '' CS Master 접수
            id = RegCSMaster(divcd, orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
        end if

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "002"
            '' CS Detail 접수
	        Call AddCSDetailByArrStr(detailitemlist, id, orderserial)

        end if


        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "003"
            if (requiremakerid<>"") then
                call RegCSMasterAddUpche(id, requiremakerid)

                ResultMsg = "맞교환 접수완료 - 업체배송"
            else
                ''맞교환 회수 접수
                newasid = RegCSMaster("A011", orderserial, reguserid, "맞교환 회수접수", contents_jupsu, gubun01, gubun02)
                Call AddCSDetailByArrStr(detailitemlist, newasid, orderserial)

                 ResultMsg = "맞교환 출고 접수 및 회수접수 완료 - 텐바이텐 배송"
            end if
        end if

        ''업체 추가 정산비 2008.11.10
        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "004"

            if (add_upchejungsandeliverypay<>"0") and (add_upchejungsandeliverypay<>"")  then
                call RegCSUpcheAddJungsanPay(id, add_upchejungsandeliverypay, add_upchejungsancause, buf_requiremakerid)
            end if
        end if


        ReturnUrl = "pop_lec_cs_register.asp?id="  + CStr(id) + "&divcd=" + divcd


        If (Err.Number = 0) and (ScanErr="") Then
            dbget.CommitTrans
        Else
            dbget.RollBackTrans
            response.write "<script>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(에러코드 : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
            'response.write "<script>history.back()</script>"
            dbget.close()	:	response.End
        End If

        ''이메일 발송 맞교환 접수
        if (isCsMailSend) then
            Call SendCsActionMail(id)

            ''맞교환 회수가 있을경우
            if (newasid>0) then
                Call SendCsActionMail(newasid)
            end if
        End If
        on error Goto 0

    elseif (divcd="A009") or (divcd="A006") or (divcd="A700") then
        ''기타사항 / 출고시유의사항 / 업체 추가 정산비
        On Error Resume Next
        dbget.beginTrans

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "001"
            '' CS Master 접수
            id = RegCSMaster(divcd, orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
        end if

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "002"
            '' CS Detail 접수
	        Call AddCSDetailByArrStr(detailitemlist, id, orderserial)

        end if

        ''업체 추가 정산비 : 2008.11.10
        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "003"

            if (add_upchejungsandeliverypay<>"0") and (add_upchejungsandeliverypay<>"")  then
                call RegCSUpcheAddJungsanPay(id, add_upchejungsandeliverypay, add_upchejungsancause, buf_requiremakerid)
            end if
        end if

        ''업체지정.
        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "004"
            if (requiremakerid<>"") then
                call RegCSMasterAddUpche(id, requiremakerid)
            end if
         end if

        ResultMsg = "등록되었습니다."
        ReturnUrl = "pop_lec_cs_register.asp?id="  + CStr(id) + "&divcd=" + divcd

        If (Err.Number = 0) and (ScanErr="") Then
            dbget.CommitTrans
        Else
            dbget.RollBackTrans
            response.write "<script>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(에러코드 : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
            'response.write "<script>history.back()</script>"
            dbget.close()	:	response.End
        End If
        on error Goto 0
    elseif (divcd="A003") or (divcd="A005") then
        ''환불접수 / 외부몰 환불접수
        On Error Resume Next
        dbget.beginTrans

        if (divcd="A005") and (Not IsExtSiteOrder(orderserial)) then
            ScanErr = "외부몰 환불접수는 외부몰 주문건만 가능합니다."
        end if

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "001"
            '' CS Master 접수
            id = RegCSMaster(divcd, orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
        end if

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "002"

            '' CS Master 취소/ 환불 관련정보
	        Call RegCSMasterRefundInfoLecture(id, returnmethod, refundrequire , orgsubtotalprice, orgitemcostsum, orgbeasongpay , orgmileagesum, orgcouponsum, orgallatdiscountsum  , canceltotal, refunditemcostsum, refundmileagesum, refundcouponsum, allatsubtractsum , refundbeasongpay, refunddeliverypay, refundadjustpay , rebankname, rebankaccount, rebankownername, paygateTid, refundmatdiv)

			'''계좌 암호화 추가.
			Call EditCSMasterRefundEncInfo(id, encmethod, rebankaccount)
	    End if


        ResultMsg = "등록되었습니다."
        ReturnUrl = "pop_lec_cs_register.asp?id="  + CStr(id) + "&divcd=" + divcd

        If (Err.Number = 0) and (ScanErr="") Then
            dbget.CommitTrans
        Else
            dbget.RollBackTrans
            response.write "<script>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(에러코드 : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
            'response.write "<script>history.back()</script>"
            dbget.close()	:	response.End
        End If

        ''이메일 발송 환불접수
        If (isCsMailSend) then
            if (divcd="A003") then
                Call SendCsActionMail(id)
            end if
        End If
        on error Goto 0

    else
        ResultMsg = "정의되지 않았습니다. : mode=" + mode + " , divcd=" + divcd
        response.write "<script>alert('" + ResultMsg + "');</script>"
        response.write "<script>history.back();</script>"
        dbget.close()	:	response.End
    end if
elseif (mode="deletecsas") then
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
        ReturnUrl = "pop_lec_cs_register.asp?id="  + CStr(id) + "&divcd=" + divcd + "&mode=editreginfo"

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
    ''접수 내역 수정
    On Error Resume Next
        dbget.beginTrans

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "001"
            '' 마스타 수정
            Call EditCSMaster(divcd, orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)

        end if

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "002"
            '' 디테일 수정
            Call EditCSDetailByArrStr(detailitemlist, id, orderserial)
        End if

        ResultMsg = ResultMsg + "->. [CS 처리건 수정] 처리\n\n"

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "003"
            '' 환불 접수 수정

            if (CheckNEditRefundInfo(id,returnmethod,rebankaccount,rebankownername,rebankname,paygateTid,refundrequire, orgsubtotalprice, orgitemcostsum, orgbeasongpay , orgmileagesum, orgcouponsum, orgallatdiscountsum , canceltotal, refunditemcostsum, refundmileagesum, refundcouponsum, allatsubtractsum , refundbeasongpay, refunddeliverypay, refundadjustpay, refundmatdiv)) Then
				'''계좌 암호화 추가.
				Call EditCSMasterRefundEncInfo(id, encmethod, rebankaccount)

                ResultMsg = ResultMsg + "->. [환불정보 수정] 처리\n\n"
            end if
        end If

        ''업체 추가 정산비 2008.11.10
        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "004"

            if (add_upchejungsandeliverypay<>"") then
               '강좌에 업체추가 배송비 정산은 없다.
                'call EditCSUpcheAddJungsanPay(id, add_upchejungsandeliverypay, add_upchejungsancause, buf_requiremakerid)
            end if
        end if

        ReturnUrl = "pop_lec_cs_register.asp?id="  + CStr(id) + "&divcd=" + divcd + "&mode=editreginfo"


        If (Err.Number = 0) and (ScanErr="") Then
            dbget.CommitTrans
        Else
            dbget.RollBackTrans
            response.write "<script>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(에러코드 : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
            'response.write "<script>history.back()</script>"
            dbget.close()	:	response.End
        End If
    On error Goto 0

elseif (mode="finishededitcsas") then
    ''완료된 내역 수정
    On Error Resume Next
        dbget.beginTrans

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "001"
            '' 마스타 수정
            Call EditCSMasterFinished(divcd, orderserial, reguserid, title, contents_jupsu, gubun01, gubun02, reguserid, contents_finish)
        end if

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "002"
            '' 디테일 수정
            Call EditCSDetailByArrStr(detailitemlist, id, orderserial)
        End if

        ResultMsg = "수정 되었습니다."
        ReturnUrl = "pop_lec_cs_register.asp?id="  + CStr(id) + "&divcd=" + divcd + "&mode=editreginfo"


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
    ''완료 처리 진행

    ''취소인경우
    if (divcd="A008") then
        On Error Resume Next
        dbget.beginTrans
        if (GC_IsOLDOrder) then ScanErr="6개월 이전 과거 내역 취소 불가 - 관리자 문의 요망"

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "001"
            Call CancelProcess(id, orderserial)
        End IF


        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "008"
            ''환불 등록건이 있는지 체크 후 환불요청/신용카드 취소요청 등록
            newasid = CheckNRegRefund(id, orderserial, reguserid)
            if (newasid>0) then
                ResultMsg = ResultMsg + "->. [환불 등록] 처리\n\n"
            end if
        End If


        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "009"
            Call FinishCSMaster(id, reguserid, contents_finish)
        End If


        ReturnUrl = "pop_lec_cs_register.asp?id="  + CStr(id) + "&divcd=" + divcd + "&mode=editfinishedinfo"


        If (Err.Number = 0) and (ScanErr="") Then
            dbget.CommitTrans
        Else
            dbget.RollBackTrans
            response.write "<script>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(에러코드 : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
            'response.write "<script>history.back()</script>"
            dbget.close()	:	response.End
        End If

        ''주문취소 완료 메일
        If (isCsMailSend) then
            Call SendCsActionMail(id)

            ''환불 접수 메일
            if (newasid>0) then
                Call SendCsActionMail(newasid)
            end if
        End IF
        On error Goto 0
    elseif (divcd="A003") or (divcd="A007") then
        ''환불요청
        dim RefreturnMethod, Refrealrefund
        On Error Resume Next
        dbget.beginTrans

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "001"
            Call FinishCSMaster(id, reguserid, contents_finish)
        End If

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "002"
            Call CheckRefundFinish(id, orderserial, RefreturnMethod, Refrealrefund)
        End If

        ResultMsg = "처리 완료"
        if (RefreturnMethod="R007") and (Refrealrefund>0) then
            ReturnUrl = "pop_lec_cs_register.asp?id="  + CStr(id) + "&divcd=" + divcd + "&finishtype=1"
        else
            ReturnUrl = "pop_lec_cs_register.asp?id="  + CStr(id) + "&divcd=" + divcd
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

        ''강좌확정후취소 접수(업체배송)  // 회수신청(텐바이텐배송) // 맞교환회수(텐바이텐배송)

        dim MinusOrderserial

        On Error Resume Next
        dbget.beginTrans


        If (modeflag2<>"norefund") and (Err.Number = 0) and (ScanErr="") Then
            errcode = "001"
            ''마이너스 주문 등록
            if (CheckNAddMinusOrder(id, orderserial, reguserid, MinusOrderserial, ScanErr)) then
                ResultMsg = ResultMsg + "->. [강좌확정 후 일부취소] 등록\n\n"
            end if
        End If


        If (modeflag2<>"norefund") and (Err.Number = 0) and (ScanErr="") Then
            errcode = "002"
            ''환불 등록건이 있는지 체크 후 환불요청/신용카드 취소요청 등록
            newasid = CheckNRegRefund(id, MinusOrderserial, reguserid)
            if (newasid>0) then
                ResultMsg = ResultMsg + "->. [취소(환불)접수] 처리\n\n"
            end if
        End If


        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "009"
            Call FinishCSMaster(id, reguserid, contents_finish)

            if (divcd="A004") then
                ResultMsg = ResultMsg + "->. 강좌확정 후 일부취소 처리 완료\n\n"
            elseif (divcd="A010") then
                ResultMsg = ResultMsg + "->. 회수 처리 완료\n\n"
            end if
        End If


        ResultMsg = ResultMsg
        ReturnUrl = "pop_lec_cs_register.asp?id="  + CStr(id) + "&divcd=" + divcd
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
    elseif  (divcd="A011") then
        On Error Resume Next
        dbget.beginTrans

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "009"
            Call FinishCSMaster(id, reguserid, contents_finish)
        End If


        ResultMsg = "처리 완료"
        ReturnUrl = "pop_lec_cs_register.asp?id="  + CStr(id) + "&divcd=" + divcd
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
            Call SendCsActionMail(id)
        End If
        On error Goto 0
    elseif  (divcd="A000") or (divcd="A001") or (divcd="A002") or (divcd="A009") or (divcd="A006") or (divcd="A005") or (divcd="A700") then
        ''맞교환 출고 / 누락 / 서비스 발송 / 기타 /  출고시 유의사항
        On Error Resume Next
        dbget.beginTrans

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "009"
            Call FinishCSMaster(id, reguserid, contents_finish)
        End If


        ResultMsg = "처리 완료"
        ReturnUrl = "pop_lec_cs_register.asp?id="  + CStr(id) + "&divcd=" + divcd
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
elseif (mode="state2jupsu") then
    if (divcd="A700") then
        '' 업체 기타정산 접수상태로 변경
        sqlStr = " select top 1 * from db_jungsan.dbo.tbl_designer_jungsan_detail"
        sqlStr = sqlStr + " where gubuncd='witakchulgo'"
        sqlStr = sqlStr + " and detailidx<>0"
        sqlStr = sqlStr + " and itemid=0"
        sqlStr = sqlStr + " and detailidx=" & id

        rsget.Open sqlStr,dbget,1
        if not rsget.Eof then
		    ResultMsg = "정산 내역이 존재합니다. 상태 변경 불가"
		else
		    ResultMsg = ""
		end if
		rsget.Close

        if (ResultMsg="") then
            sqlStr = " update db_cs.dbo.tbl_new_as_list"
            sqlStr = sqlStr + " set currstate='B001'"
            sqlStr = sqlStr + " ,finishdate=NULL"
            sqlStr = sqlStr + " where id=" & CStr(id)
            sqlStr = sqlStr + " and currstate='B007'"

            response.write sqlStr
            dbget.Execute sqlStr

            ResultMsg = "처리 완료"
            ReturnUrl = "pop_lec_cs_register.asp?id="  + CStr(id) + "&divcd=" + divcd
        else
            response.write "<script>alert('" + ResultMsg + "');</script>"
            response.write "<script>history.back();</script>"
            dbget.close()	:	response.End
        end if
    else
        ResultMsg = "정의되지 않았습니다. : mode=" + mode + " , divcd=" + divcd
        response.write "<script>alert('" + ResultMsg + "');</script>"
        response.write "<script>history.back();</script>"
        dbget.close()	:	response.End
    end if
elseif (mode="addupchejungsanEdit") then
    '' 업체 기타정산 접수상태로 변경
    sqlStr = " select top 1 * from db_jungsan.dbo.tbl_designer_jungsan_detail"
    sqlStr = sqlStr + " where gubuncd='witakchulgo'"
    sqlStr = sqlStr + " and detailidx<>0"
    sqlStr = sqlStr + " and itemid=0"
    sqlStr = sqlStr + " and detailidx=" & id

    rsget.Open sqlStr,dbget,1
    if not rsget.Eof then
	    ResultMsg = "정산 내역이 존재합니다. 수정 불가"
	else
	    ResultMsg = ""
	end if
	rsget.Close

    if (ResultMsg="") then
        if (add_upchejungsandeliverypay<>"") then
            call EditCSUpcheAddJungsanPay(id, add_upchejungsandeliverypay, add_upchejungsancause, buf_requiremakerid)
        end if

        ResultMsg = "처리 완료"
        ReturnUrl = "/cscenter/action/pop_AddUpchejungsanEdit.asp?id="  + CStr(id)
    else
        response.write "<script>alert('" + ResultMsg + "');</script>"
        response.write "<script>history.back();</script>"
        dbget.close()	:	response.End
    end if
elseif (mode="upcheconfirm2jupsu") then
    '' 업체 처리완료=>접수상태로변경
    sqlStr = " select top 1 currstate from [db_cs].[dbo].tbl_new_as_list"
    sqlStr = sqlStr + " where id=" + CStr(id)

    rsget.Open sqlStr,dbget,1
    if not rsget.Eof then
        ResultMsg = ""
        if (rsget("currstate")<>"B006") then
            ResultMsg = "업체 처리 완료 상태가 아닙니다. 수정 불가"
        end if
	else
	    ResultMsg = "코드없음. 수정 불가"

	end if
	rsget.Close

    if (ResultMsg="") then
        sqlStr = " update [db_cs].[dbo].tbl_new_as_list" + VbCrlf
        sqlStr = sqlStr + "set currstate='B001'" + VbCrlf
        sqlStr = sqlStr + ",contents_jupsu='" + (contents_jupsu) + "'" + VbCrlf
        sqlStr = sqlStr + " where id=" + CStr(id)

        dbget.Execute sqlStr

        ResultMsg = "처리 완료"
        ReturnUrl = "pop_lec_cs_register.asp?id="  + CStr(id) + "&divcd=" + divcd
    else
        response.write "<script>alert('" + ResultMsg + "');</script>"
        response.write "<script>history.back();</script>"
        dbget.close()	:	response.End
    end if
else
    ResultMsg = "정의되지 않았습니다. : mode=" + mode + " , divcd=" + divcd
    response.write "<script>alert('" + ResultMsg + "');</script>"
    response.write "<script>history.back();</script>"
    dbget.close()	:	response.End
end if

'=============================================================
'			메일 SMS 발송 관련
'=============================================================
'' 위쪽 중간에 삽입

''dim isMailProc '// 메일 발송여부
''dim isSmsProc	'// SMS 발송 여부
''
''isMailProc = False
''isSmsProc = False
''
''IF mode="regcsas" THEN
''	IF divcd ="A000" or divcd ="A001" or divcd ="A002" or divcd ="A003" or divcd ="A004" or divcd ="A007" or divcd ="A010" THEN
''		isMailProc=True
''	END IF
''ELSEIF mode="finishcsas" THEN
''	IF divcd ="A000" or divcd ="A001" or divcd ="A002" or divcd ="A003" or divcd ="A004" or divcd ="A007" or divcd ="A008" or divcd ="A010" or divcd="A900" THEN
''		isMailProc=True
''	END IF
''END IF
''
''IF mode="regcsas" THEN
''	IF divcd ="A000" or divcd ="A001" or divcd ="A002" or divcd ="A003" or divcd ="A004" or divcd ="A007" or divcd ="A010" THEN
''		isSmsProc= True
''	END IF
''
''ELSEIF mode="finishcsas" THEN
''	IF divcd ="A000" or divcd ="A001" or divcd ="A002" or divcd ="A003" or divcd ="A004" or divcd ="A007" or divcd ="A008" or divcd ="A010" or divcd="A900" THEN
''		isSmsProc= True
''	END IF
''END IF
''
'''//=======  SMS 발송 안함 =========/
''IF isSmsProc THEN
''	'oCsAction.sendSMS "",""
''	'oCsAction.sendSMS "010-8831-6240",""
''END IF


%>

<!-- #include virtual="/lib/db/dbclose.asp" -->

<%
response.write "<script>alert('" + ResultMsg + "');</script>"
response.write "<script>location.replace('" + ReturnUrl + "');</script>"
%>


<%

''''    	If (Err.Number = 0) and (ScanErr="") Then
''''            errcode = "003"
''''
''''            if (divcd="A020") then
''''        	    ''전체 취소인경우
''''        	    ''1- 전체 취소 인지 유효성 체크
''''        	    if Not (IsAllCancelRegValid(id, orderserial)) then
''''        	        ScanErr = "전체 취소 검증 오류. - 전체 취소 아님."
''''        	    end if
''''
''''
''''        	elseif (divcd="A021") then
''''        	    ''부분 취소인경우
''''        	    ''1- 부분 취소 인지 유효성 체크
''''        	    if Not (IsPartialCancelRegValid(id, orderserial)) then
''''        	        ScanErr = "전체 취소 검증 오류. - 부분 취소 아니거나 내역없음."
''''        	    end if
''''        	end if
''''        end if
%>
