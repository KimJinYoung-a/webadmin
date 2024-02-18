<%@ language=vbscript %>
<% option explicit %>
<%
session.codePage = 949
Response.CharSet = "EUC-KR"
%>
<%
'###########################################################
' Description : cs센터
' History : 2009.04.17 이상구 생성
'			2023.10.20 한용민 수정(상품쿠폰 환원 로직 추가)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_LogisticsOpen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<!-- #include virtual="/lib/util/base64unicode.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp"-->
<!-- #include virtual="/cscenter/lib/CSFunction.asp"-->
<!-- #include virtual="/cscenter/lib/cs_action_mail_Function.asp"-->
<!-- #include virtual="/lib/email/MailLib2.asp"-->
<!-- #include virtual="/lib/util/DcCyberAcctUtil.asp"-->
<!-- #include virtual="/lib/email/smsLib.asp"-->
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
'elseif (mode="finishededitcsas") then
'	'==========================================================================
'	'완료된 내역 수정
'
'elseif (mode="delfinishedcsas") then
'	'==========================================================================
'	'완료된 내역 삭제
'
'elseif (mode="realdelcsas") then
'	'==========================================================================
'	'데이타베이스에서 내역 DELETE
'
'elseif (mode="finishcsas") then
'	'==========================================================================
'	'CS 접수 내역 완료처리
'
'elseif (mode="state2jupsu") then
'	'==========================================================================
'	'업체 기타정산 접수상태로 변경
'
'elseif (mode="addupchejungsanEdit") then
'	'==========================================================================
'	'업체추가정산 수정
'
'elseif (mode="upcheconfirm2jupsu") then
'	'==========================================================================
'	'업체 처리완료 => 접수상태로변경
'
'elseif (mode="upcheconfirm2reconfirm") then
'	'==========================================================================
'	'업체 처리완료 => 업체재확인요청상태로변경
'
'elseif (mode="changeorderreg") then
'	'==========================================================================
'	'교환주문 수기생성
'
'elseif (mode="changedivcdtoa004") then
'	'==========================================================================
'	'고객 직접반품 전환(A010 -> A004)
'
'elseif (mode="changedivcdtoa010") then
'	'==========================================================================
'	'회수신청 전환(A004 -> A010)
'
'elseif (mode="restoredel") then
'	'==========================================================================
'	'삭제CS(완료이전) 복구
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
'A004			반품접수(업체배송)
'A010			회수신청(텐바이텐배송)
'
'A001			누락재발송
'A002			서비스발송

'A200			기타회수
'
'A000			맞교환출고
'
'A009			기타사항
'A006			출고시유의사항
'A700			업체기타정산
'
'A003			환불
'A005			외부몰환불요청
'
'A011			맞교환회수(텐바이텐배송)



dim mode, modeflag2, divcd, id, reguserid, ipkumdiv, fullText, failText, btnJson, manager_hp, itemName, itemCnt, affectedRows
dim title, orderserial, gubun01, gubun02, contents_jupsu, refundstr, refundresult, buyhp
dim finishuser, contents_finish, refasid

dim requireupche, requiremakerid, ForceReturnByTen
dim detailitemlist
dim csdetailitemlist

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

'고객 추가배송비(상품변경 맞교환)
dim add_customeraddmethod, add_customeradditempay, add_customeradditembuypay, add_customeraddbeasongpay, customerrealbeasongpay, customerpayordertype

'고객 추가배송비(반품)
dim addbeasongpay, addmethod

'교환주문
dim changeorderserial, changeorderstate

'품절취소 상품정보 저장
dim modifyitemstockoutyn

dim isCSServiceRefund

'쿠폰 재발행
dim copycouponinfo, copyitemcouponinfo, resultItemCouponCount
resultItemCouponCount=0
dim ocsaslist
dim oRefCSASList, refminusorderserial, refchangeorderserial

dim regDetailRows

dim needChkYN
dim orefund

dim accountdiv, acctno, acctname

dim opayordermaster, payorderserial

newasid = -1
itemCnt=0
itemName = ""
manager_hp = ""
buyhp = ""
mode        = request.Form("mode")
modeflag2   = request.Form("modeflag2")
divcd       = request.Form("divcd")
id          = request.Form("id")
ipkumdiv    = request.Form("ipkumdiv")
reguserid   = session("ssbctid")
finishuser  = reguserid
title       = html2DB(request.Form("title"))
orderserial = request.Form("orderserial")
gubun01     = request.Form("gubun01")
gubun02     = request.Form("gubun02")
contents_jupsu  	= html2DB(request.Form("contents_jupsu"))
detailitemlist  	= html2db(request.Form("detailitemlist"))
csdetailitemlist  	= html2db(request.Form("csdetailitemlist"))
contents_finish 	= html2DB(request.Form("contents_finish"))

''업체 처리 요청
requireupche = request.Form("requireupche")
requiremakerid = request.Form("requiremakerid")
ForceReturnByTen = request.Form("ForceReturnByTen")

orgitemcostsum      = request.Form("orgitemcostsum")
orgbeasongpay       = request.Form("orgbeasongpay")
orgmileagesum       = request.Form("miletotalprice")
orgcouponsum        = request.Form("tencardspend")
orgallatdiscountsum = request.Form("allatdiscountprice")
orgsubtotalprice    = request.Form("subtotalprice")

orggiftcardsum    	= request.Form("giftcardsum")			'플러스값을 저장한다.
refundgiftcardsum   = request.Form("refundgiftcardsum")
orgdepositsum    	= request.Form("depositsum")			'플러스값을 저장한다.
refunddepositsum    = request.Form("refunddepositsum")

refunditemcostsum   = request.Form("refunditemcostsum")
nextsubtotal        = request.Form("nextsubtotal")
canceltotal         = request.Form("canceltotal")

refundbeasongpay    = request.Form("refundbeasongpay")
remainbeasongpay    = request.Form("remainbeasongpay")
refunddeliverypay   = request.Form("refunddeliverypay")

refundmileagesum    = request.Form("refundmileagesum")
refundcouponsum     = request.Form("refundcouponsum")
allatsubtractsum    = request.Form("allatsubtractsum")
refundadjustpay     = request.Form("refundadjustpay")
remainitemcostsum   = request.Form("remainitemcostsum")



''환불요청액
refundrequire       = request.Form("refundrequire")
returnmethod        = request.Form("returnmethod")

rebankname          = request.Form("rebankname")
rebankaccount       = request.Form("rebankaccount")
rebankownername     = request.Form("rebankownername")

encmethod 			= "AE2" ''"PH1"

paygateTid          = request.Form("paygateTid")

add_upchejungsandeliverypay = request.Form("add_upchejungsandeliverypay")
add_upchejungsancause       = request.Form("add_upchejungsancause")
add_upchejungsancauseText   = request.Form("add_upchejungsancauseText")

buf_requiremakerid  = request.Form("buf_requiremakerid")


isCsMailSend = (request.Form("csmailsend")="on")

add_customeraddmethod   	= request.Form("add_customeraddmethod")
add_customeradditempay   	= request.Form("add_customeradditempay")
add_customeradditembuypay  	= request.Form("add_customeradditembuypay")
add_customeraddbeasongpay   = request.Form("add_customeraddbeasongpay")
customerpayordertype   		= request.Form("customerpayordertype")
customerrealbeasongpay   	= request.Form("customerrealbeasongpay")

modifyitemstockoutyn   		= request.Form("modifyitemstockoutyn")

copycouponinfo        		= request.Form("copycouponinfo")

addbeasongpay   			= request.Form("addbeasongpay")
addmethod   				= request.Form("addmethod")

needChkYN   				= request.Form("needChkYN")

accountdiv   				= request.Form("accountdiv")
acctno   					= request.Form("acctno")
acctname   					= request.Form("acctname")

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

if (Not IsNumeric(add_customeradditempay)) or (add_customeradditempay="") then add_customeradditempay = 0
if (Not IsNumeric(add_customeradditembuypay)) or (add_customeradditembuypay="") then add_customeradditembuypay = 0
if (Not IsNumeric(add_customeraddbeasongpay)) or (add_customeraddbeasongpay="") then add_customeraddbeasongpay = 0
if (Not IsNumeric(customerrealbeasongpay)) or (customerrealbeasongpay="") then customerrealbeasongpay = 0

if (Not IsNumeric(addbeasongpay)) or (addbeasongpay="") then addbeasongpay = 0

if (returnmethod="") then returnmethod="R000"
if (copycouponinfo="") then copycouponinfo="N"

''올엣카드할인.. -상품별로 차감.

dim sqlStr, errcode, i
dim ScanErr
dim ResultMsg, ReturnUrl, EtcStr
dim ProceedFinish
dim ResultCount

ScanErr = ""
ProceedFinish = False

dim IsAllCancel
dim CancelValidResultMessage

''과거 주문 내역인지 Check
GC_IsOLDOrder = CheckIsOldOrder(orderserial)



'==============================================================================
''주문 마스타
dim oordermaster

set oordermaster = new COrderMaster

oordermaster.FRectOrderSerial = orderserial

if Left(orderserial,1)="A" then
    set oordermaster.FOneItem = new COrderMasterItem
else
    oordermaster.QuickSearchOrderMaster
end if

'' 과거 6개월 이전 내역 검색
if (oordermaster.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
    oordermaster.FRectOldOrder = "on"
    oordermaster.QuickSearchOrderMaster
end if

if (IsAutoScript) and mode <> "finishcsas" then
	response.write "S_ERR|작업되지 않았습니다. mode : " & mode
	dbget.close()	:	response.End
end if

if (mode="regcsas") then
    '==========================================================================
	'CS 접수
    if (divcd="A008") then

		'----------------------------------------------------------------------
        'CS 접수 - 주문취소
        'On Error Resume Next
        dbget.beginTrans

		'// 6개월 이전 주문 취소가능(2014-03-31)
        ''if (GC_IsOLDOrder) then ScanErr = "6개월 이전 과거 내역 취소 불가 - 관리자 문의 요망"

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "001"

            'CS Master 접수
            id = RegCSMaster(divcd, orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
        end if

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "002"

            'CS Master 환불 관련정보 저장
	        Call RegCSMasterRefundInfo(id, returnmethod, refundrequire , orgsubtotalprice, orgitemcostsum, orgbeasongpay , orgmileagesum, orgcouponsum, orgallatdiscountsum  , canceltotal, refunditemcostsum, refundmileagesum, refundcouponsum, allatsubtractsum , refundbeasongpay, refunddeliverypay, refundadjustpay , rebankname, rebankaccount, rebankownername, paygateTid)
	        Call AddCSMasterRefundInfo(id, orggiftcardsum, orgdepositsum, refundgiftcardsum, refunddepositsum)

            '''계좌 암호화 추가.
	        Call EditCSMasterRefundEncInfo(id, encmethod, rebankaccount)

			'' 보너스쿠폰 재발행할지
			Call EditCSCopyCouponInfo(id, copycouponinfo)
	    End if

	    If (Err.Number = 0) and (ScanErr="") Then
            errcode = "003"

            'CS Detail 접수(관련 상품정보)
	        Call AddCSDetailByArrStr(detailitemlist, id, orderserial)

        end if

		' 상품쿠폰환급여부
		if itemCouponRefundYN="Y" then
			If (Err.Number = 0) and (ScanErr="") Then
				errcode = "011"

				' 주문 상품쿠폰 적용수 체크
				resultItemCouponCount = ItemCouponCount(id, "P", oordermaster.FOneItem.Fuserid)
				if resultItemCouponCount>0 then
					copyitemcouponinfo="Y"
				else
					copyitemcouponinfo="N"
				end if

				' 상품쿠폰 재발행할지
				Call EditCSCopyItemCouponInfo(id, copyitemcouponinfo)
			end if
		end if

	    If (Err.Number = 0) and (ScanErr="") Then
            errcode = "005"

            if (remainitemcostsum = 0) then
            	'전체취소 : 잔여상품총액이 없는 경우
            	IsAllCancel = true
            	CancelValidResultMessage = GetAllCancelRegValidResult(id, orderserial)
            else
	            IsAllCancel = false
	            CancelValidResultMessage = GetPartialCancelRegValidResult(id, orderserial)
	        end if

			if (CancelValidResultMessage <> "") then
				ScanErr = CancelValidResultMessage
			end if
        End If

        '출고완료 또는 취소된 상품이 있을 경우, 진행정지(주문취소 불가)
        '출고완료된 상품은 반품만 가능하다.
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
	            errcode = "000"

                ''마일리지 환원 체크
				Call CheckRefundMileage(id, orderserial)

	            ''금액 체크
				Call CheckRefundPrice(id, orderserial, ScanErr)
	        End If

            If (Err.Number = 0) and (ScanErr="") Then
                errcode = "008"

                Call CancelProcess(id, orderserial)

                ResultMsg = ResultMsg + "->. 주문건 취소 완료\n\n"
            End IF

            ''순서?. 위로?
            If (Err.Number = 0) and (ScanErr="") Then
                errcode = "009"

                '환불 정보가 있는지 체크 후 무통장환불/마일리지환불/신용카드취소 CS 접수 등록
                newasid = CheckNRegRefund(id, orderserial,reguserid)

                If (newasid>0) then
                    ResultMsg = ResultMsg + "->. 환불 접수 완료\n\n"
                end if
            End If

            If (Err.Number = 0) and (ScanErr="") Then
                errcode = "010"

                Call FinishCSMaster(id, reguserid, contents_finish)

				'// 취소 접수상태 내역의 금액 업데이트
				Call UpdateCancelJupsuCSPrice(id, orderserial)

                ResultMsg = ResultMsg + "->. [주문 취소 CS] 완료 처리\n\n"

				if CheckAndCopyBonusCoupon(id, reguserid) = True then
					ResultMsg = ResultMsg + "->. [보너스쿠폰 재발급] 완료 처리\n\n"
				end if
            End If

            If (Err.Number = 0) and (ScanErr="") Then
                errcode = "012"

				' 상품쿠폰재발급
				if CheckAndCopyItemCoupon(id, reguserid, "P", oordermaster.FOneItem.Fuserid) = True then
					ResultMsg = ResultMsg + "->. [상품쿠폰 재발급] 완료 처리\n\n"
				end if
            End If
        ELSE
            ResultMsg = ResultMsg + "->. 상품 준비중 상태인 상품이 존재하므로,주문 취소 접수만 진행 되었습니다.\n확인후 완료 처리하셔야 합니다.\n\n"

			If (Err.Number = 0) and (ScanErr="") Then
				errcode = "004"

				'// 상품준비중인 내역의 주문취소접수일 때, 업체 어드민에 노출
				if (requiremakerid<>"") then
					call RegCSMasterAddUpche(id, requiremakerid)
				end if
			end if
        End If

	    If (Err.Number = 0) and (ScanErr="") Then
            dbget.CommitTrans

			If (ProceedFinish) then
			Else
				if (requiremakerid<>"") then
					' 업체와 협의후 접수후 바로 취소시키는 건이 있어서 접수후 바로 알림톡이 나가는게 아니고 지연이 필요하다 해서 15분에 한번씩 스케줄로 돌게 디비화시킴.
					' sqlStr ="select replace(isnull(isnull(g.manager_hp,p.manager_hp),''),'-','') as manager_hp"
					' sqlStr = sqlStr & " from db_partner.dbo.tbl_partner p with (nolock)"
					' sqlStr = sqlStr & " join db_partner.dbo.tbl_partner_group g with (nolock)"
					' sqlStr = sqlStr & " 	on p.groupid=g.groupid"
					' sqlStr = sqlStr & " where p.id='"& requiremakerid &"'"

					' 'response.write sqlStr & "<Br>"
					' rsget.CursorLocation = adUseClient
					' rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

					' IF not rsget.EOF THEN
					' 	manager_hp = rsget("manager_hp")
					' END IF
					' rsget.close

					' if manager_hp<>"" then
					' 	sqlStr ="select max(replace(replace(replace(replace(replace(ad.itemname,char(9),''),char(10),''),char(13),''),'""',''),'''','')) as itemname, count(ad.masterid) as itemcnt"
					' 	sqlStr = sqlStr & " from db_cs.dbo.tbl_new_as_detail ad with (nolock)"
					' 	sqlStr = sqlStr & " where ad.itemid not in (0) and ad.masterid="& id &""	' 주문취소

					' 	'response.write sqlStr & "<Br>"
					' 	rsget.CursorLocation = adUseClient
					' 	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

					' 	IF not rsget.EOF THEN
					' 		itemCnt = rsget("itemcnt")

					' 		if itemName = "" then
					' 			itemName = replace(db2html(rsget("itemname")),vbcrlf,"")
					' 		end if
					' 	END IF
					' 	rsget.close

					' 	if itemCnt > 1 then
					' 		itemName = itemName & " 외 " & (itemCnt - 1) & "종"
					' 	end if

					' 	' 업체확인 중 주문취소 접수		' 2021.09.30 한용민 생성
					' 	fullText = "[10x10] 주문취소 접수 안내" & vbCrLf & vbCrLf
					' 	fullText = fullText & "판매자님의 상품에 대해 고객이 주문취소를 요청하였습니다." & vbCrLf
					' 	fullText = fullText & "확인 후 주문 취소해 주시길 바랍니다." & vbCrLf & vbCrLf
					' 	fullText = fullText & "■ 브랜드ID : "& requiremakerid &"" & vbCrLf
					' 	fullText = fullText & "■ 주문번호 : "& orderserial &"" & vbCrLf
					' 	fullText = fullText & "■ 상품명 : "& itemName &""
					' 	failText = "[텐바이텐]고객이 주문취소를 요청하였습니다.주문번호:" & orderserial
					' 	btnJson = "{""button"":[{""name"":""SCM 바로가기"",""type"":""WL"", ""url_mobile"":""https://scm.10x10.co.kr/""}]}"
					' 	call SendKakaoCSMsg_LINK("", manager_hp,"1644-6030","KC-0022",fullText,"SMS","",failText,btnJson,"","")
					' 	ResultMsg = ResultMsg + "->. 업체 담당자에게 주문취소 요청 카톡(문자)이 발송되었습니다.\n\n"
					' end if
				end if
			End If
        Else
            dbget.RollBackTrans
            response.write "<script>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(에러코드 : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
            response.write "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(에러코드 : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")"
            ''response.write "<script>history.back()</script>"
            dbget.close()	:	response.End
        End If

		'' 업배상품 상품준비중이면 업체확인 후 취소하는 두단계인데 일단은 무시한다.
		''If (ProceedFinish) then

			''가상계좌 금액/마감일 수정
			Call CheckNChangeCyberAcct(orderserial)

			if IsAllCancel = true then
				''전자보증 취소
				Call CheckNUsafeCancel(orderserial)
			end if

			''재고 및 한정수량 조절(2007-09-01 서동석 추가)
			''Call LimitItemRecover(orderserial) : 기존
			if (IsAllCancel) then
				''전체 취소인경우
				sqlStr = " exec [db_summary].[dbo].sp_Ten_RealtimeStock_cancelOrderAll '" & orderserial & "'"
				dbget.Execute sqlStr
			else
				''부분 취소인경우
				sqlStr = " select itemid,itemoption,regitemno "
				sqlStr = sqlStr & " from [db_cs].[dbo].tbl_new_as_detail "
				sqlStr = sqlStr & " where masterid=" & id
				sqlStr = sqlStr & " and orderserial='" & orderserial & "'"

				rsget.Open sqlStr,dbget,1
				if Not rsget.Eof then
					regDetailRows = rsget.getRows()
				end if
				rsget.Close

				if IsArray(regDetailRows) then
					for i=0 to UBound(regDetailRows,2)
						sqlStr = " exec [db_summary].[dbo].sp_Ten_RealtimeStock_cancelOrderPartial '" & orderserial & "'," & regDetailRows(0,i) & ",'" & regDetailRows(1,i) & "'," & regDetailRows(2,i)
						dbget.Execute sqlStr
					Next
				end if
			end if

			'' 취소 업배 상품중 품절상품의 경우 상품정보에 품절설정
			if (modifyitemstockoutyn = "Y") then
				ResultCount   = SetStockOutByCsAs(id)
				if (ResultCount > 0) then
					ResultMsg = ResultMsg + "->. [상품정보 품절 설정] 완료 처리\n\n"
				end if
			end if

		''end if

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
        'on error Goto 0

        ReturnUrl = "/cscenter/action/pop_cs_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd + "&mode=editreginfo"

    elseif (divcd="A004") or (divcd="A010") then
    	'----------------------------------------------------------------------
        'CS 접수 - 반품 접수 또는 회수신청.
        On Error Resume Next
	        dbget.beginTrans

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "001"

	            '' CS Master 접수
	            id = RegCSMaster(divcd, orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
	        end if

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "002"

	            'CS Master 환불 관련정보 저장
		        Call RegCSMasterRefundInfo(id, returnmethod, refundrequire , orgsubtotalprice, orgitemcostsum, orgbeasongpay , orgmileagesum, orgcouponsum, orgallatdiscountsum  , canceltotal, refunditemcostsum, refundmileagesum, refundcouponsum, allatsubtractsum , refundbeasongpay, refunddeliverypay, refundadjustpay , rebankname, rebankaccount, rebankownername, paygateTid)
		        Call AddCSMasterRefundInfo(id, orggiftcardsum, orgdepositsum, refundgiftcardsum, refunddepositsum)

		        '''계좌 암호화 추가.
	            Call EditCSMasterRefundEncInfo(id, encmethod, rebankaccount)

				''쿠폰 재발행할지
				Call EditCSCopyCouponInfo(id, copycouponinfo)
		    End if


		    If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "003"

	            'CS Detail 접수(관련 상품정보)
		        Call AddCSDetailByArrStr(detailitemlist, id, orderserial)
	        end if

	        '' Check - 업체배송과 자체배송을 같이 접수하지 못함.
	        ''       - 업체배송이 존재할 경우 한개의 브랜드만 존재 해야함.
	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "004"

	            if (IsReturnRegValid(id, orderserial, ScanErr, requiremakerid)) then
	                '업체배송인 경우 관련 업체 브랜드 아이디 입력(requiremakerid)
	                if ((requiremakerid<>"") and (ForceReturnByTen="")) then
	                    call RegCSMasterAddUpche(id, requiremakerid)
	                end if

	                ResultMsg = ResultMsg + "->. [반품 / 회수 CS] 접수\n\n"
	            end if
	        End if

	        ''업체 추가 정산비 2008.11.10
	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "005"

				'// 상품대금인 경우 정산금액 0원, skyer9, 2015-09-02
	            if (add_upchejungsandeliverypay<>"0") and (add_upchejungsandeliverypay<>"")  then
	                call RegCSUpcheAddJungsanPay(id, add_upchejungsandeliverypay, add_upchejungsancause, buf_requiremakerid)
	            end if
	        end If

			If (Err.Number = 0) and (ScanErr="") Then
				errcode = "006"

				'' 고객 추가배송비
				if (divcd = "A004") or (divcd = "A010")  then
					Call SetCustomerAddBeasongPay(id, addmethod, addbeasongpay, "Y", 0)
				end if

			end if

	        ReturnUrl = "/cscenter/action/pop_cs_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd

	        If (Err.Number = 0) and (ScanErr="") Then
	            dbget.CommitTrans
	        Else
	            dbget.RollBackTrans
	            response.write "<script>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(에러코드 : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
	            'response.write "<script>history.back()</script>"
	            dbget.close()	:	response.End
	        End If

	        ''이메일 발송. 반품 회수 접수
	        If (isCsMailSend) then
	            Call SendCsActionMail(id)

				buyhp=""
				sqlStr ="select replace(isnull(m.buyhp,''),'-','') as buyhp"
				sqlStr = sqlStr & " from db_order.dbo.tbl_order_master m with (nolock)"
				sqlStr = sqlStr & " where m.orderserial="& orderserial &""

				'response.write sqlStr & "<Br>"
				rsget.CursorLocation = adUseClient
				rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

				IF not rsget.EOF THEN
					buyhp = rsget("buyhp")
				END IF
				rsget.close

				if buyhp<>"" then
					itemCnt=0
					itemName = ""
					sqlStr ="select max(replace(replace(replace(replace(replace(ad.itemname,char(9),''),char(10),''),char(13),''),'""',''),'''','')) as itemname, count(ad.masterid) as itemcnt"
					sqlStr = sqlStr & " from db_cs.dbo.tbl_new_as_detail ad with (nolock)"
					sqlStr = sqlStr & " where itemid not in (0) and ad.masterid="& id &""	' 주문취소

					'response.write sqlStr & "<Br>"
					rsget.CursorLocation = adUseClient
					rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

					IF not rsget.EOF THEN
						itemCnt = rsget("itemcnt")

						itemName = replace(db2html(rsget("itemname")),vbcrlf,"")
					END IF
					rsget.close

					if itemCnt > 1 then
						itemName = itemName & " 외 " & (itemCnt - 1) & "종"
					end if

					' A004 반품접수(업체배송)
					if (divcd="A004") then
						' 반품 접수(업체 배송)		' 2021.10.12 한용민 생성
						fullText = "[10x10] 반품안내" & vbCrLf & vbCrLf
						fullText = fullText & "주문하신 상품이 반품 접수 되었습니다." & vbCrLf & vbCrLf
						fullText = fullText & "■ 주문번호 : "& orderserial &"" & vbCrLf
						fullText = fullText & "■ 상품명 : "& itemName &"" & vbCrLf & vbCrLf
						fullText = fullText & "- 신청하신 상품은 [업체개별배송]상품으로 회수서비스가 제공되지 않습니다." & vbCrLf
						fullText = fullText & "수령운송장 또는 반품주소지 확인 후 택배사 예약 및 반품하실 상품재포장 후" & vbCrLf
						fullText = fullText & "택배기사님에게 상품 전달 부탁드립니다." & vbCrLf & vbCrLf
						fullText = fullText & "※택배접수관련 고객센터로 문의주시면 조금 더 상세한 안내드리겠습니다."
						failText = "[텐바이텐]주문하신 상품이 반품 접수 되었습니다.주문번호:" & orderserial
                        '// 업배 반품시 카톡발송 안함 : 고객센터에서 택배접수 대행하고 있음, 2022-06-28, skyer9
						''call SendKakaoCSMsg_LINK("", buyhp,"1644-6030","KC-0023",fullText,"SMS","",failText,"","","")

					' A010 회수신청(텐바이텐배송)
					elseif (divcd="A010") then
						' 반품 접수(텐텐 배송)		' 2021.10.12 한용민 생성
						fullText = "[10x10] 반품안내" & vbCrLf & vbCrLf
						fullText = fullText & "주문하신 상품이 반품 접수 되었습니다." & vbCrLf & vbCrLf
						fullText = fullText & "■ 주문번호 : "& orderserial &"" & vbCrLf
						fullText = fullText & "■ 상품명 : "& itemName &"" & vbCrLf & vbCrLf
						fullText = fullText & "- 반품요청 상품이 파손되지 않도록 재포장 후 보관해주시면," & vbCrLf
						fullText = fullText & "1-3일 이내(영업일기준)에 회수방문예정입니다." & vbCrLf
						fullText = fullText & "반품 완료 후 영업일 기준으로 1-3일내에 결제수단으로 환불 처리 됩니다."

						failText = "[텐바이텐]주문하신 상품이 반품 접수 되었습니다.주문번호:" & orderserial
						call SendKakaoCSMsg_LINK("", buyhp,"1644-6030","KC-0012",fullText,"SMS","",failText,"","","")
					end if
				end if
	        end if
        on error Goto 0

    elseif (divcd="A001") or (divcd="A002") or (divcd="A200") then
    	'----------------------------------------------------------------------
        'CS 접수 - 누락재발송, 서비스발송, 기타회수
        On Error Resume Next
	        dbget.beginTrans

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "001"

	            '' CS Master 접수
	            id = RegCSMaster(divcd, orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
	        end if

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "002"

	            'CS Detail 접수(관련 상품정보)
		        Call AddCSDetailByArrStr(detailitemlist, id, orderserial)
	        end if

			'업체배송인 경우 관련 업체 브랜드 아이디 입력(requiremakerid)
	        if (requiremakerid<>"") then
	            call RegCSMasterAddUpche(id, requiremakerid)
	        end if

	        ''업체 추가 정산비 2012-06-25(skyer9)
	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "003"

	            if (add_upchejungsandeliverypay<>"0") and (add_upchejungsandeliverypay<>"")  then
	                call RegCSUpcheAddJungsanPay(id, add_upchejungsandeliverypay, add_upchejungsancause, buf_requiremakerid)
	            end if
	        end if

	        ResultMsg = "접수완료"
	        ReturnUrl = "/cscenter/action/pop_cs_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd

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
				if Not (divcd="A200") then
					'// 기타회수 메일발송 안함
	            	Call SendCsActionMail(id)
				end if
	        End If
        on error Goto 0

    elseif (divcd="A000") then
		'----------------------------------------------------------------------
        'CS 접수 - 맞교환출고
        On Error Resume Next
	        dbget.beginTrans

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "001"

	            '' CS Master 접수
	            id = RegCSMaster(divcd, orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
	        end if

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "002"

	            'CS Detail 접수(관련 상품정보)
		        Call AddCSDetailByArrStr(detailitemlist, id, orderserial)
	        end if

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "003"

				'// CS 맞교환출고(동일상품, 상품변경 - A000, A100) 접수시 출고되는 상품 한정차감
				Call ApplyLimitItemByCS(id)
	        end if

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "004"

	            if (requiremakerid<>"") then

	                '업체배송인 경우 관련 업체 브랜드 아이디 입력(requiremakerid)
	                Call RegCSMasterAddUpche(id, requiremakerid)

	                '업체배송인 경우 맞교환 회수 접수
                    newasid = RegCSMaster("A012", orderserial, reguserid, "교환회수(업체배송) 접수", contents_jupsu, gubun01, gubun02)

					'업체배송인 경우 관련 업체 브랜드 아이디 입력(requiremakerid)
                    Call RegCSMasterAddUpche(newasid, requiremakerid)

                    Call AddCSDetailByArrStr(detailitemlist, newasid, orderserial)

					'// asid 연결
					Call SetRefAsid(newasid, id)

                    ResultMsg = "교환 출고 접수 및 회수접수 완료 - 업체 배송"

	            else

	                '텐바이텐 배송의 경우 맞교환 회수 접수
	                newasid = RegCSMaster("A011", orderserial, reguserid, "교환회수 접수", contents_jupsu, gubun01, gubun02)

	                Call AddCSDetailByArrStr(detailitemlist, newasid, orderserial)

					'// asid 연결
					Call SetRefAsid(newasid, id)

	                ResultMsg = "교환 출고 접수 및 회수접수 완료 - 텐바이텐 배송"

	            end if
	        end if

	        ''업체 추가 정산비 2008.11.10
	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "005"

	            if (add_upchejungsandeliverypay<>"0") and (add_upchejungsandeliverypay<>"")  then
	                call RegCSUpcheAddJungsanPay(id, add_upchejungsandeliverypay, add_upchejungsancause, buf_requiremakerid)
	            end if
	        end if

	        ReturnUrl = "/cscenter/action/pop_cs_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd

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
	            if (newasid>0) and (requiremakerid = "") then
	            	'// 업체배송은 보내지 않는다.
   	                Call SendCsActionMail(newasid)
	            end if
	        End If
        on error Goto 0

    elseif (divcd="A009") or (divcd="A006") or (divcd="A060") or (divcd="A700") then
    	'----------------------------------------------------------------------
        'CS 접수 - 기타사항 / 출고시유의사항 / 업체긴급문의 / 업체 추가 정산비
        On Error Resume Next
	        dbget.beginTrans

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "001"

	            '' CS Master 접수
	            id = RegCSMaster(divcd, orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
	        end if

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "002"

	            'CS Detail 접수(관련 상품정보)
		        Call AddCSDetailByArrStr(detailitemlist, id, orderserial)
	        end if

	        ''업체 추가 정산비 : 2008.11.10
	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "003"

				'// 상품대금인 경우 정산금액 0원, skyer9, 2015-09-02
				if ((add_upchejungsandeliverypay<>"0") or ((divcd = "A700") And (add_upchejungsancause = "상품대금"))) and (add_upchejungsandeliverypay<>"")  then
	                call RegCSUpcheAddJungsanPay(id, add_upchejungsandeliverypay, add_upchejungsancause, buf_requiremakerid)
	            end if
	        end if

	        ''업체지정.
	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "004"

	            '업체배송인 경우 관련 업체 브랜드 아이디 입력(requiremakerid)
	            if (requiremakerid<>"") then
	                call RegCSMasterAddUpche(id, requiremakerid)
	            end if
	         end if

	        ResultMsg = "등록되었습니다."
	        ReturnUrl = "/cscenter/action/pop_cs_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd

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
    	'----------------------------------------------------------------------
        'CS 접수 - 환불접수 / 외부몰 환불접수
		isCSServiceRefund = False
        On Error Resume Next
	        dbget.beginTrans

	        if (divcd="A005") and (Not IsExtSiteOrder(orderserial)) then
	            ScanErr = "외부몰 환불접수는 외부몰 주문건만 가능합니다."
	        end if

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "001"

	            '' CS Master 접수
	            if (divcd="A003") then
					if (returnmethod = "R900") then
						title = title & "(마일리지)"
					elseif (returnmethod = "R910") then
						title = title & "(예치금)"
					elseif (returnmethod = "R100") then
						title = title & "(신용카드취소)"
					elseif (returnmethod = "R550") then
						title = title & "(기프팅취소)"
					elseif (returnmethod = "R560") then
						title = title & "(기프티콘취소)"
					elseif (returnmethod = "R120") then
						title = title & "(신용카드부분취소)"
					elseif (returnmethod = "R022") then
						title = title & "(실시간이체부분취소)"
					elseif (returnmethod = "R007") then
						title = title & "(무통장)"
					end if

					'// 서비스환불인지
					isCSServiceRefund = GetIsCSServiceRefund(id, divcd, title)

					title = GetCSRefundTitle(id, divcd, orderserial, returnmethod, title)
				end if

	            id = RegCSMaster(divcd, orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
	        end if

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "002"

	            'CS Master 환불 관련정보 저장
		        Call RegCSMasterRefundInfo(id, returnmethod, refundrequire , orgsubtotalprice, orgitemcostsum, orgbeasongpay , orgmileagesum, orgcouponsum, orgallatdiscountsum  , canceltotal, refunditemcostsum, refundmileagesum, refundcouponsum, allatsubtractsum , refundbeasongpay, refunddeliverypay, refundadjustpay , rebankname, rebankaccount, rebankownername, paygateTid)
		        Call AddCSMasterRefundInfo(id, orggiftcardsum, orgdepositsum, refundgiftcardsum, refunddepositsum)

				if (isCSServiceRefund) then
					Call SetCSServiceRefund(id)
				end if

		        '''계좌 암호화 추가.
	            Call EditCSMasterRefundEncInfo(id, encmethod, rebankaccount)
		    End if


	        ResultMsg = "등록되었습니다."
	        ReturnUrl = "/cscenter/action/pop_cs_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd

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

	elseif (divcd="A999") then
    	'----------------------------------------------------------------------
        'CS 접수 - 고객추가결제

        On Error Resume Next
	        dbget.beginTrans

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "001"

	            '' CS Master 접수
	            id = RegCSMaster(divcd, orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
	        end if

		    If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "003"

	            'CS Detail 접수(관련 상품정보)
		        Call AddCSDetailByArrStr(detailitemlist, id, orderserial)
	        end if

	        if (requiremakerid<>"") then
	            call RegCSMasterAddUpche(id, requiremakerid)
	        end if

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "005"

				'// 상품대금인 경우 정산금액 0원, skyer9, 2015-09-02
	            if (add_upchejungsandeliverypay<>"0") and (add_upchejungsandeliverypay<>"")  then
	                call RegCSUpcheAddJungsanPay(id, add_upchejungsandeliverypay, add_upchejungsancause, buf_requiremakerid)
	            end if
	        end If

			If (Err.Number = 0) and (ScanErr="") Then
				errcode = "006"

				if (requiremakerid<>"") then
					call RegCSMasterAddUpche(id, requiremakerid)
				end if
			end if

			If (Err.Number = 0) and (ScanErr="") Then
				errcode = "007"

				Call SetCustomerAddPay(id, "", add_customeradditempay, add_customeradditembuypay, add_customeraddbeasongpay, customerpayordertype, "N", 0)
			end if

			If (Err.Number = 0) and (ScanErr="") Then
				errcode = "008"
				'// 추가결제주문 생성

				payorderserial = AddPaymentOrder(id, orderserial, add_customeradditempay, add_customeraddbeasongpay, customerpayordertype, accountdiv, html2db(acctname), requiremakerid)
				ResultMsg = ResultMsg + "->. [추가결제주문 생성] 처리\n\n"

				Call SetPayOrderserial(id, payorderserial)
			end if

			If (Err.Number = 0) and (ScanErr="") Then
				errcode = "009"
				'// 가상계좌 발급
				'// 발급 안내문자 발송
				if CheckNAssignCyberAcct(id, payorderserial, acctno) = True then
					ResultMsg = ResultMsg + "->. [가상계좌 발급] 처리\n\n"
					ResultMsg = ResultMsg + "->. [가상계좌 안내 SMS 발송] 처리\n\n"
				end if
			end if

	        ''ResultMsg = "등록되었습니다."
	        ReturnUrl = "/cscenter/action/pop_cs_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd

	        If (Err.Number = 0) and (ScanErr="") Then
	            dbget.CommitTrans
	        Else
	            dbget.RollBackTrans
	            response.write "<script>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(에러코드 : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
	            'response.write "<script>history.back()</script>"
	            dbget.close()	:	response.End
	        End If
		on error Goto 0

    else
        ResultMsg = "정의되지 않았습니다[1]. : mode=" + mode + " , divcd=" + divcd
        response.write "<script>alert('" + ResultMsg + "');</script>"
        response.write "<script>history.back();</script>"
        dbget.close()	:	response.End
    end If

	If (id > 0) And needChkYN <> "" Then
		Call EditCSMasterAddInfo(id, Array( Array("needChkYN", needChkYN) ))
	End If


elseif (mode="deletecsas") then
	'==========================================================================
	'CS 삭제

	set ocsaslist = New CCSASList
	ocsaslist.FRectCsAsID = id
	ocsaslist.GetOneCSASMaster

	if (ocsaslist.FOneItem.Fdeleteyn = "Y") then
	    response.write "<script>alert(" + Chr(34) + "이미 삭제된 내역입니다." + Chr(34) + ")</script>"
	    response.write "이미 삭제된 내역입니다."
	    dbget.close()	:	response.End
	elseif (ocsaslist.FOneItem.Fcurrstate = "B007") then
		response.write "<script>alert(" + Chr(34) + "이미 완료된 내역입니다." + Chr(34) + ")</script>"
		response.write "이미 완료된 내역입니다."
		dbget.close()	:	response.End
	end if

    On Error Resume Next
        dbget.beginTrans

        ''Check Valid Delete - 현재는 B006 업체처리완료 , B007 완료 내역은 취소(삭제) 불가
        if (NOT ValidDeleteCS(id)) then
            response.write "<script>alert(" + Chr(34) + "현재 취소 가능 상태가 아닙니다. 관리자 문의 요망." + Chr(34) + ")</script>"
            response.write "<script>history.back()</script>"
            dbget.close()	:	response.End
        end if

		if (divcd = "A111") or (divcd = "A112") then
			'// 상품변경 교환회수
			Call GetChangeOrderInfo(id, changeorderserial, changeorderstate, ResultMsg)

			if (changeorderserial <> "") then
				if Not IsCancelChangeOrderValidState(changeorderserial) then
					ResultMsg = "삭제 가능 상태가 아닙니다. 먼저 교환주문[" + CStr(changeorderserial) + "] 관련 CS 를 삭제하세요."
		            response.write "<script>alert(" + Chr(34) + ResultMsg + Chr(34) + ")</script>"
		            response.write ResultMsg
		            ''response.write "<script>history.back()</script>"
		            dbget.close()	:	response.End
				end if

				Call setCancelMaster(id, changeorderserial)
			end if
			'
		end if

		if (divcd = "A003") then
			if GetDepositLogCountByAsid(orderserial, id) = 1 then
				Call AddDepositCancelLogByAsid(ocsaslist.FOneItem.Fuserid, orderserial, id)
			end if
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
        ReturnUrl = "/cscenter/action/pop_cs_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd + "&mode=editreginfo"

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
    ''접수 내역 수정

    On Error Resume Next
        dbget.beginTrans

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "001"

            '' CS Master 수정
            if (divcd = "A003") or (divcd = "A007")  then
            	title = GetCSRefundTitle(id, divcd, orderserial, returnmethod, title)
            end if

            Call EditCSMaster(id, reguserid, title, contents_jupsu, gubun01, gubun02)

            ''환불방식이 바뀐 경우.. 2011-07-20 서동석추가
            if (divcd="A007") and Not ((returnmethod="R020") or (returnmethod="R022") or (returnmethod="R080") or (returnmethod="R100") or (returnmethod="R550") or (returnmethod="R560") or (returnmethod="R120") or (returnmethod="R400") or (returnmethod="R420")) then
                sqlStr = " update [db_cs].[dbo].tbl_new_as_list"
                sqlStr = sqlStr + " set divcd='A003'"
                sqlStr = sqlStr + " where id=" + CStr(id)

                dbget.Execute sqlStr
            end if

            if (divcd="A003") and ((returnmethod="R020") or (returnmethod="R022") or (returnmethod="R080") or (returnmethod="R100") or (returnmethod="R550") or (returnmethod="R560") or (returnmethod="R120") or (returnmethod="R400") or (returnmethod="R420")) then
                sqlStr = " update [db_cs].[dbo].tbl_new_as_list"
                sqlStr = sqlStr + " set divcd='A007'"
                sqlStr = sqlStr + " where id=" + CStr(id)

                dbget.Execute sqlStr
            end if
        end if

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "002"

			'이전 내역 삭제후 전체 재입력
			''Call DeleteAllCSDetail(id, orderserial)

			if (divcd="A100") then
				'// 추가되는 상품목록 등록
				Call ModiCSDetailWithoutOrderDetailByArrStr(csdetailitemlist, id, orderserial)
			else
	            'CS Detail 접수(관련 상품정보)
		        Call ModiCSDetailByArrStr(detailitemlist, id, orderserial)
			end if

        End if

        ResultMsg = ResultMsg + "->. [CS 처리건 수정] 처리\n\n"

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "003"

            '' 환불 정보 수정
            if (CheckNEditRefundInfo(id,returnmethod,rebankaccount,rebankownername,rebankname,paygateTid,refundrequire, orgsubtotalprice, orgitemcostsum, orgbeasongpay , orgmileagesum, orgcouponsum, orgallatdiscountsum , canceltotal, refunditemcostsum, refundmileagesum, refundcouponsum, allatsubtractsum , refundbeasongpay, refunddeliverypay, refundadjustpay)) then
            	Call AddCSMasterRefundInfo(id, orggiftcardsum, orgdepositsum, refundgiftcardsum, refunddepositsum)

            	'''계좌 암호화 추가.
	            Call EditCSMasterRefundEncInfo(id, encmethod, rebankaccount)
                ResultMsg = ResultMsg + "->. [환불정보 수정] 처리\n\n"

				''쿠폰 재발행할지
				Call EditCSCopyCouponInfo(id, copycouponinfo)
            end if
        end If

        ''업체 추가 정산비 2008.11.10
        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "004"

            if (add_upchejungsandeliverypay<>"") then
                call EditCSUpcheAddJungsanPay(id, add_upchejungsandeliverypay, add_upchejungsancause, buf_requiremakerid)
            end if
        end if

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "005"

            '' 고객 추가배송비
            if (divcd = "A100") or (divcd = "A111")  then
            	Call SetCustomerAddBeasongPay(id, add_customeraddmethod, add_customeraddbeasongpay, "Y", customerrealbeasongpay)
			elseif (divcd = "A004") or (divcd = "A010") then
				Call SetCustomerAddBeasongPay(id, addmethod, addbeasongpay, "Y", 0)
            end if

        end if

        ReturnUrl = "/cscenter/action/pop_cs_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd + "&mode=editreginfo"

        If (Err.Number = 0) and (ScanErr="") Then
            dbget.CommitTrans
        Else
            dbget.RollBackTrans
            response.write "<script>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(에러코드 : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
            'response.write "<script>history.back()</script>"
            dbget.close()	:	response.End
        End If
    On error Goto 0

	If (id > 0) And needChkYN <> "" Then
		Call EditCSMasterAddInfo(id, Array( Array("needChkYN", needChkYN) ))
	End If


elseif (mode="finishededitcsas") then
	'==========================================================================
    ''완료된 내역 수정
    On Error Resume Next
        dbget.beginTrans

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "001"

            '' 마스타 수정
            Call EditCSMasterFinished(id, title, contents_jupsu, gubun01, gubun02, reguserid, contents_finish)
        end if

        If (Err.Number = 0) and (ScanErr="") Then
            errcode = "002"

            '' 디테일 수정
            Call EditCSDetailByArrStr(detailitemlist, id, orderserial)
        End if

        ResultMsg = "수정 되었습니다."
        ReturnUrl = "/cscenter/action/pop_cs_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd + "&mode=editreginfo"

        If (Err.Number = 0) and (ScanErr="") Then
            dbget.CommitTrans
        Else
            dbget.RollBackTrans
            response.write "<script>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(에러코드 : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
            'response.write "<script>history.back()</script>"
            dbget.close()	:	response.End
        End If
    On error Goto 0

elseif (mode="delfinishedcsas") then
	'==========================================================================
    ''완료된 내역 삭제

	ScanErr = ""

	set ocsaslist = New CCSASList
	ocsaslist.FRectCsAsID = id
	ocsaslist.GetOneCSASMaster

	set orefund = New CCSASList
	orefund.FRectCsAsID = ocsaslist.FOneItem.FId
	orefund.GetOneRefundInfo

	if (ocsaslist.FOneItem.Fdeleteyn = "Y") then
	    response.write "<script>alert(" + Chr(34) + "이미 삭제된 내역입니다." + Chr(34) + ")</script>"
	    response.write "이미 삭제된 내역입니다."
	    dbget.close()	:	response.End
	end if

	if (divcd="A008") then
		'// 취소완료CS 삭제

		On Error Resume Next
			dbget.beginTrans

			If (Err.Number = 0) and (ScanErr="") Then
				errcode = "001"

				set oRefCSASList = new CCSASList
				oRefCSASList.FRectCsRefAsID = id
				oRefCSASList.GetOneCSASMaster

				if (oRefCSASList.FResultCount > 0) then
					if (oRefCSASList.FOneItem.Fdeleteyn = "N") then
						ScanErr = "먼저 관련 환불CS 를 삭제하세요."
					end if
				end if
			end if

			If (Err.Number = 0) and (ScanErr="") Then
				errcode = "002"

				'// 0. 추가배송비 있는 경우, 마지막 취소건부터 모두 복구해야 함
				if (id <> 8760486) and (CheckRestoreCancelOKByAsid(id) <> True) then
					ScanErr = "마지막 취소건부터 순서대로 모두 복구해야 합니다."
				end if
			end if

			If (Err.Number = 0) and (ScanErr="") Then
				errcode = "002"

				'// 0. 정산건 있으면 삭제불가
				if (CheckJungsanExistsByAsid(id) = True) then
					ScanErr = "복구불가 - 정산내역이 있습니다."
				end if
			end if

			If (Err.Number = 0) and (ScanErr="") Then
				errcode = "003"

				If Not DeleteFinishedCSProcess(id) then
					ScanErr = "데이터 삭제시 오류"
				else
					Call AddCsMemo(ocsaslist.FOneItem.Forderserial, "1", ocsaslist.FOneItem.Fuserid, session("ssBctId"), "취소완료CS 삭제")
					ResultMsg = ResultMsg + "->. [CS처리완료건 삭제] 처리\n\n"
				End if
			end if

			If (Err.Number = 0) and (ScanErr="") Then
				errcode = "004"

				if Del_logicschulgodata(id, ocsaslist.FOneItem.Forderserial) > 0 then
					Call AddCsMemo(ocsaslist.FOneItem.Forderserial, "1", ocsaslist.FOneItem.Fuserid, session("ssBctId"), "로직스 출고 데이터 삭제")
					ResultMsg = ResultMsg + "->. [로직스 출고 데이터 삭제] 처리\n\n"
                end if
			end if

			If (Err.Number = 0) and (ScanErr="") Then
				errcode = "005"

				If Not RestoreCancelProcess(id, orderserial) then
					ScanErr = "취소주문 복구중 오류"
				else
					Call AddCsMemo(ocsaslist.FOneItem.Forderserial, "1", ocsaslist.FOneItem.Fuserid, session("ssBctId"), "취소주문 복구")
					ResultMsg = ResultMsg + "->. [취소주문 복구] 처리\n\n"
				End if
			end if

			ResultMsg = "복구 되었습니다."
			ReturnUrl = "/cscenter/action/pop_cs_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd + "&mode=editreginfo"

			If (Err.Number = 0) and (ScanErr="") Then
				dbget.CommitTrans
			Else
				dbget.RollBackTrans
				response.write "<script>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(에러코드 : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
				'response.write "<script>history.back()</script>"
				dbget.close()	:	response.End
			End If
		On error Goto 0

	elseif (divcd="A004") or (divcd="A010") then
		'// 반품완료CS 삭제

		On Error Resume Next
			dbget.beginTrans

			refminusorderserial = ""

			If (Err.Number = 0) and (ScanErr="") Then
				errcode = "001"

				set oRefCSASList = new CCSASList
				oRefCSASList.FRectCsRefAsID = id
				oRefCSASList.GetOneCSASMaster

				if (oRefCSASList.FResultCount > 0) then
					refminusorderserial = oRefCSASList.FOneItem.Frefminusorderserial

					if (oRefCSASList.FOneItem.Fdeleteyn = "N") then
						ScanErr = "먼저 관련 환불CS 를 삭제하세요."
					end if
				else
					refminusorderserial = ocsaslist.FOneItem.Frefminusorderserial
				end if
			end if

			If (Err.Number = 0) and (ScanErr="") Then
				errcode = "002"

				if (refminusorderserial = "") then ScanErr = "마이너스 주문번호 없음 - 관리자 문의 요망"
			end if

			If (Err.Number = 0) and (ScanErr="") Then
				errcode = "003"

				'// 0. 정산건 있으면 삭제불가
				if (CheckJungsanExists(refminusorderserial) = True) then
					ScanErr = "삭제불가 - 정산내역이 있습니다."
				end if
			end if

			If (Err.Number = 0) and (ScanErr="") Then
				errcode = "004"

				If Not DeleteFinishedCSProcess(id) then
					ScanErr = "데이터 삭제시 오류"
				else
					Call AddCsMemo(ocsaslist.FOneItem.Forderserial, "1", ocsaslist.FOneItem.Fuserid, session("ssBctId"), "반품완료CS 삭제")
					ResultMsg = ResultMsg + "->. [CS처리완료건 삭제] 처리\n\n"
				End if
			end if

			If (Err.Number = 0) and (ScanErr="") Then
				errcode = "005"

				if (Left(ocsaslist.FOneItem.Ffinishdate,7) < Left(Now(),7)) and ChkMaeipItemExist(id) and DateSerial(Year(Now()), Month(Now()), 4) < Now() then
					'// 매입상품 있고, 완료일자가 이전달이면
					ScanErr = "매입상품이고 반품일자가 이전달이면 반품취소불가"
				Else
					If Not CancelMinusOrderProcess(refminusorderserial) then
						ScanErr = "데이터 삭제시 오류"
					else
						Call AddCsMemo(ocsaslist.FOneItem.Forderserial, "1", ocsaslist.FOneItem.Fuserid, session("ssBctId"), "마이너스 주문 삭제")
						ResultMsg = ResultMsg + "->. [마이너스 주문 삭제] 처리\n\n"
					End if
				End If
			end if

			ResultMsg = "CS내역삭제 및 마이너스주문삭제 되었습니다."
			ReturnUrl = "/cscenter/action/pop_cs_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd + "&mode=editreginfo"

			If (Err.Number = 0) and (ScanErr="") Then
				dbget.CommitTrans
			Else
				dbget.RollBackTrans
				response.write "<script>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(에러코드 : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
				'response.write "<script>history.back()</script>"
				dbget.close()	:	response.End
			End If
		On error Goto 0

	elseif (divcd="A011") or (divcd="A012") then
		'// 맞교환회수(텐바이텐배송), 맞교환회수(업체배송)	' 2019.10.17 한용민 추가

		On Error Resume Next
			dbget.beginTrans

			If (Err.Number = 0) and (ScanErr="") Then
				errcode = "001"

				'// 0. 정산건 있으면 삭제불가
				if (CheckJungsanExistsByAsid(id) = True) then
					ScanErr = "복구불가 - 정산내역이 있습니다."
				end if
			end if

			'//텐배 교환회수 한정처리
			if divcd="A011" then
		        If (Err.Number = 0) and (ScanErr="") Then
		            errcode = "002"

					'/한정 내리고, 판매 상태 변경한다.	'/2016.07.15 한용민 생성
					call setItemLimitcs(id, orderserial, "DOWN")

					ResultMsg = ResultMsg + "->. [한정 처리] 텐배 교환회수 한정처리 복구 완료\n\n"
		        End If
			End If

			If (Err.Number = 0) and (ScanErr="") Then
				errcode = "003"

				'' 마스타 삭제
				Call DeleteFinishedCSForce(id)

				Call AddCsMemo(ocsaslist.FOneItem.Forderserial, "1", ocsaslist.FOneItem.Fuserid, session("ssBctId"), "맞교환회수 완료CS 삭제")
			end if

			ResultMsg = "맞교환회수 삭제 처리 되었습니다."
			ReturnUrl = "/cscenter/action/pop_cs_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd + "&mode=editreginfo"

			If (Err.Number = 0) and (ScanErr="") Then
				dbget.CommitTrans
			Else
				dbget.RollBackTrans
				response.write "<script>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(에러코드 : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
				'response.write "<script>history.back()</script>"
				dbget.close()	:	response.End
			End If
		On error Goto 0

	elseif (divcd="A111") or (divcd="A112") then
		'// 상품변경 맞교환회수(텐바이텐배송), 상품변경 맞교환회수(업체배송)	' 2019.10.18 한용민 추가

		On Error Resume Next
			dbget.beginTrans

			refchangeorderserial = ""

			If (Err.Number = 0) and (ScanErr="") Then
				errcode = "001"

				set oRefCSASList = new CCSASList
				oRefCSASList.FRectCsAsID = id
				'oRefCSASList.FRectCsRefAsID = id
				oRefCSASList.GetOneCSASMaster

				if (oRefCSASList.FResultCount > 0) then
					refchangeorderserial = oRefCSASList.FOneItem.Frefchangeorderserial

					' if (oRefCSASList.FOneItem.Fdeleteyn = "N") then
					' 	ScanErr = "먼저 관련 상품변경 맞교환 교환출고 를 삭제하세요."
					' end if
				else
					refchangeorderserial = ocsaslist.FOneItem.Frefchangeorderserial
				end if
			end if

			If (Err.Number = 0) and (ScanErr="") Then
				errcode = "002"

				if (refchangeorderserial = "") then ScanErr = "교환 주문번호 없음 - 관리자 문의 요망"
			end if

			If (Err.Number = 0) and (ScanErr="") Then
				errcode = "003"

				'// 0. 정산건 있으면 삭제불가
				if (CheckJungsanExists(refchangeorderserial) = True) then
					ScanErr = "복구불가 - 정산내역이 있습니다."
				end if
			end if

			If (Err.Number = 0) and (ScanErr="") Then
				errcode = "005"

				if (Left(ocsaslist.FOneItem.Ffinishdate,7) < Left(Now(),7)) and ChkMaeipItemExist(id) then
					'// 매입상품 있고, 완료일자가 이전달이면
					ScanErr = "매입상품이고 맞교환회수 일자가 이전달이면 교환불가"
				Else
					If Not CancelChangeOrderProcess(refchangeorderserial) then
						ScanErr = "데이터 삭제시 오류"
					else
						Call AddCsMemo(ocsaslist.FOneItem.Forderserial, "1", ocsaslist.FOneItem.Fuserid, session("ssBctId"), "교환 주문 삭제")
						ResultMsg = ResultMsg + "->. [교환 주문 삭제] 처리\n\n"
					End if
				End If
			end if

			'//텐배 교환회수 한정처리
			if divcd="A111" then
		        If (Err.Number = 0) and (ScanErr="") Then
		            errcode = "002"

					'/한정 내리고, 판매 상태 변경한다.	'/2016.07.15 한용민 생성
					call setItemLimitcs(id, orderserial, "DOWN")

					ResultMsg = ResultMsg + "->. [한정 처리] 텐배 교환회수 한정처리 복구 완료\n\n"
		        End If
			End If

			If (Err.Number = 0) and (ScanErr="") Then
				errcode = "001"

				'' 마스타 삭제
				Call DeleteFinishedCSForce(id)

				Call AddCsMemo(ocsaslist.FOneItem.Forderserial, "1", ocsaslist.FOneItem.Fuserid, session("ssBctId"), "상품변경 맞교환회수 완료CS 삭제")
			end if

			ResultMsg = "CS내역삭제 및 교환주문삭제 되었습니다."
			ReturnUrl = "/cscenter/action/pop_cs_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd + "&mode=editreginfo"

			If (Err.Number = 0) and (ScanErr="") Then
				dbget.CommitTrans
			Else
				dbget.RollBackTrans
				response.write "<script>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(에러코드 : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
				'response.write "<script>history.back()</script>"
				dbget.close()	:	response.End
			End If
		On error Goto 0

	elseif (divcd="A000") then
		'// 상품변경 교환출고(텐바이텐배송), 상품변경 교환출고(업체배송)	' 2019.10.18 한용민 추가

		On Error Resume Next
			dbget.beginTrans

			' 출고되는 상품 한정차감 원상복구 로직필요.

			If (Err.Number = 0) and (ScanErr="") Then
				errcode = "001"

				'// 0. 정산건 있으면 삭제불가
				if (CheckJungsanExistsByAsid(id) = True) then
					ScanErr = "복구불가 - 정산내역이 있습니다."
				end if
			end if

			If (Err.Number = 0) and (ScanErr="") Then
				errcode = "002"

				'' 마스타 삭제
				Call DeleteFinishedCSForce(id)

				Call AddCsMemo(ocsaslist.FOneItem.Forderserial, "1", ocsaslist.FOneItem.Fuserid, session("ssBctId"), "완료CS 삭제")
			end if

			ResultMsg = ResultMsg + "-> 교환출고 삭제 완료되었습니다."
			ReturnUrl = "/cscenter/action/pop_cs_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd + "&mode=editreginfo"

			If (Err.Number = 0) and (ScanErr="") Then
				dbget.CommitTrans
			Else
				dbget.RollBackTrans
				response.write "<script>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(에러코드 : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
				'response.write "<script>history.back()</script>"
				dbget.close()	:	response.End
			End If
		On error Goto 0

	elseif (divcd="A003" or divcd="A007") then
		'// 환불, (카드,이체,휴대폰취소요청)	' 2020.11.11 한용민 추가

		On Error Resume Next
			dbget.beginTrans

			if divcd="A007" and (orefund.FOneItem.Freturnmethod = "R120") then
                '// 신용카드/이체취소요청 완료건 삭제
		        If (Err.Number = 0) and (ScanErr="") Then
		            errcode = "001"

					call setRestoreEtcRealPayment(id, orderserial)

					ResultMsg = ResultMsg + "->. 잔여 결제액 복구 완료\n\n"
		        End If
			End If

			If (Err.Number = 0) and (ScanErr="") Then
				errcode = "002"

				'' 마스타 삭제
				Call DeleteFinishedCSForce(id)

				Call AddCsMemo(ocsaslist.FOneItem.Forderserial, "1", ocsaslist.FOneItem.Fuserid, session("ssBctId"), "완료CS 삭제")
			end if

			ResultMsg = ResultMsg + "-> 삭제 완료."
			ReturnUrl = "/cscenter/action/pop_cs_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd + "&mode=editreginfo"

			If (Err.Number = 0) and (ScanErr="") Then
				dbget.CommitTrans
			Else
				dbget.RollBackTrans
				response.write "<script>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(에러코드 : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
				'response.write "<script>history.back()</script>"
				dbget.close()	:	response.End
			End If
		On error Goto 0

	elseif C_ADMIN_AUTH or (((divcd = "A700") or (divcd = "A005") or (divcd = "A100")) and (Left(ocsaslist.FOneItem.Fregdate,7) = Left(Now(),7))) then

		'// 또는 업체기타정산 당월완료건
		'// 또는 입점몰환불 당월완료건
		' 또는 교환출고(상품변경.텐배)
		On Error Resume Next
			dbget.beginTrans

			if divcd="A007" and (orefund.FOneItem.Freturnmethod = "R120") then
                '// 신용카드/이체취소요청 완료건 삭제
		        If (Err.Number = 0) and (ScanErr="") Then
		            errcode = "001"

					call setRestoreEtcRealPayment(id, orderserial)

					ResultMsg = ResultMsg + "->. 잔여 결제액 복구 완료\n\n"
		        End If
			End If

			If (Err.Number = 0) and (ScanErr="") Then
				errcode = "002"

				'' 마스타 삭제
				Call DeleteFinishedCSForce(id)

				Call AddCsMemo(ocsaslist.FOneItem.Forderserial, "1", ocsaslist.FOneItem.Fuserid, session("ssBctId"), "완료CS 삭제")
			end if

			ResultMsg = ResultMsg + "-> 삭제 완료."
			ReturnUrl = "/cscenter/action/pop_cs_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd + "&mode=editreginfo"

			If (Err.Number = 0) and (ScanErr="") Then
				dbget.CommitTrans
			Else
				dbget.RollBackTrans
				response.write "<script>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(에러코드 : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
				'response.write "<script>history.back()</script>"
				dbget.close()	:	response.End
			End If
		On error Goto 0

	else
        ResultMsg = "정의되지 않았습니다[2]. : mode=" + mode + " , divcd=" + divcd
        response.write "<script>alert('" + ResultMsg + "');</script>"
        ''response.write "<script>history.back();</script>"
        dbget.close()	:	response.End
	end if


elseif (mode="realdelcsas") then
    '==========================================================================
    '데이타베이스에서 내역 DELETE

	set ocsaslist = New CCSASList
	ocsaslist.FRectCsAsID = id
	ocsaslist.GetOneCSASMaster

    sqlStr = " delete from [db_cs].[dbo].[tbl_as_refund_info] where asid = " & id
    dbget.Execute sqlStr

    sqlStr = " delete from [db_cs].[dbo].[tbl_new_as_detail] where masterid = " & id
    dbget.Execute sqlStr

    sqlStr = " delete from [db_cs].[dbo].[tbl_new_as_list] where id = " & id
    dbget.Execute sqlStr

    Call AddCsMemo(ocsaslist.FOneItem.Forderserial, "1", ocsaslist.FOneItem.Fuserid, session("ssBctId"), "CS 내역 DELETE 완료 ASID=" & id)

    response.write "삭제완료"
    dbget.close()	:	response.End

elseif (mode="finishcsas") then
    'CS 접수 내역 완료처리

	set ocsaslist = New CCSASList
	ocsaslist.FRectCsAsID = id
	ocsaslist.GetOneCSASMaster

	if (ocsaslist.FOneItem.Fdeleteyn = "Y") then
		if (IsAutoScript) then
			response.write "S_ERR|이미 삭제된 내역입니다."
		else
			response.write "<script>alert(" + Chr(34) + "이미 삭제된 내역입니다." + Chr(34) + ")</script>"
			response.write "이미 삭제된 내역입니다."
		end if
	    dbget.close()	:	response.End
	elseif (ocsaslist.FOneItem.Fcurrstate = "B007") then
		if (IsAutoScript) then
			response.write "S_ERR|이미 완료된 내역입니다."
		else
			response.write "<script>alert(" + Chr(34) + "이미 완료된 내역입니다." + Chr(34) + ")</script>"
			response.write "이미 완료된 내역입니다."
		end if
		dbget.close()	:	response.End
	end if

    if (divcd="A008") then

		'----------------------------------------------------------------------
		'CS 접수 내역 완료처리 - 주문취소
        On Error Resume Next
	    	dbget.beginTrans
			'// 취소허용(2014-03-31)
	        ''if (GC_IsOLDOrder) then ScanErr = "6개월 이전 과거 내역 취소 불가 - 관리자 문의 요망"

	        If (modeflag2<>"norefund") and (Err.Number = 0) and (ScanErr="") Then
	            errcode = "000"

                ''마일리지 환원 체크
				Call CheckRefundMileage(id, orderserial)

	            ''금액 체크
				Call CheckRefundPrice(id, orderserial, ScanErr)
	        End If

			'출고완료 또는 취소된 상품이 있을 경우, 진행정지(주문취소 불가)
			'출고완료된 상품은 반품만 가능하다.
			If (Err.Number = 0) and (ScanErr="") Then
				errcode = "000"

				''출고 완료 또는 취소된 내역이 있는지 확인
				if Not (IsCancelValidState(id, orderserial)) then
					ScanErr = "취소 검증 오류. - 출고된 내역이 있거나 취소된 내역이 있습니다."
				end if
			end if

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "001"

	            Call CancelProcess(id, orderserial)

				IsAllCancel = False
				if (remainitemcostsum = 0) then
					'전체취소 : 잔여상품총액이 없는 경우
            		IsAllCancel = True
				end if

				'// 2018-01-12, skyer9
				if (IsAllCancel) then
					''전체 취소인경우
					sqlStr = " exec [db_summary].[dbo].sp_Ten_RealtimeStock_cancelOrderAll '" & orderserial & "'"
					dbget.Execute sqlStr
				end if

				'// 여행상품
				if (oordermaster.FOneItem.Fjumundiv = "3") then
					Call AddCsMemoRequest(ocsaslist.FOneItem.Forderserial, ocsaslist.FOneItem.Fuserid, "59", session("ssBctId"), "여행상품 발권전 취소")
				end if
	        End IF

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "008"

	            '환불 정보가 있는지 체크 후 무통장환불/마일리지환불/신용카드취소 CS 접수 등록
	            newasid = CheckNRegRefund(id, orderserial, reguserid)
	            if (newasid>0) then
	                ResultMsg = ResultMsg + "->. [환불 등록] 처리\n\n"
	            end if
	        End If

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "009"

	            Call FinishCSMaster(id, reguserid, contents_finish)

				'// 취소 접수상태 내역의 금액 업데이트
				Call UpdateCancelJupsuCSPrice(id, orderserial)

				Call CheckAndCopyBonusCoupon(id, reguserid)
	        End If

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "010"

				' 상품쿠폰재발급
				call CheckAndCopyItemCoupon(id, reguserid, "P", oordermaster.FOneItem.Fuserid)
            End If

	        ReturnUrl = "/cscenter/action/pop_cs_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd + "&mode=editfinishedinfo"

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
    	'----------------------------------------------------------------------
        'CS 접수 내역 완료처리 - 환불요청 / 카드,이체,휴대폰취소요청
        dim RefreturnMethod, Refrealrefund

		set orefund = New CCSASList
			orefund.FRectCsAsID = id
			orefund.GetOneRefundInfo

		if (IsAutoScript) then
			if (divcd <> "A003") then
				response.write "S_ERR|작업되지 않았습니다. : mode=" + mode + " , divcd=" + divcd
				dbget.close()	:	response.End
			end if

			if (orefund.FOneItem.Freturnmethod <> "R910") and (orefund.FOneItem.Freturnmethod <> "R900") then
				response.write "S_ERR|마일리지 또는 예치금환불만 가능합니다. : mode=" + mode + " , divcd=" + divcd
				dbget.close()	:	response.End
			end if
		end if

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
	            Call CheckRefundFinish(id, orderserial, RefreturnMethod, Refrealrefund)
	        End If

	        ResultMsg = "처리완료\n\n"
	        if (RefreturnMethod="R007") and (Refrealrefund>0) then
	            ReturnUrl = "/cscenter/action/pop_cs_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd + "&finishtype=1"
	        else
	            ReturnUrl = "/cscenter/action/pop_cs_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd
	        end if

	        If (Err.Number = 0) and (ScanErr="") Then
	            dbget.CommitTrans
	        Else
	            dbget.RollBackTrans
	            response.write "<script>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(에러코드 : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
	            response.write "<script>history.back()</script>"
				Call SetNeedCheckToY(id)
	            dbget.close()	:	response.End
	        End If

	        ''환불 완료 메일
	        If (isCsMailSend) then
	            Call SendCsActionMail(id)

				refasid=0
				sqlStr ="select isnull(a.refasid,0) as refasid, max(replace(replace(replace(replace(replace(ad.itemname,char(9),''),char(10),''),char(13),''),'""',''),'''','')) as itemname, count(ad.masterid) as itemcnt"
				sqlStr = sqlStr & " from db_cs.dbo.tbl_new_as_list a with (nolock)"
				sqlStr = sqlStr & " join db_cs.dbo.tbl_new_as_list aa with (nolock)"
				sqlStr = sqlStr & " 	on a.orderserial = aa.orderserial"
				sqlStr = sqlStr & " 	and a.refasid = aa.id"
				sqlStr = sqlStr & " 	and a.deleteyn='N' and aa.deleteyn='N'"
				sqlStr = sqlStr & " join db_cs.dbo.tbl_new_as_detail ad with (nolock)"
				sqlStr = sqlStr & " 	on aa.id=ad.masterid"
				sqlStr = sqlStr & " 	and ad.itemid not in (0)"
				sqlStr = sqlStr & " where a.orderserial = '" & orderserial & "'"		' 주문번호
				sqlStr = sqlStr & " and a.id="& id &""	' 주문취소
				sqlStr = sqlStr & " group by isnull(a.refasid,0)"

				'response.write sqlStr & "<Br>"
				rsget.CursorLocation = adUseClient
				rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

				IF not rsget.EOF THEN
					itemCnt = rsget("itemcnt")
					refasid = rsget("refasid")

					if itemName = "" then
						itemName = replace(db2html(rsget("itemname")),vbcrlf,"")
					end if
				END IF
				rsget.close

				' 반품 cs건이 있는 경우에만 생성
				if refasid<>0 and refasid<>"" then
					if itemCnt > 1 then
						itemName = itemName & " 외 " & (itemCnt - 1) & "종"
					end if

					refundresult=0
					refunddepositsum=0
					refundmileagesum=0
					refundgiftcardsum=0
					sqlStr ="select refundresult, refunddepositsum, refundmileagesum, refundgiftcardsum"
					sqlStr = sqlStr & " from db_cs.dbo.tbl_as_refund_info r with (nolock)"
					sqlStr = sqlStr & " where asid="& id &""

					'response.write sqlStr & "<Br>"
					rsget.CursorLocation = adUseClient
					rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

					IF not rsget.EOF THEN
						refundresult = rsget("refundresult")
						refunddepositsum = rsget("refunddepositsum")
						refundmileagesum = rsget("refundmileagesum")
						refundgiftcardsum = rsget("refundgiftcardsum")
					END IF
					rsget.close
					if refunddepositsum=0 and refundmileagesum=0 and refundgiftcardsum=0 then
						refundstr=FormatNumber(refundresult,0) & "원"
					else
						refundstr=FormatNumber(refundresult,0) & "원(예치금환급 "& refunddepositsum &"원 / 마일리지환급 "& refundmileagesum &"pt / 기프트환급 "& refundgiftcardsum &"원)"
					end if

					sqlStr ="select replace(isnull(m.buyhp,''),'-','') as buyhp"
					sqlStr = sqlStr & " from db_order.dbo.tbl_order_master m with (nolock)"
					sqlStr = sqlStr & " where m.orderserial="& orderserial &""

					'response.write sqlStr & "<Br>"
					rsget.CursorLocation = adUseClient
					rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

					IF not rsget.EOF THEN
						buyhp = rsget("buyhp")
					END IF
					rsget.close

					if buyhp<>"" then
						' 반품환불		' 2021.10.05 한용민 생성
						fullText = "[10x10] 반품환불안내" & vbCrLf & vbCrLf
						fullText = fullText & "고객님" & vbCrLf
						fullText = fullText & "반품 신청하신 상품의 환불(결제취소) 처리되어 안내 드립니다." & vbCrLf & vbCrLf
						fullText = fullText & "■ 주문번호 : "& orderserial &"" & vbCrLf
						fullText = fullText & "■ 상품명 : "& itemName &"" & vbCrLf
						fullText = fullText & "■ 환불금액 : "& refundstr & vbCrLf & vbCrLf
						fullText = fullText & "※ 카드 취소 3-7일, 무통장 1-2일 (영업일기준)" & vbCrLf
						fullText = fullText & "적립금, 마일리지, 기프트카드는 유효기간내 사용가능."
						failText = "[텐바이텐]반품 신청하신 상품의 환불 처리되어 안내 드립니다.주문번호:" & orderserial
						btnJson = "{""button"":[{""name"":""반품신청내역 바로가기"",""type"":""WL"", ""url_mobile"":""https://tenten.app.link/LIJjGiqVjjb""}]}"
						call SendKakaoCSMsg_LINK("", buyhp,"1644-6030","KC-0020",fullText,"SMS","",failText,btnJson,"","")
					end if
	        	end if
			End IF
        On error Goto 0

    elseif (divcd="A004") or (divcd="A010") then
		'----------------------------------------------------------------------
        'CS 접수 내역 완료처리 - 반품접수(업체배송)  // 회수신청(텐바이텐배송)
        dim MinusOrderserial

        On Error Resume Next
	        dbget.beginTrans

	        If (modeflag2<>"norefund") and (Err.Number = 0) and (ScanErr="") Then
	            errcode = "000"

                if (divcd = "A004") or (divcd = "A010") then
                    '// 제휴몰 쿠폰가 다른경우 업데이트
				    sqlStr = " exec [db_cs].[dbo].[usp_Ten_CsAs_ChangeCouponPrice2ReturnCs] " & id
				    dbget.Execute sqlStr
                end if

	            ''금액 체크
                if (orderserial <> "20110402336") then
				    Call CheckRefundPrice(id, orderserial, ScanErr)
                end if
	        End If

	        If (modeflag2<>"norefund") and (Err.Number = 0) and (ScanErr="") Then
	            errcode = "001"

	            ''마이너스 주문 등록
	            if (CheckNAddMinusOrder(id, orderserial, reguserid, MinusOrderserial, ScanErr)) then
					'// 마이너스 주문링크 넣는 함수를 체크함수 안으로 이동, skyer9, 2015-06-24
					'' Call AddminusOrderLink(id, MinusOrderserial)
	                ResultMsg = ResultMsg + "->. [반품 주문] 등록\n\n"
	            end if
	        End If

	        If (modeflag2<>"norefund") and (Err.Number = 0) and (ScanErr="") Then
	            errcode = "002"

	            '환불 정보가 있는지 체크 후 무통장환불/마일리지환불/신용카드취소 CS 접수 등록
	            'newasid = CheckNRegRefund(id, MinusOrderserial, reguserid)

	            '원주문에 대해 CS등록한다. 마이너스 주문에는 CS를 등록할 수 없다.
                if (orderserial <> "20110402336") then
	                newasid = CheckNRegRefund(id, orderserial, reguserid)
	                call AddminusOrderLink(newasid,MinusOrderserial)
                end if

	            if (newasid>0) then
	                ResultMsg = ResultMsg + "->. [취소(환불)접수] 처리\n\n"
	            end if
	        End If

            '//텐배 반품회수신청 한정처리
            '/마이너스/환불요청 있는거
            if modeflag2<>"norefund" and divcd="A010" then
		        If (Err.Number = 0) and (ScanErr="") Then
		            errcode = "011"

					'/한정 올리고, 판매 상태 변경한다.	'/2016.07.15 한용민 생성
					''call setItemLimitcs(id, orderserial, "UP")

					''ResultMsg = ResultMsg + "->. [한정 처리] 텐배 반품회수 한정처리 완료\n\n"
                    ResultMsg = ResultMsg + "->. [한정 처리 중지] 물류 요청으로 한정처리 일시중지\n\n"
		        End If
			End If

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "009"

	            Call FinishCSMaster(id, reguserid, contents_finish)

	            if (divcd="A004") then
	                ResultMsg = ResultMsg + "->. 반품 처리 완료\n\n"
	            elseif (divcd="A010") then
	                ResultMsg = ResultMsg + "->. 회수 처리 완료\n\n"
	            end if

				if CheckAndCopyBonusCoupon(id, reguserid) = True then
					ResultMsg = ResultMsg + "->. [쿠폰 재발급] 완료 처리\n\n"
				end if
	        End If

	        ResultMsg = ResultMsg
	        ReturnUrl = "/cscenter/action/pop_cs_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd
	        If (Err.Number = 0) and (ScanErr="") Then
	            dbget.CommitTrans
	        Else
	            dbget.RollBackTrans
				if (IsAutoScript) then
					response.write "S_ERR|데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(에러코드 : " + CStr(errcode) + ":" + Err.Description + "-" + ScanErr + ")"
					Call SetNeedCheckToY(id)
				else
					response.write "<script>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(에러코드 : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
					response.write "<script>history.back()</script>"
				end if
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
    	'CS 접수 내역 완료처리 - 맞교환회수(텐바이텐배송), 맞교환회수(업체배송)

        On Error Resume Next
	        dbget.beginTrans

			'//텐배 교환회수 한정처리
			if divcd="A011" then
		        If (Err.Number = 0) and (ScanErr="") Then
		            errcode = "011"

					'/한정 올리고, 판매 상태 변경한다.	'/2016.07.15 한용민 생성
					call setItemLimitcs(id, orderserial, "UP")

					ResultMsg = ResultMsg + "->. [한정 처리] 텐배 교환회수 한정처리 완료\n\n"
		        End If
			End If

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "009"

	            Call FinishCSMaster(id, reguserid, contents_finish)
	        End If

	        ResultMsg = ResultMsg + "처리 완료"
	        ReturnUrl = "/cscenter/action/pop_cs_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd

	        If (Err.Number = 0) and (ScanErr="") Then
				dbget.CommitTrans
	        Else
	            dbget.RollBackTrans
	            response.write "<script>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(에러코드 : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
				' response.write "<script>history.back()</script>"
				Call SetNeedCheckToY(id)
	            dbget.close()	:	response.End
	        End If

	        ''맞교환 완료 메일
	        If (isCsMailSend) and (divcd <> "A012") then
	        	'// 업체배송은 보내지 않는다.
   	            Call SendCsActionMail(id)
	        End If
        On error Goto 0
    elseif (divcd="A111") or (divcd="A112") then
    	'----------------------------------------------------------------------
    	'CS 접수 내역 완료처리 - 상품변경 맞교환회수(텐바이텐배송), 상품변경 맞교환회수(업체배송)

        On Error Resume Next
	        dbget.beginTrans

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "009"

				'// 함수 안에서 CS재고보정 한다.
	            Call FinishCSMaster(id, reguserid, contents_finish)
	        End If

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "010"

				Call GetChangeOrderInfo(id, changeorderserial, changeorderstate,  ScanErr)

				if (ScanErr = "") then
					if (changeorderserial = "") then
						'// 교환주문 등록한다.
						'// 맞교환출고 상태와 무관하게 맞교환회수가 되면 교환주문 등록한다.
						'// 텐배의 경우 변심 맞교환의 경우 회수이후에 맞교환출고한다(http://logics.10x10.co.kr/v2/online/m_re_chulgo.asp 참고)
						changeorderserial = CheckAndAddChangeOrder(id, orderserial, ScanErr)

			            if (changeorderserial <> "") then
			            	Call AddChangeOrderLink(id, changeorderserial)
			                ResultMsg = ResultMsg + "->. [상품변경 맞교환 교환주문] 출고완료 등록\n\n"
			            end if
					else
						if (changeorderstate <> "8") then
							Call FinishChangeOrder(changeorderserial)
			            	Call AddChangeOrderChulgoLink(id, changeorderserial)
			                ResultMsg = ResultMsg + "->. [상품변경 맞교환 교환주문] 출고완료 전환\n\n"
						end if
					end if
				end if

	        End If

            '//텐배 교환회수 상품변경 한정처리
            if divcd="A111" then
		        If (Err.Number = 0) and (ScanErr="") Then
		            errcode = "011"

					'/한정 올리고, 판매 상태 변경한다.	'/2016.07.15 한용민 생성
					call setItemLimitcs(id, orderserial, "UP")

					ResultMsg = ResultMsg + "->. [한정 처리] 텐배 교환회수 상품변경 한정처리 완료\n\n"
		        End If
			End If

			ResultMsg = ResultMsg + "처리 완료"
	        ReturnUrl = "/cscenter/action/pop_cs_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd

	        If (Err.Number = 0) and (ScanErr="") Then
	            dbget.CommitTrans
	        Else
	            dbget.RollBackTrans
	            response.write "<script>alert(" + Chr(34) + "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(에러코드 : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
				' response.write "<script>history.back()</script>"
				Call SetNeedCheckToY(id)
	            dbget.close()	:	response.End
	        End If

	        ''맞교환 완료 메일
	        If (isCsMailSend) and (divcd <> "A112") then
	        	'// 업체배송은 보내지 않는다.
   	            Call SendCsActionMail(id)
	        End If
        On error Goto 0
    elseif  (divcd="A000") or (divcd="A100") or (divcd="A001") or (divcd="A002") or (divcd="A200") or (divcd="A009") or (divcd="A006") or (divcd="A060") or (divcd="A005") or (divcd="A700") or (divcd="A999") then
    	'----------------------------------------------------------------------
        'CS 접수 내역 완료처리 - 맞교환 출고 / 상품변경 맞교환출고 / 누락 / 서비스 발송 / 기타 /  출고시 유의사항 / 업체긴급문의

		if (IsAutoScript) and (divcd <> "A001") and (divcd <> "A000") and (divcd <> "A100") and (divcd <> "A200") then
			response.write "S_ERR|작업되지 않았습니다. : mode=" + mode + " , divcd=" + divcd
			dbget.close()	:	response.End
		end if

        On Error Resume Next
	        dbget.BeginTrans

	        If (divcd="A100") and (modeflag2<>"norefund") and (Err.Number = 0) and (ScanErr="") Then
	            errcode = "08"

	            newasid = CheckNRegRefund(id, orderserial, reguserid)

	            if (newasid>0) then
	                ResultMsg = ResultMsg + "->. [취소(환불)접수] 처리\n\n"
	            end if
	        End If

	        If (divcd="A999") and (Err.Number = 0) and (ScanErr="") Then
	            errcode = "09"

                '// 고객추가결제 상품은 출고완료처리
	            Call CheckNChulgoPaymentOrder(id, ScanErr)
	            if (ScanErr = "") then
	                ResultMsg = ResultMsg + "->. [기출고주문 상품 출고완료] 처리\n\n"
	            end if
	        End If

	        If (Err.Number = 0) and (ScanErr="") Then
	            errcode = "010"

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
				Call SetNeedCheckToY(id)
	            dbget.close()	:	response.End
	        End If

	        If (isCsMailSend) then
	            if ((divcd="A000") or (divcd="A100") or (divcd="A001") or (divcd="A002")) then
	                ''맞교환/누락/서비스 완료 메일
	                Call SendCsActionMail(id)
	            end if
	        End If
        On error Goto 0
    else
		if (IsAutoScript) then
			response.write "S_ERR|정의되지 않았습니다[3]. : mode=" + mode + " , divcd=" + divcd
		else
			ResultMsg = "정의되지 않았습니다[3]. : mode=" + mode + " , divcd=" + divcd
			response.write "<script>alert('" + ResultMsg + "');</script>"
			response.write "<script>history.back();</script>"
		end if
        dbget.close()	:	response.End
    end if



elseif (mode="state2jupsu") then
	'==========================================================================
    if (divcd="A700") then
    	'----------------------------------------------------------------------
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
            'response.write sqlStr
            dbget.Execute sqlStr

            ResultMsg = "처리 완료"
            ReturnUrl = "/cscenter/action/pop_cs_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd
        else
            response.write "<script>alert('" + ResultMsg + "');</script>"
            response.write "<script>history.back();</script>"
            dbget.close()	:	response.End
        end if
    elseif (divcd="A003") or (divcd="A005") then
        '// 환불요청(A003), 외부몰환불요청(A005)
        sqlStr = " update db_cs.dbo.tbl_new_as_list"
        sqlStr = sqlStr + " set currstate='B001'"
        sqlStr = sqlStr + " ,finishdate=NULL"
        sqlStr = sqlStr + " where id=" & CStr(id)
        sqlStr = sqlStr + " and currstate='B007'"
        'response.write sqlStr
        dbget.Execute sqlStr

        ResultMsg = "처리 완료"
        ReturnUrl = "/cscenter/action/pop_cs_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd
    else
        ResultMsg = "정의되지 않았습니다[4]. : mode=" + mode + " , divcd=" + divcd
        response.write "<script>alert('" + ResultMsg + "');</script>"
        response.write "<script>history.back();</script>"
        dbget.close()	:	response.End
    end if



elseif (mode="addupchejungsanEdit") then
	'==========================================================================
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
	'==========================================================================
    '' 업체 처리완료 => 접수상태로변경
    sqlStr = " select top 1 currstate from [db_cs].[dbo].tbl_new_as_list"
    sqlStr = sqlStr + " where id=" + CStr(id)

    rsget.Open sqlStr,dbget,1
	    if not rsget.Eof then
	        ResultMsg = ""
	        if (rsget("currstate")<>"B006") and (rsget("currstate")<>"B005") then
	            ResultMsg = "업체 처리 완료 상태가 아닙니다. 수정 불가"
	        end if
		else
		    ResultMsg = "코드없음. 수정 불가"
		end if
	rsget.Close

    if (ResultMsg="") then
        sqlStr = " update [db_cs].[dbo].tbl_new_as_list" + VbCrlf
        sqlStr = sqlStr + "set currstate='B001', confirmdate=NULL, finishuser = '" & session("ssBctId") & "' " + VbCrlf
        sqlStr = sqlStr + ",contents_jupsu='" + (contents_jupsu) + "'" + VbCrlf
        sqlStr = sqlStr + " where id=" + CStr(id)
        dbget.Execute sqlStr

        sqlStr = " update [db_cs].[dbo].tbl_new_as_detail" + VbCrlf
        sqlStr = sqlStr + "set currstate='B001'" + VbCrlf
        sqlStr = sqlStr + " where masterid=" + CStr(id)
        dbget.Execute sqlStr

		'// 이전 처리자 아이디 저장
		Call SaveCSListHistory(id)

        ResultMsg = "처리 완료"
        ReturnUrl = "/cscenter/action/pop_cs_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd
    else
        response.write "<script>alert('" + ResultMsg + "');</script>"
        response.write "<script>history.back();</script>"
        dbget.close()	:	response.End
    end if


elseif (mode="upcheconfirm2reconfirm") then
	'==========================================================================
    '' 업체 처리완료 => 업체재확인요청 상태로변경
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
        sqlStr = sqlStr + "set currstate='B005', confirmdate=NULL, finishuser = '" & session("ssBctId") & "' " + VbCrlf
        sqlStr = sqlStr + ",contents_jupsu='" + (contents_jupsu) + "'" + VbCrlf
        sqlStr = sqlStr + " where id=" + CStr(id)
        dbget.Execute sqlStr

        sqlStr = " update [db_cs].[dbo].tbl_new_as_detail" + VbCrlf
        sqlStr = sqlStr + "set currstate='B005'" + VbCrlf
        sqlStr = sqlStr + " where masterid=" + CStr(id)
        dbget.Execute sqlStr

		'// 이전 처리자 아이디 저장
		Call SaveCSListHistory(id)

        ResultMsg = "처리 완료"
        ReturnUrl = "/cscenter/action/pop_cs_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd
    else
        response.write "<script>alert('" + ResultMsg + "');</script>"
        response.write "<script>history.back();</script>"
        dbget.close()	:	response.End
    end if


elseif (mode="changeorderreg") then
	'==========================================================================
    '' 교환주문 수기생성

	Call GetChangeOrderInfo(id, changeorderserial, changeorderstate,  ResultMsg)

	if (ResultMsg="") and (changeorderserial <> "") then
		ResultMsg = "교환주문이 이미 등록되어 있습니다."
	end if

    if (ResultMsg="") then
		'// 교환주문 등록한다.
		'// 교환출고 및 회수상태와 무관하게 수기등록한다.(주문접수상태)
		'// 텐배의 경우 변심 맞교환의 경우 회수이후에 맞교환출고한다(http://logics.10x10.co.kr/v2/online/m_re_chulgo.asp 참고)
		changeorderserial = CheckAndAddChangeOrderJupsu(id, orderserial, ScanErr)

        if (changeorderserial <> "") then
        	Call AddChangeOrderJupsuLink(id, changeorderserial)
            ResultMsg = ResultMsg + "->. [상품변경 맞교환 교환주문] 주문접수 등록\n\n"
        end if

        ResultMsg = ResultMsg + "처리 완료"
        ReturnUrl = "/cscenter/action/pop_cs_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd
    else
        response.write "<script>alert('" + ResultMsg + "');</script>"
        response.write "<script>history.back();</script>"
        dbget.close()	:	response.End
    end if

elseif (mode="changedivcdtoa004") then
	'==========================================================================
    '' 고객 직접반품 전환(A010 -> A004)

    sqlStr = " select top 1 currstate, deleteyn from [db_cs].[dbo].tbl_new_as_list"
    sqlStr = sqlStr + " where id=" + CStr(id)

    rsget.Open sqlStr,dbget,1
	    if not rsget.Eof then
	        ResultMsg = ""
	        if (rsget("deleteyn")="Y") then
	            ResultMsg = "삭제된 내역입니다. 수정 불가"
	        else
		        if (rsget("currstate")<>"B001") then
		            ResultMsg = "이미 택배사에 전송된 내역입니다. 수정 불가"
		        end if
	        end if
		else
		    ResultMsg = "코드없음. 수정 불가"
		end if
	rsget.Close

    if (ResultMsg="") then
    	divcd = "A004"

        sqlStr = " update [db_cs].[dbo].tbl_new_as_list" + VbCrlf
        sqlStr = sqlStr + "set divcd='" + CStr(divcd) + "'" + VbCrlf
        sqlStr = sqlStr + ", requireupche='Y' " + VbCrlf
        sqlStr = sqlStr + ", makerid='10x10logistics' " + VbCrlf
        sqlStr = sqlStr + ", title='고객 직접반품 전환' " + VbCrlf
        sqlStr = sqlStr + ", opentitle='반품접수(업체배송)' " + VbCrlf
        sqlStr = sqlStr + " where id=" + CStr(id)
        dbget.Execute sqlStr

        ResultMsg = "처리 완료"
        ReturnUrl = "/cscenter/action/pop_cs_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd
    else
        response.write "<script>alert('" + ResultMsg + "');</script>"
        response.write "<script>history.back();</script>"
        dbget.close()	:	response.End
    end if



elseif (mode="changedivcdtoa010") then
	'==========================================================================
    '' 회수신청 전환(A004 -> A010)

    sqlStr = " select top 1 currstate, deleteyn from [db_cs].[dbo].tbl_new_as_list"
    sqlStr = sqlStr + " where id=" + CStr(id)

    rsget.Open sqlStr,dbget,1
	    if not rsget.Eof then
	        ResultMsg = ""
	        if (rsget("deleteyn")="Y") then
	            ResultMsg = "삭제된 내역입니다. 수정 불가"
	        else
		        if (rsget("currstate")<>"B001") then
		            ResultMsg = "이미 택배사에 전송된 내역입니다. 수정 불가"
		        end if
	        end if
		else
		    ResultMsg = "코드없음. 수정 불가"
		end if
	rsget.Close

    if (ResultMsg="") then
    	divcd = "A010"

        sqlStr = " update [db_cs].[dbo].tbl_new_as_list" + VbCrlf
        sqlStr = sqlStr + "set divcd='" + CStr(divcd) + "'" + VbCrlf
        sqlStr = sqlStr + ", requireupche='N' " + VbCrlf
        sqlStr = sqlStr + ", makerid=NULL " + VbCrlf
        sqlStr = sqlStr + ", title='회수신청 전환' " + VbCrlf
        sqlStr = sqlStr + ", opentitle='회수신청(텐바이텐배송)' " + VbCrlf
        sqlStr = sqlStr + " where id=" + CStr(id)
        dbget.Execute sqlStr

        ResultMsg = "처리 완료"
        ReturnUrl = "/cscenter/action/pop_cs_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd
    else
        response.write "<script>alert('" + ResultMsg + "');</script>"
        response.write "<script>history.back();</script>"
        dbget.close()	:	response.End
    end if



elseif (mode="restoredel") then

	set ocsaslist = New CCSASList
	ocsaslist.FRectCsAsID = id
	ocsaslist.GetOneCSASMaster

	if (ocsaslist.FOneItem.Fdeleteyn = "N") then
	    response.write "<script>alert(" + Chr(34) + "정상 내역입니다." + Chr(34) + ")</script>"
	    response.write "정상 내역입니다."
	    dbget.close()	:	response.End
	end if

	if (ocsaslist.FOneItem.Fcurrstate = "B007") then
        if C_ADMIN_AUTH then
            '// 허용
        else
            response.write "<script>alert(" + Chr(34) + "완료된 내역입니다." + Chr(34) + ")</script>"
	        response.write "완료된 내역입니다."
	        dbget.close()	:	response.End
        end if
	end if

	divcd = ocsaslist.FOneItem.Fdivcd
	if (_
		(divcd = "A008") or (divcd = "A004") or _
		(divcd = "A010") or (divcd = "A000") or _
		(divcd = "A100") or (divcd = "A011") or _
		(divcd = "A012") or (divcd = "A111") or _
		(divcd = "A112")_
		) then
	    ''response.write "<script>alert(" + Chr(34) + "처리불가(취소/반품/교환) 내역입니다." + Chr(34) + ")</script>"
	    ''response.write "처리불가(취소/반품/교환) 내역입니다."
	    ''dbget.close()	:	response.End
	end if

    if (ResultMsg="") then
        sqlStr = " update [db_cs].[dbo].tbl_new_as_list" + VbCrlf
        sqlStr = sqlStr + "set deleteyn = 'N', finishdate = NULL " + VbCrlf
        sqlStr = sqlStr + " where id=" + CStr(id) + " and deleteyn = 'Y' and currstate <> 'B007' "
        dbget.Execute sqlStr, affectedRows

        if (affectedRows > 0) then
		    Call AddCsMemo(ocsaslist.FOneItem.Forderserial, "1", ocsaslist.FOneItem.Fuserid, session("ssBctId"), "삭제CS(완료이전) 복구 : " & divcd)
        else
            sqlStr = " update [db_cs].[dbo].tbl_new_as_list" + VbCrlf
            sqlStr = sqlStr + "set deleteyn = 'N' " + VbCrlf
            sqlStr = sqlStr + " where id=" + CStr(id) + " and deleteyn = 'Y' and currstate = 'B007' "
            dbget.Execute sqlStr, affectedRows

            if (affectedRows > 0) then
                Call AddCsMemo(ocsaslist.FOneItem.Forderserial, "1", ocsaslist.FOneItem.Fuserid, session("ssBctId"), "삭제CS(완료내역) 복구 : " & divcd)
            end if
        end if

        ResultMsg = "처리 완료"
        ReturnUrl = "/cscenter/action/pop_cs_action_new.asp?id="  + CStr(id) + "&divcd=" + divcd
    else
        response.write "<script>alert('" + ResultMsg + "');</script>"
        response.write "<script>history.back();</script>"
        dbget.close()	:	response.End
    end if

else
	'==========================================================================
    ResultMsg = "정의되지 않았습니다[5]. : mode=" + mode + " , divcd=" + divcd
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


if (mode <> "regcsas") and (id <> "") then
	'// 이전 처리자 아이디 저장
	Call SaveCSListHistory(id)
end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db_LogisticsClose.asp" -->
<%
if (IsAutoScript) then
	response.write "S_OK"
else
	response.write "<script>alert('" + ResultMsg + "');</script>"
	response.write "<script>location.replace('" + ReturnUrl + "');</script>"
end if
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
