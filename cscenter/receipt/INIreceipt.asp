<% option Explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<% Response.Addheader "P3P","policyref='/w3c/p3p.xml', CP='NOI DSP LAW NID PSA ADM OUR IND NAV COM'"%>
<%
'###########################################################
' Description : 고객센터 현금영수증,세금계산서 밸행
' History : 이상구 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_cashreceiptcls.asp"-->
<%
dim INIpay, PInst, Tid
dim ResultCode, ResultMsg, AuthCode
dim PGAuthDate, PGAuthTime
dim ResultpCRPice, ResultSupplyPrice, ResultTax
dim ResultServicePrice, ResultUseOpt, ResultCashNoAppl
dim AckResult
dim goodname, cr_price, sup_price, tax, srvc_price, buyername
dim buyeremail, buyertel, reg_num, useopt, orderserial, userid, sitename, paymethod
dim sqlStr
dim iidx

goodname = html2db(request.Form("goodname"))
cr_price = request.Form("cr_price")
sup_price = request.Form("sup_price")
tax = request.Form("tax")
srvc_price = request.Form("srvc_price")
buyername = html2db(request.Form("buyername"))
buyeremail = html2db(request.Form("buyeremail"))
buyertel = request.Form("buyertel")
reg_num = request.Form("reg_num")
useopt = request.Form("useopt")
orderserial = request.Form("orderserial")
userid = request.Form("userid")
sitename = request.Form("sitename")
paymethod = request.Form("paymethod")

on Error resume next
sqlStr = " select * from [db_log].[dbo].tbl_cash_receipt where 1=0"
rsget.Open sqlStr,dbget,1,3
rsget.AddNew
rsget("orderserial") = orderserial
rsget("userid") = userid
rsget("sitename") = sitename
rsget("goodname") = goodname
rsget("cr_price") = cr_price
rsget("sup_price") = sup_price
rsget("tax") = tax
rsget("srvc_price") = srvc_price
rsget("buyername") = buyername
rsget("buyeremail") = buyeremail
rsget("buyertel") = buyertel
rsget("reg_num") = reg_num
rsget("useopt") = useopt
rsget("paymethod") = paymethod
rsget("cancelyn") = "N"

rsget.update
iidx = rsget("idx")
rsget.close

if Err then
	response.write "<script>alert('Error - " + Err.description + "');</script>"
	response.write "<script>history.back();</script>"
	response.end
end if

on error goto 0

'*******************************************************************************
'* INIreceipt.asp
'* 현금결제(실시간 은행계좌이체, 무통장입금)에 대한 현금결제 영수증 발행 요청한다.
'*
'* Date : 2004/12
'* Project : INIpay V4.11 for Unix
'*
'* http://www.inicis.com
'* http://support.inicis.com
'* Copyright (C) 2002 Inicis, Co. All rights reserved.
'*******************************************************************************

'###############################################################################
'# 1. 객체 생성 #
'################
Set INIpay = Server.CreateObject("INItx41.INItx41.1")

'###############################################################################
'# 2. 인스턴스 초기화 #
'######################
PInst = INIpay.Initialize("")

'###############################################################################
'# 3. 거래 유형 설정 #
'#####################
INIpay.SetActionType CLng(PInst), "receipt"

'###############################################################################
'# 4. 발급 정보 설정 #
'###############################################################################
INIpay.SetField CLng(PInst), "pgid","INIpayRECP"	'PG ID (고정)
INIpay.SetField CLng(PInst), "paymethod","CASH"		'지불방법
INIpay.SetField CLng(PInst), "spgip", "203.238.3.10" '예비 PG IP (고정)
INIpay.SetField CLng(PInst), "currency", Request("currency") '화폐단위
INIpay.SetField CLng(PInst), "admin", "1111"
INIpay.SetField CLng(PInst), "mid", Request("mid") '상점아이디
INIpay.SetField CLng(PInst), "uip", Request.ServerVariables("REMOTE_ADDR") '고객IP
INIpay.SetField CLng(PInst), "goodname", Request("goodname") '상품명
INIpay.SetField CLng(PInst), "cr_price", Request("cr_price") '총 현금 결제 금액
INIpay.SetField CLng(PInst), "sup_price", Request("sup_price") '공급가액
INIpay.SetField CLng(PInst), "tax", Request("tax") '부가세
INIpay.SetField CLng(PInst), "srvc_price", Request("srvc_price") '봉사료
INIpay.SetField CLng(PInst), "buyername", Request("buyername") '성명
INIpay.SetField CLng(PInst), "buyertel", Request("buyertel") '이동전화
INIpay.SetField CLng(PInst), "buyeremail", Request("buyeremail") '이메일
INIpay.SetField CLng(PInst), "reg_num", Request("reg_num") '현금결제자 주민등록번호
INIpay.SetField CLng(PInst), "useopt", Request("useopt") '현금영수증 발행용도 ("0" - 소비자 소득공제용, "1" - 사업자 지출증빙용)
INIpay.SetField CLng(PInst), "debug", "false" '로그모드("true"로 설정하면 상세한 로그를 남김)

'###############################################################################
'# 5. 지불 요청 #
'################
INIpay.StartAction(CLng(PInst))

'###############################################################################
'6. 발급 결과 #
'###############################################################################
'-------------------------------------------------------------------------------
' 가.모든 결제 수단에 공통되는 결제 결과 내용
'-------------------------------------------------------------------------------
Tid = INIpay.GetResult(CLng(PInst), "tid") '거래번호
ResultCode = INIpay.GetResult(CLng(PInst), "resultcode") '결과코드 ("00"이면 지불성공)
ResultMsg = INIpay.GetResult(CLng(PInst), "resultmsg") '결과내용
AuthCode = INIpay.GetResult(CLng(PInst), "authcode") '현금영수증 발생 승인번호
PGAuthDate = INIpay.GetResult(CLng(PInst), "pgauthdate") '이니시스 승인날짜
PGAuthTime = INIpay.GetResult(CLng(PInst), "pgauthtime") '이니시스 승인시각

ResultpCRPice = INIpay.GetResult(CLng(PInst), "ResultpCRPice") '결제 되는 금액
ResultSupplyPrice = INIpay.GetResult(CLng(PInst), "ResultSupplyPrice") '공급가액
ResultTax = INIpay.GetResult(CLng(PInst), "ResultTax") '부가세
ResultServicePrice = INIpay.GetResult(CLng(PInst), "ResultServicePrice") '봉사료
ResultUseOpt = INIpay.GetResult(CLng(PInst), "ResultUseOpt") '발행구분
ResultCashNoAppl = INIpay.GetResult(CLng(PInst), "ResultCashNoAppl") '승인번호

''결과 저장
AuthCode = ResultCashNoAppl   ''이것이 승인번호;;


''결과 저장
sqlStr = "update [db_log].[dbo].tbl_cash_receipt" + VbCrlf
sqlStr = sqlStr + " set tid='" + CStr(Tid) + "'" + VbCrlf
sqlStr = sqlStr + " , resultcode='" + CStr(ResultCode) + "'" + VbCrlf
sqlStr = sqlStr + " , resultmsg='" + html2db(LeftB(CStr(ResultMsg),200)) + "'" + VbCrlf
sqlStr = sqlStr + " , authcode='" + CStr(AuthCode) + "'" + VbCrlf
sqlStr = sqlStr + " , resultcashnoappl='" + CStr(ResultCashNoAppl) + "'" + VbCrlf
sqlStr = sqlStr + " where idx=" + CStr(iidx)

'response.write sqlStr
dbget.Execute sqlStr

''2016/06/30 추가. 승인일
sqlStr = "update [db_log].[dbo].tbl_cash_receipt" + VbCrlf
sqlStr = sqlStr + " SET evalDt='"&LEFT(PGAuthDate,4)&"-"&MID(PGAuthDate,5,2)&"-"&MID(PGAuthDate,7,2)&" "&LEFT(PGAuthTime,2)&":"&MID(PGAuthTime,3,2)&":"&MID(PGAuthTime,5,2)&"'" + VbCrlf
sqlStr = sqlStr + " where idx=" + CStr(iidx)
dbget.Execute sqlStr
        
''2009추가
Dim assignedRow
IF ResultCode = "00" THEN
    sqlStr = " update [db_order].[dbo].tbl_order_master" & VbCrlf
    sqlStr = sqlStr & " set authcode='" & AuthCode & "'" & VbCrlf
    sqlStr = sqlStr + " , cashreceiptreq = (case when accountdiv in ('7', '20') then 'R' else 'S' end) " + VbCrlf
    sqlStr = sqlStr & " where orderserial='" & orderserial & "'"

    sqlStr = " update [db_order].[dbo].tbl_order_master" & VbCrlf
    sqlStr = sqlStr & " set " & VbCrlf
    sqlStr = sqlStr & " authcode = (case when accountdiv in ('7', '20') then '" & AuthCode & "' else authcode end) " + VbCrlf
    sqlStr = sqlStr & " , cashreceiptreq = (case when (accountdiv in ('7', '20')) or (pggubun='NP') then 'R' else 'S' end) " + VbCrlf
    sqlStr = sqlStr & " where orderserial='" & orderserial & "'"

    dbget.Execute sqlStr,assignedRow
    
    IF (assignedRow<1) then
        sqlStr = " update [db_log].[dbo].tbl_old_order_master_2003" & VbCrlf
        sqlStr = sqlStr & " set " & VbCrlf
        sqlStr = sqlStr & " authcode = (case when accountdiv in ('7', '20') then '" & AuthCode & "' else authcode end) " + VbCrlf
        sqlStr = sqlStr & " , cashreceiptreq = (case when accountdiv in ('7', '20') then 'R' else 'S' end) " + VbCrlf
        sqlStr = sqlStr & " where orderserial='" & orderserial & "'"
    
        dbget.Execute sqlStr,assignedRow
    END IF
end if

'###############################################################################
'# 7. 결과 수신 확인 #
'#####################
'지불결과를 잘 수신하였음을 이니시스에 통보.
'[주의] 이 과정이 누락되면 모든 거래가 자동취소됩니다.
IF ResultCode = "00" THEN
	AckResult = INIpay.Ack(CLng(PInst))
	IF AckResult <> "SUCCESS" THEN '(실패)
		'=================================================================
		' 정상수신 통보 실패인 경우 이 승인은 이니시스에서 자동 취소되므로
		' 지불결과를 다시 받아옵니다(성공 -> 실패).
		'=================================================================
		ResultCode = INIpay.GetResult(CLng(PInst), "resultcode")
		ResultMsg = INIpay.GetResult(CLng(PInst), "resultmsg")
	END IF
END IF

'###############################################################################
'# 8. 인스턴스 해제 #
'####################
INIpay.Destroy CLng(PInst)

Set INIpay = Nothing
%>

<%
session("lastreceiptidx") = iidx
%>

<script type="text/javascript">
location.replace('displayreceipt.asp');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
